#!/usr/bin/env python3
import argparse
import json
import math
import statistics
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image


def safe_round(value, digits=2):
    if value is None:
        return None
    try:
        if math.isfinite(float(value)):
            return round(float(value), digits)
    except Exception:
        pass
    return None


def percentile(values, p):
    vals = sorted(v for v in values if v is not None and math.isfinite(v))
    if not vals:
        return None
    if len(vals) == 1:
        return vals[0]
    pos = (len(vals) - 1) * p
    lo = math.floor(pos)
    hi = math.ceil(pos)
    if lo == hi:
        return vals[lo]
    return vals[lo] * (hi - pos) + vals[hi] * (pos - lo)


def analyze_pdf(path: Path):
    doc = fitz.open(path)
    unique = {}
    placements = []
    filter_counts = Counter()
    page_image_counts = []

    for page_index, page in enumerate(doc):
        page_seen = 0
        for item in page.get_images(full=True):
            xref = int(item[0])
            smask = int(item[1]) if item[1] else 0
            width = int(item[2])
            height = int(item[3])
            bpc = int(item[4]) if item[4] else None
            colorspace = item[5]
            name = item[7]
            filter_name = item[8] or "unknown"
            filter_counts[filter_name] += 1

            if xref not in unique:
                try:
                    raw_stream = doc.xref_stream_raw(xref)
                    raw_bytes = len(raw_stream) if raw_stream else 0
                except Exception:
                    raw_bytes = None
                try:
                    decoded_stream = doc.xref_stream(xref)
                    decoded_bytes = len(decoded_stream) if decoded_stream else 0
                except Exception:
                    decoded_bytes = None
                unique[xref] = {
                    "xref": xref,
                    "smask": smask,
                    "width_px": width,
                    "height_px": height,
                    "bpc": bpc,
                    "colorspace": colorspace,
                    "name": name,
                    "filter": filter_name,
                    "compressed_stream_bytes": raw_bytes,
                    "decoded_stream_bytes": decoded_bytes,
                    "placements": [],
                }

            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []

            if not rects:
                placements.append({
                    "page": page_index + 1,
                    "xref": xref,
                    "width_px": width,
                    "height_px": height,
                    "bbox_pt": None,
                    "effective_ppi_x": None,
                    "effective_ppi_y": None,
                    "effective_ppi_min": None,
                })
                unique[xref]["placements"].append(placements[-1])
                page_seen += 1
                continue

            for rect in rects:
                rw = abs(rect.width)
                rh = abs(rect.height)
                ppi_x = width / (rw / 72.0) if rw > 0 else None
                ppi_y = height / (rh / 72.0) if rh > 0 else None
                ppi_min = min(ppi_x, ppi_y) if ppi_x and ppi_y else (ppi_x or ppi_y)
                rec = {
                    "page": page_index + 1,
                    "xref": xref,
                    "width_px": width,
                    "height_px": height,
                    "bbox_pt": [safe_round(rect.x0), safe_round(rect.y0), safe_round(rect.x1), safe_round(rect.y1)],
                    "effective_ppi_x": safe_round(ppi_x, 1),
                    "effective_ppi_y": safe_round(ppi_y, 1),
                    "effective_ppi_min": safe_round(ppi_min, 1),
                }
                placements.append(rec)
                unique[xref]["placements"].append(rec)
                page_seen += 1
        page_image_counts.append(page_seen)

    ppis = [p["effective_ppi_min"] for p in placements if p["effective_ppi_min"] is not None]
    compressed_sizes = [v["compressed_stream_bytes"] for v in unique.values() if v["compressed_stream_bytes"] is not None]
    decoded_sizes = [v["decoded_stream_bytes"] for v in unique.values() if v["decoded_stream_bytes"] is not None]

    for image in unique.values():
        image_ppis = [p["effective_ppi_min"] for p in image["placements"] if p["effective_ppi_min"] is not None]
        image["placement_count"] = len(image["placements"])
        image["effective_ppi_min"] = safe_round(min(image_ppis), 1) if image_ppis else None
        image["effective_ppi_max"] = safe_round(max(image_ppis), 1) if image_ppis else None
        image["effective_ppi_median"] = safe_round(statistics.median(image_ppis), 1) if image_ppis else None

    largest_images = sorted(
        unique.values(),
        key=lambda x: (x["compressed_stream_bytes"] or 0),
        reverse=True,
    )[:20]
    lowest_ppi = sorted(
        [x for x in unique.values() if x["effective_ppi_min"] is not None],
        key=lambda x: x["effective_ppi_min"],
    )[:20]

    result = {
        "file_name": path.name,
        "size_bytes": path.stat().st_size,
        "size_mb": safe_round(path.stat().st_size / 1024 / 1024, 3),
        "page_count": doc.page_count,
        "pdf_version": doc.metadata.get("format"),
        "producer": doc.metadata.get("producer"),
        "creator": doc.metadata.get("creator"),
        "unique_image_xrefs": len(unique),
        "image_placements": len(placements),
        "pages_with_image_placement_counts": page_image_counts,
        "image_filters_by_placement": dict(filter_counts),
        "unique_image_compressed_stream_bytes": sum(compressed_sizes),
        "unique_image_compressed_stream_mb": safe_round(sum(compressed_sizes) / 1024 / 1024, 3),
        "unique_image_decoded_stream_bytes": sum(decoded_sizes),
        "effective_ppi": {
            "count": len(ppis),
            "min": safe_round(min(ppis), 1) if ppis else None,
            "p10": safe_round(percentile(ppis, 0.10), 1),
            "median": safe_round(statistics.median(ppis), 1) if ppis else None,
            "p90": safe_round(percentile(ppis, 0.90), 1),
            "max": safe_round(max(ppis), 1) if ppis else None,
            "placements_lt_150": sum(1 for x in ppis if x < 150),
            "placements_lt_220": sum(1 for x in ppis if x < 220),
            "placements_ge_300": sum(1 for x in ppis if x >= 300),
            "placements_ge_600": sum(1 for x in ppis if x >= 600),
        },
        "largest_images": [strip_placements(x) for x in largest_images],
        "lowest_effective_ppi_images": [strip_placements(x) for x in lowest_ppi],
        "images": [strip_placements(x) for x in sorted(unique.values(), key=lambda x: x["xref"])],
    }
    doc.close()
    return result


def strip_placements(image):
    return {k: v for k, v in image.items() if k != "placements"}


def analyze_docx(path: Path):
    items = []
    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist():
            if not info.filename.startswith("word/media/") or info.is_dir():
                continue
            raw = zf.read(info.filename)
            rec = {
                "name": Path(info.filename).name,
                "zip_compressed_bytes": info.compress_size,
                "zip_uncompressed_bytes": info.file_size,
                "extension": Path(info.filename).suffix.lower(),
                "width_px": None,
                "height_px": None,
                "dpi": None,
                "format": None,
                "mode": None,
            }
            try:
                with Image.open(BytesIO(raw)) as img:
                    rec["width_px"], rec["height_px"] = img.size
                    rec["format"] = img.format
                    rec["mode"] = img.mode
                    dpi = img.info.get("dpi")
                    if dpi:
                        rec["dpi"] = [safe_round(dpi[0], 1), safe_round(dpi[1], 1)]
            except Exception:
                pass
            items.append(rec)

    return {
        "file_name": path.name,
        "size_bytes": path.stat().st_size,
        "size_mb": safe_round(path.stat().st_size / 1024 / 1024, 3),
        "media_count": len(items),
        "media_zip_uncompressed_bytes": sum(x["zip_uncompressed_bytes"] for x in items),
        "media_zip_uncompressed_mb": safe_round(sum(x["zip_uncompressed_bytes"] for x in items) / 1024 / 1024, 3),
        "media_extensions": dict(Counter(x["extension"] for x in items)),
        "media": sorted(items, key=lambda x: x["zip_uncompressed_bytes"], reverse=True),
    }


def build_markdown(report):
    src = report.get("source_docx")
    pdfs = report["pdfs"]
    lines = [
        "# PDF Quality Comparison",
        "",
        f"Generated: {report['generated_utc']}",
        "",
    ]
    if src:
        lines += [
            "## Source DOCX",
            "",
            f"- File: `{src['file_name']}`",
            f"- DOCX size: **{src['size_mb']} MB**",
            f"- Embedded media: **{src['media_count']} files / {src['media_zip_uncompressed_mb']} MB**",
            f"- Media types: `{src['media_extensions']}`",
            "",
        ]

    lines += [
        "## PDF comparison",
        "",
        "| Metric | " + " | ".join(p["file_name"] for p in pdfs) + " |",
        "|---|" + "|".join(["---:"] * len(pdfs)) + "|",
        "| PDF size (MB) | " + " | ".join(str(p["size_mb"]) for p in pdfs) + " |",
        "| Pages | " + " | ".join(str(p["page_count"]) for p in pdfs) + " |",
        "| Unique image XRefs | " + " | ".join(str(p["unique_image_xrefs"]) for p in pdfs) + " |",
        "| Image placements | " + " | ".join(str(p["image_placements"]) for p in pdfs) + " |",
        "| Unique compressed image streams (MB) | " + " | ".join(str(p["unique_image_compressed_stream_mb"]) for p in pdfs) + " |",
        "| Effective PPI min | " + " | ".join(str(p["effective_ppi"]["min"]) for p in pdfs) + " |",
        "| Effective PPI median | " + " | ".join(str(p["effective_ppi"]["median"]) for p in pdfs) + " |",
        "| Effective PPI p90 | " + " | ".join(str(p["effective_ppi"]["p90"]) for p in pdfs) + " |",
        "| Placements >=300 PPI | " + " | ".join(str(p["effective_ppi"]["placements_ge_300"]) for p in pdfs) + " |",
        "| Placements >=600 PPI | " + " | ".join(str(p["effective_ppi"]["placements_ge_600"]) for p in pdfs) + " |",
        "",
    ]

    for p in pdfs:
        lines += [
            f"## {p['file_name']}",
            "",
            f"- Producer: `{p.get('producer')}`",
            f"- Creator: `{p.get('creator')}`",
            f"- PDF format: `{p.get('pdf_version')}`",
            f"- Image filters: `{p.get('image_filters_by_placement')}`",
            f"- Effective PPI: `{p.get('effective_ppi')}`",
            "",
            "### Largest embedded image streams",
            "",
            "| xref | px | filter | compressed KB | placements | min PPI | median PPI |",
            "|---:|---:|---|---:|---:|---:|---:|",
        ]
        for img in p["largest_images"][:12]:
            kb = safe_round((img.get("compressed_stream_bytes") or 0) / 1024, 1)
            lines.append(
                f"| {img['xref']} | {img['width_px']}x{img['height_px']} | {img['filter']} | {kb} | {img['placement_count']} | {img['effective_ppi_min']} | {img['effective_ppi_median']} |"
            )
        lines.append("")

    return "\n".join(lines) + "\n"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source-docx", type=Path)
    ap.add_argument("--pdf", action="append", type=Path, required=True)
    ap.add_argument("--json-out", type=Path, required=True)
    ap.add_argument("--md-out", type=Path, required=True)
    args = ap.parse_args()

    report = {
        "schema_version": 1,
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "source_docx": analyze_docx(args.source_docx) if args.source_docx and args.source_docx.exists() else None,
        "pdfs": [analyze_pdf(p) for p in args.pdf],
    }
    args.json_out.parent.mkdir(parents=True, exist_ok=True)
    args.json_out.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    args.md_out.write_text(build_markdown(report), encoding="utf-8")


if __name__ == "__main__":
    main()
