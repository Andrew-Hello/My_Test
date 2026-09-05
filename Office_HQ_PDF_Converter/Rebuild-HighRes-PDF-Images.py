#!/usr/bin/env python3
import argparse
import io
import json
import math
import zipfile
from pathlib import Path

import fitz  # PyMuPDF
import imagehash
from PIL import Image, ImageOps

RASTER_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp'}


def pil_from_bytes(data: bytes):
    im = Image.open(io.BytesIO(data))
    try:
        im.seek(0)
    except Exception:
        pass
    im = ImageOps.exif_transpose(im)
    if im.mode not in ('RGB', 'RGBA'):
        im = im.convert('RGBA' if 'A' in im.getbands() else 'RGB')
    return im.copy()


def normalized_hash(im: Image.Image):
    rgb = im.convert('RGB')
    return imagehash.phash(rgb, hash_size=16)


def aspect(im):
    return im.width / max(im.height, 1)


def similarity_score(pdf_im, src_im):
    # Lower is better. pHash dominates; aspect ratio prevents obvious mismatches.
    ph = normalized_hash(pdf_im) - normalized_hash(src_im)
    ar = abs(math.log(max(aspect(pdf_im), 1e-9) / max(aspect(src_im), 1e-9)))
    return float(ph) + ar * 80.0


def load_docx_media(docx_path: Path):
    out = []
    with zipfile.ZipFile(docx_path, 'r') as zf:
        for name in zf.namelist():
            if '/media/' not in name:
                continue
            ext = Path(name).suffix.lower()
            if ext not in RASTER_EXTS:
                continue
            raw = zf.read(name)
            try:
                im = pil_from_bytes(raw)
            except Exception as exc:
                out.append({'name': name, 'error': str(exc), 'raw': raw, 'image': None})
                continue
            out.append({
                'name': name,
                'ext': ext,
                'raw': raw,
                'image': im,
                'width': im.width,
                'height': im.height,
                'hash': str(normalized_hash(im)),
            })
    return out


def collect_pdf_images(doc):
    images = {}
    for page_index in range(len(doc)):
        page = doc[page_index]
        for item in page.get_images(full=True):
            xref = item[0]
            if xref in images:
                images[xref]['pages'].append(page_index)
                continue
            try:
                extracted = doc.extract_image(xref)
                raw = extracted['image']
                im = pil_from_bytes(raw)
            except Exception as exc:
                images[xref] = {'xref': xref, 'pages': [page_index], 'error': str(exc), 'image': None}
                continue
            images[xref] = {
                'xref': xref,
                'pages': [page_index],
                'image': im,
                'width': im.width,
                'height': im.height,
                'ext': extracted.get('ext'),
                'hash': str(normalized_hash(im)),
            }
    return list(images.values())


def encode_png(im: Image.Image):
    buf = io.BytesIO()
    # PNG is lossless. optimize=False avoids spending time recompressing at the cost of fidelity.
    im.save(buf, format='PNG', optimize=False)
    return buf.getvalue()


def main():
    ap = argparse.ArgumentParser(description='Replace downsampled PDF images with original DOCX media.')
    ap.add_argument('--source-docx', type=Path, required=True)
    ap.add_argument('--input-pdf', type=Path, required=True)
    ap.add_argument('--output-pdf', type=Path, required=True)
    ap.add_argument('--report', type=Path, required=True)
    ap.add_argument('--max-score', type=float, default=45.0,
                    help='Maximum pHash/aspect matching score allowed for replacement.')
    args = ap.parse_args()

    sources = [x for x in load_docx_media(args.source_docx) if x.get('image') is not None]
    doc = fitz.open(args.input_pdf)
    pdf_images = [x for x in collect_pdf_images(doc) if x.get('image') is not None]

    # Build all candidate pairs, then greedy one-to-one assignment from best match to worst.
    pairs = []
    for pi, p in enumerate(pdf_images):
        for si, s in enumerate(sources):
            score = similarity_score(p['image'], s['image'])
            pairs.append((score, pi, si))
    pairs.sort(key=lambda x: x[0])

    used_pdf = set()
    used_src = set()
    matches = []
    for score, pi, si in pairs:
        if pi in used_pdf or si in used_src:
            continue
        p = pdf_images[pi]
        s = sources[si]
        used_pdf.add(pi)
        used_src.add(si)
        area_ratio = (s['width'] * s['height']) / max(p['width'] * p['height'], 1)
        matches.append({
            'score': round(score, 4),
            'pdf_index': pi,
            'src_index': si,
            'pdf_xref': p['xref'],
            'pdf_px': [p['width'], p['height']],
            'source_name': s['name'],
            'source_px': [s['width'], s['height']],
            'pixel_area_ratio': round(area_ratio, 3),
            'replace': bool(score <= args.max_score and area_ratio > 1.05),
        })

    replacement_errors = []
    replaced = 0
    for m in matches:
        if not m['replace']:
            continue
        p = pdf_images[m['pdf_index']]
        s = sources[m['src_index']]
        try:
            png = encode_png(s['image'])
            first_page = doc[p['pages'][0]]
            first_page.replace_image(p['xref'], stream=png)
            replaced += 1
        except Exception as exc:
            m['replace'] = False
            m['error'] = str(exc)
            replacement_errors.append({'xref': p['xref'], 'error': str(exc)})

    args.output_pdf.parent.mkdir(parents=True, exist_ok=True)
    # Garbage collection removes superseded streams; deflate is lossless.
    doc.save(args.output_pdf, garbage=4, deflate=True, clean=True)
    doc.close()

    report = {
        'source_docx': args.source_docx.name,
        'input_pdf': args.input_pdf.name,
        'output_pdf': args.output_pdf.name,
        'source_raster_media_count': len(sources),
        'pdf_unique_raster_image_count': len(pdf_images),
        'matched_count': len(matches),
        'replaced_count': replaced,
        'max_score': args.max_score,
        'matches': matches,
        'replacement_errors': replacement_errors,
        'output_size_bytes': args.output_pdf.stat().st_size if args.output_pdf.exists() else None,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    print(json.dumps({
        'source_images': len(sources),
        'pdf_images': len(pdf_images),
        'replaced': replaced,
        'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
