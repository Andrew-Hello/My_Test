#!/usr/bin/env python3
"""Restore authoritative OOXML raster media into an Office-generated PDF.

Production dependencies are intentionally limited to:
- PyMuPDF (import name: pymupdf)
- Pillow

No ImageHash / NumPy / SciPy dependency is required.  A small perceptual hash
implementation is included here so the converter remains easy to deploy on a
normal Windows Python installation.
"""

import argparse
import io
import json
import math
import statistics
import zipfile
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageOps

RASTER_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp'}
OOXML_EXTS = {'.docx', '.xlsx', '.pptx'}
PHASH_SIZE = 16
PHASH_HIGHFREQ_FACTOR = 4
PHASH_IMAGE_SIZE = PHASH_SIZE * PHASH_HIGHFREQ_FACTOR

# Precompute only the low-frequency DCT rows needed by a 16x16 pHash.
_DCT_COS = [
    [math.cos(math.pi * (2 * x + 1) * u / (2.0 * PHASH_IMAGE_SIZE))
     for x in range(PHASH_IMAGE_SIZE)]
    for u in range(PHASH_SIZE)
]
_DCT_ALPHA = [
    math.sqrt(1.0 / PHASH_IMAGE_SIZE) if u == 0 else math.sqrt(2.0 / PHASH_IMAGE_SIZE)
    for u in range(PHASH_SIZE)
]


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


def flatten_for_matching(im: Image.Image):
    """Normalize transparency against white before perceptual comparison."""
    if im.mode == 'RGBA':
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(im, mask=im.getchannel('A'))
        return bg
    if im.mode != 'RGB':
        return im.convert('RGB')
    return im


def perceptual_hash(im: Image.Image):
    """Return a 256-bit pHash compatible in spirit with ImageHash.phash.

    The implementation uses Pillow + a compact low-frequency DCT, avoiding the
    ImageHash/SciPy dependency chain.  Only relative coefficients matter for
    matching, so this is deterministic and platform-independent.
    """
    gray = flatten_for_matching(im).convert('L').resize(
        (PHASH_IMAGE_SIZE, PHASH_IMAGE_SIZE), Image.Resampling.LANCZOS
    )
    pixels = list(gray.getdata())

    # First DCT pass across each row, retaining u=0..15 only.
    row_low = [[0.0] * PHASH_SIZE for _ in range(PHASH_IMAGE_SIZE)]
    for y in range(PHASH_IMAGE_SIZE):
        offset = y * PHASH_IMAGE_SIZE
        row = pixels[offset:offset + PHASH_IMAGE_SIZE]
        for u in range(PHASH_SIZE):
            cos_u = _DCT_COS[u]
            total = 0.0
            for x, value in enumerate(row):
                total += value * cos_u[x]
            row_low[y][u] = total * _DCT_ALPHA[u]

    coeffs = []
    for v in range(PHASH_SIZE):
        cos_v = _DCT_COS[v]
        alpha_v = _DCT_ALPHA[v]
        for u in range(PHASH_SIZE):
            total = 0.0
            for y in range(PHASH_IMAGE_SIZE):
                total += row_low[y][u] * cos_v[y]
            coeffs.append(total * alpha_v)

    median = statistics.median(coeffs)
    bits = 0
    for value in coeffs:
        bits = (bits << 1) | (1 if value > median else 0)
    return bits


def hash_text(value: int):
    return format(value, '064x')


def hash_distance(a: int, b: int):
    return bin(a ^ b).count('1')


def aspect(im):
    return im.width / max(im.height, 1)


def aspect_delta_log(pdf_im, src_im):
    return abs(math.log(max(aspect(pdf_im), 1e-9) / max(aspect(src_im), 1e-9)))


def similarity_score(pdf_item, src_item):
    ph = hash_distance(pdf_item['phash'], src_item['phash'])
    ar = aspect_delta_log(pdf_item['image'], src_item['image'])
    return float(ph) + ar * 80.0


def load_ooxml_media(office_path: Path):
    if office_path.suffix.lower() not in OOXML_EXTS:
        raise ValueError(f'High-resolution media recovery supports OOXML only: {sorted(OOXML_EXTS)}')
    out = []
    with zipfile.ZipFile(office_path, 'r') as zf:
        for name in zf.namelist():
            if '/media/' not in name:
                continue
            ext = Path(name).suffix.lower()
            if ext not in RASTER_EXTS:
                continue
            raw = zf.read(name)
            try:
                im = pil_from_bytes(raw)
                phash = perceptual_hash(im)
            except Exception as exc:
                out.append({'name': name, 'error': str(exc), 'image': None})
                continue
            out.append({
                'name': name,
                'ext': ext,
                'image': im,
                'width': im.width,
                'height': im.height,
                'phash': phash,
                'hash': hash_text(phash),
            })
    return out


def collect_pdf_images(doc):
    images = {}
    for page_index in range(len(doc)):
        page = doc[page_index]
        for item in page.get_images(full=True):
            xref = int(item[0])
            if xref in images:
                images[xref]['pages'].append(page_index)
                continue
            try:
                extracted = doc.extract_image(xref)
                im = pil_from_bytes(extracted['image'])
                phash = perceptual_hash(im)
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
                'phash': phash,
                'hash': hash_text(phash),
            }
    return list(images.values())


def encode_png(im: Image.Image):
    buf = io.BytesIO()
    # Lossless re-encoding is deliberate even for a source JPEG: it prevents
    # Office from introducing an additional lossy JPEG generation in the PDF.
    im.save(buf, format='PNG', optimize=False)
    return buf.getvalue()


def save_passthrough(doc, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path, garbage=4, deflate=True, clean=True)


def main():
    ap = argparse.ArgumentParser(
        description='Replace downsampled Office-PDF image streams with authoritative OOXML source media.'
    )
    ap.add_argument('--source-office', type=Path, required=True)
    ap.add_argument('--input-pdf', type=Path, required=True)
    ap.add_argument('--output-pdf', type=Path, required=True)
    ap.add_argument('--report', type=Path, required=True)
    ap.add_argument('--max-score', type=float, default=45.0)
    ap.add_argument('--fallback-max-aspect-log', type=float, default=0.35)
    args = ap.parse_args()

    sources = [x for x in load_ooxml_media(args.source_office) if x.get('image') is not None]
    doc = fitz.open(args.input_pdf)
    pdf_images = [x for x in collect_pdf_images(doc) if x.get('image') is not None]

    if not sources or not pdf_images:
        save_passthrough(doc, args.output_pdf)
        doc.close()
        report = {
            'schema_version': 2,
            'matching_engine': 'builtin_phash_dct16',
            'source_office': args.source_office.name,
            'input_pdf': args.input_pdf.name,
            'output_pdf': args.output_pdf.name,
            'source_raster_media_count': len(sources),
            'pdf_unique_raster_image_count': len(pdf_images),
            'matched_count': 0,
            'replaced_count': 0,
            'passthrough_reason': 'no_supported_source_raster_media' if not sources else 'no_pdf_raster_images',
            'matches': [],
            'replacement_errors': [],
            'output_size_bytes': args.output_pdf.stat().st_size,
        }
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({'source_images': len(sources), 'pdf_images': len(pdf_images), 'replaced': 0,
                          'passthrough': True, 'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3)},
                         ensure_ascii=False, indent=2))
        return

    pairs = []
    for pi, p in enumerate(pdf_images):
        for si, s in enumerate(sources):
            pairs.append((similarity_score(p, s), pi, si))
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
        ar_delta = aspect_delta_log(p['image'], s['image'])
        confident_replace = bool(score <= args.max_score and area_ratio >= 0.95)
        if confident_replace and area_ratio <= 1.05:
            mode = 'same_resolution_source_preservation'
        elif confident_replace:
            mode = 'confident_hash_highres'
        elif area_ratio < 0.95:
            mode = 'source_smaller_than_pdf'
        else:
            mode = 'initially_rejected'
        matches.append({
            'score': round(score, 4),
            'pdf_index': pi,
            'src_index': si,
            'pdf_xref': p['xref'],
            'pdf_px': [p['width'], p['height']],
            'source_name': s['name'],
            'source_px': [s['width'], s['height']],
            'pixel_area_ratio': round(area_ratio, 3),
            'aspect_delta_log': round(ar_delta, 5),
            'replace': confident_replace,
            'match_mode': mode,
        })

    fallback_candidates = [
        m for m in matches
        if (not m['replace']) and m['pixel_area_ratio'] >= 1.5
        and m['aspect_delta_log'] <= args.fallback_max_aspect_log
    ]
    fallback_enabled = (
        len(sources) == len(pdf_images)
        and len(matches) == len(pdf_images)
        and 0 < len(fallback_candidates) <= 3
    )
    if fallback_enabled:
        for m in fallback_candidates:
            m['replace'] = True
            m['match_mode'] = 'unique_one_to_one_highres_fallback'

    replacement_errors = []
    replaced = 0
    for m in matches:
        if not m['replace']:
            continue
        p = pdf_images[m['pdf_index']]
        s = sources[m['src_index']]
        try:
            doc[p['pages'][0]].replace_image(p['xref'], stream=encode_png(s['image']))
            replaced += 1
        except Exception as exc:
            m['replace'] = False
            m['match_mode'] = 'replacement_error'
            m['error'] = str(exc)
            replacement_errors.append({'xref': p['xref'], 'error': str(exc)})

    args.output_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.output_pdf, garbage=4, deflate=True, clean=True)
    doc.close()

    report = {
        'schema_version': 2,
        'matching_engine': 'builtin_phash_dct16',
        'source_office': args.source_office.name,
        'input_pdf': args.input_pdf.name,
        'output_pdf': args.output_pdf.name,
        'source_raster_media_count': len(sources),
        'pdf_unique_raster_image_count': len(pdf_images),
        'matched_count': len(matches),
        'replaced_count': replaced,
        'normal_max_score': args.max_score,
        'fallback_enabled': fallback_enabled,
        'fallback_candidate_count': len(fallback_candidates),
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
        'fallback_enabled': fallback_enabled,
        'fallback_candidates': len(fallback_candidates),
        'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
