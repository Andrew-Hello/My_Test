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
OOXML_EXTS = {'.docx', '.xlsx', '.pptx'}


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


def normalized_hash(im: Image.Image):
    return imagehash.phash(flatten_for_matching(im), hash_size=16)


def aspect(im):
    return im.width / max(im.height, 1)


def aspect_delta_log(pdf_im, src_im):
    return abs(math.log(max(aspect(pdf_im), 1e-9) / max(aspect(src_im), 1e-9)))


def similarity_score(pdf_im, src_im):
    ph = normalized_hash(pdf_im) - normalized_hash(src_im)
    ar = aspect_delta_log(pdf_im, src_im)
    return float(ph) + ar * 80.0


def load_ooxml_media(office_path: Path):
    if office_path.suffix.lower() not in OOXML_EXTS:
        raise ValueError(f'High-resolution media recovery supports OOXML only: {sorted(OOXML_EXTS)}')

    out = []
    with zipfile.ZipFile(office_path, 'r') as zf:
        for name in zf.namelist():
            # DOCX -> word/media, XLSX -> xl/media, PPTX -> ppt/media.
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
    # The authoritative source pixels are re-encoded losslessly. This is
    # deliberate even for JPEG sources: Office cannot introduce another lossy
    # JPEG generation during this repair stage.
    im.save(buf, format='PNG', optimize=False)
    return buf.getvalue()


def main():
    ap = argparse.ArgumentParser(
        description='Replace downsampled Office-PDF image streams with authoritative OOXML source media.'
    )
    source_group = ap.add_mutually_exclusive_group(required=True)
    source_group.add_argument('--source-office', type=Path,
                              help='Source DOCX/XLSX/PPTX package.')
    source_group.add_argument('--source-docx', type=Path,
                              help='Deprecated alias retained for existing tests.')
    ap.add_argument('--input-pdf', type=Path, required=True)
    ap.add_argument('--output-pdf', type=Path, required=True)
    ap.add_argument('--report', type=Path, required=True)
    ap.add_argument('--max-score', type=float, default=45.0,
                    help='Maximum pHash/aspect matching score allowed for a normal replacement.')
    ap.add_argument('--fallback-max-aspect-log', type=float, default=0.35,
                    help='Maximum aspect-ratio log delta for the one-to-one high-resolution fallback.')
    args = ap.parse_args()

    source_path = args.source_office or args.source_docx
    sources = [x for x in load_ooxml_media(source_path) if x.get('image') is not None]
    doc = fitz.open(args.input_pdf)
    pdf_images = [x for x in collect_pdf_images(doc) if x.get('image') is not None]

    if not sources:
        raise RuntimeError('No supported raster media were found in the OOXML source package.')
    if not pdf_images:
        raise RuntimeError('No raster image objects were found in the PDF.')

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
        if (not m['replace'])
        and m['pixel_area_ratio'] >= 1.5
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
            png = encode_png(s['image'])
            first_page = doc[p['pages'][0]]
            first_page.replace_image(p['xref'], stream=png)
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
        'source_office': source_path.name,
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
