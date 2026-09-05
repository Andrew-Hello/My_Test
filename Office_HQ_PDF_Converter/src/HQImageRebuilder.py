#!/usr/bin/env python3
"""Restore authoritative OOXML raster media into an Office-generated PDF.

Production dependencies:
- PyMuPDF (import name: pymupdf)
- Pillow

For PPTX, the rebuilder also reads slide relationships and DrawingML srcRect
crop metadata. This lets it recreate the exact cropped view from the original
high-resolution raster before matching/replacing the PDF XObject.
"""

import argparse
import io
import json
import math
import posixpath
import statistics
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageOps

RASTER_EXTS = {'.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.gif', '.webp'}
OOXML_EXTS = {'.docx', '.xlsx', '.pptx'}
PHASH_SIZE = 16
PHASH_HIGHFREQ_FACTOR = 4
PHASH_IMAGE_SIZE = PHASH_SIZE * PHASH_HIGHFREQ_FACTOR

NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'a': NS_A, 'p': NS_P, 'r': NS_R}

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
    if im.mode == 'RGBA':
        bg = Image.new('RGB', im.size, 'white')
        bg.paste(im, mask=im.getchannel('A'))
        return bg
    if im.mode != 'RGB':
        return im.convert('RGB')
    return im


def image_pixels(gray: Image.Image):
    """Return pixels without triggering Pillow 12+ getdata deprecation warnings."""
    if hasattr(gray, 'get_flattened_data'):
        return list(gray.get_flattened_data())
    return list(gray.getdata())


def perceptual_hash(im: Image.Image):
    gray = flatten_for_matching(im).convert('L').resize(
        (PHASH_IMAGE_SIZE, PHASH_IMAGE_SIZE), Image.Resampling.LANCZOS
    )
    pixels = image_pixels(gray)

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
    return (a ^ b).bit_count()


def aspect(im):
    return im.width / max(im.height, 1)


def aspect_delta_log(pdf_im, src_im):
    return abs(math.log(max(aspect(pdf_im), 1e-9) / max(aspect(src_im), 1e-9)))


def similarity_score(pdf_item, src_item):
    ph = hash_distance(pdf_item['phash'], src_item['phash'])
    ar = aspect_delta_log(pdf_item['image'], src_item['image'])
    return float(ph) + ar * 80.0


def build_source_item(name, im, **metadata):
    phash = perceptual_hash(im)
    item = {
        'name': name,
        'image': im,
        'width': im.width,
        'height': im.height,
        'phash': phash,
        'hash': hash_text(phash),
    }
    item.update(metadata)
    return item


def read_raw_media(zf: zipfile.ZipFile):
    media = {}
    errors = []
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
            errors.append({'name': name, 'error': str(exc)})
            continue
        media[name] = {
            'name': name,
            'ext': ext,
            'image': im,
            'width': im.width,
            'height': im.height,
        }
    return media, errors


def relationship_map(zf: zipfile.ZipFile, owner_path: str):
    rel_path = posixpath.join(
        posixpath.dirname(owner_path),
        '_rels',
        posixpath.basename(owner_path) + '.rels',
    )
    if rel_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_path))
    out = {}
    base_dir = posixpath.dirname(owner_path)
    for rel in root.findall(f'{{{NS_REL}}}Relationship'):
        rid = rel.get('Id')
        target = rel.get('Target')
        target_mode = rel.get('TargetMode')
        if not rid or not target or target_mode == 'External':
            continue
        if target.startswith('/'):
            normalized = target.lstrip('/')
        else:
            normalized = posixpath.normpath(posixpath.join(base_dir, target))
        out[rid] = normalized
    return out


def parse_src_rect(node):
    if node is None:
        return {'l': 0, 't': 0, 'r': 0, 'b': 0}
    out = {}
    for key in ('l', 't', 'r', 'b'):
        try:
            out[key] = int(node.get(key, '0'))
        except Exception:
            out[key] = 0
    return out


def crop_by_src_rect(im: Image.Image, rect):
    """Apply DrawingML srcRect values (1/1000 percent, range usually 0..100000)."""
    l = max(0.0, min(0.99999, rect.get('l', 0) / 100000.0))
    t = max(0.0, min(0.99999, rect.get('t', 0) / 100000.0))
    r = max(0.0, min(0.99999, rect.get('r', 0) / 100000.0))
    b = max(0.0, min(0.99999, rect.get('b', 0) / 100000.0))

    left = int(round(im.width * l))
    top = int(round(im.height * t))
    right = int(round(im.width * (1.0 - r)))
    bottom = int(round(im.height * (1.0 - b)))

    left = max(0, min(left, im.width - 1))
    top = max(0, min(top, im.height - 1))
    right = max(left + 1, min(right, im.width))
    bottom = max(top + 1, min(bottom, im.height))

    if left == 0 and top == 0 and right == im.width and bottom == im.height:
        return im.copy(), [0, 0, im.width, im.height], False
    return im.crop((left, top, right, bottom)), [left, top, right, bottom], True


def load_pptx_candidates(zf: zipfile.ZipFile, media):
    """Build per-picture candidates, reproducing srcRect crop from slide XML."""
    candidates = []
    referenced_media = set()
    crop_count = 0
    slide_names = sorted(
        name for name in zf.namelist()
        if name.startswith('ppt/slides/slide') and name.endswith('.xml')
        and '/_rels/' not in name
    )

    for slide_name in slide_names:
        try:
            root = ET.fromstring(zf.read(slide_name))
            rels = relationship_map(zf, slide_name)
        except Exception:
            continue

        pictures = root.findall('.//p:pic', NS)
        for pic_index, pic in enumerate(pictures, start=1):
            blip = pic.find('.//a:blip', NS)
            if blip is None:
                continue
            rid = blip.get(f'{{{NS_R}}}embed')
            media_name = rels.get(rid)
            if not media_name or media_name not in media:
                continue

            referenced_media.add(media_name)
            source = media[media_name]
            full_im = source['image']

            src_rect_node = pic.find('.//a:srcRect', NS)
            rect = parse_src_rect(src_rect_node)
            candidate_im, crop_box, crop_applied = crop_by_src_rect(full_im, rect)
            if crop_applied:
                crop_count += 1

            picture_name = None
            c_nv_pr = pic.find('.//p:cNvPr', NS)
            if c_nv_pr is not None:
                picture_name = c_nv_pr.get('name') or c_nv_pr.get('id')

            display_name = (
                f"{media_name}#slide={Path(slide_name).stem};"
                f"picture={picture_name or pic_index}"
            )
            try:
                candidates.append(build_source_item(
                    display_name,
                    candidate_im,
                    ext=source['ext'],
                    media_name=media_name,
                    candidate_kind='pptx_slide_picture',
                    pptx_slide=slide_name,
                    pptx_picture_name=picture_name,
                    pptx_crop_100k=rect,
                    pptx_crop_box_px=crop_box,
                    pptx_crop_applied=crop_applied,
                    full_source_px=[full_im.width, full_im.height],
                ))
            except Exception:
                continue

    # Some media may be used by layouts, masters, charts, or other structures not
    # represented by p:pic. Keep those as conservative full-image candidates.
    for media_name, source in media.items():
        if media_name in referenced_media:
            continue
        try:
            candidates.append(build_source_item(
                media_name,
                source['image'].copy(),
                ext=source['ext'],
                media_name=media_name,
                candidate_kind='ooxml_media_unreferenced',
                pptx_crop_applied=False,
                full_source_px=[source['width'], source['height']],
            ))
        except Exception:
            continue

    # Defensive fallback for unusual PPTX files whose slide pictures were not parsed.
    if not candidates:
        for media_name, source in media.items():
            try:
                candidates.append(build_source_item(
                    media_name,
                    source['image'].copy(),
                    ext=source['ext'],
                    media_name=media_name,
                    candidate_kind='ooxml_media',
                    pptx_crop_applied=False,
                    full_source_px=[source['width'], source['height']],
                ))
            except Exception:
                continue

    return candidates, crop_count


def load_ooxml_sources(office_path: Path):
    if office_path.suffix.lower() not in OOXML_EXTS:
        raise ValueError(f'High-resolution media recovery supports OOXML only: {sorted(OOXML_EXTS)}')

    with zipfile.ZipFile(office_path, 'r') as zf:
        media, media_errors = read_raw_media(zf)

        if office_path.suffix.lower() == '.pptx':
            candidates, crop_count = load_pptx_candidates(zf, media)
        else:
            candidates = []
            crop_count = 0
            for media_name, source in media.items():
                try:
                    candidates.append(build_source_item(
                        media_name,
                        source['image'].copy(),
                        ext=source['ext'],
                        media_name=media_name,
                        candidate_kind='ooxml_media',
                        full_source_px=[source['width'], source['height']],
                    ))
                except Exception:
                    continue

    return candidates, {
        'unique_media_count': len(media),
        'candidate_count': len(candidates),
        'pptx_crop_candidate_count': crop_count,
        'media_errors': media_errors,
    }


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
    im.save(buf, format='PNG', optimize=False)
    return buf.getvalue()


def save_passthrough(doc, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path, garbage=4, deflate=True, clean=True)


def match_detail(m, s):
    for key in (
        'media_name',
        'candidate_kind',
        'pptx_slide',
        'pptx_picture_name',
        'pptx_crop_100k',
        'pptx_crop_box_px',
        'pptx_crop_applied',
        'full_source_px',
    ):
        if key in s:
            m[key] = s[key]


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

    sources, source_info = load_ooxml_sources(args.source_office)
    sources = [x for x in sources if x.get('image') is not None]

    doc = fitz.open(args.input_pdf)
    pdf_images = [x for x in collect_pdf_images(doc) if x.get('image') is not None]

    if not sources or not pdf_images:
        save_passthrough(doc, args.output_pdf)
        doc.close()
        report = {
            'schema_version': 4,
            'matching_engine': 'builtin_phash_dct16+pptx_srcRect',
            'source_office': args.source_office.name,
            'input_pdf': args.input_pdf.name,
            'output_pdf': args.output_pdf.name,
            'source_raster_media_count': source_info['unique_media_count'],
            'source_candidate_count': source_info['candidate_count'],
            'pptx_crop_candidate_count': source_info['pptx_crop_candidate_count'],
            'pdf_unique_raster_image_count': len(pdf_images),
            'matched_count': 0,
            'replaced_count': 0,
            'passthrough_reason': 'no_supported_source_raster_media' if not sources else 'no_pdf_raster_images',
            'matches': [],
            'replacement_errors': [],
            'media_errors': source_info['media_errors'],
            'output_size_bytes': args.output_pdf.stat().st_size,
        }
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({
            'source_images': source_info['unique_media_count'],
            'source_candidates': source_info['candidate_count'],
            'pptx_crop_candidates': source_info['pptx_crop_candidate_count'],
            'pdf_images': len(pdf_images),
            'replaced': 0,
            'passthrough': True,
            'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
        }, ensure_ascii=False, indent=2))
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

        if confident_replace and s.get('pptx_crop_applied'):
            mode = 'pptx_semantic_crop_match'
        elif confident_replace and area_ratio <= 1.05:
            mode = 'same_resolution_source_preservation'
        elif confident_replace:
            mode = 'confident_hash_highres'
        elif area_ratio < 0.95:
            mode = 'source_smaller_than_pdf'
        else:
            mode = 'initially_rejected'

        m = {
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
        }
        match_detail(m, s)
        matches.append(m)

    # The old unique one-to-one fallback remains only for simple OOXML media
    # sets (DOCX/XLSX and unusual PPTX without semantic expansion). It is
    # deliberately disabled when PPTX creates additional per-picture crop
    # candidates because uniqueness can no longer be inferred from counts.
    fallback_candidates = [
        m for m in matches
        if (not m['replace']) and m['pixel_area_ratio'] >= 1.5
        and m['aspect_delta_log'] <= args.fallback_max_aspect_log
    ]
    fallback_enabled = (
        source_info['candidate_count'] == source_info['unique_media_count']
        and source_info['candidate_count'] == len(pdf_images)
        and len(matches) == len(pdf_images)
        and 0 < len(fallback_candidates) <= 3
    )
    if fallback_enabled:
        for m in fallback_candidates:
            m['replace'] = True
            m['match_mode'] = 'unique_one_to_one_highres_fallback'

    replacement_errors = []
    replaced = 0
    semantic_crop_replaced = 0
    for m in matches:
        if not m['replace']:
            continue
        p = pdf_images[m['pdf_index']]
        s = sources[m['src_index']]
        try:
            doc[p['pages'][0]].replace_image(p['xref'], stream=encode_png(s['image']))
            replaced += 1
            if m['match_mode'] == 'pptx_semantic_crop_match':
                semantic_crop_replaced += 1
        except Exception as exc:
            m['replace'] = False
            m['match_mode'] = 'replacement_error'
            m['error'] = str(exc)
            replacement_errors.append({'xref': p['xref'], 'error': str(exc)})

    args.output_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.output_pdf, garbage=4, deflate=True, clean=True)
    doc.close()

    report = {
        'schema_version': 4,
        'matching_engine': 'builtin_phash_dct16+pptx_srcRect',
        'source_office': args.source_office.name,
        'input_pdf': args.input_pdf.name,
        'output_pdf': args.output_pdf.name,
        'source_raster_media_count': source_info['unique_media_count'],
        'source_candidate_count': source_info['candidate_count'],
        'pptx_crop_candidate_count': source_info['pptx_crop_candidate_count'],
        'pptx_semantic_crop_replaced_count': semantic_crop_replaced,
        'pdf_unique_raster_image_count': len(pdf_images),
        'matched_count': len(matches),
        'replaced_count': replaced,
        'normal_max_score': args.max_score,
        'fallback_enabled': fallback_enabled,
        'fallback_candidate_count': len(fallback_candidates),
        'matches': matches,
        'replacement_errors': replacement_errors,
        'media_errors': source_info['media_errors'],
        'output_size_bytes': args.output_pdf.stat().st_size if args.output_pdf.exists() else None,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    print(json.dumps({
        'source_images': source_info['unique_media_count'],
        'source_candidates': source_info['candidate_count'],
        'pptx_crop_candidates': source_info['pptx_crop_candidate_count'],
        'pptx_semantic_crop_replaced': semantic_crop_replaced,
        'pdf_images': len(pdf_images),
        'replaced': replaced,
        'fallback_enabled': fallback_enabled,
        'fallback_candidates': len(fallback_candidates),
        'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
