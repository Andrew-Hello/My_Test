#!/usr/bin/env python3
"""Restore authoritative OOXML raster media into an Office-generated PDF.

Production dependencies:
- PyMuPDF (import name: pymupdf)
- Pillow

The rebuilder understands DrawingML picture relationships and srcRect crops for
PPTX and DOCX, and reconstructs PDF soft-mask transparency before perceptual
matching. This keeps the matching threshold strict while improving correctness
for cropped or transparent Office images.
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
NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'a': NS_A, 'p': NS_P, 'pic': NS_PIC, 'w': NS_W, 'r': NS_R}

MATCHING_ENGINE = 'builtin_phash_dct16+pdf_smask+pptx_srcRect+docx_srcRect'

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
            row_low[y][u] = sum(value * cos_u[x] for x, value in enumerate(row)) * _DCT_ALPHA[u]

    coeffs = []
    for v in range(PHASH_SIZE):
        cos_v = _DCT_COS[v]
        alpha_v = _DCT_ALPHA[v]
        for u in range(PHASH_SIZE):
            coeffs.append(sum(row_low[y][u] * cos_v[y] for y in range(PHASH_IMAGE_SIZE)) * alpha_v)

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
        try:
            im = pil_from_bytes(zf.read(name))
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
        posixpath.dirname(owner_path), '_rels', posixpath.basename(owner_path) + '.rels'
    )
    if rel_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_path))
    out = {}
    base_dir = posixpath.dirname(owner_path)
    for rel in root.findall(f'{{{NS_REL}}}Relationship'):
        rid = rel.get('Id')
        target = rel.get('Target')
        if not rid or not target or rel.get('TargetMode') == 'External':
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
    """Apply DrawingML srcRect values (1/1000 percent, normally 0..100000)."""
    l = max(0.0, min(0.99999, rect.get('l', 0) / 100000.0))
    t = max(0.0, min(0.99999, rect.get('t', 0) / 100000.0))
    r = max(0.0, min(0.99999, rect.get('r', 0) / 100000.0))
    b = max(0.0, min(0.99999, rect.get('b', 0) / 100000.0))
    left = max(0, min(int(round(im.width * l)), im.width - 1))
    top = max(0, min(int(round(im.height * t)), im.height - 1))
    right = max(left + 1, min(int(round(im.width * (1.0 - r))), im.width))
    bottom = max(top + 1, min(int(round(im.height * (1.0 - b))), im.height))
    if left == 0 and top == 0 and right == im.width and bottom == im.height:
        return im.copy(), [0, 0, im.width, im.height], False
    return im.crop((left, top, right, bottom)), [left, top, right, bottom], True


def append_unreferenced_media(candidates, media, referenced_media, kind='ooxml_media_unreferenced'):
    for media_name, source in media.items():
        if media_name in referenced_media:
            continue
        try:
            candidates.append(build_source_item(
                media_name,
                source['image'].copy(),
                ext=source['ext'],
                media_name=media_name,
                candidate_kind=kind,
                full_source_px=[source['width'], source['height']],
            ))
        except Exception:
            continue


def load_pptx_candidates(zf: zipfile.ZipFile, media):
    candidates = []
    referenced_media = set()
    crop_count = 0
    slide_names = sorted(
        name for name in zf.namelist()
        if name.startswith('ppt/slides/slide') and name.endswith('.xml') and '/_rels/' not in name
    )
    for slide_name in slide_names:
        try:
            root = ET.fromstring(zf.read(slide_name))
            rels = relationship_map(zf, slide_name)
        except Exception:
            continue
        for pic_index, pic in enumerate(root.findall('.//p:pic', NS), start=1):
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
            rect = parse_src_rect(pic.find('.//a:srcRect', NS))
            candidate_im, crop_box, crop_applied = crop_by_src_rect(full_im, rect)
            crop_count += int(crop_applied)
            c_nv_pr = pic.find('.//p:cNvPr', NS)
            picture_name = c_nv_pr.get('name') if c_nv_pr is not None else None
            display_name = f"{media_name}#slide={Path(slide_name).stem};picture={picture_name or pic_index}"
            try:
                candidates.append(build_source_item(
                    display_name, candidate_im,
                    ext=source['ext'], media_name=media_name,
                    candidate_kind='pptx_slide_picture',
                    pptx_slide=slide_name, pptx_picture_name=picture_name,
                    pptx_crop_100k=rect, pptx_crop_box_px=crop_box,
                    pptx_crop_applied=crop_applied,
                    full_source_px=[full_im.width, full_im.height],
                ))
            except Exception:
                continue
    append_unreferenced_media(candidates, media, referenced_media)
    return candidates, crop_count


def docx_owner_names(zf: zipfile.ZipFile):
    owners = []
    for name in zf.namelist():
        if not name.endswith('.xml') or '/_rels/' in name:
            continue
        base = posixpath.basename(name)
        if name == 'word/document.xml':
            owners.append(name)
        elif name.startswith('word/header') or name.startswith('word/footer'):
            owners.append(name)
        elif name in {
            'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml',
            'word/glossary/document.xml',
        }:
            owners.append(name)
        elif name.startswith('word/comments') and base.endswith('.xml'):
            owners.append(name)
    return sorted(set(owners))


def load_docx_candidates(zf: zipfile.ZipFile, media):
    """Build per-picture DOCX candidates from relationships and DrawingML srcRect."""
    candidates = []
    referenced_media = set()
    crop_count = 0
    semantic_picture_count = 0

    for owner_name in docx_owner_names(zf):
        try:
            root = ET.fromstring(zf.read(owner_name))
            rels = relationship_map(zf, owner_name)
        except Exception:
            continue

        for pic_index, pic in enumerate(root.findall('.//pic:pic', NS), start=1):
            blip = pic.find('.//a:blip', NS)
            if blip is None:
                continue
            rid = blip.get(f'{{{NS_R}}}embed')
            media_name = rels.get(rid)
            if not media_name or media_name not in media:
                continue

            referenced_media.add(media_name)
            semantic_picture_count += 1
            source = media[media_name]
            full_im = source['image']
            rect = parse_src_rect(pic.find('.//a:srcRect', NS))
            candidate_im, crop_box, crop_applied = crop_by_src_rect(full_im, rect)
            crop_count += int(crop_applied)

            c_nv_pr = pic.find('.//pic:cNvPr', NS)
            picture_name = c_nv_pr.get('name') if c_nv_pr is not None else None
            display_name = f"{media_name}#owner={owner_name};picture={picture_name or pic_index}"
            try:
                candidates.append(build_source_item(
                    display_name, candidate_im,
                    ext=source['ext'], media_name=media_name,
                    candidate_kind='docx_picture',
                    docx_owner=owner_name, docx_picture_name=picture_name,
                    docx_crop_100k=rect, docx_crop_box_px=crop_box,
                    docx_crop_applied=crop_applied,
                    full_source_px=[full_im.width, full_im.height],
                ))
            except Exception:
                continue

    append_unreferenced_media(candidates, media, referenced_media)
    if not candidates:
        append_unreferenced_media(candidates, media, set(), kind='ooxml_media')
    return candidates, crop_count, semantic_picture_count


def load_generic_candidates(media):
    candidates = []
    append_unreferenced_media(candidates, media, set(), kind='ooxml_media')
    return candidates


def load_ooxml_sources(office_path: Path):
    ext = office_path.suffix.lower()
    if ext not in OOXML_EXTS:
        raise ValueError(f'High-resolution media recovery supports OOXML only: {sorted(OOXML_EXTS)}')

    with zipfile.ZipFile(office_path, 'r') as zf:
        media, media_errors = read_raw_media(zf)
        pptx_crop_count = 0
        docx_crop_count = 0
        docx_picture_count = 0
        if ext == '.pptx':
            candidates, pptx_crop_count = load_pptx_candidates(zf, media)
        elif ext == '.docx':
            candidates, docx_crop_count, docx_picture_count = load_docx_candidates(zf, media)
        else:
            candidates = load_generic_candidates(media)

    return candidates, {
        'unique_media_count': len(media),
        'candidate_count': len(candidates),
        'pptx_crop_candidate_count': pptx_crop_count,
        'docx_crop_candidate_count': docx_crop_count,
        'docx_semantic_picture_count': docx_picture_count,
        'media_errors': media_errors,
    }


def extract_pdf_image(doc, xref: int, smask: int):
    """Extract a PDF image and recompose its soft mask when Office split alpha."""
    if smask:
        base = fitz.Pixmap(doc, xref)
        mask = fitz.Pixmap(doc, smask)
        try:
            if base.alpha:
                base = fitz.Pixmap(base, 0)
            combined = fitz.Pixmap(base, mask)
            return pil_from_bytes(combined.tobytes('png')), 'png', True
        finally:
            base = None
            mask = None
    extracted = doc.extract_image(xref)
    return pil_from_bytes(extracted['image']), extracted.get('ext'), False


def collect_pdf_images(doc):
    images = {}
    soft_mask_count = 0
    for page_index in range(len(doc)):
        page = doc[page_index]
        for item in page.get_images(full=True):
            xref = int(item[0])
            smask = int(item[1]) if len(item) > 1 and item[1] else 0
            if xref in images:
                images[xref]['pages'].append(page_index)
                continue
            try:
                im, ext, smask_applied = extract_pdf_image(doc, xref, smask)
                phash = perceptual_hash(im)
                soft_mask_count += int(smask_applied)
            except Exception as exc:
                images[xref] = {
                    'xref': xref, 'smask_xref': smask, 'pages': [page_index],
                    'error': str(exc), 'image': None,
                }
                continue
            images[xref] = {
                'xref': xref, 'smask_xref': smask, 'soft_mask_applied': smask_applied,
                'pages': [page_index], 'image': im,
                'width': im.width, 'height': im.height, 'ext': ext,
                'phash': phash, 'hash': hash_text(phash),
            }
    return list(images.values()), soft_mask_count


def encode_png(im: Image.Image):
    buf = io.BytesIO()
    im.save(buf, format='PNG', optimize=False)
    return buf.getvalue()


def save_passthrough(doc, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path, garbage=4, deflate=True, clean=True)


def match_detail(m, s, p):
    for key in (
        'media_name', 'candidate_kind', 'pptx_slide', 'pptx_picture_name',
        'pptx_crop_100k', 'pptx_crop_box_px', 'pptx_crop_applied',
        'docx_owner', 'docx_picture_name', 'docx_crop_100k',
        'docx_crop_box_px', 'docx_crop_applied', 'full_source_px',
    ):
        if key in s:
            m[key] = s[key]
    m['pdf_smask_xref'] = p.get('smask_xref', 0)
    m['pdf_soft_mask_applied_for_matching'] = bool(p.get('soft_mask_applied'))


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
    collected, soft_mask_count = collect_pdf_images(doc)
    pdf_images = [x for x in collected if x.get('image') is not None]

    if not sources or not pdf_images:
        save_passthrough(doc, args.output_pdf)
        doc.close()
        report = {
            'schema_version': 5,
            'matching_engine': MATCHING_ENGINE,
            'source_office': args.source_office.name,
            'input_pdf': args.input_pdf.name,
            'output_pdf': args.output_pdf.name,
            'source_raster_media_count': source_info['unique_media_count'],
            'source_candidate_count': source_info['candidate_count'],
            'pptx_crop_candidate_count': source_info['pptx_crop_candidate_count'],
            'docx_crop_candidate_count': source_info['docx_crop_candidate_count'],
            'docx_semantic_picture_count': source_info['docx_semantic_picture_count'],
            'pdf_soft_mask_reconstructed_count': soft_mask_count,
            'pdf_unique_raster_image_count': len(pdf_images),
            'matched_count': 0,
            'replaced_count': 0,
            'passthrough_reason': 'no_supported_source_raster_media' if not sources else 'no_pdf_raster_images',
            'matches': [], 'replacement_errors': [],
            'media_errors': source_info['media_errors'],
            'output_size_bytes': args.output_pdf.stat().st_size,
        }
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({
            'source_images': source_info['unique_media_count'],
            'source_candidates': source_info['candidate_count'],
            'pptx_crop_candidates': source_info['pptx_crop_candidate_count'],
            'docx_crop_candidates': source_info['docx_crop_candidate_count'],
            'pdf_soft_masks_reconstructed': soft_mask_count,
            'pdf_images': len(pdf_images), 'replaced': 0, 'passthrough': True,
            'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
        }, ensure_ascii=False, indent=2))
        return

    pairs = []
    for pi, p in enumerate(pdf_images):
        for si, s in enumerate(sources):
            pairs.append((similarity_score(p, s), pi, si))
    pairs.sort(key=lambda x: x[0])

    used_pdf, used_src, matches = set(), set(), []
    for score, pi, si in pairs:
        if pi in used_pdf or si in used_src:
            continue
        p, s = pdf_images[pi], sources[si]
        used_pdf.add(pi)
        used_src.add(si)
        area_ratio = (s['width'] * s['height']) / max(p['width'] * p['height'], 1)
        ar_delta = aspect_delta_log(p['image'], s['image'])
        confident_replace = bool(score <= args.max_score and area_ratio >= 0.95)

        if confident_replace and s.get('pptx_crop_applied'):
            mode = 'pptx_semantic_crop_match'
        elif confident_replace and s.get('docx_crop_applied'):
            mode = 'docx_semantic_crop_match'
        elif confident_replace and s.get('candidate_kind') == 'docx_picture':
            mode = 'docx_semantic_picture_match'
        elif confident_replace and area_ratio <= 1.05:
            mode = 'same_resolution_source_preservation'
        elif confident_replace:
            mode = 'confident_hash_highres'
        elif area_ratio < 0.95:
            mode = 'source_smaller_than_pdf'
        else:
            mode = 'initially_rejected'

        m = {
            'score': round(score, 4), 'pdf_index': pi, 'src_index': si,
            'pdf_xref': p['xref'], 'pdf_px': [p['width'], p['height']],
            'source_name': s['name'], 'source_px': [s['width'], s['height']],
            'pixel_area_ratio': round(area_ratio, 3),
            'aspect_delta_log': round(ar_delta, 5),
            'replace': confident_replace, 'match_mode': mode,
        }
        match_detail(m, s, p)
        matches.append(m)

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
    pptx_semantic_crop_replaced = 0
    docx_semantic_crop_replaced = 0
    docx_semantic_picture_replaced = 0
    for m in matches:
        if not m['replace']:
            continue
        p, s = pdf_images[m['pdf_index']], sources[m['src_index']]
        try:
            doc[p['pages'][0]].replace_image(p['xref'], stream=encode_png(s['image']))
            replaced += 1
            pptx_semantic_crop_replaced += int(m['match_mode'] == 'pptx_semantic_crop_match')
            docx_semantic_crop_replaced += int(m['match_mode'] == 'docx_semantic_crop_match')
            docx_semantic_picture_replaced += int(m['match_mode'] in {
                'docx_semantic_crop_match', 'docx_semantic_picture_match'
            })
        except Exception as exc:
            m['replace'] = False
            m['match_mode'] = 'replacement_error'
            m['error'] = str(exc)
            replacement_errors.append({'xref': p['xref'], 'error': str(exc)})

    args.output_pdf.parent.mkdir(parents=True, exist_ok=True)
    doc.save(args.output_pdf, garbage=4, deflate=True, clean=True)
    doc.close()

    report = {
        'schema_version': 5,
        'matching_engine': MATCHING_ENGINE,
        'source_office': args.source_office.name,
        'input_pdf': args.input_pdf.name,
        'output_pdf': args.output_pdf.name,
        'source_raster_media_count': source_info['unique_media_count'],
        'source_candidate_count': source_info['candidate_count'],
        'pptx_crop_candidate_count': source_info['pptx_crop_candidate_count'],
        'pptx_semantic_crop_replaced_count': pptx_semantic_crop_replaced,
        'docx_semantic_picture_count': source_info['docx_semantic_picture_count'],
        'docx_crop_candidate_count': source_info['docx_crop_candidate_count'],
        'docx_semantic_picture_replaced_count': docx_semantic_picture_replaced,
        'docx_semantic_crop_replaced_count': docx_semantic_crop_replaced,
        'pdf_soft_mask_reconstructed_count': soft_mask_count,
        'pdf_unique_raster_image_count': len(pdf_images),
        'matched_count': len(matches), 'replaced_count': replaced,
        'normal_max_score': args.max_score,
        'fallback_enabled': fallback_enabled,
        'fallback_candidate_count': len(fallback_candidates),
        'matches': matches, 'replacement_errors': replacement_errors,
        'media_errors': source_info['media_errors'],
        'output_size_bytes': args.output_pdf.stat().st_size if args.output_pdf.exists() else None,
    }
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')

    print(json.dumps({
        'source_images': source_info['unique_media_count'],
        'source_candidates': source_info['candidate_count'],
        'pptx_crop_candidates': source_info['pptx_crop_candidate_count'],
        'pptx_semantic_crop_replaced': pptx_semantic_crop_replaced,
        'docx_semantic_pictures': source_info['docx_semantic_picture_count'],
        'docx_crop_candidates': source_info['docx_crop_candidate_count'],
        'docx_semantic_picture_replaced': docx_semantic_picture_replaced,
        'docx_semantic_crop_replaced': docx_semantic_crop_replaced,
        'pdf_soft_masks_reconstructed': soft_mask_count,
        'pdf_images': len(pdf_images), 'replaced': replaced,
        'fallback_enabled': fallback_enabled,
        'fallback_candidates': len(fallback_candidates),
        'output_mb': round(args.output_pdf.stat().st_size / 1024 / 1024, 3),
    }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
