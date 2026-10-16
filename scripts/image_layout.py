"""以來源 PDF 的實際字級和 DPI 決定整科共同倍率，避免短題被放大。"""
from collections import Counter
from pathlib import Path
import re

from PIL import Image
import pymupdf


def source_body_font(source, records):
    weights = Counter()
    for record in records:
        for part in record['parts']:
            clip = pymupdf.Rect(part['rect'])
            for block in source[part['page']-1].get_text('dict', clip=clip)['blocks']:
                for line in block.get('lines', []):
                    for span in line['spans']:
                        count = len(re.findall(r'[A-Za-z\u3400-\u9fff]', span['text']))
                        if count >= 2 and 8 <= span['size'] <= 24:
                            weights[round(span['size']*2)/2] += count
    return weights.most_common(1)[0][0] if weights else None


def image_size_cm(path, dpi):
    with Image.open(path) as image:
        return image.width/dpi*2.54, image.height/dpi*2.54


def plan_question_scale(source, records, folder, dpi, width_cm, height_cm, target_pt=11):
    """同一本題本只採一個倍率；題圖過寬時整科一起縮小，不逐題塞滿。"""
    if dpi <= 0:
        raise ValueError('截圖 DPI 必須大於零')
    body_pt = source_body_font(source, records)
    desired = target_pt/body_pt if body_pt else 1.0
    limits = [desired]
    sizes = {}
    for record in records:
        width, height = image_size_cm(Path(folder)/record['image'], dpi)
        limits.extend([width_cm/width, height_cm/height])
        sizes[record['id']] = (width, height)
    scale = min(limits)
    return {'mode':'uniform_source_scale', 'source_body_font_pt':body_pt,
            'target_body_font_pt':target_pt, 'effective_body_font_pt':body_pt*scale if body_pt else None,
            'source_scale':scale, 'dpi':dpi, 'sizes':sizes,
            'font_detection':'pdf_text' if body_pt else 'natural_size_fallback'}
