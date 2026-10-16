# /// script
# requires-python = ">=3.11"
# dependencies = ["pymupdf>=1.24", "pillow"]
# ///
"""依題號輸出題本截圖，並擷取既有成果報告的題目欄作為排版參考。

python question_screenshots.py --source both
python question_screenshots.py --source booklets --pdf 題本.pdf
所有頁碼均為 PDF 的實體頁碼（從 1 開始），座標單位為 PDF point。
"""
from __future__ import annotations

import argparse
import csv
import fnmatch
import html
import io
import json
import re
from pathlib import Path

import pymupdf as fitz
from PIL import Image, ImageChops, ImageOps, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
CODE = re.compile(r"(?P<year>\d{3})-(?P<subject>[CME])(?P<grade>\d)-[A-Z]\d-(?P<number>\d{2})")
QSTART = re.compile(r"^\s*(\d{1,2})\s*[.．、](?!\d)")
SUBJECTS = {"C": "國語文", "M": "數學", "E": "英語文"}
GRADES = dict(zip("三四五六七八九", range(3, 10)))


def lines(page):
    return [(fitz.Rect(line["bbox"]), "".join(s["text"] for s in line["spans"]).strip())
            for block in page.get_text("dict")["blocks"]
            for line in block.get("lines", [])]


def identity(path):
    name = path.stem
    year = re.search(r"\d{3}", name)
    grade = re.search(r"([三四五六七八九\d])年級", name)
    if grade is None and path.is_file():
        # 「國中新生」檔名可能省略年級，封面才載明七年級。
        with fitz.open(path) as document:
            if not document.needs_pass and len(document):
                grade = re.search(r"([三四五六七八九\d])年級", re.sub(r"\s+", "", document[0].get_text()))
    subject = next((s for s in ["國語文", "數學", "英語", "理化", "生物", "自然"] if s in name), "未分類")
    subject = "英語文" if subject == "英語" else subject
    g = GRADES.get(grade[1], grade[1]) if grade else "未知"
    region = "桃園市_" if "桃園市" in name else ""
    return f"{year[0] if year else '未知'}_{region}{subject}_{g}年級"


def trim(image, padding=16):
    image = image.convert("RGB")
    diff = ImageChops.difference(image, Image.new("RGB", image.size, "white"))
    # 忽略接近白色的掃描底色，但保留抗鋸齒字緣。
    bbox = diff.convert("L").point(lambda v: 255 if v > 12 else 0).getbbox()
    if bbox:
        image = image.crop(bbox)
    return ImageOps.expand(image, border=padding, fill="white")


def render(doc, parts, dpi):
    images = []
    for part in parts:
        page = doc[part["page"] - 1]
        rect = fitz.Rect(part["rect"]) & page.rect
        if rect.is_empty or rect.width < 2 or rect.height < 2:
            raise ValueError(f"無效裁切範圍: {part}")
        pix = page.get_pixmap(dpi=dpi, clip=rect, alpha=False)
        image = Image.open(io.BytesIO(pix.tobytes("png"))).convert('RGB')
        draw = ImageDraw.Draw(image)
        for excluded in part.get('exclude_rects', []):
            box = fitz.Rect(excluded) & rect
            if not box.is_empty:
                # Pixmap 原點為取整後的裝置座標；只移除題號字元，不裁掉同列題幹。
                draw.rectangle((int(box.x0*dpi/72-pix.x), int(box.y0*dpi/72-pix.y),
                                int(box.x1*dpi/72-pix.x)+1, int(box.y1*dpi/72-pix.y)+1), fill='white')
        images.append(trim(image))
    width = max(im.width for im in images)
    result = Image.new("RGB", (width, sum(im.height for im in images) + 24 * (len(images) - 1)), "white")
    top = 0
    for im in images:
        result.paste(im, (0, top))
        top += im.height + 24
    return result


def exclude_question_number(doc, record):
    """精確定位第一個裁切區中的題號字元，保留題幹數字、選項與圖框。"""
    if record.get('kind') != 'question' or not record.get('parts'):
        return
    part = record['parts'][0]
    clip = fitz.Rect(part['rect'])
    candidates = []
    for block in doc[part['page']-1].get_text('rawdict')['blocks']:
        for line in block.get('lines', []):
            chars = [c for span in line['spans'] for c in span['chars']]
            text = ''.join(c['c'] for c in chars)
            match = QSTART.match(text)
            if not match or int(match[1]) != record['number']:
                continue
            box = fitz.Rect(chars[match.start(1)]['bbox'])
            for char in chars[match.start(1)+1:match.end()]:
                if char['c'].strip():
                    box |= fitz.Rect(char['bbox'])
            if clip.contains(box):
                candidates.append(box)
    if candidates:
        box = min(candidates, key=lambda r: (r.y0, r.x0))
        part['exclude_rects'] = [list(box + (-.2, -.2, .2, .2))]
        record['question_number_removed'] = True


def exclude_section_heading(doc, record):
    """排除題尾框外、獨立成行的下一大題標題；保留文章內相同文字。"""
    if record.get('kind') != 'question':
        return
    pattern=re.compile(r'^[一二三四五六七八九十]+[、．.](?:閱讀理解|語文知識)(?:題|測驗)?[：:。]?$')
    for part in record.get('parts',[]):
        page=doc[part['page']-1]
        clip=fitz.Rect(part['rect'])
        frames=text_frames(page)
        for box,text in lines(page):
            if (pattern.fullmatch(re.sub(r'\s+','',text)) and clip.contains(box)
                    and box.y0>clip.y0+15 and not any(frame.contains(box) for frame in frames)):
                part['excluded_section_heading']={'text':text,'rect':list(box),'original_bottom':clip.y1}
                part['rect'][3]=box.y0-2
                break


def exclude_exam_transition(doc, record):
    """排除聽力結束／閱讀開始的考場提示，不能將它當成最後一題內容。"""
    if record.get('kind') != 'question':
        return
    for part in record.get('parts', []):
        clip = fitz.Rect(part['rect'])
        for box, text in lines(doc[part['page']-1]):
            compact = re.sub(r'\s+', '', text)
            if ('聽力試題已播音完畢' in compact and '閱讀測驗' in compact
                    and clip.intersects(box) and box.y0 > clip.y0):
                part['excluded_exam_instruction'] = {'text':text, 'original_bottom':clip.y1}
                part['rect'][3] = min(clip.y1, box.y0-3)
                break


def question_cell(page, label):
    """以題目欄的實際表格邊線定位，不把垂直置中的「題目」當上緣。"""
    mx, my = (label.x0 + label.x1) / 2, (label.y0 + label.y1) / 2
    horizontal, vertical = [], []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "re":
                r = fitz.Rect(item[1])
            elif item[0] == "l":
                r = fitz.Rect(item[1], item[2]); r.normalize()
            else:
                continue
            if r.height <= 2 and r.width > 20 and r.x0 <= mx <= r.x1:
                horizontal.append((r.y0 + r.y1) / 2)
            if r.width <= 2 and r.height > 8 and r.y0 <= my <= r.y1:
                vertical.append((r.x0 + r.x1) / 2)
    above = [y for y in horizontal if y < label.y0]
    below = [y for y in horizontal if y > label.y1]
    right = [x for x in vertical if x > label.x1]
    if not above or not below or len(right) < 2:
        raise ValueError("無法辨認題目欄邊線，請在 overrides.json 指定範圍")
    return [min(right) + 1.5, max(above) + 1.5, max(right) - 1.5, min(below) - 1.5]


def report_records(path, doc):
    records = []
    for i, page in enumerate(doc):
        matches = list(CODE.finditer(page.get_text()))
        if not matches:
            continue
        for m in matches:
            labels = [r for r, t in lines(page) if t == "題目"]
            record = {"id": f"q{int(m['number'])}", "number": int(m["number"]),
                      "code": m[0], "kind": "question", "parts": [], "warnings": []}
            if len(matches) != 1 or len(labels) != 1:
                record["warnings"].append("本頁非單一試題表格；需指定裁切範圍")
            else:
                try:
                    rect = fitz.Rect(question_cell(page, labels[0]))
                    # 有些原稿圖片超出表格右框（如數學六年級第 5 題）。
                    # 保留該列完整圖片，不能只按框線截掉末端選項或圖形。
                    for im in page.get_image_info():
                        bbox = fitz.Rect(im["bbox"])
                        if im["width"] > 100 and im["height"] > 35 and bbox.intersects(rect) and bbox.y0 >= rect.y0-2 and bbox.y1 <= rect.y1+2:
                            rect |= bbox
                    record["parts"] = [{"page": i + 1, "rect": list(rect)}]
                except ValueError as exc:
                    record["warnings"].append(str(exc))
            records.append(record)
    # 閱讀題的共用文章常放在獨立頁，不在任一「題目」欄中。
    # 保留試題分析章節內的獨立大圖，供核對與插入；不混入題號序列。
    if records:
        pages = [part["page"] for r in records for part in r["parts"]]
        if pages:
            occupied = set(pages)
            for page_number in range(min(pages), max(pages) + 1):
                if page_number in occupied:
                    continue
                page = doc[page_number - 1]
                images = [fitz.Rect(im["bbox"]) for im in page.get_image_info()
                          if im["width"] > 100 and im["height"] > 100]
                if not images:
                    continue
                rect = images[0]
                for other in images[1:]:
                    rect |= other
                records.append({"id": f"material_p{page_number:03d}", "kind": "material", "number": None,
                                "parts": [{"page": page_number, "rect": list(rect + (-1, -1, 1, 1))}],
                                "warnings": ["獨立圖文頁，請核對適用題號；可能為共用文章或教學圖例"]})
    return records


def page_bounds(page):
    bottom = page.rect.height * .93
    for rect, text in lines(page):
        if re.fullmatch(r"(?:\d+[.．、]\s*)?(?:作答完畢|試題結束)[。！!]?", text):
            bottom = min(bottom, rect.y0 - 3)
        if rect.y0 > page.rect.height * .85 and (
            (re.fullmatch(r"[-－—\s]*\d{1,3}[-－—\s]*", text)
             and rect.y0 > page.rect.height * .90
             and .40 < (rect.x0+rect.x1)/2/page.rect.width < .60)
            or any(s in text for s in ["請翻頁", "請繼續", "作答完畢", "試題結束"])):
            bottom = min(bottom, rect.y0 - 3)
    return 45, bottom


def text_frames(page):
    """找出封閉文字框；Word PDF 常將四邊分別輸出為細長填色矩形。"""
    horizontal, vertical, frames = [], [], []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "re":
                rect = fitz.Rect(item[1])
                if rect.width >= 60 and rect.height >= 18 and "s" in drawing["type"]:
                    frames.append(rect)
            elif item[0] == "l":
                rect = fitz.Rect(item[1], item[2])
                rect.normalize()
            else:
                continue
            if rect.width >= 60 and rect.height <= 2:
                horizontal.append(rect)
            elif rect.height >= 18 and rect.width <= 2:
                vertical.append(rect)
    for top in horizontal:
        for bottom in horizontal:
            # 原稿邊角可能有約 2pt 的伸出／錯位，不要求四邊端點完全相等。
            if bottom.y0 < top.y1 + 18 or abs(top.x0-bottom.x0) > 2.5 or abs(top.x1-bottom.x1) > 2.5:
                continue
            def side_at(x):
                return [v for v in vertical if abs((v.x0+v.x1)/2-x) <= 2.5
                        and v.y0 <= top.y1+1.5 and v.y1 >= bottom.y0-1.5]
            left, right = side_at(top.x0), side_at(top.x1)
            if left and right:
                frames.append(top | bottom | left[0] | right[0])
    return frames


def question_top(start, frames, x0, x1):
    """題號旁的文字框可能比題號高；整個框應歸入本題，不能留碎線給前題。"""
    y = start["rect"].y0
    boundary = y - 5
    for frame in frames:
        if frame.x0 >= x0 and frame.x1 <= x1 and y-24 <= frame.y0 <= y+2 and frame.y1 > y+10:
            boundary = min(boundary, frame.y0 - 4)
    return max(0, boundary)


def link_reading_passages(records):
    """沒有題組標題的獨立文章頁，與次頁續文合併，依閱讀順序連到小題。"""
    questions = sorted((r for r in records if r["kind"] == "question"), key=lambda r: r["number"])
    question_pages = {p["page"] for r in questions for p in r["parts"]}
    candidates = sorted((r for r in records if r["id"].startswith("material_p")),
                        key=lambda r: r["parts"][0]["page"])
    runs, current = [], []
    for material in candidates:
        page = material["parts"][0]["page"]
        if current and page != current[-1]["parts"][-1]["page"]+1:
            runs.append(current)
            current = []
        if not current and page in question_pages:
            continue  # 單純題前說明不能僅憑位置視為閱讀文章。
        current.append(material)
        if page in question_pages:
            runs.append(current)
            current = []
    if current:
        runs.append(current)
    linked = []
    for run in runs:
        last = run[-1]["parts"][-1]
        following = [r for r in questions if any(p["page"] > last["page"] or
            (p["page"] == last["page"] and p["rect"][1] >= last["rect"][3]-.1) for p in r["parts"])]
        if following:
            linked.append((following[0]["number"], run))
    consumed = set()
    for i, (start, run) in enumerate(linked):
        later_starts = [s for s, _ in linked[i+1:]]
        later_starts += [r["question_range"][0] for r in records if r.get("question_range") and r["question_range"][0] > start]
        end = min(later_starts)-1 if later_starts else questions[-1]["number"]
        group_id = f"group_{start}-{end}"
        records.append({"id": group_id, "kind": "material", "number": None,
            "question_range": [start, end], "parts": [part for r in run for part in r["parts"]],
            "warnings": ["依獨立文章頁及後續小題順序建立關聯，請核對適用題號"]})
        consumed.update(r["id"] for r in run)
        for question in questions:
            if start <= question["number"] <= end:
                question.setdefault("materials", []).append(group_id)
    return [r for r in records if r["id"] not in consumed]


def booklet_records(path, doc):
    records, expected, active_group = [], 1, None
    group_records = {}
    for i, page in enumerate(doc):
        top, bottom = page_bounds(page)
        frames = text_frames(page) if "國語" in path.name else []
        page_groups = []
        for r, text in lines(page):
            compact = re.sub(r"\s+", "", text)
            match = re.search(r"(?:第)?(\d{1,2})[-~～至－—](\d{1,2})", compact)
            if match and ("請閱讀" in compact or compact.startswith("◆") or compact.startswith("題組")):
                page_groups.append({"start": int(match[1]), "end": int(match[2]), "rect": r})
        candidates = []
        for r, text in lines(page):
            m = QSTART.match(text)
            if m and not re.fullmatch(r"\d+[.．、]\s*(作答完畢|試題結束)[。！!]?", text) and top <= r.y0 < bottom and (r.x0 < page.rect.width * .18 or .47 < r.x0 / page.rect.width < .60):
                candidates.append({"number": int(m[1]), "rect": r, "text": text})
        double = any(c["rect"].x0 > page.rect.width * .47 for c in candidates)
        candidates.sort(key=lambda c: (int(double and c["rect"].x0 > page.rect.width / 2), c["rect"].y0))
        starts = []
        for candidate in candidates:
            if candidate["number"] == expected:
                starts.append(candidate)
                expected += 1
        for start in starts:
            right_column = double and start["rect"].x0 > page.rect.width/2
            start["crop_top"] = question_top(start, frames,
                page.rect.width/2+2 if right_column else 18,
                page.rect.width/2-2 if double and not right_column else page.rect.width-18)
        for group in page_groups:
            following = [s["crop_top"] for s in starts if s["rect"].y0 > group["rect"].y0 + 3]
            following += [g["rect"].y0 - 5 for g in page_groups if g["rect"].y0 > group["rect"].y0 + 3]
            end = min(following) if following else bottom
            group_id = f"group_{group['start']}-{group['end']}"
            active_group = {"id": group_id, **group}
            record = group_records.setdefault(group_id, {"id": group_id, "kind": "material", "number": None,
                "question_range": [group["start"], group["end"]], "parts": [],
                "warnings": ["自動辨認題組材料，請檢查起訖與跨頁內容"]})
            record["parts"].append({"page": i + 1, "rect": [18, max(top, group["rect"].y0-5), page.rect.width-18, end]})
        # 原始題本中無法辨認的文章頁另外保存，避免把整頁閱讀材料漏掉。
        if not starts:
            if page_groups:
                continue
            if active_group and expected <= active_group["end"]:
                group_records[active_group["id"]]["parts"].append(
                    {"page": i+1, "rect": [18, top, page.rect.width-18, bottom]})
            elif expected > 1 and i < len(doc) - 1:
                records.append({"id": f"material_p{i+1:03d}", "kind": "material", "number": None,
                                "parts": [{"page": i + 1, "rect": [18, top, page.rect.width - 18, bottom]}],
                                "warnings": ["未含新題號的頁面；請確認是否為共用文章或跨頁續文"]})
            continue
        for col in range(2 if double else 1):
            column = [s for s in starts if int(double and s["rect"].x0 > page.rect.width / 2) == col]
            x0 = page.rect.width / 2 + 2 if col else 18
            x1 = page.rect.width / 2 - 2 if double and not col else page.rect.width - 18
            boundaries = [s["crop_top"] for s in column]
            for j, start in enumerate(column):
                y0 = boundaries[j]
                y1 = boundaries[j+1] if j+1 < len(column) else bottom
                next_groups = [g["rect"].y0 - 5 for g in page_groups if y0 < g["rect"].y0 - 5 < y1]
                if next_groups:
                    y1 = min(next_groups)
                records.append({"id": f"q{start['number']}", "kind": "question", "number": start["number"],
                                "parts": [{"page": i + 1, "rect": [x0, y0, x1, y1]}],
                                "warnings": ["題本自動裁切：請核對選項、題組材料與跨頁內容"]})
        # 含題號的頁面，上方也可能是跨頁文章或聽力說明。
        first_y = min(s["crop_top"] for s in starts)
        first_header = min((g["rect"].y0-5 for g in page_groups), default=bottom)
        prefix_end = min(first_y, first_header)
        if prefix_end - top > 65:
            if active_group and not page_groups and starts[0]["number"] <= active_group["end"]:
                group_records[active_group["id"]]["parts"].append(
                    {"page": i+1, "rect": [18, top, page.rect.width-18, prefix_end]})
            else:
                records.append({"id": f"material_p{i+1:03d}", "kind": "material", "number": None,
                            "parts": [{"page": i + 1, "rect": [18, top, page.rect.width - 18, prefix_end]}],
                            "warnings": ["題目前的共用材料／說明，需人工確認適用題號"]})
    for record in records:
        if record["number"] is not None:
            record["materials"] = [g["id"] for g in group_records.values()
                                   if g["question_range"][0] <= record["number"] <= g["question_range"][1]]
    records.extend(group_records.values())
    if "國語" in path.name:
        records = link_reading_passages(records)
    return sorted(records, key=lambda r: (r["number"] is None, r["number"] or r["id"]))


def write_preview(folder, records, title):
    by_id = {r["id"]: r for r in records}
    ordered, emitted = [], set()
    for r in records:
        if r["kind"] != "question":
            continue
        for material_id in r.get("materials", []):
            if material_id in by_id and material_id not in emitted:
                ordered.append(by_id[material_id])
                emitted.add(material_id)
        ordered.append(r)
        emitted.add(r["id"])
    ordered.extend(r for r in records if r["id"] not in emitted)
    cards = []
    navigation = []
    for r in ordered:
        image = r.get("image")
        warnings = "；".join(r["warnings"])
        heading = r["id"]
        aliases = ""
        if r.get("question_range"):
            start, end = r["question_range"]
            heading = f"共用文章／題組材料（第 {start}～{end} 題）"
            aliases = " ".join(f"q{n}" for n in range(start, end+1))
            navigation.append(f'<a href="#{r["id"]}">{start}～{end} 題文章</a>')
        links = '、'.join(f'<a href="#{key}">查看共用文章／材料</a>' for key in r.get("materials", []) if key in by_id)
        cards.append(f'<article id="{html.escape(r["id"])}" data-key="{html.escape(r["id"]+" "+r.get("code", "")+" "+aliases)}">'
                     f'<h2>{html.escape(heading)} <small>{html.escape(r.get("code", ""))}</small></h2>'
                     f'<p>PDF 頁碼：{", ".join(str(p["page"]) for p in r["parts"])}</p>'
                     + (f'<p>{links}</p>' if links else '')
                     + f'<p class="warn">{html.escape(warnings)}</p>'
                     + (f'<a href="{image}"><img class="{"passage" if r["kind"] == "material" else "question"}" loading="lazy" src="{image}"></a>' if image else '<p>尚無圖片，請修正裁切設定。</p>')
                     + '</article>')
    (folder / "index.html").write_text('''<!doctype html><meta charset="utf-8"><title>截圖檢視</title>
<style>body{font:16px system-ui;margin:30px;background:#f3f5f7;color:#172332}header{position:sticky;top:0;background:white;padding:15px;border-bottom:2px solid #234}input{padding:10px;width:300px}article{background:white;margin:20px 0;padding:20px;border-radius:8px;scroll-margin-top:260px}img{max-width:100%;max-height:850px;object-fit:contain;object-position:left}img.passage{max-height:none}nav a{display:inline-block;margin:4px 16px 4px 0}small{font-size:15px}.warn{color:#925407}button{padding:10px}@media print{header,input,button,.warn{display:none}article{break-inside:avoid}}</style>
<header><h1>''' + html.escape(title) + '''</h1><input placeholder="搜尋 q12 或試題代碼" oninput="document.querySelectorAll('article').forEach(a=>a.hidden=!a.dataset.key.toLowerCase().includes(this.value.toLowerCase()))"><button onclick="window.print()">列印／另存 PDF</button><p>圖片依實際題號命名。點圖片可開原尺寸；黃色文字代表待確認事項。</p></header>'''
        + '<nav>文章快速連結：'+' '.join(navigation)+'</nav>' + "\n".join(cards), encoding="utf-8")


def export(path, source, output, dpi, overrides, reference):
    folder = output / source / identity(path)
    folder.mkdir(parents=True, exist_ok=True)
    manifest_path = folder / "manifest.json"
    if manifest_path.exists():
        previous = json.loads(manifest_path.read_text(encoding="utf-8"))
        if previous["source_pdf"] != str(path.resolve()):
            raise ValueError(f"資料夾識別名稱重複: {path.name}；請指定另一個 --output")
    with fitz.open(path) as doc:
        if doc.needs_pass:
            raise ValueError("PDF 受密碼保護，請提供可讀版本")
        records = report_records(path, doc) if source == "reports" else booklet_records(path, doc)
        if not records:
            raise ValueError("找不到可辨認的題號／試題代碼，可能需要 OCR")
        mapping = reference.get(identity(path), {})
        known_numbers = {r["number"] for r in records if r["number"] is not None}
        missing = sorted(set(mapping) - known_numbers) if source == "booklets" else []
        for n in missing:
            records.append({"id": f"q{n}", "number": n, "kind": "question", "parts": [],
                            "warnings": ["報告有此題，但題本未偵測到；請指定裁切範圍"]})
        settings = overrides.get(path.name, {})
        existing_ids = {r["id"] for r in records}
        for item_id, setting in settings.items():
            if item_id not in existing_ids:
                match = re.fullmatch(r"q(\d+)", item_id)
                records.append({"id": item_id, "number": int(match[1]) if match else None,
                                "kind": "question" if match else "material", "parts": [], "warnings": []})
        for record in records:
            item = settings.get(record["id"])
            if item:
                record["parts"] = item["parts"]
                record["warnings"] = [] if item.get("reviewed") else ["使用自訂座標，尚未標記完成檢查"]
            ref = mapping.get(record["number"], {})
            if source == "booklets" and ref:
                record["code"] = ref["code"]
                record["report_page"] = ref["page"]
            record["image"] = None
            if record["parts"]:
                try:
                    if source == 'booklets':
                        exclude_exam_transition(doc, record)
                        if '國語' in path.name:
                            exclude_section_heading(doc, record)
                        exclude_question_number(doc, record)
                    im = render(doc, record["parts"], dpi)
                    name = record["id"] + ".png"
                    im.save(folder / name, dpi=(dpi, dpi))
                    record.update(image=name, width_px=im.width, height_px=im.height,
                                  width_cm=round(im.width / dpi * 2.54, 2),
                                  height_cm=round(im.height / dpi * 2.54, 2))
                except (ValueError, IndexError) as exc:
                    record["warnings"].append(str(exc))
            if record["image"] is None and (folder / (record["id"] + ".png")).exists():
                # 防止前次成功的圖片掩蓋本次裁切失敗。
                (folder / (record["id"] + ".png")).unlink()
        manifest = {"source_pdf": str(path.resolve()), "source_type": source, "dpi": dpi,
                    "identity": identity(path), "missing_in_booklet": missing, "records": records}
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        # 只清除上次清單列出的舊輸出，例如結尾「31. 作答完畢」被修正後的假題號。
        if 'previous' in locals():
            current_images = {r["image"] for r in records if r["image"]}
            for old in previous["records"]:
                name = old.get("image")
                if name and Path(name).name == name and name not in current_images:
                    (folder / name).unlink(missing_ok=True)
        with (folder / "對照清單.csv").open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.writer(stream)
            writer.writerow(["題號", "類型", "試題代碼", "來源PDF頁碼", "報告PDF頁碼", "圖片", "共用材料", "寬cm", "高cm", "待確認"])
            for r in records:
                writer.writerow([r["number"], r["kind"], r.get("code", ""),
                                 ";".join(str(p["page"]) for p in r["parts"]), r.get("report_page", ""),
                                 r["image"], ";".join(r.get("materials", [])), r.get("width_cm", ""), r.get("height_cm", ""), "；".join(r["warnings"])])
        write_preview(folder, records, identity(path))
        questions = [r for r in records if r["kind"] == "question"]
        return {"name": identity(path), "source": source, "path": str((folder / 'index.html').relative_to(output)).replace('\\', '/'),
                "questions": len(questions), "images": sum(bool(r["image"]) for r in questions),
                "review": sum(bool(r["warnings"]) for r in records), "missing": missing}


def load_reference(report_paths):
    result = {}
    for path in report_paths:
        mapping = {}
        with fitz.open(path) as doc:
            for i, page in enumerate(doc):
                for m in CODE.finditer(page.get_text()):
                    number = int(m["number"])
                    if number in mapping:
                        raise ValueError(f"報告試題代碼重複: {path.name} q{number}")
                    mapping[number] = {"code": m[0], "page": i + 1}
        result[identity(path)] = mapping
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", choices=["booklets", "reports", "both"], default="booklets")
    parser.add_argument("--pdf", type=Path, help="單一 PDF（需指定 --source booklets 或 reports）")
    parser.add_argument("--input-dir", type=Path, help="遞迴處理指定資料夾中的 PDF")
    parser.add_argument("--match", default="*", help="只處理檔名符合此萬用字元模式的 PDF")
    parser.add_argument("--output", type=Path, default=ROOT / "output" / "screenshots")
    parser.add_argument("--dpi", type=int, default=220)
    parser.add_argument("--overrides", type=Path, default=ROOT / "crop_overrides.json")
    args = parser.parse_args()
    if args.dpi < 72 or args.dpi > 600:
        parser.error("dpi 應介於 72 到 600")
    if args.pdf and args.source == "both":
        parser.error("--pdf 需搭配 --source booklets 或 reports")
    if args.input_dir and (args.pdf or args.source == "both"):
        parser.error("--input-dir 需搭配單一來源類型，且不能與 --pdf 同時指定")
    if args.input_dir and not args.input_dir.is_dir():
        parser.error("--input-dir 資料夾不存在")
    reports = sorted((ROOT / "成果報告").glob("*.pdf"))
    reference = load_reference(reports)
    overrides = json.loads(args.overrides.read_text(encoding="utf-8")) if args.overrides.exists() else {}
    summary, failures = [], []
    for source in (["reports", "booklets"] if args.source == "both" else [args.source]):
        paths = [args.pdf] if args.pdf else (reports if source == "reports" else
                sorted((ROOT / "題本").glob("*縣市*.pdf")) + sorted(ROOT.glob("*題本.pdf")))
        if args.input_dir:
            paths = sorted(args.input_dir.rglob("*.pdf"))
        paths = [path for path in paths if fnmatch.fnmatch(path.name, args.match)]
        for path in paths:
            try:
                item = export(path, source, args.output, args.dpi, overrides, reference)
                summary.append(item)
                print(f"{source} {item['name']}: {item['images']}/{item['questions']} 題, 待確認 {item['review']} 項", flush=True)
            except Exception as exc:
                failures.append({"file": str(path), "error": str(exc)})
                print(f"失敗 {path.name}: {exc}", flush=True)
    args.output.mkdir(parents=True, exist_ok=True)
    summary_file = args.output / "summary.json"
    if summary_file.exists():
        previous_summary = json.loads(summary_file.read_text(encoding="utf-8"))
        updated = {(r["source"], r["name"]) for r in summary}
        summary = [r for r in previous_summary["results"] if (r["source"], r["name"]) not in updated] + summary
        summary.sort(key=lambda r: (r["source"], r["name"]))
    (args.output / "summary.json").write_text(json.dumps({"results": summary, "failures": failures}, ensure_ascii=False, indent=2), encoding="utf-8")
    with (args.output / "全部題目對照.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.writer(stream)
        writer.writerow(["年度科目年級", "題號", "試題代碼", "題本PDF", "題本PDF頁碼", "報告PDF頁碼", "圖片路徑", "共用材料", "待確認"])
        for item in summary:
            if item["source"] != "booklets":
                continue
            folder = args.output / Path(item["path"]).parent
            manifest = json.loads((folder / "manifest.json").read_text(encoding="utf-8"))
            for record in manifest["records"]:
                if record["kind"] == "question":
                    writer.writerow([item["name"], record["number"], record.get("code", ""), manifest["source_pdf"],
                        ";".join(str(p["page"]) for p in record["parts"]), record.get("report_page", ""),
                        str((folder / record["image"]).resolve()) if record["image"] else "",
                        ";".join(record.get("materials", [])), "；".join(record["warnings"])])
    rows = ''.join(f'<tr><td>{r["source"]}</td><td><a href="{html.escape(r["path"])}">{html.escape(r["name"])}</a></td><td>{r["images"]}/{r["questions"]}</td><td>{r["review"]}</td></tr>' for r in summary)
    (args.output / "index.html").write_text('<!doctype html><meta charset="utf-8"><title>各科截圖索引</title><style>body{font:17px system-ui;margin:40px}td,th{padding:12px;border-bottom:1px solid #ddd}a{color:#145a92}</style><h1>各科截圖索引</h1><p><a href="全部題目對照.csv">下載全部題目對照清單</a></p><p>reports：既有報告題目欄參考圖；booklets：題本原始題目。題本結果需檢查共用材料與跨頁內容。</p><table><tr><th>來源</th><th>科目／年級</th><th>圖片／題數</th><th>待確認項目</th></tr>' + rows + '</table>' + ''.join('<p>'+html.escape(str(e))+'</p>' for e in failures), encoding="utf-8")
    if failures or any(r["images"] < r["questions"] for r in summary):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
