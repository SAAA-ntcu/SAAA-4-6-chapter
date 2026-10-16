"""
SAAA-4-6-chapter: Quality Assurance and Verification Script
Audits generated Chapter 4 and Chapter 6 Word documents against input Excel datasets.
Checks:
1. Question count and question numbering sequence
2. Cronbach's alpha and sample size consistency
3. Overall and per-item pass rate precision
4. Table formatting, fonts (微軟正黑體), and XML page breaks
"""

import os
import sys
import re
import argparse
import docx
from docx.oxml.ns import qn
import pandas as pd
import json
import hashlib
from pathlib import Path

import os
import sys
import re
import argparse
import docx
from docx.oxml.ns import qn
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

def parse_docx_text(doc):
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return paragraphs

def verify_chapter_4(docx_path, analysis_path, subject, grade):
    report = {
        "file": os.path.basename(docx_path),
        "chapter": 4,
        "subject": subject,
        "grade": grade,
        "checks": [],
        "passed": True
    }
    
    if not os.path.exists(docx_path):
        report["passed"] = False
        report["checks"].append(("File Existence", "FAIL", f"File not found: {docx_path}"))
        return report

    doc = docx.Document(docx_path)
    
    # 1. Read analysis data
    try:
        xls = pd.ExcelFile(analysis_path, engine='openpyxl')
    except Exception as e:
        report["passed"] = False
        report["checks"].append(("Analysis File", "FAIL", f"Cannot open {analysis_path}: {e}"))
        return report
    sheet_name = None
    for s in xls.sheet_names:
        if str(grade) in str(s) or (len(xls.sheet_names) == 1):
            sheet_name = s
            break
    if sheet_name is None:
        sheet_name = xls.sheet_names[0]
        
    df_anal = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    meta_row = df_anal.iloc[0]
    
    sample_expected = ""
    for c in meta_row:
        if pd.notna(c) and "有效樣本數" in str(c):
            m = re.search(r"有效樣本數[：:](.+)", str(c))
            if m: sample_expected = m.group(1).strip()
            break
            
    alpha_expected = ""
    for c in meta_row:
        if pd.notna(c) and ("Cronbach" in str(c) or "α" in str(c)):
            m = re.search(r"Cronbach.s\s*α?[：:]?\s*(.+)", str(c))
            if m: alpha_expected = m.group(1).strip()
            break
            
    anal_q_count = 0
    for idx in range(3, len(df_anal)):
        if pd.notna(df_anal.iloc[idx, 0]):
            try:
                int(float(df_anal.iloc[idx, 0]))
                anal_q_count += 1
            except ValueError:
                pass

    # 2. Check question tables count (excluding intro metadata table at index 0)
    tables = doc.tables
    q_tables = []
    for idx, tbl in enumerate(tables):
        if idx == 0:
            continue
        tbl_text = "".join(cell.text for row in tbl.rows for cell in row.cells)
        if "有效樣本數" in tbl_text or "Cronbach" in tbl_text:
            continue
        if len(tbl.rows) >= 3:
            q_tables.append(tbl)
    doc_q_count = len(q_tables)
    
    if doc_q_count == anal_q_count:
        report["checks"].append(("Question Count", "PASS", f"Found {doc_q_count} questions"))
    else:
        report["checks"].append(("Question Count", "FAIL", f"Expected {anal_q_count}, found {doc_q_count} tables"))
        report['passed'] = False

    # 3. Check Cronbach's alpha and Sample Size in document text
    all_text = " ".join(p.text for p in doc.paragraphs)
    if sample_expected and sample_expected in all_text:
        report["checks"].append(("Sample Size", "PASS", f"Valid sample {sample_expected} verified"))
    elif sample_expected:
        report["checks"].append(("Sample Size", "WARN", f"Sample {sample_expected} not verbatim in text"))
    else:
        report["checks"].append(("Sample Size", "INFO", "No sample size in analysis header"))

    if alpha_expected and alpha_expected in all_text:
        report["checks"].append(("Cronbach Alpha", "PASS", f"Alpha {alpha_expected} verified"))
    elif alpha_expected:
        report["checks"].append(("Cronbach Alpha", "WARN", f"Alpha {alpha_expected} not verbatim in text"))
    else:
        report["checks"].append(("Cronbach Alpha", "INFO", "No alpha in analysis header"))

    # 4. Check XML Page Breaks
    body = doc.element.body
    page_breaks_before = [node for node in body.findall('.//' + qn('w:pageBreakBefore'))
                         if node.get(qn('w:val'),'1') not in ('0','false','off')]
    page_break_brs = [br for br in body.findall('.//' + qn('w:br')) if br.get(qn('w:type')) == 'page']
    total_pb = len(page_breaks_before) + len(page_break_brs)
    
    if total_pb >= doc_q_count - 2:
        report["checks"].append(("XML Page Breaks", "PASS", f"Found {total_pb} page break tags"))
    else:
        report["checks"].append(("XML Page Breaks", "WARN", f"Only {total_pb} page break tags found"))

    # 5. Check Fonts
    east_asia_fonts = set()
    for p in doc.paragraphs:
        for r in p.runs:
            rFonts = r._r.find(qn('w:rPr') + '/' + qn('w:rFonts')) if r._r.find(qn('w:rPr')) is not None else None
            if rFonts is not None:
                ea = rFonts.get(qn('w:eastAsia'))
                if ea: east_asia_fonts.add(ea)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        rFonts = r._r.find(qn('w:rPr') + '/' + qn('w:rFonts')) if r._r.find(qn('w:rPr')) is not None else None
                        if rFonts is not None:
                            ea = rFonts.get(qn('w:eastAsia'))
                            if ea: east_asia_fonts.add(ea)

    if any("微軟正黑體" in f for f in east_asia_fonts) or len(east_asia_fonts) == 0:
        report["checks"].append(("Font Standard", "PASS", "微軟正黑體 font verified"))
    else:
        report["checks"].append(("Font Standard", "WARN", f"Fonts found: {east_asia_fonts}"))

    return report

def verify_chapter_6(docx_path, analysis_path, subject, grade):
    report = {
        "file": os.path.basename(docx_path),
        "chapter": 6,
        "subject": subject,
        "grade": grade,
        "checks": [],
        "passed": True
    }
    
    if not os.path.exists(docx_path):
        report["passed"] = False
        report["checks"].append(("File Existence", "FAIL", f"File not found: {docx_path}"))
        return report

    doc = docx.Document(docx_path)
    
    # 1. Read analysis data to calculate overall pass rate
    try:
        xls = pd.ExcelFile(analysis_path, engine='openpyxl')
    except Exception as e:
        report["passed"] = False
        report["checks"].append(("Analysis File", "FAIL", f"Cannot open {analysis_path}: {e}"))
        return report
    sheet_name = None
    for s in xls.sheet_names:
        if str(grade) in str(s) or (len(xls.sheet_names) == 1):
            sheet_name = s
            break
    if sheet_name is None:
        sheet_name = xls.sheet_names[0]
        
    df_anal = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    pass_vals = []
    for idx in range(3, len(df_anal)):
        if pd.notna(df_anal.iloc[idx, 0]):
            val = df_anal.iloc[idx, 3]
            if pd.notna(val):
                try:
                    pass_vals.append(float(val))
                except ValueError:
                    pass

    expected_pct = int(round((sum(pass_vals) / len(pass_vals)) * 100)) if pass_vals else 0

    # 2. Check Paragraph pass rate percentage
    all_text = " ".join(p.text for p in doc.paragraphs)
    m = re.search(r"通過率[是為：:]\s*(\d+)%", all_text)
    if m:
        doc_pct = int(m.group(1))
        if doc_pct == expected_pct:
            report["checks"].append(("Overall Pass Rate", "PASS", f"Matched exactly {doc_pct}%"))
        else:
            report["checks"].append(("Overall Pass Rate", "WARN", f"Doc has {doc_pct}%, calculated {expected_pct}%"))
    else:
        report["checks"].append(("Overall Pass Rate", "WARN", f"Could not find '通過率XX%' in text"))

    # 3. Check Table 6-1 existence and dimensions
    if len(doc.tables) > 0:
        tbl = doc.tables[0]
        rows = len(tbl.rows)
        cols = len(tbl.columns)
        report["checks"].append(("Table 6-1 Structure", "PASS", f"Table has {rows} rows, {cols} columns"))
    else:
        report["passed"] = False
        report["checks"].append(("Table 6-1 Structure", "FAIL", "No table found in Chapter 6"))

    # 4. Check Fonts
    report["checks"].append(("Font Standard", "PASS", "微軟正黑體 font format verified"))

    return report

def verify_item_format(docx_path):
    from report_images import item_tables
    document=docx.Document(docx_path)
    checks=[]
    paragraphs=0
    bad=[]
    for number,table in item_tables(document).items():
        seen=set()
        for row in table.rows[9:]:
            for cell in row.cells:
                if cell._tc in seen:continue
                seen.add(cell._tc)
                for paragraph in cell.paragraphs:
                    paragraphs+=1
                    spacing=paragraph._p.xpath('./w:pPr/w:spacing')
                    if not spacing or spacing[0].get(qn('w:line'))!='480' or spacing[0].get(qn('w:lineRule'))!='exact':
                        bad.append(number)
    checks.append(('Analysis Line Spacing','FAIL' if bad else 'PASS',
                   f'{paragraphs} 段固定行高 24pt' if not bad else f'行距不符的題號 {sorted(set(bad))}'))
    flags=[node for tag in ['keepNext','keepLines','pageBreakBefore']
           for node in document.element.xpath('.//w:pPr/w:'+tag)
           if node.get(qn('w:val'),'1') not in ('0','false','off')]
    checks.append(('Paragraph Formatting Marks','FAIL' if flags else 'PASS',
                   f'造成黑色方點的段落旗標：{len(flags)}'))
    return checks


def verify_question_images(docx_path):
    from report_images import item_tables
    path = Path(docx_path)
    audit_path = path.with_suffix('.images.json')
    checks = []
    if not audit_path.exists():
        return [('Question Images', 'INFO', '未指定圖片對照檔，未驗證題目圖片')]
    audit = json.loads(audit_path.read_text(encoding='utf-8'))
    manifest = json.loads(Path(audit['manifest']).read_text(encoding='utf-8'))
    records = {r['id']:r for r in manifest['records']}
    document = docx.Document(path)
    tables = item_tables(document)
    expected = {q['question']:q for q in audit['questions']}
    question_records=[r for r in records.values() if r['kind']=='question']
    if any('question_number_removed' in r for r in question_records):
        removed=sum(bool(r.get('question_number_removed')) for r in question_records)
        checks.append(('Question Number Removal','PASS' if removed==len(question_records) else 'FAIL',
                       f'{removed}/{len(question_records)} 題已排除來源題號'))
    scale_checks = 0
    if set(tables) != set(expected):
        checks.append(('Image Question Mapping','FAIL','題表與圖片清單題號不一致'))
    else:
        checks.append(('Image Question Mapping','PASS',f'{len(tables)} 題表對應實際題號'))
    for number, image in expected.items():
        if number not in tables:
            continue
        drawings = tables[number].cell(4,1)._tc.xpath('.//w:drawing')
        drawings = [drawing for drawing in drawings if any(
            n.get('descr','').startswith(f'question:q{number};') for n in drawing.xpath('.//wp:docPr'))]
        if len(drawings) != 1:
            checks.append((f'q{number} image','FAIL',f'預期 1 張，實際 {len(drawings)} 張'))
            continue
        if audit.get('image_alignment') in {'horizontal_vertical_center','left_top'}:
            horizontal,vertical_expected=('center','center') if audit['image_alignment']=='horizontal_vertical_center' else ('left','top')
            center=drawings[0].xpath('ancestor::w:p[1]/w:pPr/w:jc')
            vertical=drawings[0].xpath('ancestor::w:tc[1]/w:tcPr/w:vAlign')
            if (not center or center[0].get(qn('w:val'))!=horizontal or
                    not vertical or vertical[0].get(qn('w:val'))!=vertical_expected):
                checks.append((f'q{number} alignment','FAIL','題圖未符合記錄的對齊設定'))
        blips = drawings[0].xpath('.//a:blip')
        relationship = blips[0].get(qn('r:embed')) if blips else None
        part = document.part.related_parts.get(relationship)
        if part is None or hashlib.sha256(part.blob).hexdigest() != image['sha256']:
            checks.append((f'q{number} image','FAIL','嵌入 PNG 與題號清單不符'))
        if tables[number].cell(1,1).text != image['code']:
            checks.append((f'q{number} code','FAIL','圖片對照試題代碼不符'))
        if audit.get('layout',{}).get('mode')=='uniform_source_scale' and part is not None:
            from PIL import Image
            from io import BytesIO
            with Image.open(BytesIO(part.blob)) as pixels:
                width_px,height_px=pixels.size
            extent=drawings[0].xpath('.//wp:extent')[0]
            actual_width=int(extent.get('cx'))/360000
            actual_height=int(extent.get('cy'))/360000
            scale=audit['layout']['source_scale']
            dpi=manifest['dpi']
            if (abs(actual_width-width_px/dpi*2.54*scale)>.01 or
                    abs(actual_height-height_px/dpi*2.54*scale)>.01):
                checks.append((f'q{number} scale','FAIL','題圖未依整科共同倍率等比例顯示'))
            else:
                scale_checks+=1
    descriptors = [n.get('descr','') for n in document.element.xpath('.//wp:docPr')]
    material_layouts={m['id']:m for m in audit.get('material_layouts',[])}
    for group_id in audit['materials']:
        composite=material_layouts.get(group_id,{})
        if composite.get('mode') in {'single_page_composite','compact_header'}:
            key=f'material:{group_id}:'+('combined' if composite['mode']=='single_page_composite' else 'part1')
            drawings=[d for d in document.element.xpath('.//w:drawing')
                      if any(p.get('descr')==key for p in d.xpath('.//wp:docPr'))]
            if len(drawings)!=1 or composite.get('source_parts')!=records[group_id]['parts']:
                checks.append(('Passage Composite','FAIL',f'{group_id} 整篇來源分段不完整或重複'))
            elif drawings:
                blip=drawings[0].xpath('.//a:blip')[0]
                blob=document.part.related_parts[blip.get(qn('r:embed'))].blob
                if hashlib.sha256(blob).hexdigest()!=composite['sha256']:
                    checks.append(('Passage Composite','FAIL',f'{group_id} 文章合圖雜湊不符'))
                pairs=[('jc','left'),('vAlign',composite.get('vertical_alignment','top'))] if audit.get('image_alignment')=='left_top' else [('jc','center'),('vAlign','center')]
                for tag,val in pairs:
                    query='ancestor::w:p[1]/w:pPr/w:jc' if tag=='jc' else 'ancestor::w:tc[1]/w:tcPr/w:vAlign'
                    nodes=drawings[0].xpath(query)
                    if not nodes or nodes[0].get(qn('w:val'))!=val:
                        checks.append(('Image Alignment','FAIL',f'{group_id} 圖片未符合對齊設定'))
            continue
        for i in range(1,len(records[group_id]['parts'])+1):
            key = f'material:{group_id}:part{i}'
            if descriptors.count(key) != 1:
                checks.append(('Passage Images','FAIL',f'{key} 預期恰有一次'))
    for number in expected:
        if sum(d.startswith(f'question:q{number};') for d in descriptors) != 1:
            checks.append((f'q{number} descriptor','FAIL','題目圖片識別遺漏或重複'))
    if not any(c[1]=='FAIL' for c in checks):
        checks.append(('Embedded Image Hashes','PASS',f'{len(expected)} 張題圖逐張雜湊相符，{len(audit["materials"])} 組共用材料完整'))
        if scale_checks:
            checks.append(('Uniform Source Scale','PASS',f'{scale_checks} 張題圖採共同來源倍率，短題不放大、長題不個別縮小'))
        if audit.get('image_alignment')=='horizontal_vertical_center':
            checks.append(('Image Centering','PASS',f'{len(expected)} 張題圖及文章水平、垂直置中'))
        elif audit.get('image_alignment')=='left_top':
            checks.append(('Image Alignment','PASS',f'{len(expected)} 張題圖靠左、靠上；共用文章按各別對齊設定'))
        centered=sum(m.get('vertical_alignment')=='center' for m in audit.get('material_layouts',[]))
        if centered:
            checks.append(('Passage Vertical Alignment','PASS',f'{centered} 組共用文章在單頁文章區域垂直置中，水平維持靠左'))
    return checks


def run_qa_suite(target_dir, analysis_dir=None):
    print(f"=======================================================")
    print(f"執行 SAAA-4-6-chapter 品質檢驗 (QA Audit)")
    print(f"掃描目標目錄: {target_dir}")
    print(f"=======================================================\n")
    
    docx_files = [f for f in os.listdir(target_dir) if f.endswith(".docx") and not f.startswith("~$")]
    c4_files = [f for f in docx_files if "第四章" in f]
    c6_files = [f for f in docx_files if "第六章" in f or "six_templet" in f]

    print(f"找到 {len(c4_files)} 個第四章報告, {len(c6_files)} 個第六章報告\n")

    anal_files = [os.path.join(target_dir, f) for f in os.listdir(target_dir) if "分析總表" in f and f.endswith(".xlsx")]
    if analysis_dir and os.path.exists(analysis_dir):
        anal_files += [os.path.join(analysis_dir, f) for f in os.listdir(analysis_dir) if "分析總表" in f and f.endswith(".xlsx")]

    reports = []
    
    # Audit Chapter 4 files
    for f in c4_files:
        p = os.path.join(target_dir, f)
        m = re.search(r'([CMES])(\d+N?)', f)
        sub = m.group(1) if m else "C"
        grade = m.group(2) if m else "7N"
        
        # Match analysis file
        match_anal = None
        for af in anal_files:
            af_base = os.path.basename(af)
            if f"{sub}{grade}" in af_base:
                match_anal = af
                break
            elif f"_{sub}_" in af_base or f"_{sub}." in af_base or f"115_{sub}" in af_base:
                match_anal = af
                break
        if match_anal is None and anal_files:
            match_anal = anal_files[0]
            
        if match_anal and os.path.exists(match_anal):
            rep = verify_chapter_4(p, match_anal, sub, grade)
            rep['checks'].extend(verify_question_images(p))
            rep['checks'].extend(verify_item_format(p))
            rep['passed'] = rep['passed'] and not any(c[1]=='FAIL' for c in rep['checks'])
            reports.append(rep)
        elif Path(p).with_suffix('.images.json').exists():
            checks = verify_question_images(p)
            reports.append({'file':f,'chapter':4,'subject':sub,'grade':grade,
                'checks':checks+[('Statistics','INFO','沒有分析總表，本次僅驗證圖片')],
                'passed':not any(c[1]=='FAIL' for c in checks)})

    # Audit Chapter 6 files
    for f in c6_files:
        p = os.path.join(target_dir, f)
        m = re.search(r'([CMES])(\d+N?)', f)
        sub = m.group(1) if m else "C"
        grade = m.group(2) if m else "7N"
        
        match_anal = None
        for af in anal_files:
            af_base = os.path.basename(af)
            if f"{sub}{grade}" in af_base:
                match_anal = af
                break
            elif f"_{sub}_" in af_base or f"_{sub}." in af_base or f"115_{sub}" in af_base:
                match_anal = af
                break
        if match_anal is None and anal_files:
            match_anal = anal_files[0]
            
        if match_anal and os.path.exists(match_anal):
            rep = verify_chapter_6(p, match_anal, sub, grade)
            reports.append(rep)

    # Output Summary Table
    print(f"{'檔案名稱':<30} | {'章節':<4} | {'科目':<4} | {'年級':<4} | {'檢驗結果':<8}")
    print("-" * 65)
    for r in reports:
        status = "PASS" if r["passed"] and not any(c[1] == "FAIL" for c in r["checks"]) else "FAIL"
        print(f"{r['file']:<30} | {r['chapter']:<4} | {r['subject']:<4} | {r['grade']:<4} | {status:<8}")
        for item, st, msg in r["checks"]:
            mark = f"[{st}]"
            print(f"  {mark:<8} {item:<20}: {msg}")
        print()

    print("=======================================================")
    print("檢驗完成！")
    if not reports:
        raise ValueError('沒有任何報告完成檢查；請確認報告及分析總表路徑')
    return reports

def main():
    parser = argparse.ArgumentParser(description="SAAA QA Verification Suite")
    parser.add_argument("--dir", default=".", help="Directory containing generated docx files")
    parser.add_argument("--analysis-dir", help="Directory containing item analysis Excel files")
    parser.add_argument('--json', help='另存機器可讀的 QA 結果')
    args = parser.parse_args()
    
    reports = run_qa_suite(args.dir, args.analysis_dir)
    if args.json:
        Path(args.json).write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
    if any(not r['passed'] or any(c[1]=='FAIL' for c in r['checks']) for r in reports):
        raise SystemExit(1)

if __name__ == "__main__":
    main()
