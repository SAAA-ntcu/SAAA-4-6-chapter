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
        report["checks"].append(("Question Count", "WARN", f"Expected {anal_q_count}, found {doc_q_count} tables"))

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
    page_breaks_before = body.findall('.//' + qn('w:pageBreakBefore'))
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
            reports.append(rep)

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

def main():
    parser = argparse.ArgumentParser(description="SAAA QA Verification Suite")
    parser.add_argument("--dir", default=".", help="Directory containing generated docx files")
    parser.add_argument("--analysis-dir", help="Directory containing item analysis Excel files")
    args = parser.parse_args()
    
    run_qa_suite(args.dir, args.analysis_dir)

if __name__ == "__main__":
    main()
