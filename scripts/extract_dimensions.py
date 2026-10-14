"""
SAAA-4-6-chapter: Dimension and Answer Extractor / Normalizer
Utility to extract, normalize, and validate assessment answer sheets and dimension indicators.
Supports:
1. Converting raw horizontal answer sheets & PPTX slide tables into standard 2-sheet Excel format (答案 + 向度).
2. Inspecting and diagnosing sheet structures in answer / analysis Excel files.
"""

import os
import sys
import re
import argparse
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

FONT_CHINESE = "微軟正黑體"

def inspect_excel(path):
    if not os.path.exists(path):
        print(f"[錯誤] 檔案不存在: {path}")
        return
    xls = pd.ExcelFile(path)
    print(f"\n=======================================================")
    print(f"檢視 Excel 檔案結構: {os.path.basename(path)}")
    print(f"工作表清單 ({len(xls.sheet_names)} 個): {xls.sheet_names}")
    print(f"=======================================================")
    for s in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=s, header=None)
        print(f"\n工作表: [{s}] - 行數: {len(df)}, 欄數: {len(df.columns)}")
        # Print first 3 rows
        preview_rows = min(len(df), 4)
        for r in range(preview_rows):
            row_vals = [str(df.iloc[r, c]) if pd.notna(df.iloc[r, c]) else "" for c in range(min(len(df.columns), 8))]
            print(f"  Row {r}: {row_vals}")

def apply_styling(ws, header_fill_color=None):
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    header_fill = PatternFill(start_color=header_fill_color, end_color=header_fill_color, fill_type='solid') if header_fill_color else None

    for col in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row, col)
            val = cell.value
            if row == 1:
                cell.font = Font(name=FONT_CHINESE, size=11, bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if header_fill:
                    cell.fill = header_fill
            else:
                cell.font = Font(name=FONT_CHINESE, size=11, bold=False)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
            
            s = str(val) if val is not None else ""
            cjk_count = sum(1 for ch in s if ord(ch) > 127)
            length = len(s) + cjk_count
            if length > max_len:
                max_len = length
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

def normalize_single_cohort(ans_dict, dim_dict, cohort_code, out_path, subject="C"):
    wb = openpyxl.Workbook()
    
    # Sheet 1: 答案
    ws_ans = wb.active
    ws_ans.title = "答案"
    ws_ans.cell(1, 1, "題號")
    ws_ans.cell(1, 2, cohort_code)
    
    q_nums = sorted(list(set(ans_dict.keys()) | set(dim_dict.keys())))
    for idx, q in enumerate(q_nums, start=1):
        ws_ans.cell(idx + 1, 1, q)
        ws_ans.cell(idx + 1, 2, ans_dict.get(q, ""))
    apply_styling(ws_ans, "D9E1F2")

    # Sheet 2: 向度 or 評量指標代
    sheet_title = "評量指標代" if subject.upper().startswith("C") else "向度"
    ws_dim = wb.create_sheet(title=sheet_title)
    ws_dim.cell(1, 1, "題號")
    
    if subject.upper().startswith("C"):
        ws_dim.cell(1, 2, f"{cohort_code}評量指標")
        for idx, q in enumerate(q_nums, start=1):
            ws_dim.cell(idx + 1, 1, q)
            ws_dim.cell(idx + 1, 2, dim_dict.get(q, ""))
    else:
        ws_dim.cell(1, 2, f"{cohort_code}1_向度")
        ws_dim.cell(1, 3, f"{cohort_code}2_向度")
        for idx, q in enumerate(q_nums, start=1):
            ws_dim.cell(idx + 1, 1, q)
            d_val = dim_dict.get(q, ("", ""))
            if isinstance(d_val, tuple):
                d1, d2 = d_val
            else:
                d1, d2 = d_val, ""
            ws_dim.cell(idx + 1, 2, d1)
            ws_dim.cell(idx + 1, 3, d2)

    apply_styling(ws_dim, "E2EFDA")
    wb.save(out_path)
    print(f"成功輸出標準格式答案向度檔: {out_path} ({len(q_nums)} 題)")

def main():
    parser = argparse.ArgumentParser(description="SAAA Dimension & Answer Extractor")
    parser.add_argument("--inspect", help="Inspect sheet and column structure of an Excel file")
    args = parser.parse_args()

    if args.inspect:
        inspect_excel(args.inspect)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()
