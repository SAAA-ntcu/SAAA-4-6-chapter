# SAAA-4-6-chapter: Troubleshooting & Error Recovery Guide

This guide helps diagnose and resolve common formatting, XML, and data issues encountered during assessment report generation.

---

## 1. Word XML Page Break Issues

### Problem: Extra Blank Pages Between Question Tables
- **Cause**: An empty paragraph `<w:p>` remains between `<w:tbl>` elements, or an extra `<w:br w:type="page"/>` is placed right before a table that already has `w:pageBreakBefore`.
- **Solution**:
  The script automatically sweeps the document body and removes empty paragraphs between tables:
  ```python
  for child in list(body):
      if child.tag.endswith('}p'):
          text = ''.join(t.text or '' for t in child.findall('.//' + qn('w:t')))
          if not text.strip():
              body.remove(child)
  ```

### Problem: Table Split Across Pages (Page Break in Table Middle)
- **Cause**: Word defaults allow rows to split across page breaks.
- **Solution**:
  Enforce `w:cantSplit` on all table rows:
  ```python
  for row in table.rows:
      trPr = row._tr.get_or_add_trPr()
      if trPr.find(qn('w:cantSplit')) is None:
          trPr.append(OxmlElement('w:cantSplit'))
  ```

### Problem: Chinese Heading Separated from First Question Table
- **Cause**: Applying `pageBreakBefore` unconditionally to all tables causes the table to jump to the next page, leaving the subheading `(一)形音知識` alone at the bottom of the previous page.
- **Solution**:
  Use `prev_was_heading` state tracking in `add_pagebreaks_in_tables_chinese`:
  - When a heading like `(一)形音知識` is encountered, set `prev_was_heading = True`.
  - When processing the immediately following table, do **NOT** insert `w:pageBreakBefore`.
  - Reset `prev_was_heading = False` so subsequent question tables in the same domain get their page break.

---

## 2. Table 6-1 Formatting & Merging Issues

### Problem: Diagonal Border Missing on Empty Intersections
- **Cause**: Word OpenXML requires the `tl2br` (top-left to bottom-right) border element inside `tcBorders`.
- **Solution**:
  Use the built-in helper:
  ```python
  def set_cell_diagonal_border(cell):
      tcPr = cell._tc.get_or_add_tcPr()
      tcBorders = tcPr.get_or_add_tcBorders()
      tl2br = OxmlElement('w:tl2br')
      tl2br.set(qn('w:val'), 'single')
      tl2br.set(qn('w:sz'), '4')
      tl2br.set(qn('w:color'), 'BFBFBF')
      tcBorders.append(tl2br)
  ```

### Problem: Row Heights in Table 6-1 Vary or Look Compressed
- **Cause**: Word table rows without explicit row height rules resize based on font and padding.
- **Solution**:
  Set row heights explicitly:
  ```python
  for row in table.rows:
      row.height = Pt(24)
      row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
  ```

---

## 3. Font and Encoding Issues

### Problem: Text Rendered in SimSun or Times New Roman instead of 微軟正黑體
- **Cause**: In Word XML, Chinese characters fall under `w:eastAsia` rather than `w:ascii` or `w:hAnsi`. If `w:eastAsia` is not explicitly set, Word falls back to default system fonts.
- **Solution**:
  Always set all three font targets in run properties:
  ```python
  rFonts = OxmlElement('w:rFonts')
  rFonts.set(qn('w:eastAsia'), '微軟正黑體')
  rFonts.set(qn('w:ascii'), '微軟正黑體')
  rFonts.set(qn('w:hAnsi'), '微軟正黑體')
  rPr.append(rFonts)
  ```

### Problem: Python Console UnicodeEncodeError on Windows
- **Cause**: Windows command prompt defaults to CP950 or OEM code pages, which cannot output non-ASCII characters like `✓`.
- **Solution**:
  Include `sys.stdout.reconfigure(encoding='utf-8', errors='replace')` at the top of scripts and use ASCII status markers `[PASS]`, `[WARN]`, `[FAIL]`.

---

## 4. Input File Anomalies

### Problem: Sheet Names Differ from Standard (e.g. `115_M_試題分析總表 (1).xlsx`)
- **Resolution**: `generate_reports.py` and `qa_verify.py` use fuzzy sheet name matching:
  - If sheet name contains the grade number (e.g. `"7"` in `"7年級"` or `"7"`), it is matched automatically.
  - If the workbook only contains a single sheet (e.g., in `C7N_試題分析總表.xlsx`), that sheet is used regardless of name.
