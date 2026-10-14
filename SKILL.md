---
name: SAAA-4-6-chapter
description: >-
  Automates the generation of Chapter 4 (肆、試題分析與教學建議) and Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)
  for Taiwanese academic diagnostic assessments (學力檢測 / 新生學力測驗).
  Supports 4 subjects: Chinese (國語文 C), Math (數學 M), English (英語文 E), and Science (自然科學 S)
  across standard elementary/junior high grades (3~8) and special entry cohorts (C7N, M7N, E7N, S7N).
  Uses deterministic scripts to parse Excel datasets, render docxtpl Word templates, inject XML page breaks,
  and generate Table 6-1 matrices with automated QA validation.
---

# SAAA-4-6-chapter Skill

This skill provides an end-to-end automated workflow for producing official diagnostic assessment report chapters:
- **Chapter 4 (肆、試題分析與教學建議)**: Complete item analysis tables with discrimination indices, distractor distributions, high/low scoring group pass rates, sub-domain codes, and layout-preserving XML page breaks.
- **Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)**: Structured multi-dimensional pass rate matrices, high/overall/low group breakdowns, cognitive level intersections, diagonal empty cell borders, and customized column widths.

---

## Supported Subjects and Cohorts

| Subject | Code | Standard Grades | Special Cohorts | Templates Used |
| :--- | :--- | :--- | :--- | :--- |
| **國語文** (Chinese) | `C` | 3, 4, 5, 6, 7, 8 | `C7N` | `pydocxTemplet_C_new.docx`, `C_demo.docx` |
| **數學** (Math) | `M` | 3, 4, 5, 6, 7, 8 | `M7N` | `pydocxTemplet_M.docx`, `S_M_demo.docx` |
| **英語文** (English) | `E` | 4, 5, 6, 7, 8 | `E7N` | `pydocxTemplet_E.docx`, `six_templet_E.docx` |
| **自然科學** (Science) | `S` | 5, 7, 8 | `S7N` | `pydocxTemplet_S.docx`, `S_M_demo.docx` |

---

## Directory Structure

```text
SAAA-4-6-chapter/
├── SKILL.md                          # Main skill instruction file (this document)
├── scripts/
│   ├── generate_reports.py          # Unified multi-subject report generator
│   ├── qa_verify.py                 # Automated verification and auditing tool
│   └── extract_dimensions.py        # Answer sheet normalizer and Excel inspector
├── resources/
│   └── templates/                   # Standard Word templates for all subjects
│       ├── pydocxTemplet_C_new.docx # Chinese Chapter 4 template
│       ├── pydocxTemplet_M.docx     # Math Chapter 4 template
│       ├── pydocxTemplet_E.docx     # English Chapter 4 template
│       ├── pydocxTemplet_S.docx     # Science Chapter 4 template
│       ├── C_demo.docx              # Chinese Chapter 6 Table 6-1 template
│       ├── S_M_demo.docx            # Math & Science Table 6-1 template
│       └── six_templet_E.docx       # English Table 6-1 template
└── references/
    ├── subject_specifications.md    # Detailed data models and field schemas
    └── troubleshooting.md           # Word XML page break rules and error recovery
```

---

## Standard Workflow

### Step 1: Input Discovery and Validation
Identify the input files:
1. **Answers & Dimensions Excel**:
   - Standard format: `115_<Subject>_ans.xlsx` (contains sheets `答案` and `向度` or `評量指標代`).
   - Single-cohort format: `<Cohort>_ans.xlsx` (e.g., `C7N_ans.xlsx`, `M7N_ans.xlsx`, `E7N_ans.xlsx`).
2. **Item Analysis Excel**:
   - Standard format: `115_<Subject>_試題分析總表.xlsx` (sheets named by grade `3`, `4`, `5`, etc.).
   - Single-cohort format: `<Cohort>_試題分析總表.xlsx`.
3. **Raw Unrounded Statistics (High-Precision Mode)**:
   - Format: `<Cohort>_分析結果.xlsx` (contains unrounded 4-decimal floating rates for high/low/overall groups).
   - If present alongside the summary table, the engine automatically prioritizes these raw floating numbers when computing Chapter 6 domain averages to avoid secondary round-off error propagation.

> [!TIP]
> If dealing with raw unstructured Excel tables or PPTX slides (e.g. horizontal answer sheets), use `scripts/extract_dimensions.py` to inspect and convert them into the standard 2-sheet format.

### Step 2: Generate Reports
Run the unified generator script:

```bash
# Generate both Chapter 4 and Chapter 6 for a specific cohort:
python scripts/generate_reports.py \
  --ans-file "<path_to_ans.xlsx>" \
  --analysis-file "<path_to_analysis.xlsx>" \
  --subject <C|M|E|S> \
  --grade <grade_or_cohort> \
  --year 115 \
  --chapter both \
  --out-dir "<output_directory>"

# Examples:
# Chinese 7N:
python scripts/generate_reports.py --ans-file C7N_ans.xlsx --analysis-file C7N_試題分析總表.xlsx --subject C --grade 7N --out-dir .

# Math 7N:
python scripts/generate_reports.py --ans-file M7N_ans.xlsx --analysis-file M7N_試題分析總表.xlsx --subject M --grade 7N --out-dir .

# English 7N:
python scripts/generate_reports.py --ans-file E7N_ans.xlsx --analysis-file E7N_試題分析總表.xlsx --subject E --grade 7N --out-dir .

# Batch auto-scan all subjects in a directory:
python scripts/generate_reports.py --auto-scan "<directory_path>" --out-dir "<output_directory>"
```

### Step 3: Run Automated QA Verification
Verify the generated Word documents against the input data:

```bash
python scripts/qa_verify.py --dir "<output_directory>" --analysis-dir "<analysis_files_directory>"
```
The QA suite automatically validates:
- Question table count matches analysis items.
- Cronbach's alpha and sample size match analysis metadata.
- Overall pass rate percentage in Chapter 6 text matches calculated student average.
- Table 6-1 row/column structure satisfies subject criteria.
- XML page break elements (`w:pageBreakBefore`) exist between question tables.
- Standard font `微軟正黑體` is applied to all text runs.

---

## Detailed Subject Rules

Refer to [subject_specifications.md](./references/subject_specifications.md) for full details:
- **Chinese (`C`)**: Groups questions by domain code (`A`~`J`). Subheadings `(一)形音知識` stay together with their first question table (`prev_was_heading` logic). Chapter 6 Table 6-1 has 5 columns and 13 rows.
- **Math (`M`)**: Continuous questions Q1~QN. Chapter 6 Table 6-1 has 13 columns with 3 rows per domain (高, 全, 低) and empty intersection diagonal borders (`w:tl2br`).
- **English (`E`)**: Continuous questions Q1~QN. Chapter 6 Table 6-1 has 6 columns with vertical merges for `聽力` and `閱讀`.
- **Science (`S`)**: Continuous questions Q1~QN. Supports unscored items (`不予計分`). Chapter 6 Table 6-1 has 16 columns (4 cognitive domains + overall).

---

## Troubleshooting

For common issues regarding Word XML page breaks, table row heights, or encoding errors, consult the [troubleshooting guide](./references/troubleshooting.md).
