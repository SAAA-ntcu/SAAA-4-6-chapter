---
name: saaa-4-6-chapter
description: >-
  Automates the generation of Chapter 4 (肆、試題分析與教學建議) and Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)
  for Taiwanese academic diagnostic assessments (學力檢測 / 新生學力測驗).
  Supports 4 subjects: Chinese (國語文 C), Math (數學 M), English (英語文 E), and Science (自然科學 S)
  across standard elementary/junior high grades (3~8) and special entry cohorts (C7N, M7N, E7N, S7N).
  Uses deterministic scripts to parse Excel datasets, render docxtpl Word templates, inject XML page breaks,
  and generate Table 6-1 matrices with automated QA validation. Can crop booklet PDFs,
  preserve Chinese text frames and shared reading passages, and insert verified images into Chapter 4.
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

> [!NOTE]
> **Special Rules for 7th Grade Entrants (7N / 桃園市七年級新生)**:
> - **Internal Item Codes**: Use grade digit `6` (e.g. `115-C6-A1-01`, `115-M6-D1-01`, `115-E6-A1-01`, `115-S6-A1-01`), representing Grade 6 curriculum scope.
> - **Exam Title & Preamble**: Dynamically formatted as `[年]年桃園市[科目]七年級新生學習能力檢測`.
> - **Math Table 6-1 Header**: Top-left merged cell displays `六`.
> - **Table 6-1 Title**: Standardized to `表6-1　整體學生及高分組、低分組學生於各評量向度之通過率`.
> - **Expert Fields**: Preserve learning-focus and qualitative-analysis slots for domain experts. With a supplied booklet, automatically fill the Chapter 4 question-image row; leave it blank when no booklet is supplied.

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

### Booklet-to-report workflow

When existing Word reports or drafts are supplied, read [word_reference_review.md](./references/word_reference_review.md). Use `inspect_reference_docs.py` to inventory logical questions in merged tables, source images, workbooks, comments and revisions. Matching item codes alone do not establish that general-grade and entrant-cohort questions are the same. Use `verify_editorial.py` to distinguish data/image completion from pending expert content.

When the user supplies subject folders with booklet PDF, horizontal answer Excel and assessment-framework PPTX, read [screenshot_integration.md](./references/screenshot_integration.md) and use `scripts/build_from_source.py`. Supply the actual analysis directory; do not treat an answer sheet, historical item difficulty or a missing workbook as student results. Drafts with missing statistics require explicit `--draft` and visible pending fields.

For 桃園市七年級新生 validation, use the three subject folders and their matching C7N/M7N/E7N analysis files. The integrated entry point defaults to `_AI產出` filenames. It inserts every question by actual number, includes shared articles before their first related question, and records source/PNG hashes. Run both data/image QA and rendered-page inspection before delivery.

Keep question text at a consistent reading size: use the shared source scale from `image_layout.py`, never stretch every trimmed image to fill the cell. Verify actual embedded dimensions as well as image hashes; the layout audit records the source font estimate and effective reading size.
For English, target 13pt effective booklet text; retain the 11pt target for other subjects. Shared passages use their own proportional page-fit scale. Use `build_from_source.py --subject E` to revise English alone after another subject's layout has been accepted.

Remove the leading question number from each booklet image without removing question text, options, diagrams or shared articles. From 「一、試題分析」 onward, use exactly 24pt line spacing. Avoid paragraph keep flags that produce black square formatting marks in Word; use explicit page breaks outside item tables.

The standard normalized-Excel workflow below remains available. Add `--booklet-pdf` or `--images-manifest` to its single-subject command to fill the question-image row.

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
- Explicit page breaks exist between question tables.
- Standard font `微軟正黑體` is applied to all text runs.
- With `.images.json`, every question image matches the recorded PNG hash and every shared article part appears exactly once.
- Chapter 4 analysis paragraphs have exactly 24pt line spacing; no paragraph keep flags remain that cause black square formatting marks. Newly cropped images record removal of the leading question number.

QA is structural/data evidence, not proof of page layout. Render and inspect the final reports; preserve the source Word templates. On Windows without LibreOffice, `scripts/render_word.ps1` can export read-only QA PDFs through a hidden independent Microsoft Word instance.

After rendering, run `python scripts/verify_pagination.py --dir "<output_directory>" --pdf-dir "<rendered_pdf_directory>" --json "<output_directory>/分頁檢驗_AI產出.json"`. Every question's selection-rate table and analysis opening must occur on the same PDF page; long analyses may continue naturally on the next page. Keep the subject-wide question image scale unchanged when repairing pagination. Combine all parts of each Chinese/English shared passage into one complete image and fit it proportionally on one page with its heading; use a separate page-fit scale for passages. Keep ordinary question images left/top aligned. Chinese and English shared passages must fit completely on one page and be vertically centered in the page article area below the heading, with horizontal alignment left. Range-only listening headers remain inside the first item. Retain all image scales and aspect ratios. Measure reference Word layouts with `scripts/measure_reference_layout.py`; see `references/word_reference_review.md` for interpretation limits.

---

County/city names in report text (for example 桃園市) use a proper-name mark: a single underline applied only to the name. Preserve all wording and other formatting; apply `scripts/proper_names.py` after final document assembly.

## Detailed Subject Rules

Screenshots supplied as layout references control spacing, image placement and table geometry only. Preserve the existing Chinese introduction text, grade wording, answer/option labels, assessment-target text and knowledge-dimension fields. For the reference layout, put each category introduction together with its first item table and use matching list indentation for both categories. This moves page boundaries without rewriting content.

Refer to [subject_specifications.md](./references/subject_specifications.md) for full details:
- **Chinese (`C`)**: Groups questions by domain code (`A`~`J`). Subheadings `(一)形音知識` stay together with their first question table (`prev_was_heading` logic). Chapter 6 Table 6-1 has 5 columns and 13 rows.
- **Math (`M`)**: Continuous questions Q1~QN. Chapter 6 Table 6-1 has 13 columns with 3 rows per domain (高, 全, 低) and empty intersection diagonal borders (`w:tl2br`).
- **English (`E`)**: Continuous questions Q1~QN. Chapter 6 Table 6-1 has 6 columns with vertical merges for `聽力` and `閱讀`.
- **Science (`S`)**: Continuous questions Q1~QN. Supports unscored items (`不予計分`). Chapter 6 Table 6-1 has 16 columns (4 cognitive domains + overall).

---

## Troubleshooting

For common issues regarding Word XML page breaks, table row heights, or encoding errors, consult the [troubleshooting guide](./references/troubleshooting.md).
