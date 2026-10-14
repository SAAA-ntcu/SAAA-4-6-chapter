# SAAA-4-6-chapter: Subject Specifications Reference

This document details the data structures, template variables, column mappings, and layout specifications for generating Chapter 4 and Chapter 6 reports across all four subjects.

---

## 1. Subject Overview & Cohort Naming

| Subject | Code | Chinese Name | Standard Grades | Special Cohorts | Primary Question Types |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **國語文** | `C` | 國語 | 3, 4, 5, 6, 7, 8 | `C7N` | 4-choice single answer (1~4 or A~D) |
| **數學** | `M` | 數學 | 3, 4, 5, 6, 7, 8 | `M7N` | 4-choice single answer (1~4 or A~D) |
| **英語文** | `E` | 英文 / 英語文 | 4, 5, 6, 7, 8 | `E7N` | 聽力測驗 + 閱讀測驗 (1~4 or A~D) |
| **自然科學** | `S` | 自然 / 自然科學 | 5, 7, 8 | `S7N` | 4-choice single answer, supports invalid items |

---

## 2. Chapter 4 (肆、試題分析與教學建議)

### 2.1 Context Variables (docxtpl)
All subject templates utilize the following variables in jinja2 context:

| Variable | Type | Description | Example |
| :--- | :--- | :--- | :--- |
| `year` | `str` | Academic school year | `"115"` |
| `subject` | `str` | Chinese subject name | `"國語"`, `"數學"`, `"英文"`, `"自然"` |
| `grade_cn` | `str` | Chinese numeral for grade | `"三"`, `"四"`, `"五"`, `"六"`, `"七"`, `"八"` |
| `grade` | `str` | Grade code / cohort | `"7"`, `"7N"`, `"5"` |
| `cronbach_s_alpha` | `str` | Reliability index (2 decimals) | `"0.83"` |
| `valid_sample` | `str` | Number of test takers | `"21,400"` |
| `domain_legend` | `str` | Semicolon-separated legend | `"A：數與量；B：空間與形狀..."` |
| `questions` | `list[dict]` | List of question dictionaries | See question structure below |

### 2.2 Question Dictionary Fields
Each entry in `questions` includes:
- `number`: Question sequence number (1, 2, ... N).
- `code`: Item identification code, e.g. `115-C7-A1-01`, `115-M7-A1-01`, `115-E7-A1-01`, `115-S5-A1-01`.
- `ans`: Correct answer text (`A`, `B`, `C`, `D` or `①`~`④`, or `不予計分`).
- `domain`: Content domain name (e.g. `形音知識`, `數與量`, `聽力`, `物質與能量`).
- `cognitive`: Cognitive level or sub-domain name (e.g. `概念理解`, `程序執行`, `解題思考`, `辨識句意`).
- `pass_rate`: Percentage pass rate formatted as integer without `%` sign (e.g. `"68"`).
- `discrimination`: Item discrimination index formatted to 2 decimal places (e.g. `"0.45"`).
- `hi_pass`, `lo_pass`: High-group and low-group pass rates (e.g. `"88"`, `"42"`).
- `h1`~`h4`: Option labels with asterisk attached to correct answer (e.g. `h1="A*"`, `h2="B"`, `h3="C"`, `h4="D"`).
- `all_1`~`all_4`, `all_9`: Option choice proportions for all students (2 decimals). `9` represents missing/invalid responses.
- `hi_1`~`hi_4`, `hi_9`: Choice proportions for the high-scoring group (upper 27%).
- `lo_1`~`lo_4`, `lo_9`: Choice proportions for the low-scoring group (lower 27%).
- `w1_ans`~`w3_ans`, `w1_rate`~`w3_rate`: The three incorrect distractor options sorted in order of option sequence with percentage choice rates.

### 2.3 Subject-Specific Layout Rules (Chapter 4)
- **Chinese (`C`)**:
  - Questions are grouped and ordered by domain prefix: Category A (`一、語文知識：` [A~F]) then Category B (`二、閱讀理解：` [G~J]).
  - Subheadings `(一)形音知識` must be kept on the same page as the first question table of that section.
  - Subsequent question tables under the same section receive an XML page break (`w:pageBreakBefore`).
  - Major headings `二、閱讀理解：` and subheadings `([二三四五六七八九十])` receive `w:pageBreakBefore`.
  - Specific paragraph indents: `評量的` gets left indent 840 twips / 3.5 char and first line indent 480 twips / 2.0 char.
- **Math (`M`), English (`E`), Science (`S`)**:
  - Continuous question sequence from Q1 to QN.
  - Every question table starting from Q2 receives `w:pageBreakBefore`.
  - Empty paragraphs between tables are automatically purged to prevent stray blank pages.
- **Science Invalid Items**:
  - If `is_invalid == True` (unscored questions), rates are replaced with `"-"`, and answer is displayed as `"不予計分"`.

---

## 3. Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)

### 3.1 Chinese (`C`)
- **Template**: `C_demo.docx`.
- **Columns (5)**:
  1. `向度名稱` (Domain Name)
  2. `題號` (Question Numbers, e.g. `1、2、5`)
  3. `高分組` (High Group Pass Rate, 2 decimals)
  4. `全體` (All Students Pass Rate, 2 decimals)
  5. `低分組` (Low Group Pass Rate, 2 decimals)
- **Rows**: 2 header rows + 10 domain rows + 1 overall row (`整體`) = 13 rows total.
  (If original template has 11 domain rows, row 12 `國學知識` is removed).
- **Paragraphs**:
  - P1: Lists all domain names: `國語文X年級學力檢測，以「選擇題型」的方式，檢測「...」等10項能力。`
  - P2: Overall pass rate: `國語文X年級學力檢測全體學生通過率是XX%，各向度施測結果統計如下表6-1所示：`

### 3.2 Math (`M`)
- **Template**: `S_M_demo.docx`.
- **Columns (13)**:
  - Col 0: Grade (`三`~`八` / `七`)
  - Cols 1~3: `概念理解` (題號, 分組, 通過率)
  - Cols 4~6: `程序執行` (題號, 分組, 通過率)
  - Cols 7~9: `解題思考` (題號, 分組, 通過率)
  - Cols 10~12: `整體` (題數, 分組, 通過率)
- **Rows**: 2 header rows + 3 rows per content domain (高, 全, 低) + 3 overall summary rows.
  - Grade 3~6: 5 content domains = 2 + 15 + 3 = 20 rows.
  - Grade 7~8 / 7N: 4 content domains = 2 + 12 + 3 = 17 rows.
- **Empty Intersections**: Merged 3x3 block with diagonal border (`w:tl2br`).

### 3.3 English (`E`)
- **Template**: `six_templet_E.docx`.
- **Columns (6)**:
  1. `大向度` (Main Domain: 聽力 / 閱讀)
  2. `次向度` (Sub-domain: 辨識句意, 基本常用字詞...)
  3. `題號` (Question Numbers)
  4. `高分組` (High Group Pass Rate)
  5. `全體` (All Students Pass Rate)
  6. `低分組` (Low Group Pass Rate)
- **Merging**: Column 0 is vertically merged across all `聽力` rows, and merged across all `閱讀` rows.
- **Overall Row**: Bottom row merges Cols 0 and 1 with label `整體`.

### 3.4 Science (`S`)
- **Template**: `S_M_demo.docx` (reconstructed to 16 columns).
- **Columns (16)**:
  - Col 0: Grade
  - Cols 1~3: `記憶`
  - Cols 4~6: `理解`
  - Cols 7~9: `應用`
  - Cols 10~12: `分析`
  - Cols 13~15: `整體` (題數, 分組, 通過率)
- **Rows**: 2 header rows + 3 rows per content domain + 3 overall rows.
- **Layout**: 12pt Chinese font with 2-by-2 character wrapping for domain names.
