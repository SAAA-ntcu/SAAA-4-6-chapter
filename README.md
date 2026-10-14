# 國中小學力檢測 第四章與第六章報告自動化生成 Skill (SAAA-4-6-chapter)

本專案為專門用於臺灣國中小學力檢測、新生學力測驗之 **「第四章（肆、試題分析與教學建議）與第六章（陸、整體教學建議 / 表6-1 各評量向度通過率）」自動化產出與排版技能包（Skill）**。相容於各類 AI Agent 體系（Google Antigravity、Claude Desktop / Projects、Cursor、Codex 等）與獨立 Python CLI 執行。

---

## 🌟 核心功能與特色

1. **第四章（試題分析與教學建議）完整產出**：
   - 自動解析 Excel 答案檔與試題分析總表，逐題繪製通過率、鑑別度、高低分組劃記率矩陣。
   - 注入 Word 底層 XML 分頁標籤（`w:pageBreakBefore`），確保每一題獨立分頁且小標題（如 `(一)形音知識`）不孤行。
2. **第六章（表 6-1 各評量向度通過率）矩陣生成**：
   - 支援 4 科不同的向度維度架構（國語單向度、數學/自然內容與認知雙向度交集、英語聽力/閱讀垂直跨行合併）。
   - 數學/自然自動注入空白向度交集之灰色斜線邊框（`w:tl2br`）。
   - 固定列高與行距（24pt）及微軟正黑體排版。
3. **高精度統計學運算（消除二次捨入累積誤差）**：
   - 優先讀取未截斷之原始分析結果（`*_分析結果.xlsx`），進行真實浮點數之向度平均。
   - 內建嚴格四捨五入（Round Half-Up）格式化演算法，徹底消弭二進位浮點數（如 `0.805` 被誤捨為 `0.80`）的計算誤差。
4. **全學科與年級廣泛支援**：
   - 支援 4 大學科：國語文 (`C`)、數學 (`M`)、英語文 (`E`)、自然科學 (`S`)。
   - 支援國小至國中全年級（3~8 年級）及新生特定專案代碼（如 `C7N`、`M7N`、`E7N`、`S7N`）。
5. **內建自動化 QA 稽核檢驗套件**：
   - 內建 `qa_verify.py`，一鍵自動核對題數、信度、樣本數、總體通過率與 Word XML 排版規範。

---

## 📁 專案結構說明

```text
SAAA-4-6-chapter/
├── SKILL.md                          # Agent 技能定義與操作規範
├── README.md                         # 專案完整說明文件
├── requirements.txt                  # Python 相依套件清單
├── scripts/                          # 核心處理腳本
│   ├── generate_reports.py          # 統一多學科報表生成引擎（支援 CLI 與批次處理）
│   ├── qa_verify.py                 # 自動化稽核與數據驗證工具
│   └── extract_dimensions.py        # 答案卷向度提取與標準化轉換工具
├── resources/
│   └── templates/                   # 標準 Word 官方樣板
│       ├── pydocxTemplet_C_new.docx # 國語文第四章樣板
│       ├── pydocxTemplet_M.docx     # 數學第四章樣板
│       ├── pydocxTemplet_E.docx     # 英語文第四章樣板
│       ├── pydocxTemplet_S.docx     # 自然科學第四章樣板
│       ├── C_demo.docx              # 國語文第六章表 6-1 樣板
│       ├── S_M_demo.docx            # 數學與自然第六章表 6-1 樣板
│       └── six_templet_E.docx       # 英語文第六章表 6-1 樣板
└── references/                       # 技術規範與參考文件
    ├── subject_specifications.md    # 各學科欄位規格與資料模型詳細定義
    └── troubleshooting.md           # Word XML 排版與故障排除指引
```

---

## 🚀 快速開始

### 1. 安裝環境需求

```bash
git clone https://github.com/SAAA-ntcu/SAAA-4-6-chapter.git
cd SAAA-4-6-chapter
pip install -r requirements.txt
```

### 2. 單一學科與年級產出

```bash
# 同時產生第四章與第六章（以桃園七年級新生國語 C7N 為例）：
python scripts/generate_reports.py \
  --ans-file "國語文/C7N_ans.xlsx" \
  --analysis-file "C7N_試題分析總表.xlsx" \
  --subject C \
  --grade 7N \
  --year 115 \
  --chapter both \
  --out-dir "./output"

# 數學 M7N：
python scripts/generate_reports.py --ans-file M7N_ans.xlsx --analysis-file M7N_試題分析總表.xlsx --subject M --grade 7N --out-dir ./output

# 英語 E7N：
python scripts/generate_reports.py --ans-file E7N_ans.xlsx --analysis-file E7N_試題分析總表.xlsx --subject E --grade 7N --out-dir ./output
```

### 3. 目錄自動掃描與批次產出

```bash
# 自動掃描指定目錄下的所有學科試卷與總表並批次產生：
python scripts/generate_reports.py --auto-scan "./data" --out-dir "./output"
```

### 4. 執行品質稽核驗證

```bash
python scripts/qa_verify.py --dir "./output" --analysis-dir "./data"
```

---

## 🤖 如何掛載至各類 AI Agent？

### 方式 1：在 Google Antigravity / Gemini CLI 中使用
將本技能目錄放置於 `.agents/skills/SAAA-4-6-chapter` 或全域配置 `~/.gemini/config/skills/SAAA-4-6-chapter`，Agent 即會自動發現並調用本技能。

### 方式 2：在 Cursor / Claude Desktop / 一般 LLM 專案中使用
將此資料夾放入專案根目錄，並在系統提示詞（System Prompt / Rules / Memory）中加入：
> 「請參考 `@SAAA-4-6-chapter/SKILL.md` 規範，依照精確統計學精度與官方樣板排版規則，執行學力檢測第四章與第六章報告之生成與校驗任務。」

---

## 📄 授權條款

本專案由 [SAAA-ntcu](https://github.com/SAAA-ntcu) 維護，專供教育評量資料分析、學力檢測報告編制與學術研究使用。
