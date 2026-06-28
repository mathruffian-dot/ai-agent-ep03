# 🤖 自動安裝指南：教學檔案處理工具包

> **這份檔案是寫給 AI Agent 讀的。**
> 使用者把這份檔案交給你（Claude Code / Codex / Antigravity / OpenCode 等）時，
> 代表他想安裝以下 Python 套件，請依「執行步驟」自動完成安裝與驗證。

---

## 🎯 給 Agent 的指令（請照做）

當使用者要求「依這份檔案安裝」時，請依序執行：

1. **確認環境**：偵測作業系統與 Python／pip 版本（`python --version`、`pip --version`）。若無 Python，提醒使用者先安裝 Python 3.10+。
2. **安裝 Python 套件**：執行下方「核心安裝指令」。建議在虛擬環境（venv）或使用者既有環境安裝，安裝前先告知使用者你要裝什麼。
3. **處理系統相依**：依使用者作業系統，提示或安裝下方「系統相依套件」（OCR 引擎、Poppler、ffmpeg）。這些不是 pip 套件，需個別處理。
4. **驗證**：執行下方「驗證指令」，回報每個套件是否 import 成功，列出失敗清單。
5. **回報**：給使用者一份清單——✅ 成功 / ⚠️ 需手動處理（系統相依）/ ❌ 失敗與原因。

> 權限提醒：安裝會用到 `pip install` 與（可能）系統套件管理器，請依使用者的權限設定執行；若為半自動模式，先說明再進行。

---

## 📦 核心安裝指令（所有平台）

```bash
pip install python-docx docxcompose openpyxl xlsxwriter pandas python-pptx pypdf PyMuPDF pdfplumber pdf2image reportlab fpdf2 pillow matplotlib qrcode markitdown ocrmypdf docx2pdf
```

### 僅 Windows 追加（COM 自動化：控制 Word/PPT、Office 轉 PDF）

```bash
pip install pywin32
```

### 影音工具（選用：旁白、影片下載、抓現成字幕）

```bash
pip install edge-tts yt-dlp youtube-transcript-api
```

> 註：本工具包**刻意不含 `openai-whisper`**。它需下載大型語音模型、吃 CPU/GPU，
> 一般使用者常裝不起來也跑不動。需要影片逐字稿時，改用輕量的
> `youtube-transcript-api`（直接抓 YouTube 現成字幕，免模型、免 ffmpeg）。

---

## 🧩 套件用途對照（安裝時可向使用者說明）

| 套件 | 用途 |
|------|------|
| `python-docx` | 生成／讀寫 Word（出考卷、套印） |
| `docxcompose` | 合併多份 Word |
| `openpyxl` | **讀取** Excel（成績計算、拆表必備） |
| `xlsxwriter` | 生成美化 Excel |
| `pandas` | 成績分析、統計、拆表 |
| `python-pptx` | 生成／改寫 PowerPoint |
| `pypdf` | PDF 合併、拆分、浮水印 |
| `PyMuPDF` | PDF 抽頁、抽文字、轉圖 |
| `pdfplumber` | PDF 表格／文字精準抽取 |
| `pdf2image` | PDF 轉圖片（需 Poppler） |
| `reportlab` | 生成 PDF、做浮水印圖層 |
| `fpdf2` | 輕量生成 PDF |
| `pillow` | 圖片處理（裁切、去白邊、合成） |
| `matplotlib` | 數據圖表（成績落點、答對率） |
| `qrcode` | 生成 QR Code |
| `markitdown` | 各種檔案轉乾淨 Markdown（餵 AI 前處理） |
| `ocrmypdf` | 掃描 PDF 做 OCR（需 Tesseract） |
| `docx2pdf` | Word 轉 PDF（Windows/Mac，需安裝 Office/Word） |
| `pywin32` | Windows COM 自動化（控制 Office、轉檔） |
| `edge-tts` | 文字轉語音（免費、雲端、輕量） |
| `yt-dlp` | 下載 YouTube 影片／音訊 |
| `youtube-transcript-api` | 抓 YouTube 現成字幕做逐字稿（輕量、免模型） |

---

## ⚙️ 系統相依套件（非 pip，需個別安裝）

有些套件需要系統層級的程式才能運作，請依作業系統處理：

### 1. Tesseract OCR（`ocrmypdf` 需要）
- **Windows**：`winget install UB-Mannheim.TesseractOCR`
- **macOS**：`brew install tesseract tesseract-lang`
- **Linux (Debian/Ubuntu)**：`sudo apt install tesseract-ocr tesseract-ocr-chi-tra`
- **繁中語言包（重要）**：Windows 用 winget 靜默安裝只含 `eng`+`osd`，要 OCR 中文需自行補繁中包：
  下載 [`chi_tra.traineddata`](https://github.com/tesseract-ocr/tessdata_best/raw/main/chi_tra.traineddata)
  放進 `C:\Program Files\Tesseract-OCR\tessdata\`（寫入需系統管理員權限）。
  驗證：`tesseract --list-langs` 應出現 `chi_tra`。

### 2. Poppler（`pdf2image` 需要）
- **Windows**：`winget install oschwartz10612.Poppler`
- **macOS**：`brew install poppler`
- **Linux**：`sudo apt install poppler-utils`

### 3. ffmpeg（`yt-dlp` 下載與音訊合併需要）
- **Windows**：`winget install Gyan.FFmpeg`
- **macOS**：`brew install ffmpeg`
- **Linux**：`sudo apt install ffmpeg`

### 4. Microsoft Word／PowerPoint（`docx2pdf`、`pywin32` COM 轉檔需要）
- 需本機已安裝 Microsoft Office。無 Office 者，Word 轉 PDF 改用 LibreOffice：`soffice --headless --convert-to pdf <檔案>`。

> ⚠️ **裝完系統工具務必重開終端機！** Tesseract / Poppler / ffmpeg 安裝後會寫進 PATH，
> 但**目前開著的視窗（含你這個 Agent）抓的是舊 PATH**，必須重開一個新的終端機（或重啟 Agent），
> `ocrmypdf`、`pdf2image`、`yt-dlp` 才找得到它們。安裝完請提醒使用者這一點。

---

## ✅ 驗證指令

安裝後執行，確認套件可正常匯入（回報失敗清單）：

```bash
python -c "import docx, docxcompose, openpyxl, xlsxwriter, pandas, pptx, pypdf, fitz, pdfplumber, pdf2image, reportlab, fpdf, PIL, matplotlib, qrcode, markitdown, ocrmypdf; print('✅ 全部核心套件匯入成功')"
```

> 註：部分套件的「匯入名稱」與「安裝名稱」不同——
> `python-docx` → `import docx`；`PyMuPDF` → `import fitz`；`python-pptx` → `import pptx`；`pillow` → `import PIL`；`fpdf2` → `import fpdf`。

---

## 📌 給 Agent 的最終回報範本

```
安裝完成回報：
✅ Python 套件：全部 / 部分成功（列出失敗者）
⚠️ 系統相依：Tesseract / Poppler / ffmpeg / Office —— 已裝 or 需使用者手動處理
❌ 失敗項目與原因：...
下一步建議：...
```

---

> 來源：三師爸 Sense Bar｜AI Agent 基本功 EP03｜youtube.com/@sensebar
</content>
