# EP03 ・ Python 能力：教學現場老師的「檔案處理」工具地圖

> 來源：實際掃描三師爸這台電腦（Python 3.14.3，179 個套件）+ 2026.06 上網查找。
> 受眾（TA / Target Audience）：**教學現場的老師**。
> 用途：EP03「能力② 執行程式與裝工具」的 demo 素材。
> 主軸：聚焦老師天天在處理的 **Word／Excel／PPT／PDF**，痛點都是「人工做很煩、量一大就爆炸」。
> 圖例：✅＝這台已裝　🆕＝建議再裝（`pip install`）

---

## 一、本機已安裝套件總覽（依用途分類）

| 類別 | 代表套件 | 能做什麼 |
|------|----------|----------|
| 📄 Word | `python-docx`、`docxcompose` | 生成/讀寫/合併 Word |
| 📊 Excel | `xlsxwriter`、`odfpy` | 生成 Excel（讀取需另裝 openpyxl） |
| 📑 PPT | `python-pptx` | 生成/改寫 PowerPoint |
| 📕 PDF | `PyMuPDF`、`pypdf`、`pdfplumber`、`pdf2image`、`fpdf2`、`reportlab` | 合併/拆分/抽文字/轉圖/浮水印/生成 |
| 🔄 轉檔 | `markitdown`、`pdf2image`、`pywin32`(COM) | 各檔互轉、Office 轉 PDF |
| 🖼️ 圖像/圖表 | `pillow`、`matplotlib`、`qrcode` | 圖片處理、數據圖、QR Code |
| 🎙️ 語音影音 | `openai-whisper`、`edge-tts`、`yt-dlp` | 字幕、旁白、影片下載 |

---

## 二、教學檔案處理：老師最有感的專案（核心 demo）

每個都附：痛點 → 用到的套件 → 對 Agent 說的一句話。

### 📄 Word 篇

**W1. 🏆 套印個人化通知單／獎狀／成績單（mail merge）**　✅＋🆕`openpyxl`
- **痛點**：30 個學生要 30 張獎狀／30 份個別通知單，手動換名字換到天荒地老
- **套件**：`python-docx`（套模板）+ `openpyxl`（讀名單）
- **一句話**：「讀這份班級名單 Excel，套進這個獎狀 Word 模板，每個學生產一份，存成 PDF」

**W2. 📝 出考卷／學習單，題目卷與答案卷分開**　✅
- **痛點**：出完題還要手動做一份去掉答案的版本
- **套件**：`python-docx`
- **一句話**：「把這些題目做成 Word，產出『學生卷（無答案）』和『教師卷（含詳解）』兩份」

**W3. 📚 多份講義合併成一份、批次轉 PDF**　✅`docxcompose`＋`pywin32`
- **痛點**：各課 Word 講義散落，要併成一本、再轉 PDF 發給學生
- **套件**：`docxcompose`（合併）+ `pywin32` COM（轉 PDF）
- **一句話**：「把這資料夾的 Word 講義依檔名順序合併成一份，再另存成 PDF」

### 📊 Excel 篇

**E1. 🧮 成績計算 + 排名 + 及格標示 + 各班統計**　✅＋🆕`openpyxl`
- **痛點**：算平均、加權、排名、標紅不及格，每次段考重來一遍
- **套件**：`openpyxl`（讀寫）+ `xlsxwriter`（美化輸出）
- **一句話**：「讀 grades.xlsx，算總分與排名，不及格標紅，各班平均放最後一列，存成新檔」

**E2. 📈 段考成績分析（各題答對率、平均、落點圖）**　🆕`pandas`＋✅`matplotlib`
- **痛點**：想看哪一題全班錯最多、班級落點，但不會做統計圖
- **套件**：`pandas`（分析）+ `matplotlib`（畫圖）
- **一句話**：「分析這份答題明細，畫出各題答對率長條圖和全班分數分布圖」

**E3. ✂️ 把總成績單「拆成」各班／各學生個別檔**　🆕`openpyxl`
- **痛點**：一份大表要拆成各班導師各自的檔、或每生一張個人成績單
- **套件**：`openpyxl`
- **一句話**：「把這份全校成績表，依『班級』欄拆成一個班一個 Excel 檔」

**E4. 🪑 自動產生座位表 / 隨機分組**　✅`xlsxwriter`
- **痛點**：每次調座位、分組都要手喬
- **套件**：`xlsxwriter`（+ 內建 random）
- **一句話**：「用這份名單隨機排一張 6×5 的座位表，輸出成 Excel」

### 📑 PPT 篇

**P1. 🪄 教材大綱 → 自動生成整份上課簡報**　✅
- **痛點**：把講義重點一頁一頁貼進 PPT 很花時間
- **套件**：`python-pptx`
- **一句話**：「把這份教材大綱，每個重點做成一頁投影片，套用這個範本」

**P2. 🖼️ 一堆圖片 → 自動排成圖卡簡報**　✅`python-pptx`＋`pillow`
- **痛點**：單字卡、生物圖鑑、作品集，要一張張貼進 PPT
- **套件**：`python-pptx` + `pillow`
- **一句話**：「把這資料夾的圖片，每張做成一頁投影片，下方加上檔名當標題」

**P3. 🎨 統一整份簡報的字型／字級／加上校徽**　✅
- **痛點**：別人給的 PPT 字體亂七八糟，要逐頁改
- **套件**：`python-pptx`
- **一句話**：「把這份 PPT 全部字型改成標楷體、每頁右下角加上這張校徽」

### 📕 PDF 篇

**D1. 📎 考卷合併 / 拆分 / 重新排序**　✅
- **痛點**：考古題、學生作業散在幾十個 PDF
- **套件**：`pypdf`、`PyMuPDF`
- **一句話**：「把這些 PDF 合併成一份，並把第 5～8 頁單獨抽出來另存」

**D2. 💧 PDF 加浮水印（班級／姓名／防外流）**　✅
- **痛點**：講義或考卷想加上「僅供 ○ 班使用」防止外流
- **套件**：`pypdf`、`reportlab`
- **一句話**：「幫這份 PDF 每頁加上淡灰色浮水印『302 班 期中複習』」

**D3. 🔍 掃描的紙本講義 OCR → 可搜尋、可複製**　🆕`ocrmypdf`（需 Tesseract）
- **痛點**：掃描的考古題是圖片，無法選取文字、無法餵 AI
- **套件**：`ocrmypdf` / `pytesseract`
- **一句話**：「把這份掃描 PDF 做 OCR，變成可以複製文字的 PDF」

**D4. 🧩 抽課本某幾頁 / PDF 轉圖貼到學習單**　✅`PyMuPDF`＋`pdf2image`
- **痛點**：只要課本某張圖、某幾頁貼進學習單
- **套件**：`PyMuPDF`、`pdf2image`
- **一句話**：「把這份課本 PDF 的第 12 頁轉成圖片，去掉白邊」

---

## 三、其他已驗證的老師小工具（✅ 本機已裝）

| 工具 | 一句話 |
|------|--------|
| 🔳 連結轉 QR Code（`qrcode`） | 「把這 5 個連結各生一張 QR Code，貼到學習單」 |
| 📹 影片轉字幕（`yt-dlp`+`openai-whisper`） | 「把這支影片轉成 SRT 字幕」 |
| 🔊 講稿轉語音旁白（`edge-tts`） | 「把這份講稿轉成中文語音 mp3」 |
| 📦 課本轉乾淨文字（`markitdown`） | 「把這份 PDF/PPT 轉成 Markdown 餵 AI」 |

---

## 四、需要先補裝的套件（安裝缺口提醒）

| 套件 | 為什麼需要 | 安裝 |
|------|-----------|------|
| `openpyxl` | 目前只有 `xlsxwriter`（**只能寫不能讀**），讀成績 Excel 必備 | `pip install openpyxl` |
| `pandas` | 成績分析、統計、拆表的主力 | `pip install pandas` |
| `ocrmypdf` | 掃描講義 OCR（另需系統裝 Tesseract 引擎） | `pip install ocrmypdf` |
| `docx2pdf`（可選） | Word 轉 PDF 的簡單方案（或直接用已裝的 pywin32 COM） | `pip install docx2pdf` |

> 一鍵補齊：`pip install openpyxl pandas ocrmypdf docx2pdf`

---

## 五、對應 EP03 的鋪陳建議

- **三大頭牌 demo（最有感、最好炒氣氛）**：
  1. **W1 套印獎狀／通知單**——「30 張獎狀一秒生」最震撼，幾乎每位老師都做過。
  2. **E1+E2 成績全套**——計算、排名、拆班、畫分析圖，命中行政庶務痛點。
  3. **P1 教材變整份 PPT**——備課時間直接砍半。
- **節奏**：每個格式（Word→Excel→PPT→PDF）各挑 1 個現場 demo，剛好對應「能力② 執行程式」一段。
- **誠實提醒**：示範前先 `pip install openpyxl pandas`，順便讓觀眾看到「Agent 自己會把缺的套件裝起來」——正好呼應本集「執行程式與裝工具」的主題。

---

## 參考來源（2026.06 上網查找）
- Word Mail Merge with Python：https://learndataanalysis.org/automate-microsoft-word-mail-merge-with-python/
- 大量證書/獎狀套印（教師實例）：https://frankbuck.org/certificate-creation/
- 從 Excel/Google Sheets 生成證書：https://certifier.io/blog/how-to-create-certificates-from-google-sheets-and-excel
</content>
