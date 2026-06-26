# -*- coding: utf-8 -*-
"""產生 EP03 教學檔案處理工具列表的一頁式 PDF（繁體中文）。
使用 reportlab 內建 MSung-Light CID 字型，免外部字型檔。
重新產生：python make_pdf.py
"""
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle)

FONT = "MSung-Light"
pdfmetrics.registerFont(UnicodeCIDFont(FONT))

styles = getSampleStyleSheet()
P = ParagraphStyle("body", parent=styles["Normal"], fontName=FONT,
                   fontSize=7.2, leading=9.2)
PH = ParagraphStyle("cell", parent=P, fontSize=7, leading=8.8)
TITLE = ParagraphStyle("title", parent=styles["Title"], fontName=FONT,
                       fontSize=18, leading=21)
SUB = ParagraphStyle("sub", parent=styles["Normal"], fontName=FONT,
                     fontSize=8, leading=11, textColor=colors.HexColor("#555555"))
CAT = ParagraphStyle("cat", parent=styles["Normal"], fontName=FONT,
                     fontSize=10, leading=13, textColor=colors.white)

ACCENT = colors.HexColor("#0d6efd")
HEADBG = colors.HexColor("#0d1530")
CATBG = colors.HexColor("#1f6feb")
ZEBRA = colors.HexColor("#eef4ff")


def cell(t):
    return Paragraph(t, PH)


# 每組：(分類, [ (項目, 痛點, 套件, 一句話) ... ])
DATA = [
    ("📄 Word", [
        ("套印獎狀/通知單/成績單 🏆", "30 人 30 張，手動換名換到崩潰",
         "python-docx + openpyxl", "讀班級名單 Excel，套進獎狀 Word 模板，每人產一份並存成 PDF"),
        ("出考卷（學生卷/教師卷分開）", "出完題還要手動刪答案做另一版",
         "python-docx", "把題目做成 Word，產出『學生卷(無答案)』和『教師卷(含詳解)』兩份"),
        ("講義合併 + 批次轉 PDF", "各課 Word 散落，要併本再轉 PDF",
         "docxcompose + pywin32", "把資料夾的 Word 講義依檔名合併成一份，再另存成 PDF"),
    ]),
    ("📊 Excel", [
        ("成績計算 + 排名 + 及格標紅", "平均、加權、排名、標紅，每次段考重來",
         "openpyxl + xlsxwriter", "讀 grades.xlsx 算總分與排名，不及格標紅，各班平均放最後一列"),
        ("段考分析（答對率/落點圖）", "想看哪題全班錯最多，但不會做統計圖",
         "pandas + matplotlib", "分析答題明細，畫各題答對率長條圖和全班分數分布圖"),
        ("總成績單拆成各班/各生", "一張大表要拆給各導師、或每生一張",
         "openpyxl", "把全校成績表依『班級』欄拆成一個班一個 Excel 檔"),
        ("隨機座位表/分組", "每次調座位、分組都要手喬",
         "xlsxwriter", "用名單隨機排一張 6×5 座位表，輸出成 Excel"),
    ]),
    ("📑 PowerPoint", [
        ("教材大綱 → 整份簡報 🪄", "把重點一頁頁貼進 PPT 很花時間",
         "python-pptx", "把教材大綱每個重點做成一頁投影片，套用指定範本"),
        ("圖片 → 圖卡簡報", "單字卡、圖鑑要一張張貼",
         "python-pptx + pillow", "把資料夾圖片每張做成一頁投影片，下方加檔名當標題"),
        ("統一字型/加校徽", "別人給的 PPT 字體亂，要逐頁改",
         "python-pptx", "把 PPT 全部字型改成標楷體，每頁右下角加上校徽"),
    ]),
    ("📕 PDF", [
        ("考卷合併/拆分/重排", "考古題、作業散在幾十個 PDF",
         "pypdf + PyMuPDF", "把這些 PDF 合併成一份，並把第 5~8 頁單獨抽出另存"),
        ("加浮水印（防外流）", "講義想加『僅供 ○ 班使用』",
         "pypdf + reportlab", "幫 PDF 每頁加上淡灰色浮水印『302 班 期中複習』"),
        ("掃描講義 OCR → 可編輯", "掃描檔是圖片，無法選字、無法餵 AI",
         "ocrmypdf", "把掃描 PDF 做 OCR，變成可以複製文字的 PDF"),
        ("抽課本某幾頁轉圖", "只要課本某張圖貼進學習單",
         "PyMuPDF + pdf2image", "把課本 PDF 第 12 頁轉成圖片，去掉白邊"),
    ]),
    ("🧰 其他常用", [
        ("連結轉 QR Code", "要學生掃 Padlet/表單，手做 QR 很煩",
         "qrcode + pillow", "把 5 個連結各生一張 QR Code，貼到學習單"),
        ("影片轉字幕", "自己聽打逐字稿要好幾小時",
         "yt-dlp + openai-whisper", "把這支影片轉成 SRT 字幕"),
        ("講稿轉語音旁白（免費）", "請人配音或買 AI 語音很貴",
         "edge-tts", "把講稿轉成中文語音 mp3，語速慢一點"),
        ("課本轉乾淨文字（餵 AI）", "課本格式亂，丟 AI 效果差",
         "markitdown", "把 PDF/PPT 轉成 Markdown"),
    ]),
]

COLW = [40 * mm, 52 * mm, 42 * mm, 132 * mm]
HEAD = [cell(f"<b>{h}</b>") for h in ("項目", "痛點", "套件", "一句話（複製給 Agent）")]


def build():
    doc = SimpleDocTemplate(
        "EP03_教學檔案處理_工具列表.pdf", pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm,
        topMargin=9 * mm, bottomMargin=8 * mm)
    flow = []
    flow.append(Paragraph("EP03 教學檔案處理工具列表｜老師的 Python 神器清單", TITLE))
    flow.append(Paragraph(
        "三師爸 Sense Bar・AI Agent 基本功 EP03　|　用法：把「一句話」直接複製貼給你的 AI Agent，它就會幫你做。完整資料與簡報見 GitHub。",
        SUB))
    flow.append(Spacer(1, 3 * mm))

    rows = [HEAD]
    style = [
        ("FONTNAME", (0, 0), (-1, -1), FONT),
        ("BACKGROUND", (0, 0), (-1, 0), HEADBG),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cccccc")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    r = 1
    for catname, items in DATA:
        rows.append([Paragraph(f"<b>{catname}</b>", CAT), "", "", ""])
        style.append(("SPAN", (0, r), (-1, r)))
        style.append(("BACKGROUND", (0, r), (-1, r), CATBG))
        r += 1
        for item, pain, pkg, line in items:
            rows.append([cell(item), cell(pain),
                         cell(f"<font color='#0d6efd'>{pkg}</font>"), cell(line)])
            if r % 2 == 0:
                style.append(("BACKGROUND", (0, r), (-1, r), ZEBRA))
            r += 1

    t = Table(rows, colWidths=COLW, repeatRows=1)
    t.setStyle(TableStyle(style))
    flow.append(t)
    doc.build(flow)
    print("PDF 產生完成：EP03_教學檔案處理_工具列表.pdf")


if __name__ == "__main__":
    build()
