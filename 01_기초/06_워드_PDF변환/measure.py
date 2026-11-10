"""워드 기초 6강 — PDF 로 저장 실측. 이 PC 워드(Microsoft 365)를 따로 띄워(DispatchEx) 연습 문서를 만들고 여러 방법으로 PDF 를 만든 뒤
PyMuPDF 로 PDF 를 읽어 쪽 수 · 크기 · 링크 · 책갈피(목차) · 글꼴 포함 여부 · 그림 픽셀 · 메모 글자가 찍혔는지를 잰다.
- A: 다른 이름으로 저장 → PDF(SaveAs2 FileFormat 17)
- B: 내보내기 표준(ExportAsFixedFormat OptimizeFor 0) · C: 최소 크기(OptimizeFor 1)
- D: 제목으로 책갈피 만들기(CreateBookmarks 1) · E: 변경 내용 표시 문서(Item 7, 메모 포함) · F: 2쪽만(Range 3, From 2 To 2)
- G: PDF/A(UseISO19005_1)
연습 문서: 가상 '동아리 11월 활동 보고서' — 제목 1 세 개 · 제목 2 둘 · 링크 하나 · 큰 사진 하나(만든 그림) · 표 하나 · 메모 하나, 3쪽.
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 · pymupdf · pillow)"""
import datetime, json, os, random, sys
import pymupdf
import pythoncom
import win32com.client as win32
from PIL import Image, ImageDraw, ImageFilter

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "word06" + ("_en" if EN else ""))
os.makedirs(WORK, exist_ok=True)
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}

if EN:
    T = {"title": "Club report — November", "h": ["1. What we did", "2. Photos", "3. Next month"], "h2": ["Members", "Budget"],
         "link_text": "club notice board", "comment": "Check the date before sending", "body": "This is a practice report for a made-up club. "}
else:
    T = {"title": "동아리 11월 활동 보고서", "h": ["1. 한 일", "2. 사진", "3. 다음 달 계획"], "h2": ["참가 인원", "예산"],
         "link_text": "동아리 공지 게시판", "comment": "보내기 전에 날짜 확인", "body": "연습용으로 만든 가상 동아리의 보고서입니다. "}
LINK = "https://www.coding-now.com/"

# 큰 사진 대신 쓸 그림(사진처럼 색이 고르지 않게 — 압축이 잘 안 되는 그림)
PIC = os.path.join(WORK, "photo.png")
random.seed(7)
im = Image.new("RGB", (2400, 1600))
d = ImageDraw.Draw(im)
for y in range(1600):
    d.line([(0, y), (2400, y)], fill=(40 + y // 12, 90 + y // 20, 160 - y // 16))
for _ in range(1800):
    x, y, r = random.randint(0, 2400), random.randint(0, 1600), random.randint(6, 60)
    d.ellipse([x - r, y - r, x + r, y + r], fill=(random.randint(0, 255), random.randint(0, 255), random.randint(0, 255)))
im = im.filter(ImageFilter.GaussianBlur(1.2))
im.save(PIC)
R["picture"] = {"px": im.size, "png_bytes": os.path.getsize(PIC)}


def pdf_info(p):
    with pymupdf.open(p) as doc:
        fonts = {}
        for pg in doc:
            for f in pg.get_fonts(full=True):
                # (xref, ext, type, basefont, name, encoding, ...) — ext 'n/a' 면 글꼴 파일이 안 들어 있음
                fonts[f[3]] = f[1]
        imgs = []
        for pg in doc:
            for x in pg.get_images(full=True):
                imgs.append([x[2], x[3]])
        links = sum(1 for pg in doc for ln in pg.get_links() if ln.get("uri"))
        text = "".join(pg.get_text() for pg in doc)
        return {"bytes": os.path.getsize(p), "pages": doc.page_count, "toc": doc.get_toc(), "uri_links": links,
                "fonts": fonts, "images_px": imgs, "has_comment_text": T["comment"] in text,
                "metadata_title": doc.metadata.get("title"), "metadata_author": doc.metadata.get("author"), "pdfa_hint": "pdfaid" in (doc.xref_get_key(doc.pdf_catalog(), "Metadata")[1] or "") or None}


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["word_build"] = w.Build
    doc = w.Documents.Add()
    sel = w.Selection

    def para(text, style=None):
        sel.Style = doc.Styles(style) if style else doc.Styles(-1)  # -1 = wdStyleNormal
        sel.TypeText(text)
        sel.TypeParagraph()

    para(T["title"], -63)  # wdStyleTitle
    para(T["h"][0], -2)    # 제목 1
    for _ in range(6):
        para(T["body"] * 4)
    sel.TypeText(("More notices on the " if EN else "공지는 ") )
    start = sel.Range.Start
    sel.TypeText(T["link_text"])
    rng = doc.Range(start, sel.Range.Start)
    doc.Hyperlinks.Add(rng, LINK)
    sel.EndKey(6)
    sel.TypeText("." if EN else "에서 봅니다.")
    sel.TypeParagraph()
    para(T["h2"][0], -3)   # 제목 2
    t = doc.Tables.Add(sel.Range, 3, 3, 1, 0)
    cells = [["Week", "Event", "People"], ["1", "Book swap", "12"], ["2", "Clean-up", "9"]] if EN else \
            [["주", "활동", "인원"], ["1주", "책 바꿔 읽기", "12"], ["2주", "공원 정리", "9"]]
    for i in range(3):
        for j in range(3):
            t.Cell(i + 1, j + 1).Range.Text = cells[i][j]
    sel.EndKey(6)
    sel.TypeParagraph()
    para(T["h2"][1], -3)
    for _ in range(3):
        para(T["body"] * 3)
    para(T["h"][1], -2)
    sel.InlineShapes.AddPicture(PIC)
    sel.TypeParagraph()
    para(T["h"][2], -2)
    for _ in range(8):
        para(T["body"] * 4)
    # 메모: 제목에
    cm = doc.Comments.Add(doc.Paragraphs(1).Range, T["comment"])
    cm.Author = "CODINGNOW"  # 이 PC 사용자 이름 · 이니셜이 메모 풍선에 찍히지 않게
    cm.Initial = "CN"
    for k in ("Author", "Last Author"):
        doc.BuiltinDocumentProperties(k).Value = "CODINGNOW"
    doc.Repaginate()
    R["doc"] = {"pages": doc.ComputeStatistics(2), "comments": doc.Comments.Count, "hyperlinks": doc.Hyperlinks.Count,
                "headings": sum(1 for p in doc.Paragraphs if p.OutlineLevel in (1, 2)),
                "font": doc.Styles(-1).Font.Name}
    docx = os.path.join(WORK, "report.docx")
    if os.path.exists(docx):
        os.remove(docx)
    doc.SaveAs2(docx, 16)
    R["docx_bytes"] = os.path.getsize(docx)

    out = {}
    pA = os.path.join(WORK, "A_saveas.pdf")
    doc.SaveAs2(pA, 17)
    out["A_saveas"] = pdf_info(pA)
    doc = w.Documents.Open(docx)  # SaveAs2 PDF 뒤엔 문서 이름이 PDF 로 바뀌어 다시 연다

    def exp(name, **kw):
        p = os.path.join(WORK, name + ".pdf")
        args = dict(OutputFileName=p, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, Range=0, From=1, To=1,
                    Item=0, IncludeDocProps=True, KeepIRM=True, CreateBookmarks=0, DocStructureTags=True,
                    BitmapMissingFonts=True, UseISO19005_1=False)
        args.update(kw)
        doc.ExportAsFixedFormat(**args)
        out[name] = pdf_info(p)

    exp("B_standard")
    exp("C_minimum", OptimizeFor=1)
    exp("D_heading_bookmarks", CreateBookmarks=1)
    exp("E_with_markup", Item=7)
    exp("F_page2_only", Range=3, From=2, To=2)
    exp("G_pdfa", UseISO19005_1=True)
    R["pdf"] = out
    doc.Close(0)
finally:
    w.Quit(0)
json.dump(R, open(os.path.join(HERE, "measure_en.json" if EN else "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k, v in R["pdf"].items():
    print(k, v["bytes"], "쪽", v["pages"], "링크", v["uri_links"], "목차", len(v["toc"]), "메모", v["has_comment_text"], "그림", v["images_px"], "글꼴", v["fonts"])
print(R["doc"], R["picture"], R["docx_bytes"])
