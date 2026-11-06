"""shots/ 원본과 실습·정답 파일의 PDF 에서 강의 페이지용 그림을 만들어 coding-now public/lectures/word/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄).
영어판은 화면 그림 대신 practice.docx · answer.docx 를 PDF 로 바꿔 쪽을 나란히 놓은 그림만 쓴다(이 PC 워드 메뉴가 한국어).
실행: python crop_web.py  (pywin32 · pymupdf · Pillow 필요)"""
import os, tempfile
import pymupdf
import win32com.client as win32
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
out = {
    "basic-04-enters.png": S("s2_practice_enters.png").crop((8, 44, 1192, 845)),   # 편집 기호 켬 — 첫 단락 뒤 줄줄이 ¶ · 상태 표시줄 1/3 페이지
    "basic-04-break.png": S("s3_answer_break.png").crop((8, 44, 1192, 560)),       # 정답 — '페이지 나누기' 표시 한 줄
    "basic-04-pages.png": S("s1_practice_pages.png").crop((40, 255, 1160, 775)),   # 실습 세 쪽 — 3쪽에 ¶ 하나
    "basic-04-end1pt.png": S("s4_answer_end_1pt.png").crop((8, 44, 1192, 845)),    # 표 뒤 단락 글꼴 크기 1 · 2/2 페이지
}

# 실습(3쪽) · 정답(2쪽)을 PDF 로 바꿔 나란히
tmp = tempfile.mkdtemp(prefix="wd04crop_")
w = win32.DispatchEx("Word.Application")
pdfs = {}
try:
    w.Visible = False
    w.DisplayAlerts = 0
    for nm in ("실습.docx", "정답.docx", "practice.docx", "answer.docx"):
        d = w.Documents.Open(os.path.join(HERE, nm), ReadOnly=True)
        p = os.path.join(tmp, nm.replace(".docx", ".pdf"))
        d.ExportAsFixedFormat(p, 17)
        d.Close(0)
        pdfs[nm] = p
finally:
    w.Quit(0)

font = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 24)
small = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 20)


def tiles(pdf):
    doc = pymupdf.open(pdf)
    ims = []
    for pg in doc:
        pix = pg.get_pixmap(dpi=40)
        ims.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    doc.close()
    return ims


def compare(prac, ans, labels):
    a, b = tiles(prac), tiles(ans)
    W, H = a[0].size
    gap, pad, head = 18, 24, 44
    width = pad * 2 + 3 * W + 2 * gap
    im = Image.new("RGB", (width, pad + (head + H + 36) * 2 + pad), "white")
    d = ImageDraw.Draw(im)
    for row, (ims, lab) in enumerate(((a, labels[0]), (b, labels[1]))):
        y = pad + row * (head + H + 36)
        d.text((pad, y), lab, fill=(30, 41, 59), font=font)
        for i, t in enumerate(ims):
            x = pad + i * (W + gap)
            im.paste(t, (x, y + head))
            d.rectangle([x, y + head, x + W - 1, y + head + H - 1], outline=(148, 163, 184), width=2)
            d.text((x + 4, y + head + H + 6), labels[2].format(n=i + 1), fill=(100, 116, 139), font=small)
    return im


out["basic-04-compare.png"] = compare(pdfs["실습.docx"], pdfs["정답.docx"], ["실습 — 3쪽(마지막 쪽은 빈 쪽)", "정답 — 2쪽", "{n}쪽"])
out["basic-04-compare-en.png"] = compare(pdfs["practice.docx"], pdfs["answer.docx"], ["Practice — 3 pages (last one blank)", "Answer — 2 pages", "Page {n}"])
for n, m in out.items():
    m.save(os.path.join(DST, n), optimize=True)
    print(n, m.size)
