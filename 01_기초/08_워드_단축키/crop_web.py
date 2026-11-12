"""강의 페이지용 그림 — 정답.docx 를 숨긴 워드로 PDF 로 내보내 그 쪽을 그림으로(11/12: 바탕화면을 다른 세션이 쓰는 중이라 보이는 워드 창을 띄우지 않음).
실습 · 정답 두 쪽의 위쪽을 나란히 놓아 단축키로 바뀐 곳(제목 · 내어쓰기 · Shift+Enter 줄)을 빨간 테두리로.
결과 coding-now public/lectures/word/basic-08-before-after.png. 실행: python crop_web.py (pywin32 · pymupdf · pillow)"""
import io, os, tempfile
import pymupdf
import pythoncom
import win32com.client as win32
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\content_hub\homepage\public\word_tmp"
DST = r"D:\WebSite\content_hub\homepage\public\lectures\word"
Z = 1.5


def to_pdf(name):
    out = os.path.join(tempfile.mkdtemp(), name.replace(".docx", ".pdf"))
    doc = wd.Documents.Open(os.path.join(HERE, name), False, True, False)
    doc.ExportAsFixedFormat(out, 17)
    doc.Close(False)
    return out


def top_png(pdf, boxes_fn):
    d = pymupdf.open(pdf)
    pg = d[0]
    pix = pg.get_pixmap(matrix=pymupdf.Matrix(Z, Z))
    im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    words = pg.get_text("words")
    dr = ImageDraw.Draw(im)
    for r in boxes_fn(words):
        x0, y0, x1, y1 = (v * Z for v in r)
        dr.rectangle((x0 - 5, y0 - 4, x1 + 5, y1 + 4), outline=(220, 30, 30), width=3)
    last = max(w[3] for w in words)
    left = min(w[0] for w in words)
    return im.crop((int((left - 24) * Z), int(40 * Z), int((left + 400) * Z), int((last + 20) * Z)))


def lines(words, pred):
    """줄(block · line)별로 묶어 pred(첫 낱말 글자) 가 맞는 줄의 사각형."""
    from collections import defaultdict
    g = defaultdict(list)
    for w in words:
        g[(w[5], w[6])].append(w)
    out = []
    for ws in g.values():
        if pred(" ".join(x[4] for x in ws)):
            out.append((min(x[0] for x in ws), min(x[1] for x in ws), max(x[2] for x in ws), max(x[3] for x in ws)))
    return out


pythoncom.CoInitialize()
wd = win32.DispatchEx("Word.Application")
try:
    wd.Visible = False
    wd.DisplayAlerts = 0
    a = to_pdf("실습.docx")
    b = to_pdf("정답.docx")
finally:
    wd.Quit()
im_a = top_png(a, lambda ws: [])
im_b = top_png(b, lambda ws: lines(ws, lambda t: t.startswith("사진") or t.startswith("비가") or t.startswith("일시")))
pad, head = 24, 40
W = im_a.width + im_b.width + pad * 3
H = max(im_a.height, im_b.height) + head + pad
out = Image.new("RGB", (W, H), "white")
try:
    f = ImageFont.truetype("malgunbd.ttf", 22)
except OSError:
    f = ImageFont.load_default()
d = ImageDraw.Draw(out)
d.text((pad, 10), "실습.docx (단축키 누르기 전)", fill=(60, 60, 60), font=f)
d.text((im_a.width + pad * 2, 10), "정답.docx (Ctrl+B · Ctrl+E · Ctrl+] · Ctrl+T · Shift+Enter)", fill=(60, 60, 60), font=f)
out.paste(im_a, (pad, head))
out.paste(im_b, (im_a.width + pad * 2, head))
d.rectangle((pad - 1, head - 1, pad + im_a.width, head + im_a.height), outline=(200, 200, 200), width=1)
d.rectangle((im_a.width + pad * 2 - 1, head - 1, im_a.width + pad * 2 + im_b.width, head + im_b.height), outline=(200, 200, 200), width=1)
out.save(os.path.join(DST, "basic-08-before-after.png"))
print("basic-08-before-after.png", out.size)
