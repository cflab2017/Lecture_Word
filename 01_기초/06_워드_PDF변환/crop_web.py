"""shots/ 원본에서 강의 페이지용 그림 — 창 그림은 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙).
[다른 이름으로 저장] 창은 왼쪽 폴더 목록(이 PC 폴더 이름)이 보이지 않게 아래쪽(파일 이름 · 형식 · 최적화 · 옵션)만 쓴다.
PDF 비교 그림은 measure.py 가 만든 PDF 첫 쪽(KR · EN 따로). 결과는 coding-now public/lectures/word/basic-06-*.png."""
import os
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
TOP = 44
RED = (220, 30, 30)
FONT = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 26)


def boxed(im, boxes):
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=RED, width=3)
    return im


s1 = Image.open(os.path.join(HERE, "shots", "s1_practice.png")).convert("RGB")
boxed(s1, [(878, 180, 1185, 360), (232, 252, 680, 320)]).crop((0, TOP, s1.width, s1.height)).save(os.path.join(DST, "basic-06-practice.png"))
s2 = Image.open(os.path.join(HERE, "shots", "s2_save_as_pdf.png")).convert("RGB").crop((0, 730, 720, 971))
boxed(s2, [(18, 36, 712, 62), (72, 100, 548, 172)]).save(os.path.join(DST, "basic-06-save-as-pdf.png"))
s3 = Image.open(os.path.join(HERE, "shots", "s3_pdf_options.png")).convert("RGB")
boxed(s3, [(10, 176, 335, 236), (10, 254, 335, 314)]).save(os.path.join(DST, "basic-06-pdf-options.png"))
for p in ("basic-06-practice.png", "basic-06-save-as-pdf.png", "basic-06-pdf-options.png"):
    print(p, Image.open(os.path.join(DST, p)).size)


def compare(lang):
    work = os.path.join(os.environ["TEMP"], "claude", "word06" + ("_en" if lang == "en" else ""))
    labels = (("Save As → PDF (as is)", "Export: Document only") if lang == "en" else ("다른 이름으로 저장 → PDF 그대로", "옵션에서 '문서'만"))
    pages = []
    for name in ("A_saveas", "B_standard"):
        with pymupdf.open(os.path.join(work, name + ".pdf")) as doc:
            pix = doc[0].get_pixmap(dpi=80, clip=pymupdf.Rect(0, 0, 595, 520))
            p = os.path.join(work, name + "_top.png")
            pix.save(p)
            pages.append(Image.open(p).convert("RGB"))
    gap, head = 24, 48
    W = sum(i.width for i in pages) + gap * 3
    H = max(i.height for i in pages) + head + gap
    out = Image.new("RGB", (W, H), (244, 244, 244))
    d = ImageDraw.Draw(out)
    x = gap
    for im, lab in zip(pages, labels):
        d.text((x, 10), lab, font=FONT, fill=(30, 30, 30))
        out.paste(im, (x, head))
        d.rectangle((x - 1, head - 1, x + im.width, head + im.height), outline=(180, 180, 180), width=1)
        x += im.width + gap
    name = "basic-06-pdf-compare" + ("-en" if lang == "en" else "") + ".png"
    out.save(os.path.join(DST, name))
    print(name, out.size)


compare("ko")
compare("en")
