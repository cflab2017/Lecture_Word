"""shots/ 원본과 measure.py 가 만든 PDF 에서 강의 페이지용 그림을 만들어 coding-now public/lectures/word/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄).
usage: python crop_web.py <measure.py 가 만든 wd03_ 임시 폴더(KR)> [영어판 임시 폴더 — 주면 basic-03-pages-en.png 도]"""
import os, sys
import pymupdf
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
TMP = sys.argv[1]
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
out = {}
out["basic-03-footer.png"] = S("s2_footer_edit.png").crop((8, 44, 1192, 845))     # 머리글/바닥글 탭 + 2쪽 바닥글 '1' + 상태 표시줄 2/4 페이지
out["basic-03-dialog.png"] = S("s3_format_dialog.png")                             # 페이지 번호 서식 — 시작 번호 0
# 정답 문서 PDF 1~4쪽 아래 40mm — 표지는 번호 없음, 2쪽부터 1 · 2 · 3
k = 72 / 25.4
font = ImageFont.truetype(r"C:\Windows\Fonts\malgun.ttf", 22)


def pages(tmp, labels):
    doc = pymupdf.open(os.path.join(tmp, "answer.pdf"))
    tiles = []
    for i in range(4):
        pg = doc[i]; h = pg.rect.height / k
        pix = pg.get_pixmap(dpi=150, clip=pymupdf.Rect(40 * k, (h - 40) * k, 170 * k, (h - 10) * k))
        tiles.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    W, H = tiles[0].size
    im = Image.new("RGB", (W * 2 + 24, (H + 40) * 2 + 12), "white")
    d = ImageDraw.Draw(im)
    for i, t in enumerate(tiles):
        x = (i % 2) * (W + 24); y = (i // 2) * (H + 52)
        d.text((x + 4, y + 4), labels[i], fill=(71, 85, 105), font=font)
        im.paste(t, (x, y + 36))
        d.rectangle([x, y + 36, x + W - 1, y + 36 + H - 1], outline=(148, 163, 184), width=2)
    return im


out["basic-03-pages.png"] = pages(TMP, ["1쪽(표지) 아래", "2쪽 아래", "3쪽 아래", "4쪽 아래"])
if len(sys.argv) > 2:
    out["basic-03-pages-en.png"] = pages(sys.argv[2], ["Page 1 (cover), bottom", "Page 2, bottom", "Page 3, bottom", "Page 4, bottom"])
for n, m in out.items():
    m.save(os.path.join(DST, n), optimize=True); print(n, m.size)
