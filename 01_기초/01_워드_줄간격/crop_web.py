"""shots/*.png(원본 화면)을 강의 페이지용 그림으로 자른다 — 제목 표시줄(계정 사진)은 잘라 낸다.
결과는 인자로 받은 폴더(사이트의 public/lectures/word)에 basic-01-*.png 로 저장. 실행: python crop_web.py <저장 폴더>"""
import os, sys
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
S = os.path.join(HERE, "shots")
OUT = sys.argv[1]
RED = (220, 38, 38)
FONT = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 15)


def box(d, xy, n=None):
    d.rounded_rectangle(xy, radius=4, outline=RED, width=3)
    if n:
        x, y = xy[2], xy[1]  # 번호는 상자 오른쪽 위
        d.ellipse((x - 11, y - 11, x + 9, y + 9), fill=RED)
        d.text((x - 1, y - 1), str(n), fill="white", font=FONT, anchor="mm")


# ① 리본 — 홈 탭 단락 그룹: 줄 간격 단추(1) · 단락 대화상자 열기(2)
im = Image.open(os.path.join(S, "s1_compare.png")).crop((0, 48, 620, 182)).convert("RGB")
d = ImageDraw.Draw(im)
box(d, (544, 109 - 48, 582, 132 - 48), 1)
box(d, (586, 162 - 48, 603, 178 - 48), 2)
im = im.resize((im.width * 3 // 2, im.height * 3 // 2), Image.LANCZOS)
im.save(os.path.join(OUT, "basic-01-ribbon.png"), optimize=True)

# ② 같은 글, 줄 간격만 다르게
im = Image.open(os.path.join(S, "s1_compare.png")).crop((170, 300, 800, 745))
im.save(os.path.join(OUT, "basic-01-compare.png"), optimize=True)

# ③ 단락 대화상자 — 간격 칸들
im = Image.open(os.path.join(S, "s2_dialog.png")).convert("RGB")
d = ImageDraw.Draw(im)
box(d, (26, 337, 190, 390), 1)      # 단락 앞 · 단락 뒤
box(d, (238, 337, 450, 390), 2)     # 줄 간격 · 값
box(d, (26, 394, 300, 414), 3)      # 같은 스타일 단락 사이 공백 없음
box(d, (26, 414, 375, 436), 4)      # 문서 격자 맞춤
box(d, (128, 644, 278, 671), 5)     # 기본값으로 설정
im.save(os.path.join(OUT, "basic-01-dialog.png"), optimize=True)

# ④ 줄 격자 함정 — 위: 격자 칸 켜짐(두 배로 벌어짐), 아래: 칸을 끈 뒤
a = Image.open(os.path.join(S, "s3_grid.png")).crop((170, 300, 800, 545))
b = Image.open(os.path.join(S, "s3b_grid_fixed.png")).crop((170, 300, 800, 505))
lab_h = 30
c = Image.new("RGB", (a.width, lab_h + a.height + lab_h + b.height), "white")
d = ImageDraw.Draw(c)
d.rectangle((0, 0, c.width, lab_h), fill=(254, 226, 226))
d.text((10, 6), "문서 격자 칸 켜짐 — 아래 구역(줄 격자)만 줄 간격 1.0 이 두 배로", fill=(153, 27, 27), font=FONT)
c.paste(a, (0, lab_h))
y = lab_h + a.height
d.rectangle((0, y, c.width, y + lab_h), fill=(220, 252, 231))
d.text((10, y + 6), "단락 → '[페이지 설정]에서 지정된 문자 수에 맞춰 문자 간격 조정' 끈 뒤", fill=(22, 101, 52), font=FONT)
c.paste(b, (0, y + lab_h))
c.save(os.path.join(OUT, "basic-01-grid.png"), optimize=True)
for f in sorted(os.listdir(OUT)):
    if f.startswith("basic-01"):
        print(f, Image.open(os.path.join(OUT, f)).size)

# ⑤ 줄 간격 단추를 누르면 나오는 메뉴(실제 화면, 접근성 정보로 연 것 — menu_names.py)
im = Image.open(os.path.join(S, "s5_menu.png")).convert("RGB")
im = im.resize((im.width * 3 // 2, im.height * 3 // 2), Image.LANCZOS)
im.save(os.path.join(OUT, "basic-01-menu.png"), optimize=True)

# ⑥ 단락 대화상자 '줄 간격' 칸의 여섯 가지(dialog_rules.py)
im = Image.open(os.path.join(S, "s6_rules.png")).convert("RGB")
im.save(os.path.join(OUT, "basic-01-rules.png"), optimize=True)
print("menu", "rules")
