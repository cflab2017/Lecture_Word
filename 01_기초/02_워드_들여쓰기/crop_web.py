"""shots/ 원본에서 강의 페이지용 그림을 잘라 coding-now public/lectures/word/ 에 넣는다(제목 표시줄·계정 사진은 잘라 냄)."""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
S = lambda n: Image.open(os.path.join(HERE, "shots", n))
out = {}
out["basic-02-doc.png"] = S("s1_first.png").crop((8, 182, 1005, 525))          # 눈금자 + 세 단락(편집 기호 표시)
a = S("s1_first.png").crop((180, 180, 600, 215)); b = S("s2_hanging.png").crop((180, 180, 600, 215))
r = Image.new("RGB", (420, 78), "white"); r.paste(a, (0, 0)); r.paste(b, (0, 43))
out["basic-02-rulers.png"] = r                                                  # 위: 첫 줄 2글자 · 아래: 둘째 줄 이하 2글자
rb = S("s1_first.png").crop((405, 82, 600, 178)); d = ImageDraw.Draw(rb)
d.rectangle((133, 2, 186, 26), outline=(220, 38, 38), width=3)                  # 홈 → 단락의 '내어쓰기' · '들여쓰기' 단추
out["basic-02-ribbon.png"] = rb
out["basic-02-dialog.png"] = S("s4_dialog_hanging.png").crop((0, 0, 463, 315))  # 단락 창 — 첫 줄: 둘째 줄 이하, 값 2 글자
for n, im in out.items():
    im.save(os.path.join(DST, n), optimize=True); print(n, im.size)
