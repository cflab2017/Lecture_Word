"""shots/ 원본에서 강의 페이지용 그림 — 창 그림은 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙).
s4 는 제목 표시줄의 '호환성 모드' 글자를 보여 줘야 해서 자르지 않고 계정 사진 자리만 바탕색으로 덮는다.
한글 '호환 문서' 창은 한글 편집 창이 열리지 않아(구독 없음) 창만 있는 그림이다.
결과는 coding-now public/lectures/word/basic-07-*.png (KR 만)."""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
TOP = 44
RED = (220, 30, 30)


def boxed(im, boxes):
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=RED, width=3)
    return im


def shot(name):
    return Image.open(os.path.join(HERE, "shots", name)).convert("RGB")


def save(im, name):
    im.save(os.path.join(DST, name))
    print(name, im.size)


# 1 [파일 변환] 창 — 목록에서 워드가 고른 '인코딩된 텍스트' · 맨 아래 '아래아 한글 2.0-97'
im = shot("s1_convert_dialog.png")
save(boxed(im, [(463, 340, 718, 357), (463, 475, 718, 492)]).crop((0, TOP, im.width, im.height)), "basic-07-convert-dialog.png")
# 2 그냥 열기 — 알아볼 수 없는 글자 · 글꼴 굴림체 · 상태 표시줄 글자 수
im = shot("s2_hwp_garbage.png")
save(boxed(im, [(328, 262, 772, 302), (122, 84, 296, 106), (12, 824, 232, 845)]).crop((0, TOP, im.width, im.height)), "basic-07-hwp-garbage.png")
# 3 아래아 한글 2.0-97 변환기로 열기 — 오류 창
im = shot("s3_ara20_error.png")
save(boxed(im, [(453, 345, 746, 507)]).crop((0, TOP, im.width, im.height)), "basic-07-ara20-error.png")
# 4 한글이 만든 DOCX — 제목 표시줄 '호환성 모드' · 글꼴 함초롬바탕 10 · 넓은 줄 간격
im = shot("s4_hwp_to_docx.png")
d = ImageDraw.Draw(im)
d.rectangle((995, 6, 1052, 44), fill=im.getpixel((975, 25)))  # 계정 사진 덮기
save(boxed(im, [(415, 13, 512, 37), (122, 84, 345, 106), (318, 380, 890, 476)]), "basic-07-hwp-to-docx.png")
# 5 한글 2024 '호환 문서' 창(OOXML 로 저장할 때)
im = shot("s5_hangul_compat.png")
save(boxed(im, [(34, 46, 370, 104)]), "basic-07-hangul-compat.png")
