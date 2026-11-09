"""shots/ 원본에서 강의 페이지용 그림 — 창 그림은 제목 표시줄(계정 사진)을 잘라 내고 빨간 테두리로 볼 곳 표시(11/6 규칙).
결과는 coding-now public/lectures/word/basic-05-*.png. 영어판은 화면 그림 대신 애니메이션(이 PC 워드 메뉴가 한국어)."""
import os
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
DST = r"D:\WebSite\WebSite\Website_Codingnow\public\lectures\word"
TOP = 44


def win(src, dst, boxes, bottom=None):
    im = Image.open(os.path.join(HERE, "shots", src)).convert("RGB")
    d = ImageDraw.Draw(im)
    for b in boxes:
        d.rectangle(b, outline=(220, 30, 30), width=3)
    im = im.crop((0, TOP, im.width, bottom or im.height))
    im.save(os.path.join(DST, dst))
    print(dst, im.size)


def plain(src, dst):
    im = Image.open(os.path.join(HERE, "shots", src)).convert("RGB")
    im.save(os.path.join(DST, dst))
    print(dst, im.size)


win("s1_practice_tabs.png", "basic-05-tabs.png", [(298, 352, 560, 418)], bottom=620)
plain("s2_text_to_table.png", "basic-05-text-to-table.png")
win("s3_answer_page2.png", "basic-05-repeat-header.png", [(300, 740, 540, 770)])
plain("s4_row_repeat.png", "basic-05-row-repeat.png")
win("s5_comma_trap.png", "basic-05-comma-trap.png", [(301, 340, 909, 370)], bottom=520)
