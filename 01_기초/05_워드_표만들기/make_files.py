"""워드 기초 5강(표 만들기) 실습 · 정답 파일 — 이 PC 워드를 따로 띄워(DispatchEx) 만든다. 내용은 연습용 가상 회원 명단(전화번호는 010-0000-xxxx 꼴 가짜).
실습: 탭으로 나눈 명단 45줄(번호 · 이름 · 연락처) — 2쪽을 넘긴다. 표가 아니라 글.
정답: 같은 명단을 '텍스트를 표로 변환'(탭) → 표 구분선 스타일 · 첫 행 '페이지마다 머리글 행 반복' · '내용에 자동으로 맞춤'.
argv 'en' 이면 영어(practice.docx · answer.docx). 작성자 · 마지막 저장자는 strip_author 로 CODINGNOW.
실행: python make_files.py [en]  (pywin32 · pymupdf)"""
import os, sys
import pythoncom
import win32com.client as win32
import fitz

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "Book club members" if EN else "독서 모임 회원 명단"
HEAD = ["No", "Name", "Phone"] if EN else ["번호", "이름", "연락처"]
NAMES = (["Kim", "Lee", "Park", "Choi", "Jung", "Kang", "Cho", "Yoon", "Jang"] if EN
         else ["김민준", "이서연", "박지후", "최하은", "정도윤", "강수아", "조예준", "윤지아", "장시우"])
N = 45
rows = [HEAD] + [[str(i), NAMES[(i - 1) % len(NAMES)] + ("" if i <= len(NAMES) else f" {(i - 1) // len(NAMES) + 1}"), f"010-0000-{i:04d}"] for i in range(1, N + 1)]

pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
out = {}
try:
    w.Visible = False
    w.DisplayAlerts = 0
    for kind in ("practice", "answer"):
        d = w.Documents.Add()
        d.Range().Text = TITLE + "\r" + "\r".join("\t".join(r) for r in rows)
        d.Paragraphs(1).Range.Font.Size = 14
        d.Paragraphs(1).Range.Font.Bold = True
        if kind == "answer":
            rng = d.Range(d.Paragraphs(2).Range.Start, d.Range().End)
            # DefaultTableBehavior=1(wdWord9TableBehavior) = 화면 [텍스트를 표로 변환] 처럼 '표 구분선' 스타일
            t = rng.ConvertToTable(Separator=1)
            # COM 은 '표준 표'(테두리 없음)로 만든다(11/7) — 화면 [삽입 → 표] 의 기본인 '표 구분선'(measure.json insert_default)으로 맞춘다
            t.Style = "표 구분선"
            out["answer_style"] = str(t.Style)
            t.Rows(1).HeadingFormat = True
            t.Rows(1).Range.Font.Bold = True
            t.AutoFitBehavior(2)
            t.AutoFitBehavior(1)
        name = (kind + ".docx") if EN else ("실습.docx" if kind == "practice" else "정답.docx")
        p = os.path.join(HERE, name)
        if os.path.exists(p):
            os.remove(p)
        d.SaveAs2(p, 16)
        pdfp = os.path.join(os.environ["TEMP"], f"word05_{kind}.pdf")
        d.ExportAsFixedFormat(pdfp, 17)
        with fitz.open(pdfp) as f:
            out[name] = {"pages": f.page_count, "pages_with_head": [i + 1 for i in range(f.page_count) if HEAD[2] in f[i].get_text()],
                         "tables": d.Tables.Count}
        d.Close(0)
finally:
    w.Quit()
print(out)
