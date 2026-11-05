"""워드 기초 2강 실습·정답 파일 — 가상 '도서관 이용 안내'. 실습 = 들여쓰기 없음(둘째 단락은 Tab 으로 민 상태),
정답 = 본문 단락 첫 줄 2글자 · 번호 단락 둘째 줄 이하 2글자. 영어판(practice·answer)은 0.5in · 0.25in.
실행: python make_files.py  (pywin32 필요) — 저장 뒤 작성자 정보는 strip_author.py 로 CODINGNOW 로 바꾼다."""
import os, sys
import pythoncom
import win32com.client as win32
from measure_text import P1, P2, ITEM

HERE = os.path.dirname(os.path.abspath(__file__))
ITEMS_KO = [ITEM, "2.\t연장: 대출 중인 책은 반납일 전에 홈페이지나 안내 데스크에서 한 번, 7일까지 늘릴 수 있습니다."]
EN = {
    "title": "Library notice for November",
    "p1": ("During November the second-floor reading room is open from 9 a.m. to 9 p.m. every day. Seats fill up quickly "
           "during exam weeks, so please come early if you need a desk."),
    "p2": ("You can borrow up to five books for fourteen days. If a book comes back late, borrowing is paused for as many "
           "days as it was late. You can renew once, for up to seven days."),
    "items": ["1.\tReturns: the self-service return box on the first floor is open 24 hours, including the days the library is closed.",
              "2.\tRenewals: renew a book once before its due date, online or at the front desk, for up to seven more days."],
}


def build(w, title, p1, p2, items, answer, en):
    doc = w.Documents.Add()
    doc.Content.InsertAfter("\r".join([title, p1, ("" if answer else "\t") + p2] + items))
    for i in range(1, doc.Paragraphs.Count + 1):
        doc.Paragraphs(i).Reset(); doc.Paragraphs(i).SpaceAfter = 8
    doc.Paragraphs(1).Range.Font.Bold = True; doc.Paragraphs(1).Range.Font.Size = 14
    if answer:
        for k in (2, 3):
            if en:
                doc.Paragraphs(k).FirstLineIndent = 36.0
            else:
                doc.Paragraphs(k).CharacterUnitFirstLineIndent = 2
        for k in range(4, doc.Paragraphs.Count + 1):
            if en:
                doc.Paragraphs(k).LeftIndent = 18.0; doc.Paragraphs(k).FirstLineIndent = -18.0
            else:
                doc.Paragraphs(k).CharacterUnitFirstLineIndent = -2
    return doc


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False; w.DisplayAlerts = 0
    for name, answer in (("실습.docx", False), ("정답.docx", True)):
        d = build(w, "11월 도서관 이용 안내", P1, P2, ITEMS_KO, answer, False)
        d.SaveAs2(os.path.join(HERE, name), 16); d.Close(0)
    for name, answer in (("practice.docx", False), ("answer.docx", True)):
        d = build(w, EN["title"], EN["p1"], EN["p2"], EN["items"], answer, True)
        d.SaveAs2(os.path.join(HERE, name), 16); d.Close(0)
    print("saved")
finally:
    w.Quit(0)
