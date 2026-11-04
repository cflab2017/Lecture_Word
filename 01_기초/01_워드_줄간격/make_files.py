"""워드 기초 1강 실습·정답 파일을 만든다(KR 실습.docx·정답.docx, EN practice.docx·answer.docx).
워드를 따로(보이지 않게) 띄워 새 문서로 만들고 저장한다. 내용은 모두 연습용 가상 글이다. Normal.dotm 은 건드리지 않는다.
실행: python make_files.py  (pip install pywin32)"""
import os
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
FMT_DOCX = 16  # wdFormatXMLDocument

KO_TITLE = "11월 팀 회의 안내"
KO_PARAS = [
    "이번 주 팀 회의는 화요일 오전 10시에 3층 대회의실에서 합니다. 지난달 매출 보고와 다음 분기 계획을 나누고, 각 팀은 진행 상황을 5분씩 발표합니다.",
    "발표 자료는 월요일 오후 6시까지 팀 공유 폴더에 올려 주세요. 파일 이름은 '팀이름_발표.pptx'로 맞춰 주시면 순서대로 띄우기 쉽습니다.",
    "회의가 끝나면 12시부터 같은 층 휴게실에서 점심을 함께합니다. 못 오시는 분은 금요일까지 총무팀에 알려 주세요.",
    "주차는 지하 2층을 이용하시고, 방문 등록은 1층 안내 데스크에서 하면 됩니다.",
]
EN_TITLE = "Why Short Breaks Help You Study"
EN_PARAS = [
    "Most students try to study for hours without stopping, but research on attention suggests that short, planned breaks help you remember more. A ten-minute walk or a glass of water gives your brain time to rest before the next session.",
    "One simple method is to study for twenty-five minutes and then rest for five. After four rounds, take a longer break of about twenty minutes. Keeping a timer on your desk makes the routine easier to follow.",
    "Breaks work best when you step away from the screen. Checking your phone during a break often turns five minutes into twenty, so leave it in another room until the timer rings.",
    "Try the method for one week and write down how much you finished each day. Many students find that they get more done in less time.",
]


def build(w, title, paras, mode, path, lang):
    doc = w.Documents.Add()
    doc.Content.Text = title + "\r" + "\r".join(paras)
    if lang == "en":
        doc.Content.LanguageID = 1033
        doc.Content.Font.Name = "Times New Roman"
        doc.Content.Font.Size = 12
    t = doc.Paragraphs(1)
    t.Range.Font.Bold = True
    t.Range.Font.Size = 16 if lang == "ko" else 14
    t.Alignment = 0 if lang == "ko" else 1
    body = [doc.Paragraphs(i) for i in range(2, doc.Paragraphs.Count + 1)]
    if mode == "practice":
        if lang == "ko":
            # 줄 간격이 문단마다 제각각 + 문단 사이를 빈 줄(Enter 두 번)로 띄움
            rules = [(0, None), (2, None), (0, None), (5, 36)]
            for p, (rule, val) in zip(body, rules):
                p.LineSpacingRule = rule
                if val:
                    p.LineSpacing = val
                p.SpaceAfter = 0
            for p in reversed(body[:-1]):
                p.Range.InsertParagraphAfter()
        else:
            # 학교 과제 형식: 줄 간격 1.0 + 단락 뒤 8pt (두 줄 간격으로 바꿔야 함)
            for p in body:
                p.LineSpacingRule = 0
                p.SpaceAfter = 8
    else:
        if lang == "ko":
            for p in body:
                p.LineSpacingRule = 5
                p.LineSpacing = 13.8  # 1.15 배수
                p.SpaceAfter = 10
        else:
            for p in doc.Paragraphs:
                p.LineSpacingRule = 2  # Double
                p.SpaceAfter = 0
                p.SpaceBefore = 0
            for p in body:
                p.FirstLineIndent = 36  # 0.5 inch
    doc.BuiltInDocumentProperties("Author").Value = "CODINGNOW"
    doc.BuiltInDocumentProperties("Title").Value = title
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, FMT_DOCX)
    doc.Close(0)
    print("saved", os.path.basename(path))


w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    build(w, KO_TITLE, KO_PARAS, "practice", os.path.join(HERE, "실습.docx"), "ko")
    build(w, KO_TITLE, KO_PARAS, "answer", os.path.join(HERE, "정답.docx"), "ko")
    build(w, EN_TITLE, EN_PARAS, "practice", os.path.join(HERE, "practice.docx"), "en")
    build(w, EN_TITLE, EN_PARAS, "answer", os.path.join(HERE, "answer.docx"), "en")
finally:
    w.Quit(0)
