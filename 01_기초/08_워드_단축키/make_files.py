"""워드 기초 8강(단축키) 실습 · 정답 파일 — 가상 사진 동아리 11월 모임 안내(실제 모임 아님).
실습: 서식 없는 제목 + 본문 5단락(모두 기본 스타일). 정답: 제목 Ctrl+B · Ctrl+E · Ctrl+] 세 번(14pt), 일시~문의 네 단락 Ctrl+T(내어쓰기),
일시 단락 끝에 Shift+Enter 로 '비가 오면 …' 줄. 작성자 · 마지막 수정한 사람은 CODINGNOW.
실행: python make_files.py [en]  (pywin32)"""
import os, re, shutil, sys, tempfile, zipfile
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
EN = len(sys.argv) > 1 and sys.argv[1] == "en"

KO = ["사진 동아리 11월 정기 모임 안내",
      "이번 달 정기 모임은 야외 촬영(photo walk)으로 진행합니다. 카메라가 없으면 휴대폰으로 찍어도 됩니다.",
      "일시: 11월 21일(토) 오후 2시",
      "장소: 시민공원 정문 앞",
      "준비물: 카메라 또는 휴대폰, 여분 배터리",
      "문의: 총무 김하늘"]
EN_T = ["Photo Club November Meetup",
        "This month we are going on a photo walk. If you do not have a camera, a phone is fine.",
        "When: Saturday, November 21, 2 PM",
        "Where: City Park, main gate",
        "Bring: a camera or phone, a spare battery",
        "Questions: Ava (club secretary)"]
RAIN = "비가 오면 28일(토)로 미룹니다" if not EN else "If it rains, we move to November 28"


def fix_core(path):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "docProps/core.xml":
                s = data.decode("utf-8")
                for tag in ("dc:creator", "cp:lastModifiedBy"):
                    s = re.sub(f"<{tag}>[^<]*</{tag}>", f"<{tag}>CODINGNOW</{tag}>", s)
                data = s.encode("utf-8")
            zout.writestr(info, data, compress_type=info.compress_type)
    shutil.move(tmp, path)


pythoncom.CoInitialize()
wd = win32.DispatchEx("Word.Application")
try:
    wd.Visible = False
    wd.DisplayAlerts = 0
    for answer, name in ((False, "practice.docx" if EN else "실습.docx"), (True, "answer.docx" if EN else "정답.docx")):
        doc = wd.Documents.Add()
        doc.Content.Text = "\r".join(EN_T if EN else KO)
        if answer:
            t = doc.Paragraphs(1).Range
            t.Font.Bold = True
            t.ParagraphFormat.Alignment = 1
            t.Font.Size = 14  # Ctrl+] 세 번(11 → 14)
            for i in (3, 4, 5, 6):
                pf = doc.Paragraphs(i).Range.ParagraphFormat
                pf.LeftIndent = 40  # Ctrl+T 한 번 = 내어쓰기(실측 measure.json: 왼쪽 40pt · 첫 줄 -40pt, 기본 탭 간격 40pt)
                pf.FirstLineIndent = -40
            r = doc.Paragraphs(3).Range
            r.MoveEnd(1, -1)
            r.InsertAfter("\v" + RAIN)
        doc.BuiltInDocumentProperties("Author").Value = "CODINGNOW"
        path = os.path.join(HERE, name)
        if os.path.exists(path):
            os.remove(path)
        doc.SaveAs2(path, 16)
        doc.Close(False)
        fix_core(path)
        print("saved", name)
finally:
    wd.Quit()
