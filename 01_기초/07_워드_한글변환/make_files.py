"""워드 기초 7강(워드 ↔ 한글 변환) 실습 · 정답 파일 — measure.py 가 %TEMP%/claude/word07[_en] 에 만든 파일을 가져온다(먼저 python measure.py [en]).
실습.docx = 워드로 만든 가상 '사진 동아리 12월 모임 안내'(맑은 고딕, 표 3×3, 그림, 바닥글 쪽 번호)
실습.hwp  = 같은 내용을 한글 2024 로 만든 것(함초롬바탕 10pt · 줄 간격 160%, 표 3×3, 그림, 쪽 번호 매기기 '- 1 -')
정답.hwp  = 실습.docx 를 한글 2024 로 열어 HWP 로 저장한 것 · 정답.docx = 실습.hwp 를 한글 2024 '다른 이름으로 저장 → OOXML(.docx)' 한 것
argv 'en' 이면 practice.* · answer.*. 작성자: DOCX 는 core.xml(creator · lastModifiedBy), HWP 는 한글 문서 정보(HSummaryInfo Author)를 CODINGNOW 로.
실행: python make_files.py [en]  (pywin32)"""
import os, re, shutil, sys, tempfile, zipfile
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "word07" + ("_en" if EN else ""))
P, A = ("practice", "answer") if EN else ("실습", "정답")
USER = os.environ.get("USERNAME", "")


def docx_author(path):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "docProps/core.xml":
                s = data.decode("utf-8")
                s = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)", r"\1CODINGNOW\2", s)
                s = re.sub(r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)", r"\1CODINGNOW\2", s)
                data = s.encode("utf-8")
            zout.writestr(info, data)
    shutil.move(tmp, path)
    core = zipfile.ZipFile(path).read("docProps/core.xml").decode("utf-8")
    return re.findall(r"<dc:creator>([^<]*)<", core) + re.findall(r"<cp:lastModifiedBy>([^<]*)<", core)


# DOCX
shutil.copyfile(os.path.join(WORK, "practice_word.docx"), os.path.join(HERE, P + ".docx"))
shutil.copyfile(os.path.join(WORK, "hwp_to.docx"), os.path.join(HERE, A + ".docx"))
print(P + ".docx", docx_author(os.path.join(HERE, P + ".docx")))
print(A + ".docx", docx_author(os.path.join(HERE, A + ".docx")))

# HWP — 한글로 열어 문서 정보 작성자를 바꾸고 %TEMP% 에 저장한 뒤 복사('파일 접근 허용' 창을 피함)
hwp = win32.Dispatch("HWPFrame.HwpObject")
hwp.XHwpWindows.Item(0).Visible = False
try:
    for src, name in (("native.hwp", P + ".hwp"), ("docx_to.hwp", A + ".hwp")):
        assert hwp.Open(os.path.join(WORK, src), "HWP", "")
        si = hwp.HParameterSet.HSummaryInfo
        hwp.HAction.GetDefault("DocSummaryInfo", si.HSet)
        before = si.Author
        si.Author = "CODINGNOW"
        ok = hwp.HAction.Execute("DocSummaryInfo", si.HSet)
        out = os.path.join(WORK, "out_" + name)
        if os.path.exists(out):
            os.remove(out)
        assert hwp.SaveAs(out, "HWP", "")
        hwp.Clear(1)
        shutil.copyfile(out, os.path.join(HERE, name))
        raw = open(out, "rb").read()
        print(name, "작성자 전", repr(before), "→ Execute", bool(ok), "| PC 사용자 이름 남음", bool(USER) and USER.encode("utf-16-le") in raw,
              "| CODINGNOW 들어감", "CODINGNOW".encode("utf-16-le") in raw)
finally:
    hwp.Clear(1)
    hwp.Quit()
