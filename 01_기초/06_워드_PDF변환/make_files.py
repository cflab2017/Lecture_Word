"""워드 기초 6강(PDF 변환) 실습 · 정답 파일 — measure.py 가 만든 연습 문서(%TEMP%/claude/word06[_en]/report.docx)를 실습 파일로 복사하고,
같은 문서를 '문서만(메모 빼고) + 제목으로 책갈피 + 표준'으로 내보낸 PDF 를 정답 파일로 만든다. 먼저 python measure.py [en] 를 돌릴 것.
argv 'en' 이면 practice.docx · answer.pdf. 작성자 · 메모 작성자는 CODINGNOW.
실행: python make_files.py [en]  (pywin32 · pymupdf)"""
import os, shutil, sys
import pymupdf
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "word06" + ("_en" if EN else ""))
src = os.path.join(WORK, "report.docx")
assert os.path.exists(src), "먼저 measure.py 를 돌리세요"
dst = os.path.join(HERE, "practice.docx" if EN else "실습.docx")
shutil.copyfile(src, dst)
pdf = os.path.join(HERE, "answer.pdf" if EN else "정답.pdf")
pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    doc = w.Documents.Open(dst, ReadOnly=True)
    doc.ExportAsFixedFormat(OutputFileName=pdf, ExportFormat=17, OpenAfterExport=False, OptimizeFor=0, Range=0, From=1, To=1,
                            Item=0, IncludeDocProps=True, KeepIRM=True, CreateBookmarks=1, DocStructureTags=True,
                            BitmapMissingFonts=True, UseISO19005_1=False)
    doc.Close(0)
finally:
    w.Quit(0)
with pymupdf.open(pdf) as d:
    txt = "".join(p.get_text() for p in d)
    print("ok", dst, pdf, "쪽", d.page_count, "목차", len(d.get_toc()), "작성자", d.metadata.get("author"),
          "메모 글자", ("Check the date" in txt) or ("날짜 확인" in txt))
