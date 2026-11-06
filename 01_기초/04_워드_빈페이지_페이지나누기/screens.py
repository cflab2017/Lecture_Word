"""워드 기초 4강 화면 — measure.py 가 저장한 실습.docx · 정답.docx 를 보이는 워드(따로 띄운 것)에서 읽기 전용으로 열고
편집 기호(모두 표시, Ctrl+Shift+8)를 켜서 PrintWindow 로 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 실습 세 쪽 나란히 — 1쪽 아래 빈 단락(¶) 줄, 3쪽에 ¶ 하나 ② 실습 100% — 제목 뒤 이어지는 ¶ 줄
③ 정답 100% — '페이지 나누기' 표시 ④ 정답 문서 끝(표 바로 뒤 단락) — 글꼴 크기 1 ⑤ 정답 두 쪽 나란히.
결과: shots/*.png (원본 크기, 저장소에는 올리지 않음). 실행: python screens.py"""
import os, sys, time
import win32com.client as win32
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02_워드_들여쓰기"))
from screens_util import grab, find_windows

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)

w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0

    def open_doc(nm):
        d = w.Documents.Open(os.path.join(HERE, nm), ReadOnly=True)
        w.Visible = True
        w.WindowState = 0
        w.Width, w.Height = 900, 640
        win = d.ActiveWindow
        win.View.Type = 3            # 인쇄 모양
        win.View.ShowAll = True      # 편집 기호 모두 표시
        time.sleep(1.5)
        return d, win

    hwnd = None
    # ① ② 실습
    d, win = open_doc("실습.docx")
    hwnd = find_windows("OpusApp")[0]
    win.View.Zoom.PageColumns = 3
    win.View.Zoom.PageRows = 1
    d.Range(0, 0).Select()
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s1_practice_pages.png"))
    win.View.Zoom.Percentage = 100
    d.Range(0, 0).Select()
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s2_practice_enters.png"))
    d.Close(0)
    # ③ ④ ⑤ 정답
    d, win = open_doc("정답.docx")
    hwnd = find_windows("OpusApp")[0]
    win.View.Zoom.Percentage = 100
    d.Range(0, 0).Select()
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s3_answer_break.png"))
    end = d.Paragraphs.Last.Range
    end.Select()
    w.Selection.Collapse(1)
    win.ScrollIntoView(end)
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s4_answer_end_1pt.png"))
    win.View.Zoom.PageColumns = 3
    win.View.Zoom.PageRows = 1
    d.Range(0, 0).Select()
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s5_answer_pages.png"))
    d.Close(0)
finally:
    w.Quit(0)
