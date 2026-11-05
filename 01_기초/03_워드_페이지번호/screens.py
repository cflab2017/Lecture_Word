"""워드 기초 3강 화면 — 보이는 워드(따로 띄운 것)에 정답과 같은 문서(아래쪽 가운데 번호 · 첫 페이지 다르게 · 시작 번호 0)를 만들고 PrintWindow 로 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 1·2쪽 나란히(표지는 번호 없음, 2쪽 아래 '1') ② 2쪽 바닥글 편집 상태(리본 '머리글/바닥글' 탭 — 첫 페이지를 다르게 지정)
③ 페이지 번호 서식 대화상자(시작 번호 0).  결과: shots/*.png (원본 크기, 저장소에는 올리지 않음). 실행: python screens.py"""
import ctypes, os, sys, threading, time
import win32com.client as win32
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02_워드_들여쓰기"))
from screens_util import grab, find_windows, pid_of
from measure_text import KO, build

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
u32 = ctypes.windll.user32

w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0
    doc = w.Documents.Add()
    build(doc, KO)
    pn = doc.Sections(1).Footers(1).PageNumbers
    pn.Add(1, True)
    doc.PageSetup.DifferentFirstPageHeaderFooter = True
    pn = doc.Sections(1).Footers(1).PageNumbers
    pn.RestartNumberingAtSection = True; pn.StartingNumber = 0
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 900, 640
    win = doc.ActiveWindow
    win.View.Type = 3
    time.sleep(1.5)
    hwnd = find_windows("OpusApp")[0]; pid = pid_of(hwnd)
    # ① 1·2쪽 나란히
    win.View.Zoom.PageColumns = 2; win.View.Zoom.PageRows = 1
    doc.Range(0, 0).Select(); time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s1_two_pages.png"))
    # ② 2쪽 바닥글 편집
    win.View.Zoom.Percentage = 100
    doc.Paragraphs(4).Range.Select(); w.Selection.Collapse(1)
    win.ActivePane.View.SeekView = 10   # wdSeekCurrentPageFooter
    time.sleep(2.5)
    grab(hwnd, os.path.join(OUT, "s2_footer_edit.png"))
    # ③ 페이지 번호 서식 대화상자

    def shoot_dialog(name):
        for _ in range(60):
            time.sleep(0.25)
            d = find_windows("bosa_sdm_msword", pid)
            if d:
                time.sleep(1.5)
                grab(d[0], os.path.join(OUT, name))
                u32.PostMessageW(d[0], 0x0010, 0, 0)  # WM_CLOSE = 취소
                return
        print("dialog not found")
    th = threading.Thread(target=shoot_dialog, args=("s3_format_dialog.png",)); th.start()
    w.Dialogs(298).Show()  # wdDialogFormatPageNumber
    th.join()
    win.ActivePane.View.SeekView = 0
    doc.Close(0)
finally:
    w.Quit(0)
