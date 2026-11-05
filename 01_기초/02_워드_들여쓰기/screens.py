"""워드 기초 2강 화면 — 보이는 워드(따로 띄운 것)에 시험 문서를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 첫 줄 2글자 · Tab · 번호 내어쓰기 세 단락(눈금자 · 편집 기호 표시, 커서는 첫 단락) ② 같은 문서, 커서를 번호 단락에(눈금자 표식이 바뀜)
③ 단락 대화상자 — 첫 줄 2글자 ④ 단락 대화상자 — 내어쓰기.
결과: shots/*.png (원본 크기, 저장소에는 올리지 않음). 실행: python screens.py"""
import ctypes, os, threading, time
import win32com.client as win32
from screens_util import grab, find_windows, pid_of
from measure_text import P1, P2, ITEM

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
u32 = ctypes.windll.user32

w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0
    doc = w.Documents.Add()
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 760, 640
    win = doc.ActiveWindow
    win.View.Type = 3; win.View.Zoom.Percentage = 100
    win.DisplayRulers = True
    win.View.ShowAll = True  # 편집 기호(탭 →, 띄어쓰기 ·, 단락 ¶)
    time.sleep(1.5)
    hwnd = find_windows("OpusApp")[0]; pid = pid_of(hwnd)
    doc.Content.InsertAfter(f"{P1}\r\t{P2}\r{ITEM}")
    for i in range(1, 4):
        doc.Paragraphs(i).Reset(); doc.Paragraphs(i).SpaceAfter = 12
    doc.Paragraphs(1).CharacterUnitFirstLineIndent = 2
    doc.Paragraphs(3).CharacterUnitFirstLineIndent = -2
    win.ScrollIntoView(doc.Paragraphs(1).Range)
    doc.Paragraphs(1).Range.Characters(1).Select(); w.Selection.Collapse(1)
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s1_first.png"))
    doc.Paragraphs(3).Range.Characters(3).Select(); w.Selection.Collapse(1)
    time.sleep(1.5)
    grab(hwnd, os.path.join(OUT, "s2_hanging.png"))

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
    for k, name in ((1, "s3_dialog_first.png"), (3, "s4_dialog_hanging.png")):
        doc.Paragraphs(k).Range.Characters(2).Select(); w.Selection.Collapse(1)
        th = threading.Thread(target=shoot_dialog, args=(name,)); th.start()
        w.Dialogs(175).Show()  # wdDialogFormatParagraph
        th.join()
    doc.Close(0)
finally:
    w.Quit(0)
