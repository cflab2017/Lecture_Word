"""워드 기초 5강(표 만들기) 화면 — 실습 · 정답 파일을 보이는 워드(따로 띄운 것)에서 읽기 전용으로 열어 PrintWindow 로 찍는다. 저장하지 않는다.
s1 실습: 편집 기호(→ 탭)를 켠 명단 · s2 [텍스트를 표로 변환] 대화상자(명단 선택 뒤 — 열 수 3 · 탭) · s3 정답 2쪽 맨 위(머리글 행 반복)
s4 표 속성 '행' 탭(머리글 행으로 페이지마다 반복) · s5 쉼표 함정 — '서울, 강남,2' 를 쉼표로 변환한 표(새 문서, 저장 안 함)
결과 shots/*.png. 실행: python screens.py (pywin32 · pillow)"""
import os, sys, threading, time
import win32com.client as win32, win32gui, win32process, win32con

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02_워드_들여쓰기"))
from screens_util import grab, find_windows  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)


def dialog_shot(pid, main, name):
    for _ in range(60):
        time.sleep(0.3)
        found = []

        def cb(h, _):
            if h != main and win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid \
                    and win32gui.GetClassName(h) in ("bosa_sdm_msword", "#32770", "NUIDialog"):
                found.append(h)
            return True
        win32gui.EnumWindows(cb, 0)
        if found:
            time.sleep(1.2)
            print(name, "|", win32gui.GetClassName(found[0]), "|", win32gui.GetWindowText(found[0]))
            grab(found[0], os.path.join(OUT, name))
            win32gui.PostMessage(found[0], win32con.WM_CLOSE, 0, 0)
            return
    print(name, "dialog not found")


w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0

    def open_doc(nm, show_all=False):
        d = w.Documents.Open(os.path.join(HERE, nm), ReadOnly=True)
        w.Visible = True
        w.WindowState = 0
        w.Width, w.Height = 900, 640
        win = d.ActiveWindow
        win.View.Type = 3
        win.View.ShowAll = show_all
        win.View.Zoom.Percentage = 100
        time.sleep(1.5)
        return d, win

    # s1 · s2 실습
    d, win = open_doc("실습.docx", show_all=True)
    hwnd = find_windows("OpusApp")[0]
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    d.Range(0, 0).Select()
    time.sleep(1.5)
    grab(hwnd, os.path.join(OUT, "s1_practice_tabs.png"))
    d.Range(d.Paragraphs(2).Range.Start, d.Range().End).Select()
    th = threading.Thread(target=dialog_shot, args=(pid, hwnd, "s2_text_to_table.png")); th.start()
    try:
        w.CommandBars.ExecuteMso("ConvertTextToTable")
    except Exception as e:  # noqa: BLE001
        print("convert:", str(e)[:100])
    th.join()
    d.Close(0)
    # s3 · s4 정답
    d, win = open_doc("정답.docx")
    hwnd = find_windows("OpusApp")[0]
    t = d.Tables(1)
    # 2쪽 첫 줄로
    p2 = d.GoTo(1, 1, 2)  # wdGoToPage, wdGoToAbsolute, 2
    p2.Select()
    win.ScrollIntoView(p2)
    win.SmallScroll(Down=0)
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s3_answer_page2.png"))
    t.Cell(1, 1).Range.Select()
    w.Selection.Collapse(1)
    dlg = w.Dialogs(861)  # wdDialogTableProperties
    try:
        dlg.DefaultTab = 1800001  # 행 탭
    except Exception as e:  # noqa: BLE001
        print("tab:", str(e)[:80])
    th = threading.Thread(target=dialog_shot, args=(pid, hwnd, "s4_row_repeat.png")); th.start()
    try:
        dlg.Show()
    except Exception as e:  # noqa: BLE001
        print("props:", str(e)[:100])
    th.join()
    d.Close(0)
    # s5 쉼표 함정(새 문서)
    d = w.Documents.Add()
    w.Visible = True
    win = d.ActiveWindow
    win.View.Type = 3
    win.View.Zoom.Percentage = 100
    d.Range().Text = "지역,인원\r서울, 강남,2\r부산,3"
    t = d.Range().ConvertToTable(Separator=2)
    t.Style = "표 구분선"
    d.Range(0, 0).Select()
    time.sleep(1.5)
    grab(find_windows("OpusApp")[0], os.path.join(OUT, "s5_comma_trap.png"))
    print("comma table:", t.Rows.Count, "x", t.Columns.Count)
    d.Close(0)
finally:
    w.Quit()
