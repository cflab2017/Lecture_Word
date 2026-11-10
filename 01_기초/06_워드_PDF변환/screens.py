"""워드 기초 6강(PDF 변환) 화면 — 실습 파일을 보이는 워드(따로 띄운 것)에서 읽기 전용으로 열어 PrintWindow 로 찍는다. 저장하지 않는다.
s1 실습 문서(메모 풍선이 보이는 화면) · s2 [다른 이름으로 저장] 대화상자 — 파일 형식 PDF(표준 · 최소 크기 · 옵션 단추)
s3 PDF [옵션] 대화상자(페이지 범위 · 게시 대상 문서/변경 내용 표시 문서 · 책갈피 만들기 · PDF/A) — 둘 다 취소로 닫음
s4 · s5 measure.py 가 만든 PDF 첫 쪽(다른 이름으로 저장 그대로 · 내보내기 문서만) — PyMuPDF 로 그림.
결과 shots/*.png. 실행: python screens.py (pywin32 · pillow · pymupdf) — 먼저 measure.py · make_files.py"""
import os, sys, threading, time
import pymupdf
import win32com.client as win32, win32gui, win32process, win32con

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02_워드_들여쓰기"))
from screens_util import grab  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
WORK = os.path.join(os.environ["TEMP"], "claude", "word06")
BM_CLICK = 0x00F5


def dialogs(pid, main):
    found = []

    def cb(h, _):
        if h != main and win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid \
                and win32gui.GetClassName(h) in ("bosa_sdm_msword", "#32770", "NUIDialog"):
            found.append(h)
        return True
    win32gui.EnumWindows(cb, 0)
    return found


def children(h):
    out = []
    win32gui.EnumChildWindows(h, lambda c, _: out.append((c, win32gui.GetClassName(c), win32gui.GetWindowText(c))) or True, 0)
    return out


def save_as_flow(pid, main):
    seen = set()
    for _ in range(80):
        time.sleep(0.3)
        ds = dialogs(pid, main)
        if ds:
            break
    else:
        print("save-as not found"); return
    time.sleep(2.0)
    sa = ds[0]
    seen.add(sa)
    print("s2 |", win32gui.GetClassName(sa), "|", win32gui.GetWindowText(sa), "|", [t for _, _, t in children(sa) if t][:20])
    grab(sa, os.path.join(OUT, "s2_save_as_pdf.png"))
    btn = [c for c, cls, t in children(sa) if cls == "Button" and "옵션" in t]
    print("옵션 단추:", len(btn))
    if btn:
        win32gui.PostMessage(btn[0], BM_CLICK, 0, 0)
        for _ in range(40):
            time.sleep(0.3)
            new = [h for h in dialogs(pid, main) if h not in seen]
            if new:
                time.sleep(1.2)
                print("s3 |", win32gui.GetWindowText(new[0]), "|", [t for _, _, t in children(new[0]) if t][:30])
                grab(new[0], os.path.join(OUT, "s3_pdf_options.png"))
                win32gui.PostMessage(new[0], win32con.WM_CLOSE, 0, 0)
                time.sleep(1.0)
                break
    win32gui.PostMessage(sa, win32con.WM_CLOSE, 0, 0)


w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0
    d = w.Documents.Open(os.path.join(HERE, "실습.docx"), ReadOnly=True)
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 900, 640
    win = d.ActiveWindow
    win.View.Type = 3
    win.View.Zoom.Percentage = 100
    time.sleep(2)
    hwnd = win.Hwnd
    top = hwnd
    while win32gui.GetParent(top):
        top = win32gui.GetParent(top)
    _, pid = win32process.GetWindowThreadProcessId(top)
    grab(top, os.path.join(OUT, "s1_practice.png"))
    th = threading.Thread(target=save_as_flow, args=(pid, top)); th.start()
    dlg = w.Dialogs(84)  # wdDialogFileSaveAs
    try:
        dlg.Format = 17  # PDF
    except Exception as e:  # noqa: BLE001
        print("format:", str(e)[:80])
    try:
        dlg.Show()
    except Exception as e:  # noqa: BLE001
        print("show:", str(e)[:80])
    th.join()
    d.Close(0)
finally:
    w.Quit(0)

# PDF 첫 쪽 그림(이 PC 워드가 만든 PDF 그대로)
for name, out in (("A_saveas", "s4_pdf_saveas_p1.png"), ("B_standard", "s5_pdf_export_p1.png")):
    with pymupdf.open(os.path.join(WORK, name + ".pdf")) as doc:
        doc[0].get_pixmap(dpi=90).save(os.path.join(OUT, out))
        print(out)
