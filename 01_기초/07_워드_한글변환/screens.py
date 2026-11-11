"""워드 기초 7강(워드 ↔ 한글 변환) 화면 — 보이는 워드(따로 띄운 것)에서 PrintWindow 로 찍는다. 파일은 읽기 전용으로 열고 저장하지 않는다.
대화상자는 워드 창 그림 위 제자리에 겹쳐 붙인다(11/6 규칙: 창 전체가 보이게).
s1 실습.hwp 를 '변환 확인'으로 열 때 워드가 띄우는 [파일 변환] 창(Documents.Open ConfirmConversions=True) — 취소
s2 실습.hwp 를 그냥 열었을 때(ConfirmConversions=False) — 알아볼 수 없는 글자로 열린 화면
s3 '아래아 한글 2.0-97' 변환기로 열게 하면 뜨는 오류 창
s4 정답.docx(한글 2024 가 HWP → OOXML 로 저장) 를 워드로 연 화면 — 글꼴 칸 함초롬바탕 · 본문 커서
한글 쪽 화면: 이 PC 한글 2024 는 편집 창이 열리지 않아(구독 없음) '호환 문서' 창만 measure 중에 찍었다(shots/s5_hangul_compat.png, 11/11 10:0x).
결과 shots/*.png, 읽은 값 screens.json. 실행: python screens.py (pywin32 · pillow) — 먼저 measure.py · make_files.py"""
import json, os, sys, threading, time
import win32com.client as win32, win32con, win32gui, win32process
from PIL import Image

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02_워드_들여쓰기"))
from screens_util import grab  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots")
os.makedirs(OUT, exist_ok=True)
R = {}
DLG = ("bosa_sdm_msword", "#32770", "NUIDialog")


def dialogs(pid, main):
    found = []

    def cb(h, _):
        if h != main and win32gui.IsWindowVisible(h) and win32process.GetWindowThreadProcessId(h)[1] == pid \
                and win32gui.GetClassName(h) in DLG:
            found.append(h)
        return True
    win32gui.EnumWindows(cb, 0)
    return found


def texts(h):
    out = []
    win32gui.EnumChildWindows(h, lambda c, _: out.append(win32gui.GetWindowText(c)) or True, 0)
    return [t for t in out if t]


def composite(main, dlg, out):
    mr, dr = win32gui.GetWindowRect(main), win32gui.GetWindowRect(dlg)
    mw, mh, dw, dh = mr[2] - mr[0], mr[3] - mr[1], dr[2] - dr[0], dr[3] - dr[1]
    win32gui.SetWindowPos(dlg, 0, mr[0] + max(0, (mw - dw) // 2), mr[1] + max(60, (mh - dh) // 2), 0, 0,
                          win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)
    time.sleep(0.8)
    a, b = out + ".main.png", out + ".dlg.png"
    grab(main, a)
    grab(dlg, b)
    m, d = Image.open(a), Image.open(b)
    mr, dr = win32gui.GetWindowRect(main), win32gui.GetWindowRect(dlg)
    m.paste(d, (dr[0] - mr[0], dr[1] - mr[1]))
    m.save(out)
    os.remove(a); os.remove(b)
    return {"title": win32gui.GetWindowText(dlg), "class": win32gui.GetClassName(dlg), "texts": texts(dlg)[:14],
            "offset": (dr[0] - mr[0], dr[1] - mr[1]), "size": d.size}


def open_and_shoot(path, confirm, fmt, out):
    """Documents.Open 이 대화상자에서 멈추는 동안 다른 스레드가 찾아 찍고 닫는다."""
    res = {}

    def shoot():
        for _ in range(80):
            time.sleep(0.3)
            ds = dialogs(PID, HWND)
            if ds:
                time.sleep(1.5)
                res.update(composite(HWND, ds[0], os.path.join(OUT, out)))
                win32gui.PostMessage(ds[0], win32con.WM_CLOSE, 0, 0)
                return
        res["err"] = "dialog not found"
    th = threading.Thread(target=shoot)
    th.start()
    try:
        d = w.Documents.Open(path, confirm, True, False, "", "", False, "", "", fmt)
        res["opened"] = True
        d.Close(0)
    except Exception as e:  # noqa: BLE001
        res["open_error"] = (getattr(e, "excepinfo", None) or [None, None, str(e)])[2]
    th.join()
    return res


w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0
    base = w.Documents.Add()  # 창을 띄우려고 빈 문서 하나
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 900, 640
    time.sleep(1.5)
    HWND = base.ActiveWindow.Hwnd
    PID = win32process.GetWindowThreadProcessId(HWND)[1]
    HWP = os.path.join(HERE, "실습.hwp")
    # s1 변환 확인 창
    R["s1"] = open_and_shoot(HWP, True, 0, "s1_convert_dialog.png")
    # s3 아래아 한글 2.0-97 변환기 오류
    ara = next((c.OpenFormat for c in w.FileConverters if c.ClassName == "ARA20"), None)
    R["s3"] = open_and_shoot(HWP, False, ara, "s3_ara20_error.png")
    # s2 그냥 열기 — 알아볼 수 없는 글자
    d = w.Documents.Open(HWP, False, True, False)
    d.ActiveWindow.View.Type = 3
    d.ActiveWindow.View.Zoom.Percentage = 100
    time.sleep(3)
    R["s2"] = {"pages": d.Content.Information(4), "chars": len(d.Content.Text)}
    grab(d.ActiveWindow.Hwnd, os.path.join(OUT, "s2_hwp_garbage.png"))
    d.Close(0)
    # s4 한글이 만든 DOCX
    d = w.Documents.Open(os.path.join(HERE, "정답.docx"), False, True, False)
    d.ActiveWindow.View.Type = 3
    d.ActiveWindow.View.Zoom.Percentage = 100
    tp = next(i for i in range(1, d.Paragraphs.Count + 1) if len(d.Paragraphs(i).Range.Text) > 40)  # 첫 본문 단락
    p = d.Paragraphs(tp)
    d.Range(p.Range.Start, p.Range.Start + 5).Select()  # 본문 첫 낱말을 골라 글꼴 칸에 그 글꼴이 보이게
    time.sleep(2.5)
    R["s4"] = {"pages": d.Content.Information(4), "font": p.Range.Font.Name, "size": p.Range.Font.Size,
               "line_rule": p.Format.LineSpacingRule, "line_pt": p.Format.LineSpacing,
               "footer": d.Sections(1).Footers(1).Range.Text.strip("\r"), "compat_mode": d.CompatibilityMode,
               "first_para_empty": d.Paragraphs(1).Range.Text == "\r", "body_para_index": tp,
               "selected_font_box": w.Selection.Font.Name}
    grab(d.ActiveWindow.Hwnd, os.path.join(OUT, "s4_hwp_to_docx.png"))
    d.Close(0)
    try:
        base.Close(0)
    except Exception:  # noqa: BLE001  빈 문서는 다른 문서를 열 때 워드가 바꿔 끼워 이미 없을 수 있다
        pass
finally:
    w.Quit(0)
json.dump(R, open(os.path.join(HERE, "screens.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
