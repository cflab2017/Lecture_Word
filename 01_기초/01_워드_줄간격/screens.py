"""워드 기초 1강 화면 — 보이는 워드(따로 띄운 것)에 시험 문서를 만들고 PrintWindow 로 창을 찍는다. 저장하지 않고 닫는다.
찍는 것: ① 줄 간격 1.0·1.15·1.5·2.0 비교 문서(리본 포함) ② 단락 대화상자(들여쓰기 및 간격) ③ 줄 격자 함정(같은 1.0 이 두 배로).
결과: shots/*.png (원본 크기, 저장소에는 올리지 않음). 실행: python screens.py"""
import ctypes, ctypes.wintypes as W, os, threading, time
import win32com.client as win32
from PIL import Image
from screens_util import grab, find_windows, pid_of

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
u32, g32 = ctypes.windll.user32, ctypes.windll.gdi32
u32.SetProcessDPIAware()
TEXT = ("이번 주 팀 회의는 화요일 오전 10시에 3층 대회의실에서 합니다. 지난달 매출 보고와 다음 분기 계획을 "
        "나누고, 각 팀은 진행 상황을 5분씩 발표합니다.")


w = win32.DispatchEx("Word.Application")
try:
    w.DisplayAlerts = 0
    doc = w.Documents.Add()
    w.Visible = True
    w.WindowState = 0  # 보통 창
    w.Width, w.Height = 760, 640
    win = doc.ActiveWindow
    win.View.Type = 3
    win.View.Zoom.Percentage = 100
    time.sleep(1.5)
    main = [h for h in find_windows("OpusApp")]
    hwnd = main[0]
    pid = pid_of(hwnd)

    # ① 줄 간격 비교 — 같은 글을 1.0 · 1.15 · 1.5 · 2.0 으로
    rng = doc.Content
    parts = [("줄 간격 1.0", 0, None), ("줄 간격 1.15", 5, 13.8), ("줄 간격 1.5", 1, None), ("줄 간격 2.0", 2, None)]
    rng.Text = "\r".join(f"{lab}\r{TEXT}" for lab, _, _ in parts)
    for i, (lab, rule, val) in enumerate(parts):
        head = doc.Paragraphs(2 * i + 1); body = doc.Paragraphs(2 * i + 2)
        head.Range.Font.Bold = True; head.Range.Font.Color = 0x00309C  # 진한 빨강(BGR)
        head.LineSpacingRule = 0; head.SpaceAfter = 2
        body.LineSpacingRule = rule
        if val:
            body.LineSpacing = val
        body.SpaceAfter = 14
    doc.Paragraphs(2).Range.Select()
    w.Selection.Collapse(1)
    win.ScrollIntoView(doc.Paragraphs(1).Range)
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s1_compare.png"))

    # ② 단락 대화상자 — 다른 스레드가 창을 찍고 닫는다(Show 는 닫힐 때까지 멈춘다)
    def shoot_dialog(name="s2_dialog.png"):
        for _ in range(60):
            time.sleep(0.25)
            d = find_windows("bosa_sdm_msword", pid)
            if d:
                time.sleep(1.2)
                grab(d[0], os.path.join(OUT, name))
                u32.PostMessageW(d[0], 0x0010, 0, 0)  # WM_CLOSE = 취소
                return
        print("dialog not found")
    shoot_dialog_as = shoot_dialog
    doc.Paragraphs(4).Range.Select()  # '줄 간격 1.15' 본문 단락에 커서만 둔다(범위를 고르면 값이 섞여 빈칸으로 보인다)
    w.Selection.Collapse(1)
    th = threading.Thread(target=shoot_dialog); th.start()
    dlg = w.Dialogs(175)  # wdDialogFormatParagraph
    dlg.Show()
    th.join()

    # ③ 줄 격자 함정 — 위 구역은 격자 없음, 아래 구역은 '줄 격자'. 둘 다 줄 간격 1.0
    doc.Content.Text = ""
    doc.Content.Text = f"격자 없음 · 줄 간격 1.0\r{TEXT}\r"
    end = doc.Content; end.Collapse(0)
    end.InsertBreak(2)  # 다음 페이지가 아니라 같은 페이지에서 이어지게 아래서 바꾼다
    doc.Sections(2).PageSetup.SectionStart = 0  # 이어서
    end = doc.Content; end.Collapse(0)
    end.InsertAfter(f"줄 격자 · 줄 간격 1.0\r{TEXT}")
    doc.Sections(2).PageSetup.LayoutMode = 2
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        p.LineSpacingRule = 0; p.SpaceAfter = 6
        if p.Range.Text.startswith(("격자", "줄 격자")):
            p.Range.Font.Bold = True; p.Range.Font.Color = 0x00309C
    doc.Repaginate()
    win.ScrollIntoView(doc.Paragraphs(1).Range)
    time.sleep(2.0)
    grab(hwnd, os.path.join(OUT, "s3_grid.png"))
    # 고친 뒤 — 줄 격자 구역 단락만 '문서 격자에 맞춤'을 끈다(DisableLineHeightGrid)
    for i in range(1, doc.Paragraphs.Count + 1):
        if doc.Paragraphs(i).Range.Information(2) >= 1 and doc.Paragraphs(i).Range.Sections(1).Index == 2:
            doc.Paragraphs(i).DisableLineHeightGrid = True
    doc.Repaginate(); time.sleep(1.5)
    grab(hwnd, os.path.join(OUT, "s3b_grid_fixed.png"))
    # 같은 단락을 '배수 1.15'로 두고 대화상자 — 줄 간격 칸 이름과 격자 칸이 꺼진 모습 확인
    last = doc.Paragraphs(doc.Paragraphs.Count)
    last.LineSpacingRule = 5; last.LineSpacing = 13.8
    last.Range.Select(); w.Selection.Collapse(1)
    th = threading.Thread(target=shoot_dialog_as, args=("s4_dialog_multiple.png",)); th.start()
    w.Dialogs(175).Show()
    th.join()
    doc.Close(0)
finally:
    w.Quit(0)
