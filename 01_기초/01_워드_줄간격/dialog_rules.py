"""단락 대화상자의 '줄 간격' 칸 이름을 확인한다 — 규칙(1.5줄·2줄·최소·고정)마다 대화상자를 열어 찍고 닫는다.
결과: shots/s6_rule_<번호>.png (줄 간격 칸만 잘라 붙인 것은 shots/s6_rules.png). 저장하지 않는다. 실행: python dialog_rules.py"""
import os, threading, time
import win32com.client as win32
from PIL import Image
from screens_util import grab, find_windows, pid_of

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
RULES = [(0, None), (1, None), (2, None), (3, 20), (4, 20), (5, 36)]

w = win32.DispatchEx("Word.Application")
try:
    doc = w.Documents.Add()
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 760, 640
    doc.Content.Text = "줄 간격 이름 확인용 글"
    time.sleep(1.5)
    pid = pid_of(find_windows("OpusApp")[0])
    shots = []
    for rule, val in RULES:
        p = doc.Paragraphs(1)
        p.LineSpacingRule = rule
        if val:
            p.LineSpacing = val
        p.Range.Select(); w.Selection.Collapse(1)
        name = os.path.join(OUT, f"s6_rule_{rule}.png")

        def shoot():
            for _ in range(60):
                time.sleep(0.25)
                d = find_windows("bosa_sdm_msword", pid)
                if d:
                    time.sleep(2.0); grab(d[0], name)
                    import ctypes; ctypes.windll.user32.PostMessageW(d[0], 0x0010, 0, 0)
                    return
        th = threading.Thread(target=shoot); th.start()
        w.Dialogs(175).Show(); th.join()
        shots.append(name)
    doc.Close(0)
finally:
    w.Quit(0)
strip = [Image.open(s).crop((230, 335, 463, 392)) for s in shots]
c = Image.new("RGB", (strip[0].width, sum(i.height for i in strip)), "white")
y = 0
for i in strip:
    c.paste(i, (0, y)); y += i.height
c.save(os.path.join(OUT, "s6_rules.png"))
print("ok", len(shots))
