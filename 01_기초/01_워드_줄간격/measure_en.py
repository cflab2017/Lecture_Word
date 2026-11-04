"""Word Basics 1 (EN) — line spacing measured in Word for Microsoft 365 with English text. Output: measure_en.json
Starts its own hidden Word, makes a new document, closes it without saving. Normal.dotm is not touched.
Measures: distance between lines (pt) for Single / 1.15 / 1.5 / Double per font, paragraph gap with Space After,
Shift+Enter vs Enter, and "Don't add space between paragraphs of the same style".
Run: python measure_en.py  (pip install pywin32)"""
import json, os
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
R = {}
TEXT = ("The team meeting is on Tuesday at 10 a.m. in the third-floor conference room. We will go over last month's "
        "sales report and next quarter's plan, and each team will give a five-minute update. Please upload your slides "
        "to the shared team folder by 6 p.m. on Monday.")
POS = 6  # wdVerticalPositionRelativeToPage


def line_tops(rng):
    tops = []
    for i in range(1, rng.Characters.Count + 1):
        y = round(rng.Characters(i).Information(POS), 2)
        if not tops or abs(y - tops[-1]) > 0.5:
            tops.append(y)
    return tops


w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["version"] = f"{w.Version} build {w.Build}"
    doc = w.Documents.Add()
    doc.ActiveWindow.View.Type = 3
    st = doc.Styles(-1)
    doc.Content.Text = TEXT + "\r" + TEXT + "\r" + TEXT
    doc.Content.LanguageID = 1033  # English (US)

    def measure(font, size, rule, value=None, after=0):
        doc.Content.Font.Name = font
        doc.Content.Font.Size = size
        p = doc.Content.ParagraphFormat
        p.LineSpacingRule = rule
        if value is not None:
            p.LineSpacing = value
        p.SpaceAfter = after
        p.SpaceBefore = 0
        doc.Repaginate()
        t1 = line_tops(doc.Paragraphs(1).Range)
        t2 = line_tops(doc.Paragraphs(2).Range)
        span = t1[-1] - t1[0]
        return {"lines": len(t1), "avg_line_pitch": round(span / (len(t1) - 1), 2), "par_gap": round(t2[0] - t1[-1], 2),
                "font_used": doc.Paragraphs(1).Range.Font.Name}

    R["by_font"] = {}
    for font, size in [("Aptos", 12), ("Calibri", 11), ("Times New Roman", 12), ("Arial", 11)]:
        R["by_font"][f"{font} {size}pt"] = {
            "Single": measure(font, size, 0),
            "1.15": measure(font, size, 5, 13.8),
            "1.5": measure(font, size, 1),
            "Double": measure(font, size, 2),
        }
    R["space_after_TNR12_double"] = {
        "after 0": measure("Times New Roman", 12, 2, after=0),
        "after 8pt": measure("Times New Roman", 12, 2, after=8),
    }

    # Shift+Enter (line break, code 11) vs Enter (new paragraph) with Space After 8pt
    doc.Content.Text = "First line\vSecond line (Shift+Enter)\rThird line (Enter)"
    doc.Content.Font.Name = "Aptos"; doc.Content.Font.Size = 12
    doc.Content.ParagraphFormat.LineSpacingRule = 0
    doc.Content.ParagraphFormat.SpaceAfter = 8
    doc.Repaginate()
    p1 = doc.Paragraphs(1).Range
    ys = [round(p1.Characters(1).Information(POS), 2), round(p1.Characters(12).Information(POS), 2),
          round(doc.Paragraphs(2).Range.Characters(1).Information(POS), 2)]
    R["shift_enter"] = {"paragraphs": doc.Paragraphs.Count, "gap_shift_enter": round(ys[1] - ys[0], 2), "gap_enter": round(ys[2] - ys[1], 2)}

    # Don't add space between paragraphs of the same style
    doc.Content.Text = TEXT + "\r" + TEXT
    doc.Content.Font.Name = "Aptos"; doc.Content.Font.Size = 12
    doc.Content.ParagraphFormat.LineSpacingRule = 0
    doc.Content.ParagraphFormat.SpaceAfter = 8
    doc.Repaginate()
    g0 = round(line_tops(doc.Paragraphs(2).Range)[0] - line_tops(doc.Paragraphs(1).Range)[-1], 2)
    st.NoSpaceBetweenParagraphsOfSameStyle = True  # this document's Normal style only
    doc.Repaginate()
    g1 = round(line_tops(doc.Paragraphs(2).Range)[0] - line_tops(doc.Paragraphs(1).Range)[-1], 2)
    R["no_space_same_style"] = {"gap_off": g0, "gap_on": g1}
    doc.Close(0)
finally:
    w.Quit(0)

json.dump(R, open(os.path.join(HERE, "measure_en.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
