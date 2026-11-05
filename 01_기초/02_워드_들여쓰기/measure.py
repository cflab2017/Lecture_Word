"""워드 기초 2강 — 들여쓰기 실측. 이 PC 워드(Microsoft 365)를 숨긴 채 따로 띄워(DispatchEx)
첫 줄 들여쓰기 · Tab · 띄어쓰기 · 내어쓰기 · 왼쪽 들여쓰기에서 줄마다 글자가 시작하는 가로 위치(pt)를 잰다.
결과 measure.json. 데이터는 모두 연습용 가상 '도서관 이용 안내'.
실행: python measure.py  (pywin32 필요)"""
import datetime, json, os, sys
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
X = 5  # wdHorizontalPositionRelativeToPage
LANG = sys.argv[1] if len(sys.argv) > 1 else "ko"
if LANG == "en":
    P1 = ("During November the second-floor reading room is open from 9 a.m. to 9 p.m. every day. Seats fill up quickly "
          "during exam weeks, so please come early if you need a desk.")
    P2 = ("You can borrow up to five books for fourteen days. If a book comes back late, borrowing is paused for as many "
          "days as it was late.")
    ITEM = ("1.\tReturns: the self-service return box on the first floor is open 24 hours, including the days the library is closed.")
else:
    P1 = ("11월 한 달 동안 도서관 2층 열람실은 매일 오전 9시부터 오후 9시까지 문을 엽니다. 시험 기간에는 자리가 빨리 "
          "차니 자리가 필요하면 일찍 오시기 바랍니다.")
    P2 = ("대출은 한 사람에 5권까지, 대출 기간은 14일입니다. 반납이 늦어지면 늦은 날만큼 대출이 정지됩니다. "
          "연장은 한 번, 7일까지 할 수 있습니다.")
    ITEM = ("1.\t반납: 1층 무인 반납기는 24시간 쓸 수 있으며, 도서관이 쉬는 날에도 책을 넣을 수 있습니다.")
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": LANG}


def lines_x(p):
    """단락 p 의 줄마다 첫 글자(띄어쓰기·탭 아닌)의 가로 위치(pt). Range.Information 은 들여쓰기를 반영하지 않아
    (선택하지 않은 범위는 72 그대로였다) 글자를 하나씩 선택해 Selection.Information 으로 잰다."""
    out = []
    doc = p.Range.Document
    p.Range.Characters(1).Select(); w.Selection.Collapse(1)
    end = p.Range.End
    for _ in range(4):
        w.Selection.HomeKey(5)  # wdLine
        s0 = w.Selection.Start
        if s0 >= end - 1:
            break
        r = doc.Range(s0, s0 + 1)
        while r.Text in (" ", "	") and r.End < end:
            r = doc.Range(r.Start + 1, r.Start + 2)
        r.Select()
        out.append(round(w.Selection.Information(X), 2))
        w.Selection.Collapse(1)
        if w.Selection.MoveDown(5, 1) == 0 or w.Selection.Start >= end - 1:
            break
    return out


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = True; w.DisplayAlerts = 0  # 보이는 창이어야 Selection.Information 이 맞다
    w.WindowState = 0; w.Width, w.Height = 760, 600
    R["build"] = w.Build
    doc = w.Documents.Add()
    doc.ActiveWindow.View.Type = 3
    R["default_tab_stop"] = doc.DefaultTabStop
    st = doc.Styles(-1)  # 표준
    R["normal"] = {"font": st.Font.Name, "size": st.Font.Size, "left": st.ParagraphFormat.LeftIndent, "first": st.ParagraphFormat.FirstLineIndent}
    R["tab_indent_key_option"] = w.Options.TabIndentKey
    R["margin_left"] = doc.PageSetup.LeftMargin
    labels = {}
    for idm in ("IndentIncreaseWord", "IndentDecreaseWord", "ParagraphIndentLeft", "ParagraphIndentRight", "IndentLeft", "IndentRight",
                "ParagraphDialog", "ParagraphMarks", "ViewRulerWord", "ParagraphOptionsDialog", "FormattingMarks", "ShowAllFormatting"):
        try:
            labels[idm] = w.CommandBars.GetLabelMso(idm).replace("​", "")
        except Exception:
            labels[idm] = None
    R["labels"] = labels
    # 단축키 — 한 가지 보조키만(BuildKeyCode 에 두 개를 주면 이 PC 에서 오류)
    R["shortcuts"] = {}
    for name, code in (("Ctrl+M", 77), ("Ctrl+T", 84), ("Ctrl+Q", 81)):
        try:
            R["shortcuts"][name] = w.FindKey(w.BuildKeyCode(512, code)).Command
        except Exception as e:
            R["shortcuts"][name] = f"err {e}"

    cases = {}

    def para(text, **fmt):
        doc.Content.Delete(); doc.Content.InsertAfter(text)
        p = doc.Paragraphs(1); p.Reset()  # 앞 시험의 들여쓰기가 남지 않게
        p.Alignment = fmt.pop("align", 3)  # 양쪽 맞춤(이 PC 표준 스타일 기본 확인용으로 따로 적음)
        for k, v in fmt.items():
            setattr(p, k, v)
        doc.Repaginate(); w.ScreenRefresh()
        return p

    p = para(P1); R["normal_alignment"] = doc.Styles(-1).ParagraphFormat.Alignment
    cases["없음"] = {"x": lines_x(p), "first": p.FirstLineIndent, "left": p.LeftIndent}
    if LANG == "en":
        p = para(P1, FirstLineIndent=36.0)  # 0.5"
        cases["첫 줄 0.5in"] = {"x": lines_x(p), "first": p.FirstLineIndent}
    else:
        p = para(P1, CharacterUnitFirstLineIndent=2)
        cases["첫 줄 2글자"] = {"x": lines_x(p), "first": p.FirstLineIndent, "char_unit": p.CharacterUnitFirstLineIndent}
        p = para(P1, FirstLineIndent=28.35)  # 1cm = 28.35pt
        cases["첫 줄 1cm"] = {"x": lines_x(p), "first": p.FirstLineIndent}
    p = para("\t" + P1)
    cases["Tab 한 번"] = {"x": lines_x(p), "first_char_x": round(p.Range.Characters(2).Information(X), 2)}
    p = para("    " + P1)
    cases["띄어쓰기 4칸"] = {"x": lines_x(p), "first_char_x": round(p.Range.Characters(5).Information(X), 2)}
    p = para("    " + P2)
    cases["띄어쓰기 4칸(다른 단락)"] = {"x": lines_x(p), "first_char_x": round(p.Range.Characters(5).Information(X), 2)}
    p = para("    " + P1, align=0)
    cases["띄어쓰기 4칸 왼쪽 맞춤"] = {"x": lines_x(p), "first_char_x": round(p.Range.Characters(5).Information(X), 2)}
    if LANG == "en":
        p = para(P1, LeftIndent=36.0)
        cases["왼쪽 0.5in"] = {"x": lines_x(p), "left": p.LeftIndent}
        p = para(P1, LeftIndent=36.0, FirstLineIndent=-36.0)
        cases["내어쓰기 0.5in"] = {"x": lines_x(p), "left": p.LeftIndent, "first": p.FirstLineIndent}
        p = para(ITEM, LeftIndent=18.0, FirstLineIndent=-18.0)
        cases["번호 내어쓰기 0.25in"] = {"x": lines_x(p), "tab_after_number_x": round(p.Range.Characters(4).Information(X), 2)}
        p = para(ITEM)
        cases["번호 내어쓰기 없음"] = {"x": lines_x(p)}
    else:
        p = para(P1, CharacterUnitLeftIndent=2)
        cases["왼쪽 2글자"] = {"x": lines_x(p), "left": p.LeftIndent}
        p = para(P1, CharacterUnitLeftIndent=0, CharacterUnitFirstLineIndent=-2)
        cases["내어쓰기 2글자"] = {"x": lines_x(p), "left": p.LeftIndent, "first": p.FirstLineIndent,
                                "char_first": p.CharacterUnitFirstLineIndent, "char_left": p.CharacterUnitLeftIndent}
        p = para(ITEM, CharacterUnitFirstLineIndent=-2)
        cases["번호 내어쓰기 2글자"] = {"x": lines_x(p), "tab_after_number_x": round(p.Range.Characters(4).Information(X), 2),
                                 "left": p.LeftIndent, "first": p.FirstLineIndent}
        p = para(ITEM)
        cases["번호 내어쓰기 없음"] = {"x": lines_x(p)}
    # 단축키가 부르는 명령(Ctrl+M = Indent, Ctrl+T = HangingIndent) — WordBasic 은 속성처럼 읽기만 해도 실행된다
    for cmd in ("Indent", "HangingIndent", "UnIndent"):
        p = para(P1)
        if cmd == "UnIndent":
            p.LeftIndent = 80.0
        p.Range.Select(); w.Selection.Collapse(1)
        getattr(w.WordBasic, cmd)
        q = doc.Paragraphs(1)
        cases[f"명령 {cmd}"] = {"left": round(q.LeftIndent, 2), "first": round(q.FirstLineIndent, 2), "x": lines_x(q)}
    # 홈 탭 단락 그룹의 두 단추(이름표 '들여쓰기' · '내어쓰기')
    for idm in ("IndentIncreaseWord", "IndentDecreaseWord"):
        p = para(P1)
        if idm == "IndentDecreaseWord":
            p.LeftIndent = 80.0
        p.Range.Select(); w.Selection.Collapse(1)
        before = (round(p.LeftIndent, 2), round(p.FirstLineIndent, 2))
        w.CommandBars.ExecuteMso(idm)
        q = doc.Paragraphs(1)
        cases[f"단추 {idm}"] = {"before": before, "left": round(q.LeftIndent, 2), "first": round(q.FirstLineIndent, 2)}
    # 글꼴을 바꾸면 띄어쓰기 4칸의 폭이 달라지는가(글자 단위 들여쓰기와 비교)
    for font in (("바탕", "Times New Roman") if LANG == "ko" else ("Times New Roman", "Courier New")):
        p = para("    " + P1); p.Range.Font.Name = font; doc.Repaginate(); w.ScreenRefresh()
        sp = lines_x(p)[0]
        p = para(P1); p.Range.Font.Name = font
        if LANG == "ko":
            p.CharacterUnitFirstLineIndent = 2
        else:
            p.FirstLineIndent = 36.0
        doc.Repaginate(); w.ScreenRefresh()
        cases[f"글꼴 {font}"] = {"spaces4_x": sp, "indent_x": lines_x(p)[0], "first": round(p.FirstLineIndent, 2)}
    R["cases"] = cases
    doc.Close(0)
finally:
    w.Quit(0)

out = "measure.json" if LANG == "ko" else "measure_en.json"
json.dump(R, open(os.path.join(HERE, out), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
