"""워드 기초 1강 줄 간격 — 강의에 적은 값을 워드(Microsoft 365)로 직접 잰다. 결과: measure.json
워드를 따로(보이지 않게) 띄워 새 문서를 만들고, 저장하지 않고 닫는다. Normal.dotm(기본 서식 파일)은 건드리지 않는다.
잰 것: 기본 서식(글꼴·크기·줄 간격·단락 뒤), 줄 간격 설정마다 줄과 줄 사이 거리(pt), 단락 사이 거리,
Shift+Enter(줄 바꿈) · '같은 스타일 단락 사이 공백 없음' · 줄 격자 · 단축키(Ctrl+1/2/5/0)가 무슨 명령인지.
실행: python measure.py  (pip install pywin32)"""
import json, os, sys
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
R = {}

TEXT = ("이번 주 팀 회의는 화요일 오전 10시에 3층 대회의실에서 합니다. 지난달 매출 보고와 다음 분기 계획을 "
        "나누고, 각 팀은 진행 상황을 5분씩 발표합니다. 발표 자료는 월요일 오후 6시까지 팀 공유 폴더에 올려 주세요.")
POS = 6  # wdVerticalPositionRelativeToPage


def line_tops(rng):
    """단락 범위 안 글자마다 세로 위치를 재서 줄 머리 위치 목록을 돌려준다(pt)."""
    tops = []
    for i in range(1, rng.Characters.Count + 1):
        y = round(rng.Characters(i).Information(POS), 2)
        if not tops or abs(y - tops[-1]) > 0.5:
            tops.append(y)
    return tops


def pitch(tops):
    d = [round(b - a, 2) for a, b in zip(tops, tops[1:])]
    return d


w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["version"] = f"{w.Version} build {w.Build}"
    doc = w.Documents.Add()
    doc.ActiveWindow.View.Type = 3  # 인쇄 모양
    st = doc.Styles(-1)  # wdStyleNormal
    pf = st.ParagraphFormat
    R["normal"] = {
        "font": st.Font.Name, "font_fareast": st.Font.NameFarEast, "size": st.Font.Size,
        "rule": pf.LineSpacingRule, "line_spacing_pt": pf.LineSpacing, "lines": round(pf.LineSpacing / 12, 3),
        "space_before": pf.SpaceBefore, "space_after": pf.SpaceAfter,
        "no_space_same_style": bool(st.NoSpaceBetweenParagraphsOfSameStyle),
        "disable_line_grid": bool(pf.DisableLineHeightGrid),
        "layout_mode": doc.PageSetup.LayoutMode,  # 0 기본, 1 격자, 2 줄 격자, 3 글자 격자
        "label": w.CommandBars.GetLabelMso("LineSpacingMenu") if True else None,
    }

    # 시험 문단 3개
    doc.Content.Text = TEXT + "\r" + TEXT + "\r" + TEXT
    doc.Repaginate()

    def apply(rule, value=None, after=None, font=None, size=None):
        p = doc.Content.ParagraphFormat
        p.LineSpacingRule = rule
        if value is not None:
            p.LineSpacing = value
        if after is not None:
            p.SpaceAfter = after
        if font:
            doc.Content.Font.Name = font
            try:  # 한글 글꼴 칸에는 영문 글꼴을 넣을 수 없다(오류) — 그때는 한글 칸을 그대로 둔다
                doc.Content.Font.NameFarEast = font
            except Exception:
                pass
        if size:
            doc.Content.Font.Size = size
        doc.Repaginate()
        p1 = line_tops(doc.Paragraphs(1).Range)
        p2 = line_tops(doc.Paragraphs(2).Range)
        return {"lines_in_par1": len(p1), "line_pitch": pitch(p1), "par_gap": round(p2[0] - p1[-1], 2)}

    # 기본 상태 그대로
    R["default"] = apply(pf.LineSpacingRule, pf.LineSpacing)
    # 단락 뒤 간격 0 으로 두고 줄 간격만 바꾼다
    cases = [
        ("1.0 (한 줄)", 0, None), ("1.15 (배수)", 5, round(12 * 1.15, 2)), ("1.5 줄", 1, None),
        ("2.0 (두 줄)", 2, None), ("3.0 (배수)", 5, round(12 * 3, 2)),
        ("고정 값 20pt", 4, 20), ("최소 20pt", 3, 20), ("고정 값 8pt(글자보다 작게)", 4, 8),
    ]
    R["by_rule_after0"] = {}
    for name, rule, val in cases:
        R["by_rule_after0"][name] = apply(rule, val, after=0)
    # 단락 뒤 간격
    R["after"] = {
        "1.0 + 뒤 0": apply(0, None, after=0),
        "1.0 + 뒤 8pt": apply(0, None, after=8),
        "1.0 + 뒤 12pt": apply(0, None, after=12),
        "2.0 + 뒤 8pt": apply(2, None, after=8),
    }
    # 글꼴에 따라 '한 줄' 높이가 다르다
    R["single_by_font"] = {}
    for font, size in [("맑은 고딕", 10), ("맑은 고딕", 11), ("바탕", 11), ("굴림", 11), ("돋움", 11)]:
        r = apply(0, None, after=0, font=font, size=size)
        r2 = apply(2, None, after=0)
        R["single_by_font"][f"{font} {size}pt"] = {"single": r["line_pitch"][:1], "double": r2["line_pitch"][:1],
                                                     "font_used": doc.Paragraphs(1).Range.Font.NameFarEast}
    # 글꼴 원래대로
    doc.Content.Font.Name = st.Font.Name; doc.Content.Font.NameFarEast = st.Font.NameFarEast; doc.Content.Font.Size = st.Font.Size

    # Enter 대신 Shift+Enter(줄 바꿈 글자, 코드 11) — 단락 뒤 간격이 안 붙는다
    doc.Content.Text = "첫째 줄\v둘째 줄(Shift+Enter)\r셋째 줄(Enter)"
    doc.Content.ParagraphFormat.LineSpacingRule = 0
    doc.Content.ParagraphFormat.SpaceAfter = 8
    doc.Repaginate()
    ys = [round(doc.Paragraphs(1).Range.Characters(1).Information(POS), 2),
          round(doc.Paragraphs(1).Range.Characters(6).Information(POS), 2),
          round(doc.Paragraphs(2).Range.Characters(1).Information(POS), 2)]
    R["shift_enter"] = {"paragraphs": doc.Paragraphs.Count, "tops": ys,
                        "gap_shift_enter": round(ys[1] - ys[0], 2), "gap_enter": round(ys[2] - ys[1], 2)}

    # 같은 스타일 단락 사이 공백 없음
    doc.Content.Text = TEXT + "\r" + TEXT
    p = doc.Content.ParagraphFormat
    p.LineSpacingRule = 0; p.SpaceAfter = 8; p.SpaceBefore = 0
    doc.Repaginate()
    g0 = round(line_tops(doc.Paragraphs(2).Range)[0] - line_tops(doc.Paragraphs(1).Range)[-1], 2)
    try:
        st.NoSpaceBetweenParagraphsOfSameStyle = True  # 이 문서의 표준 스타일만(Normal.dotm 아님)
        doc.Repaginate()
        g1 = round(line_tops(doc.Paragraphs(2).Range)[0] - line_tops(doc.Paragraphs(1).Range)[-1], 2)
    except Exception as e:
        g1 = f"error {e}"
    R["no_space_same_style"] = {"gap_off": g0, "gap_on": g1}
    st.NoSpaceBetweenParagraphsOfSameStyle = False

    # 줄 격자(페이지 설정 → 문서 격자 → 줄 격자만) — 한 줄을 격자 칸에 맞춘다
    doc.Content.ParagraphFormat.SpaceAfter = 0
    grid = {}
    for lm, label in [(0, "격자 없음"), (2, "줄 격자")]:
        doc.PageSetup.LayoutMode = lm
        if lm == 2:
            grid["lines_page"] = doc.PageSetup.LinesPage
        for dis in (False, True):
            doc.Content.ParagraphFormat.DisableLineHeightGrid = dis
            for rule, val, nm in [(0, None, "1.0"), (5, round(12 * 1.15, 2), "1.15"), (1, None, "1.5")]:
                doc.Content.ParagraphFormat.LineSpacingRule = rule
                if val:
                    doc.Content.ParagraphFormat.LineSpacing = val
                doc.Repaginate()
                grid[f"{label} · 격자에 맞춤 {'끔' if dis else '켬'} · {nm}"] = pitch(line_tops(doc.Paragraphs(1).Range))[:1]
    doc.PageSetup.LayoutMode = 0
    R["line_grid"] = grid

    # 단축키가 어떤 명령에 묶여 있는지
    K = 512  # wdKeyControl
    keys = {"Ctrl+1": 49, "Ctrl+2": 50, "Ctrl+5": 53, "Ctrl+0": 48}
    R["shortcuts"] = {}
    for name, code in keys.items():
        try:
            R["shortcuts"][name] = w.FindKey(w.BuildKeyCode(K, code)).Command
        except Exception as e:
            R["shortcuts"][name] = f"error {e}"

    # 리본 이름(한국어 화면)
    R["labels"] = {}
    for mso in ["LineSpacingMenu", "ParagraphDialog", "GroupParagraph", "LineSpacingOptions", "AddSpaceBeforeParagraph", "RemoveSpaceAfterParagraph", "SetAsDefault"]:
        try:
            R["labels"][mso] = w.CommandBars.GetLabelMso(mso)
        except Exception as e:
            R["labels"][mso] = None

    doc.Close(0)
finally:
    w.Quit(0)

json.dump(R, open(os.path.join(HERE, "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
