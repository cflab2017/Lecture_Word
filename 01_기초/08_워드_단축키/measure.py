"""워드 기초 8강(단축키) 실측 — 창을 띄우지 않는다.
(11/12 처음엔 보이는 워드 창에 키를 보냈는데, 워드는 Shift 가 자주 빠져 Ctrl+Shift+V 가 Ctrl+V 로 들어가 클립보드 글이 붙는 일이 있었다 → 이 방식으로 바꿈)
① 이 PC 워드의 키 지정표(Application.FindKey)에서 단축키마다 묶인 명령 이름을 읽고
② 그 명령을 WordBasic 으로 그대로 실행해(키를 누르면 워드가 실행하는 바로 그 명령) 실습 파일 사본에서 결과를 잰다. 저장하지 않는다.
결과 measure.json(·measure_en.json). 실행: python measure.py [en]"""
import datetime, json, os, shutil, sys, tempfile
import pythoncom
import win32com.client as win32

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
EN = len(sys.argv) > 1 and sys.argv[1] == "en"
SRC = os.path.join(HERE, "practice.docx" if EN else "실습.docx")
R = {"when": datetime.datetime.now().isoformat(timespec="seconds"), "file": os.path.basename(SRC), "how": "FindKey + WordBasic(숨긴 워드)"}
C, S, A = 512, 256, 1024  # wdKeyControl · wdKeyShift · wdKeyAlt
KEYS = {  # 표시 이름: (수식키, 키 코드)
    "Ctrl+B": ((C,), 66), "Ctrl+I": ((C,), 73), "Ctrl+U": ((C,), 85), "Ctrl+E": ((C,), 69), "Ctrl+L": ((C,), 76), "Ctrl+R": ((C,), 82), "Ctrl+J": ((C,), 74),
    "Ctrl+]": ((C,), 221), "Ctrl+[": ((C,), 219), "Ctrl+Shift+>": ((C, S), 190), "Ctrl+Shift+<": ((C, S), 188),
    "Ctrl+Space": ((C,), 32), "Ctrl+Q": ((C,), 81), "Ctrl+1": ((C,), 49), "Ctrl+2": ((C,), 50), "Ctrl+5": ((C,), 53),
    "Ctrl+M": ((C,), 77), "Ctrl+Shift+M": ((C, S), 77), "Ctrl+T": ((C,), 84), "Ctrl+Shift+T": ((C, S), 84),
    "Shift+Enter": ((S,), 13), "Ctrl+Enter": ((C,), 13), "Ctrl+Shift+8": ((C, S), 56), "F4": ((), 115), "Shift+F3": ((S,), 114),
    "Ctrl+Shift+C": ((C, S), 67), "Ctrl+Shift+V": ((C, S), 86), "Ctrl+H": ((C,), 72), "Ctrl+A": ((C,), 65), "Ctrl+Z": ((C,), 90), "Ctrl+Y": ((C,), 89),
}


def main():
    pythoncom.CoInitialize()
    work = os.path.join(tempfile.mkdtemp(), os.path.basename(SRC))
    shutil.copy(SRC, work)
    wd = win32.DispatchEx("Word.Application")
    try:
        wd.Visible = False
        wd.DisplayAlerts = 0
        doc = wd.Documents.Open(work, False, False, False)
        doc.Activate()
        bind = {}
        for name, (mod, code) in KEYS.items():
            try:
                kb = wd.FindKey(wd.BuildKeyCode(*mod, code) if mod else code)  # 수식키는 따로따로 넘겨야 한다(합치면 '매개 변수가 잘못')
                bind[name] = kb.Command
            except Exception as e:
                bind[name] = f"(읽기 실패 {e.args[0] if e.args else e})"
        R["bindings"] = bind
        print("bindings", bind, flush=True)
        WB = wd.WordBasic

        def run(cmd):
            # 글자에 묶인 키(Shift+Enter → 수동 줄 바꿈 문자 \x0b, 키 지정표 종류 '기호')는 그 글자를 넣는다
            if len(cmd) == 1:
                wd.Selection.TypeText(cmd)
                return
            # 숨긴 창에서는 WordBasic 으로 불러도 반영되지 않는 명령이 있었다(Bold · CenterPara · SpacePara · ResetChar · ResetPara · ChangeCase …)
            # → 같은 명령의 개체 모델 메서드로(워드 도움말의 명령 설명과 같은 동작). 시스템 클립보드는 건드리지 않는다.
            sel = wd.Selection
            EQUIV = {
                "Bold": lambda: setattr(sel.Font, "Bold", 9999998), "Italic": lambda: setattr(sel.Font, "Italic", 9999998),
                "Underline": lambda: setattr(sel.Font, "Underline", 1 if sel.Font.Underline == 0 else 0),
                "CenterPara": lambda: setattr(sel.ParagraphFormat, "Alignment", 1), "LeftPara": lambda: setattr(sel.ParagraphFormat, "Alignment", 0),
                "RightPara": lambda: setattr(sel.ParagraphFormat, "Alignment", 2), "JustifyPara": lambda: setattr(sel.ParagraphFormat, "Alignment", 3),
                "SpacePara1": lambda: sel.ParagraphFormat.Space1(), "SpacePara2": lambda: sel.ParagraphFormat.Space2(), "SpacePara15": lambda: sel.ParagraphFormat.Space15(),
                "ResetChar": lambda: sel.Font.Reset(), "ResetPara": lambda: sel.ParagraphFormat.Reset(),
                "ChangeCase": lambda: setattr(sel.Range, "Case", -1),  # wdNextCase
                "ShowAll": lambda: setattr(wd.ActiveWindow.View, "ShowAll", not wd.ActiveWindow.View.ShowAll),
                "EditRedoOrRepeat": lambda: wd.Repeat(),
            }
            if cmd in EQUIV:
                EQUIV[cmd]()
                return
            # WordBasic 은 늦은 바인딩에서 '속성'으로 읽히는 이름이 있어(Bold 등) 메서드로 직접 부른다
            o = WB._oleobj_
            try:
                o.Invoke(o.GetIDsOfNames(cmd), 0, pythoncom.DISPATCH_METHOD, True)
            except Exception as e:
                R.setdefault("run_errors", []).append([cmd, str(e.args[2][2] if len(e.args) > 2 and e.args[2] else e)])

        P = lambda i: doc.Paragraphs(i).Range
        S_ = wd.Selection

        def sel_text(i):
            P(i).Select(); S_.MoveEnd(1, -1)

        font = lambda i: {"bold": P(i).Font.Bold, "size": P(i).Font.Size, "underline": P(i).Font.Underline}
        sel_text(1)
        R["title_before"] = {**font(1), "align": P(1).ParagraphFormat.Alignment}
        run(bind["Ctrl+B"]); run(bind["Ctrl+E"])
        sizes = []
        for _ in range(3):
            run(bind["Ctrl+]"]); sizes.append(P(1).Font.Size)
        R["ctrl_bracket_right_sizes_from_11"] = sizes
        run(bind["Ctrl+["]); R["ctrl_bracket_left_size"] = P(1).Font.Size
        steps = []
        for _ in range(4):
            run(bind["Ctrl+Shift+>"]); steps.append(P(1).Font.Size)
        R["ctrl_shift_gt_steps_from_13"] = steps
        run(bind["Ctrl+Shift+<"]); R["ctrl_shift_lt_after"] = P(1).Font.Size
        R["title_after"] = {**font(1), "align": P(1).ParagraphFormat.Alignment}
        run(bind["Ctrl+Space"]); R["ctrl_space_title"] = {**font(1), "align": P(1).ParagraphFormat.Alignment}
        run(bind["Ctrl+Q"]); R["ctrl_q_title"] = {**font(1), "align": P(1).ParagraphFormat.Alignment}
        sel_text(2); al = {}
        for k in ("Ctrl+R", "Ctrl+J", "Ctrl+E", "Ctrl+L"):
            run(bind[k]); al[k] = P(2).ParagraphFormat.Alignment
        R["align_p2"] = al
        ls = {}
        for k in ("Ctrl+2", "Ctrl+5", "Ctrl+1"):
            run(bind[k]); ls[k] = {"rule": P(2).ParagraphFormat.LineSpacingRule, "spacing": P(2).ParagraphFormat.LineSpacing}
        R["line_spacing_p2"] = ls
        sel_text(3); ind = []
        for k in ("Ctrl+M", "Ctrl+M", "Ctrl+Shift+M", "Ctrl+Shift+M", "Ctrl+T", "Ctrl+T", "Ctrl+Shift+T"):
            run(bind[k]); ind.append({"key": k, "left": round(P(3).ParagraphFormat.LeftIndent, 2), "first": round(P(3).ParagraphFormat.FirstLineIndent, 2)})
        R["indent_p3"] = ind
        run(bind["Ctrl+Q"]); R["ctrl_q_p3"] = {"left": P(3).ParagraphFormat.LeftIndent, "first": P(3).ParagraphFormat.FirstLineIndent}
        R["default_tab_stop"] = doc.DefaultTabStop
        n0 = doc.Paragraphs.Count
        P(3).Select(); S_.MoveEnd(1, -1); S_.Collapse(0)
        run(bind["Shift+Enter"])
        R["shift_enter"] = {"paras_before": n0, "paras_after": doc.Paragraphs.Count, "char_code": ord(doc.Range(S_.Start - 1, S_.Start).Text)}
        S_.TypeText("비가 오면 28일(토)로 미룹니다" if not EN else "If it rains, we move to November 28")
        R["p3_after_shift_enter"] = P(3).Text.replace("\x0b", "↵").replace("\r", "¶")
        P(4).Select(); S_.MoveEnd(1, -1); S_.Collapse(0)
        n1 = doc.Paragraphs.Count
        S_.TypeParagraph()
        R["enter"] = {"paras_before": n1, "paras_after": doc.Paragraphs.Count, "space_after_pt": P(4).ParagraphFormat.SpaceAfter}
        doc.Undo(1)
        pages0 = doc.ComputeStatistics(2)
        S_.EndKey(6)
        run(bind["Ctrl+Enter"])
        R["ctrl_enter"] = {"pages_before": pages0, "pages_after": doc.ComputeStatistics(2), "char_code": ord(doc.Range(S_.Start - 1, S_.Start).Text)}
        doc.Undo(1)
        R["undo_after_ctrl_enter_pages"] = doc.ComputeStatistics(2)
        v = wd.ActiveWindow.View
        b0 = v.ShowAll; run(bind["Ctrl+Shift+8"]); b1 = v.ShowAll; run(bind["Ctrl+Shift+8"])
        R["ctrl_shift_8"] = {"before": b0, "after": b1, "again": v.ShowAll}
        sel_text(1); S_.Font.Bold = True; S_.Font.Size = 14; S_.Font.Color = 0x0000C0
        S_.Collapse(1); S_.MoveRight(1, 2)
        run(bind["Ctrl+Shift+C"])
        sel_text(6); run(bind["Ctrl+Shift+V"])
        R["format_paste_p6"] = {"bold": P(6).Font.Bold, "size": P(6).Font.Size, "color": P(6).Font.Color, "align": P(6).ParagraphFormat.Alignment}
        rng = doc.Content
        rng.Find.Execute("photo walk")
        rng.Select(); cs = []
        for _ in range(4):
            run(bind["Shift+F3"]); cs.append(S_.Text)
        R["shift_f3_cycle"] = cs
        # F4 는 새로 연 사본에서(앞 단계에서 단락이 바뀌어 있으므로)
        doc.Close(False); doc = wd.Documents.Open(work, False, False, False); doc.Activate(); S_ = wd.Selection
        sel_text(4); run(bind["Ctrl+U"])
        sel_text(5); run(bind["F4"])
        R["f4_repeat"] = {"p4_underline": P(4).Font.Underline, "p5_underline_after_F4": P(5).Font.Underline, "p3_untouched": P(3).Font.Underline}
        R["doc_text_end"] = doc.Content.Text.replace("\x0b", "↵").replace("\r", "¶").replace("\x0c", "[쪽]")
        doc.Close(False)
    finally:
        wd.Quit()
    name = "measure_en.json" if EN else "measure.json"
    with open(os.path.join(HERE, name), "w", encoding="utf-8") as fp:
        json.dump(R, fp, ensure_ascii=False, indent=2, default=str)
    print(json.dumps(R, ensure_ascii=False, indent=1, default=str))


main()
