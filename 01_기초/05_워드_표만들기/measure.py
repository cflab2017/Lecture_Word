"""워드 기초 5강 — 표 만들기 실측. 이 PC 워드(Microsoft 365)를 따로 띄워(DispatchEx) 표를 넣고 값을 읽는다.
- 표 삽입: Tables.Add(행, 열) — 기본 스타일 · 열 너비(pt)
- 텍스트를 표로: 탭 · 쉼표로 나눈 글을 ConvertToTable → 행 · 열 수
- 머리글 행 반복: 45행 표(2쪽)를 PDF 로 저장 → 머리글 글자가 몇 쪽에 나오는지(HeadingFormat 끔/켬)
- 자동 맞춤: AutoFitBehavior 0(고정) · 1(내용) · 2(창) 뒤 열 너비
- 문서 맨 위 표 앞에 줄: 첫 칸에서 SplitTable(= Ctrl+Shift+Enter) → 표 앞 단락 수
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 · pymupdf 필요)"""
import datetime, json, os, sys, tempfile
import pythoncom
import win32com.client as win32
import fitz

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}
clean = lambda s: (s or "").replace("​", "").replace("\r\x07", "").replace("\r", "").strip()
HEAD = ["No", "Name", "Phone"] if EN else ["번호", "이름", "연락처"]
TMP = tempfile.mkdtemp()


def pdf_pages_with(doc, word):
    p = os.path.join(TMP, "t.pdf")
    doc.ExportAsFixedFormat(p, 17)
    with fitz.open(p) as f:
        return {"pages": f.page_count, "pages_with_head": [i + 1 for i in range(f.page_count) if word in f[i].get_text()]}


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["build"] = w.Build
    R["labels"] = {}
    for k in ("TableInsertGallery", "TableInsert", "ConvertTextToTable", "TableRepeatHeaderRows", "TableAutoFitContents",
              "TableAutoFitWindow", "TableAutoFitFixedColumnWidth", "TableSplitTable", "TableAutoFitMenu"):
        try:
            R["labels"][k] = clean(w.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            R["labels"][k] = "ERR " + str(e)[:60]

    # ① 표 삽입 기본값
    d = w.Documents.Add()
    t0 = d.Tables.Add(d.Range(0, 0), 3, 4)  # COM 기본(wdWord8TableBehavior)
    R["insert_com_default"] = {"style": str(t0.Style), "borders": bool(t0.Borders.Enable)}
    d.Close(0)
    d = w.Documents.Add()
    t = d.Tables.Add(d.Range(0, 0), 3, 4, 1, 0)  # 화면 삽입과 같은 wdWord9TableBehavior · 고정 너비
    R["insert_default"] = {"style": str(t.Style), "rows": t.Rows.Count, "cols": t.Columns.Count,
                           "col_widths_pt": [round(t.Columns(i).Width, 1) for i in range(1, 5)],
                           "page_width_pt": round(d.PageSetup.PageWidth, 1),
                           "text_width_pt": round(d.PageSetup.PageWidth - d.PageSetup.LeftMargin - d.PageSetup.RightMargin, 1),
                           "borders": bool(t.Borders.Enable)}
    d.Close(0)

    # ② 텍스트를 표로 — 탭 / 쉼표
    d = w.Documents.Add()
    rows = [HEAD] + [[str(i), f"{'Kim' if EN else '김'}{i}", f"010-0000-{i:04d}"] for i in range(1, 4)]
    d.Range().Text = "\r".join("\t".join(r) for r in rows)
    t = d.Range().ConvertToTable(Separator=1)  # wdSeparateByTabs
    R["text_to_table_tab"] = {"rows": t.Rows.Count, "cols": t.Columns.Count, "cell_1_2": clean(t.Cell(1, 2).Range.Text)}
    d.Close(0)
    d = w.Documents.Add()
    rows2 = [["Seoul, Korea", "2"], ["Busan", "3"]] if EN else [["서울, 강남", "2"], ["부산", "3"]]
    d.Range().Text = "\r".join(",".join(r) for r in rows2)  # 쉼표가 내용 안에도 있는 경우
    t = d.Range().ConvertToTable(Separator=2)  # wdSeparateByCommas
    R["text_to_table_comma_inside"] = {"source": [",".join(r) for r in rows2], "rows": t.Rows.Count, "cols": t.Columns.Count,
                                       "row1": [clean(t.Cell(1, c).Range.Text) for c in range(1, t.Columns.Count + 1)]}
    d.Close(0)

    # ③ 머리글 행 반복 — 45행
    d = w.Documents.Add()
    n = 45
    t = d.Tables.Add(d.Range(0, 0), n + 1, 3)
    for c, h in enumerate(HEAD, start=1):
        t.Cell(1, c).Range.Text = h
    for i in range(1, n + 1):
        t.Cell(i + 1, 1).Range.Text = str(i)
        t.Cell(i + 1, 2).Range.Text = f"{'Name' if EN else '이름'} {i}"
        t.Cell(i + 1, 3).Range.Text = f"010-0000-{i:04d}"
    hw = HEAD[2]
    off = pdf_pages_with(d, hw)
    t.Rows(1).HeadingFormat = True
    on = pdf_pages_with(d, hw)
    R["repeat_header"] = {"rows": n + 1, "off": off, "on": on}
    # 머리글 반복이 둘째 행부터면 안 먹는다 — 1행 끄고 2행만 켬
    t.Rows(1).HeadingFormat = False
    try:
        t.Rows(2).HeadingFormat = True
        R["repeat_header"]["row2_only"] = pdf_pages_with(d, hw)
        R["repeat_header"]["row2_only"]["row1_flag"] = bool(t.Rows(1).HeadingFormat)
        R["repeat_header"]["row2_only"]["row2_flag"] = bool(t.Rows(2).HeadingFormat)
    except Exception as e:  # noqa: BLE001
        R["repeat_header"]["row2_only"] = "ERR " + str(e)[:80]
    d.Close(0)

    # ④ 자동 맞춤
    d = w.Documents.Add()
    t = d.Tables.Add(d.Range(0, 0), 3, 3, 1, 0)
    vals = [HEAD, ["1", "Kim" if EN else "김민준", "010-1234-5678"], ["2", "Lee" if EN else "이서연", "010-9876-5432"]]
    for r, row in enumerate(vals, start=1):
        for c, v in enumerate(row, start=1):
            t.Cell(r, c).Range.Text = v
    res = {}
    # 11/7: 새 표(고정 너비)에 바로 AutoFitBehavior(1) 을 부르면 너비가 안 바뀐다(150.2 그대로) — 창(2) 뒤에 내용(1)이면 줄어든다.
    # 화면의 [자동 맞춤 → 내용에 자동으로 맞춤] 결과를 재려고 창 → 내용 → 고정(지금 너비로 고정) 순서로 잰다.
    for name, beh in (("window", 2), ("content", 1), ("fixed_after_content", 0)):
        t.AutoFitBehavior(beh)
        d.Repaginate()
        res[name] = [round(t.Cell(1, i).Width, 1) for i in range(1, 4)]
        res[name + "_total"] = round(sum(res[name]), 1)
    R["autofit"] = res
    d.Close(0)

    # ⑤ 문서 맨 위 표 앞에 줄 넣기
    d = w.Documents.Add()
    t = d.Tables.Add(d.Range(0, 0), 2, 2)
    before = d.Range(0, t.Range.Start).Paragraphs.Count if t.Range.Start > 0 else 0
    t.Cell(1, 1).Range.Select()
    w.Selection.Collapse(1)
    w.Selection.SplitTable()
    t2 = d.Tables(1)
    R["text_above_table"] = {"table_start_before": 0, "table_start_after": t2.Range.Start,
                             "paragraphs_before_table": d.Range(0, t2.Range.Start).Paragraphs.Count if t2.Range.Start > 0 else 0,
                             "tables": d.Tables.Count}
    d.Close(0)
finally:
    w.Quit()
out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
with open(out, "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1, default=str)
print(json.dumps(R, ensure_ascii=False, indent=1, default=str))
