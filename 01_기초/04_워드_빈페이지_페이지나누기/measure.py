"""워드 기초 4강 — 빈 페이지 삭제 · 페이지 나누기 실측. 이 PC 워드(Microsoft 365)를 숨긴 채 따로 띄워(DispatchEx)
빈 페이지가 생기는 네 가지(빈 단락 여러 개 · 끝의 페이지 나누기 · 끝의 구역 나누기 · 쪽을 꽉 채운 표 뒤의 마지막 단락)를 만들고,
쪽수(ComputeStatistics 2)와 PDF 쪽수를 잰 뒤 고쳐서 다시 잰다. 단축키는 워드 키 연결표(FindKey)에서 읽는다.
결과 measure.json(argv 'en' 이면 영어 글로 measure_en.json) · 실습.docx · 정답.docx (practice.docx · answer.docx).
실행: python measure.py [en]  (pywin32 · pymupdf 필요)"""
import datetime, json, os, sys, tempfile
import pythoncom
import win32com.client as win32
import pymupdf

sys.stdout.reconfigure(encoding="utf-8")
from measure_text import KO, EN, build

LANG_EN = "en" in sys.argv[1:]
T = EN if LANG_EN else KO
HERE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix="wd04_")
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if LANG_EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()


def pages(d):
    d.Repaginate()
    return d.ComputeStatistics(2)   # wdStatisticPages


def pdf_pages(d, tag):
    p = os.path.join(TMP, tag + ".pdf")
    d.ExportAsFixedFormat(p, 17)
    doc = pymupdf.open(p)
    n = len(doc)
    last_text = doc[n - 1].get_text().strip()
    doc.close()
    return {"pdf_pages": n, "last_page_chars": len(last_text)}


def brk(rng, kind):
    """범위 앞에 나누기를 넣는다 — 펼친 범위에 InsertBreak 를 부르면 그 글이 나누기로 바뀐다(11/6 첫 실행에서 제목이 사라짐)."""
    r = rng.Duplicate
    r.Collapse(1)                   # wdCollapseStart
    r.InsertBreak(kind)


def page_of(rng):
    return rng.Information(3)       # wdActiveEndPageNumber


def replace_all(d, what):
    """찾아 바꾸기(모두 바꾸기) — 이름 붙인 인수(FindText=…)로 부르면 11/6 첫 실행에서 아무것도 안 바뀌어 위치 인수로 부른다."""
    f = d.Content.Find
    f.ClearFormatting()
    f.Replacement.ClearFormatting()
    n = d.Content.Text.count(chr(12))
    ok = f.Execute(what, False, False, False, False, False, True, 1, False, "", 2)   # wdFindStop · wdReplaceAll
    return {"found": bool(ok), "chr12_before": n, "chr12_after": d.Content.Text.count(chr(12))}


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["build"] = w.Build
    # 0) 단축키 — 키 연결표
    keys = {}
    for name, codes in (("Ctrl+Enter", (512, 13)), ("Ctrl+Shift+8", (512, 256, 56)), ("Ctrl+Shift+Enter", (512, 256, 13))):
        try:
            kc = w.BuildKeyCode(*codes)
            keys[name] = w.FindKey(kc).Command
        except Exception as e:  # noqa: BLE001
            keys[name] = "ERR " + str(e)[:60]
    R["keys"] = keys
    R["labels"] = {}
    for k in ("PageBreakInsertWord", "ParagraphMarks", "BreaksGallery"):
        try:
            R["labels"][k] = clean(w.CommandBars.GetLabelMso(k))
        except Exception as e:  # noqa: BLE001
            R["labels"][k] = "ERR " + str(e)[:50]

    def fresh(rows=0):
        d = w.Documents.Add()
        t = build(d, T, rows)
        return d, t

    # 1) 페이지 나누기(Ctrl+Enter) — '2. 쓴 곳' 앞에 넣으면 2쪽 맨 위에서 시작
    d, _ = fresh()
    base = pages(d)
    brk(d.Paragraphs(4).Range, 7)    # wdPageBreak — '2. 쓴 곳' 앞
    p_after = pages(d)
    h2 = d.Paragraphs(5).Range      # 나누기 문자가 앞 단락에 들어가 번호가 한 칸 밀림
    R["page_break"] = {"pages_before": base, "pages_after": p_after, "heading_page": page_of(h2),
                       "heading_top_pt": round(h2.Information(6), 1), "heading_text": h2.Text.strip()[:12]}
    # 위에 다섯 줄을 더 써도 2쪽 맨 위 그대로인지
    d.Paragraphs(3).Range.InsertParagraphAfter()
    d.Paragraphs(4).Range.Text = "\r".join(["추가한 줄"] * 5 if not LANG_EN else ["Added line"] * 5)
    hh = [p for p in d.Paragraphs if p.Range.Text.strip().startswith(T["p2"][0][:4])][0].Range
    R["page_break"]["after_adding_5_lines"] = {"heading_page": page_of(hh), "heading_top_pt": round(hh.Information(6), 1), "pages": pages(d)}
    # 페이지 나누기 지우기 = 나누기 문자(chr 12) 지우기
    rep = replace_all(d, "^m")
    R["page_break"]["after_delete"] = {"pages": pages(d), "heading_page": page_of(hh), **rep}
    d.Close(0)
    # 1b) 엔터로 밀어서 다음 쪽 맨 위에 맞춘 경우 — 위에 다섯 줄 더 쓰면 제목이 밀려 내려감
    d, _ = fresh()
    h = d.Paragraphs(4).Range
    n_enter = 0
    while page_of(d.Paragraphs(4 + n_enter).Range) == 1 and n_enter < 80:
        d.Paragraphs(3 + n_enter).Range.InsertParagraphAfter()
        n_enter += 1
    hh = d.Paragraphs(4 + n_enter).Range
    R["enter_push"] = {"enters_needed": n_enter, "heading_page": page_of(hh), "heading_top_pt": round(hh.Information(6), 1)}
    d.Paragraphs(3).Range.InsertParagraphAfter()
    d.Paragraphs(4).Range.Text = "\r".join(["추가한 줄"] * 5 if not LANG_EN else ["Added line"] * 5)
    hh = [p for p in d.Paragraphs if p.Range.Text.strip().startswith(T["p2"][0][:4])][0].Range
    R["enter_push"]["after_adding_5_lines"] = {"heading_page": page_of(hh), "heading_top_pt": round(hh.Information(6), 1), "pages": pages(d)}
    d.Close(0)

    # 2) 빈 페이지 원인 ① 빈 단락 여러 개
    d, _ = fresh()
    b = pages(d)
    for _ in range(60):
        d.Content.InsertParagraphAfter()
    R["blank_enters"] = {"pages_before": b, "pages_with_60_empty": pages(d), **pdf_pages(d, "enters")}
    # 끝의 빈 단락 지우기(글자 있는 마지막 단락 뒤 전부)
    last_text = max(i for i in range(1, d.Paragraphs.Count + 1) if d.Paragraphs(i).Range.Text.strip())
    d.Range(d.Paragraphs(last_text).Range.End - 1, d.Content.End - 1).Delete()
    R["blank_enters"]["after_delete"] = pages(d)
    d.Close(0)
    # 2) ② 끝의 페이지 나누기
    d, _ = fresh()
    b = pages(d)
    d.Content.Paragraphs.Last.Range.InsertParagraphAfter()
    brk(d.Paragraphs.Last.Range, 7)
    R["blank_page_break"] = {"pages_before": b, "pages_after_break": pages(d), **pdf_pages(d, "pb")}
    rep = replace_all(d, "^m")
    R["blank_page_break"]["after_delete"] = {"pages": pages(d), **rep}
    d.Close(0)
    # 2) ③ 끝의 구역 나누기(다음 페이지부터)
    d, _ = fresh()
    b = pages(d)
    d.Content.Paragraphs.Last.Range.InsertParagraphAfter()
    brk(d.Paragraphs.Last.Range, 2)  # wdSectionBreakNextPage
    R["blank_section_break"] = {"pages_before": b, "pages_after_break": pages(d), "sections": d.Sections.Count, **pdf_pages(d, "sb")}
    rep = replace_all(d, "^b")
    R["blank_section_break"]["after_delete"] = {"pages": pages(d), "sections": d.Sections.Count, **rep}
    d.Close(0)
    # 2) ④ 쪽을 꽉 채운 표 뒤의 마지막 단락 — 줄 하나(27.7pt)가 끝 단락(약 23pt)과 비슷해 줄 수만 늘려서는 안 생긴다
    #    (11/6 첫 실행: 10~59줄 모두 끝 단락이 표와 같은 쪽). 그래서 줄을 쪽 끝까지 더한 뒤 표 위 단락 아래 간격을 1pt 씩 늘린다.
    def table_end_blank(d, t):
        """표 줄을 쪽 끝까지 더하고(넘치면 한 줄 뺌) 표 바로 위 단락 아래 간격을 1pt 씩 늘려
        '표 마지막 줄은 앞 쪽, 끝 단락만 다음 쪽' 을 만든다. 성공하면 (줄 수, 더한 간격) 아니면 None."""
        def last_row_page():
            d.Repaginate()
            return page_of(t.Rows(t.Rows.Count).Range)
        p0 = last_row_page()
        for _ in range(80):
            n = t.Rows.Count
            row = t.Rows.Add()
            vals = [T["table_row"][0].format(d=1 + (n - 1) % 30), T["table_row"][1], T["table_row"][2]]
            for c, v in enumerate(vals, 1):
                t.Cell(n + 1, c).Range.Text = v
            p1 = last_row_page()
            if p1 > p0:
                t.Rows(t.Rows.Count).Delete()
                break
            p0 = p1
        spacer = [p for p in d.Paragraphs if p.Range.Text.strip().startswith(T["p2"][1][:8])][0]
        base_sa = spacer.SpaceAfter
        for add in range(1, 40):
            spacer.SpaceAfter = base_sa + add
            lr = last_row_page()
            ep = page_of(d.Paragraphs.Last.Range)
            if ep > lr:
                return {"table_rows": t.Rows.Count - 1, "spacer_space_after": base_sa + add, "table_last_row_page": lr, "end_para_page": ep}
            if lr > p0:
                return None
        return None

    def thin_end_para(d):
        pf = d.Paragraphs.Last
        pf.Range.Font.Size = 1
        pf.SpaceBefore = 0
        pf.SpaceAfter = 0
        pf.LineSpacingRule = 4    # wdLineSpaceExactly
        pf.LineSpacing = 1

    d, t = fresh(5)
    rec = table_end_blank(d, t)
    if rec:
        rec.update({"pages": pages(d), **pdf_pages(d, "tbl")})
        # 지워 보기 — 끝 단락(표 뒤) 을 Delete · Backspace(TypeBackspace) 로
        before_len = len(d.Content.Text)
        d.Paragraphs.Last.Range.Delete()
        rec["after_delete_try"] = {"pages": pages(d), "content_len_change": len(d.Content.Text) - before_len}
        # 고치기 1: 글자 크기 1pt 만
        d.Paragraphs.Last.Range.Font.Size = 1
        rec["fix_font_1pt"] = {"pages": pages(d), "end_para_page": page_of(d.Paragraphs.Last.Range)}
        # 고치기 2: 단락 앞뒤 간격 0 + 줄 간격 고정 1pt 까지
        thin_end_para(d)
        rec["fix_spacing_too"] = {"pages": pages(d), "end_para_page": page_of(d.Paragraphs.Last.Range), **pdf_pages(d, "tbl_fixed")}
        # 고치기 3(다른 길): 끝 단락을 숨김 글자로 — 화면·인쇄 쪽수
        d2, t2 = fresh(5)
        r2 = table_end_blank(d2, t2)
        if r2:
            d2.Paragraphs.Last.Range.Font.Hidden = True
            rec["fix_hidden"] = {"pages": pages(d2), **pdf_pages(d2, "tbl_hidden")}
        d2.Close(0)
        R["blank_table_end"] = rec
    else:
        R["blank_table_end"] = {"error": "표 끝 사례를 못 찾음"}
    d.Close(0)

    # 실습·정답 파일 — 실습: '2. 쓴 곳' 을 엔터로 2쪽에 밀어 둠 + 2쪽을 꽉 채운 표 뒤 끝 단락만 3쪽(빈 페이지)
    #                    정답: 엔터 대신 페이지 나누기(Ctrl+Enter) + 끝 단락 1pt·간격 0 → 2쪽
    names = ("practice.docx", "answer.docx") if LANG_EN else ("실습.docx", "정답.docx")
    for kind, nm in zip(("practice", "answer"), names):
        d, t = fresh(5)
        if kind == "practice":
            n = 0
            while page_of(d.Paragraphs(4 + n).Range) == 1 and n < 80:
                d.Paragraphs(3 + n).Range.InsertParagraphAfter()
                n += 1
                d.Repaginate()
        else:
            brk(d.Paragraphs(4).Range, 7)
        info = table_end_blank(d, t)
        if kind == "answer":
            thin_end_para(d)
        d.BuiltInDocumentProperties("Author").Value = "CODINGNOW"
        d.BuiltInDocumentProperties("Last Author").Value = "CODINGNOW"
        # 쪽수는 PDF 로 센다 — 같은 실행 안의 ComputeStatistics 가 끝 단락을 줄인 뒤에도 3 으로 남은 적이 있다(11/6, 다시 열면 2)
        R.setdefault("files", {})[nm] = {"made": info, "enters": n if kind == "practice" else 0,
                                          **pdf_pages(d, kind)}
        d.SaveAs2(os.path.join(HERE, nm), 16)
        d.Close(0)
finally:
    w.Quit(0)

json.dump(R, open(os.path.join(HERE, "measure_en.json" if LANG_EN else "measure.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k in ("keys", "labels", "page_break", "enter_push", "blank_enters", "blank_page_break", "blank_section_break", "blank_table_end", "files"):
    print(k, json.dumps(R.get(k), ensure_ascii=False)[:500])
