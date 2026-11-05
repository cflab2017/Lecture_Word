"""워드 기초 3강 — 페이지 번호 실측. 이 PC 워드(Microsoft 365)를 숨긴 채 따로 띄워(DispatchEx) 표지 1쪽 + 본문 3쪽 보고서에
페이지 번호를 여러 방법으로 넣고 PDF 로 저장한 뒤, 쪽마다 아래 여백(바닥글)에 찍힌 글자와 위치(mm)를 읽는다. 결과 measure.json.
실습·정답 파일도 이 폴더에 만든다(실습 = 번호 없음, 정답 = 아래쪽 가운데 · 첫 페이지 다르게 · 시작 번호 0).
실행: python measure.py  (pywin32 · pymupdf 필요) — argv 'en' 이면 영어 글로 measure_en.json · practice.docx · answer.docx."""
import datetime, json, os, sys, tempfile
import pythoncom
import win32com.client as win32
import pymupdf

sys.stdout.reconfigure(encoding="utf-8")
from measure_text import KO, EN, build

LANG_EN = "en" in sys.argv[1:]
T = EN if LANG_EN else KO
HERE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix="wd03_")
MM = 25.4 / 72
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if LANG_EN else "ko"}
clean = lambda s: (s or "").replace("​", "").strip()


def footer_text(pdf):
    """쪽마다 아래 30mm 안의 글자(바닥글)와 그 가운데 x(mm)."""
    doc = pymupdf.open(pdf)
    out = []
    for pg in doc:
        h = pg.rect.height
        ws = [w for w in pg.get_text("words") if w[3] > h - 30 / MM]
        txt = " ".join(w[4] for w in sorted(ws, key=lambda w: w[0]))
        x = round(((min(w[0] for w in ws) + max(w[2] for w in ws)) / 2) * MM, 1) if ws else None
        y = round((h - max(w[3] for w in ws)) * MM, 1) if ws else None
        out.append({"text": txt, "x_mid_mm": x, "from_bottom_mm": y})
    pw = round(doc[0].rect.width * MM, 1)
    doc.close()
    return {"pages": len(out), "page_w_mm": pw, "footer": out}


def fresh(w):
    d = w.Documents.Add()
    build(d, T)
    return d


def save_pdf(d, tag):
    p = os.path.join(TMP, tag + ".pdf")
    d.ExportAsFixedFormat(p, 17)  # wdExportFormatPDF
    return footer_text(p)


pythoncom.CoInitialize()
w = win32.DispatchEx("Word.Application")
try:
    w.Visible = False; w.DisplayAlerts = 0
    R["build"] = w.Build
    R["labels"] = {idm: clean(w.CommandBars.GetLabelMso(idm)) for idm in
                   ("PageNumberFormat", "PageNumbersRemove", "HeaderFooterDifferentFirstPageWord", "HeaderFooterClose", "FooterInsertGallery", "GoToFooter", "HeaderFooterLinkToPrevious")}
    # 0) 번호 없음
    d = fresh(w); R["none"] = save_pdf(d, "none"); d.Close(0)
    # 1) 아래쪽 가운데(PageNumbers.Add 가운데 · 첫 쪽에도)
    d = fresh(w)
    pn = d.Sections(1).Footers(1).PageNumbers
    pn.Add(1, True)  # wdAlignPageNumberCenter, FirstPage=True
    R["bottom_center"] = save_pdf(d, "bottom_center")
    # 2) 첫 페이지를 다르게 지정
    d.PageSetup.DifferentFirstPageHeaderFooter = True
    R["diff_first"] = save_pdf(d, "diff_first")
    # 3) + 페이지 번호 서식 — 시작 번호 0
    pn = d.Sections(1).Footers(1).PageNumbers
    pn.RestartNumberingAtSection = True
    pn.StartingNumber = 0
    R["start0"] = save_pdf(d, "start0")
    # 4) 시작 번호 5(앞 문서에서 이어 붙일 때) — 첫 페이지 다르게는 끈다
    d.PageSetup.DifferentFirstPageHeaderFooter = False
    pn.StartingNumber = 5
    R["start5"] = save_pdf(d, "start5")
    d.Close(0)
    # 5) 함정 — 바닥글에 숫자 '1' 을 직접 친 경우
    d = fresh(w)
    d.Sections(1).Footers(1).Range.Text = "1"
    d.Sections(1).Footers(1).Range.ParagraphFormat.Alignment = 1
    R["typed_1"] = save_pdf(d, "typed_1")
    d.Close(0)
    # 6) X / Y 쪽 — PAGE 필드 + ' / ' + NUMPAGES 필드, 그다음 첫 페이지 다르게 + 시작 번호 0
    d = fresh(w)
    fr = d.Sections(1).Footers(1).Range
    fr.Text = ""
    fr.Fields.Add(fr, 33)            # wdFieldPage
    fr = d.Sections(1).Footers(1).Range; fr.Collapse(0); fr.InsertAfter(" / ")
    fr = d.Sections(1).Footers(1).Range; fr.Collapse(0)
    d.Fields  # noqa
    fr.Fields.Add(fr, 26)            # wdFieldNumPages
    d.Sections(1).Footers(1).Range.ParagraphFormat.Alignment = 1
    d.Sections(1).Footers(1).Range.Fields.Update()
    R["page_x_of_y"] = save_pdf(d, "page_x_of_y")
    d.PageSetup.DifferentFirstPageHeaderFooter = True
    pn = d.Sections(1).Footers(1).PageNumbers
    pn.RestartNumberingAtSection = True; pn.StartingNumber = 0
    d.Sections(1).Footers(1).Range.Fields.Update()
    R["page_x_of_y_start0"] = save_pdf(d, "page_x_of_y_start0")
    R["page_x_of_y_codes"] = [f.Code.Text.strip() for f in d.Sections(1).Footers(1).Range.Fields]
    # 7) 고치기 — 전체 쪽수 자리를 { = { NUMPAGES } - 1 } 계산 필드로(표지 1쪽을 뺌)
    ftr = d.Sections(1).Footers(1).Range
    last = ftr.Fields(ftr.Fields.Count)          # NUMPAGES
    rr = last.Result; st = last.Code.Start - 1
    last.Delete()
    r2 = d.Sections(1).Footers(1).Range; r2.SetRange(st, st)
    f = r2.Fields.Add(r2, -1, "= - 1", False)    # wdFieldEmpty, 계산 필드 껍데기
    c = f.Code; ins = c.Duplicate; ins.SetRange(c.Start + 2, c.Start + 2)   # ⚠️ d.Range(...) 는 본문 이야기(story)라 바닥글 안이 아니다
    ins.Fields.Add(ins, 26)                       # 안쪽에 NUMPAGES
    d.Sections(1).Footers(1).Range.Fields.Update()
    R["page_x_of_y_minus1"] = save_pdf(d, "page_x_of_y_minus1")
    R["page_x_of_y_minus1_codes"] = [ff.Code.Text.strip() for ff in d.Sections(1).Footers(1).Range.Fields]
    d.Close(0)
    # (갤러리 건물 블록 '일반 번호 2' 를 자동화로 바닥글에 넣어 보려 했으나 바닥글에 안 들어가고 워드 연결이 끊겼다 — 11/5. 그래서 같은 모양은 PageNumbers.Add 로 잰다)
    # 실습·정답
    names = ("practice.docx", "answer.docx") if LANG_EN else ("실습.docx", "정답.docx")
    d = fresh(w); d.SaveAs2(os.path.join(HERE, names[0]), 16); d.Close(0)
    d = fresh(w)
    pn = d.Sections(1).Footers(1).PageNumbers
    pn.Add(1, True)
    d.PageSetup.DifferentFirstPageHeaderFooter = True
    pn = d.Sections(1).Footers(1).PageNumbers
    pn.RestartNumberingAtSection = True; pn.StartingNumber = 0
    d.SaveAs2(os.path.join(HERE, names[1]), 16)
    R["answer"] = save_pdf(d, "answer")
    d.Close(0)
finally:
    w.Quit(0)

out = os.path.join(HERE, "measure_en.json" if LANG_EN else "measure.json")
json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k, v in list(R.items()) + [("gallery:" + g, x) for g, x in R.get("gallery", {}).items()]:
    if isinstance(v, dict) and "footer" in v:
        print(k, v["pages"], [f["text"] for f in v["footer"]], [f["x_mid_mm"] for f in v["footer"]][:2], v["footer"][-1]["from_bottom_mm"])
print(R["labels"], R.get("page_x_of_y_codes"), R.get("page_x_of_y_minus1_codes"))
