"""워드 기초 7강 — 워드 ↔ 한글(HWP) 변환 실측. 이 PC 워드(Microsoft 365)와 한글 2024 를 자동화로 띄워(둘 다 창 숨김) 같은 내용의 문서를
양쪽에서 만들고 서로 바꾼 뒤 다시 열어 쪽 수 · 글꼴 · 글자 크기 · 줄 간격 · 표 · 그림 · 쪽 번호를 잰다.
- 워드 파일 변환기 목록(FileConverters) — HWP 를 여는 변환기가 있는지
- W: 워드로 만든 연습 문서(practice.docx) — 한글 2024 로 열기 → HWP · HWPX · PDF 로 저장(docx_to.*)
- H: 한글로 만든 연습 문서(native.hwp) — 한글 2024 '다른 이름으로 저장' 형식 OOXML(.docx) 로 저장(hwp_to.docx) → 워드로 열기
- 워드로 .hwp(한글 2024 저장) · .hwpx 를 직접 열기 — 성공/오류 문구
한글 쪽 값은 HWPX(zip 안 XML)로 저장해 읽는다(header.xml 글꼴 · 글자 크기 · 줄 간격, section0.xml 표 · 그림 · 쪽 번호 · 용지).
⚠️ 한글은 %TEMP% 밖에 저장하면 '파일 접근 허용' 창, HWP 가 아닌 형식(OOXML)으로 저장하면 '호환 문서'(계속/취소) 창이 뜬다 —
   감시 스레드가 '호환 문서'에 Enter(계속)를 보낸다. 이 PC 한글은 구독이 없어 HAction.Run 편집 동작은 False 지만
   InsertText · TableCreate · CharShape · PageNumPos 의 HAction.Execute 와 SetPos(표 칸 목록) · InsertPicture 는 된다(11/11).
⚠️ pywin32 늦은 바인딩에서 이름 붙인 인자는 무시될 수 있어 선택 인자는 자리로 넘긴다.
결과 measure.json (argv 'en' 이면 measure_en.json). 실행: python measure.py [en]  (pywin32 · pymupdf · pillow)"""
import datetime, json, os, re, sys, threading, time, zipfile
import pymupdf
import pythoncom
import win32com.client as win32
import win32con, win32gui, win32process
from PIL import Image, ImageDraw

sys.stdout.reconfigure(encoding="utf-8")
EN = "en" in sys.argv[1:]
HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(os.environ["TEMP"], "claude", "word07" + ("_en" if EN else ""))
os.makedirs(WORK, exist_ok=True)
R = {"run": datetime.datetime.now().isoformat(timespec="seconds"), "lang": "en" if EN else "ko"}

if EN:
    T = {"title": "Photo Club — December Meeting",
         "body": "This is a practice notice for a made-up photo club. We meet once a month to share pictures and plan the next outing. ",
         "head": "Schedule",
         "table": [["Date", "Place", "Bring"], ["Dec 6", "Club room", "Camera"], ["Dec 20", "City park", "Tripod"]],
         "tail": "Please reply by Friday if you can come."}
else:
    T = {"title": "사진 동아리 12월 모임 안내",
         "body": "연습용으로 만든 가상 사진 동아리의 안내문입니다. 한 달에 한 번 모여 찍은 사진을 함께 보고 다음 출사 계획을 정합니다. ",
         "head": "일정",
         "table": [["날짜", "장소", "준비물"], ["12월 6일", "동아리방", "카메라"], ["12월 20일", "시민 공원", "삼각대"]],
         "tail": "참석할 수 있는 분은 금요일까지 알려 주세요."}
N_BEFORE, N_AFTER = 8, 9   # 두 쪽이 되게 본문 문단 수

PIC = os.path.join(WORK, "photo.png")
im = Image.new("RGB", (900, 450), (64, 120, 170))
d = ImageDraw.Draw(im)
d.ellipse((330, 60, 570, 300), fill=(250, 200, 70))
d.polygon([(0, 450), (300, 220), (520, 450)], fill=(40, 90, 60))
d.polygon([(380, 450), (700, 180), (900, 450)], fill=(30, 70, 50))
im.save(PIC)


def rm(p):
    if os.path.exists(p):
        os.remove(p)
    return p


# ── 대화상자 감시 ─────────────────────────────────────────────
SEEN = []          # (앱, 제목, 처리)
WATCH_PIDS = {}    # pid -> 앱 이름
STOP = False


def watch():
    while not STOP:
        def cb(h, _):
            pid = win32process.GetWindowThreadProcessId(h)[1]
            if pid in WATCH_PIDS and win32gui.IsWindowVisible(h):
                title, cls = win32gui.GetWindowText(h), win32gui.GetClassName(h)
                if title == "호환 문서":
                    SEEN.append((WATCH_PIDS[pid], title, "Enter(계속)"))
                    win32gui.PostMessage(h, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                    win32gui.PostMessage(h, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    time.sleep(1)
                elif cls in ("bosa_sdm_msword", "#32770", "NUIDialog") and WATCH_PIDS[pid] == "word":
                    kids = []
                    win32gui.EnumChildWindows(h, lambda c, _: kids.append(win32gui.GetWindowText(c)) or True, 0)
                    SEEN.append(("word", title, [k for k in kids if k][:12]))
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                    time.sleep(1)
            return True
        try:
            win32gui.EnumWindows(cb, 0)
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.4)


# ── 읽기 ───────────────────────────────────────────────────
def word_info(doc):
    doc.Repaginate()
    fonts, sizes = {}, {}
    for i in range(1, doc.Paragraphs.Count + 1):
        rg = doc.Paragraphs(i).Range
        if not rg.Text.strip("\r\x07 "):
            continue
        for k in (rg.Font.Name, rg.Font.NameFarEast):
            if k:
                fonts[k] = fonts.get(k, 0) + 1
        sizes[str(rg.Font.Size)] = sizes.get(str(rg.Font.Size), 0) + 1

    def pf(i):
        p = doc.Paragraphs(i)
        f = p.Format
        return {"text": p.Range.Text[:18], "font": p.Range.Font.Name, "far_east": p.Range.Font.NameFarEast, "size": p.Range.Font.Size,
                "bold": p.Range.Font.Bold, "line_rule": f.LineSpacingRule, "line_pt": round(f.LineSpacing, 2),
                "before": f.SpaceBefore, "after": f.SpaceAfter}
    body_i = next(i for i in range(2, doc.Paragraphs.Count + 1) if len(doc.Paragraphs(i).Range.Text) > 40)
    sec = doc.Sections(1)
    ps = sec.PageSetup
    foot = sec.Footers(1)
    title_p = next((i for i in range(1, doc.Paragraphs.Count + 1) if T["title"] in doc.Paragraphs(i).Range.Text), None)
    return {"pages": doc.Content.Information(4), "paragraphs": doc.Paragraphs.Count,
            "title_found": title_p is not None,
            "tables": [[t.Rows.Count, t.Columns.Count, t.Cell(1, 1).Range.Text.strip("\r\x07")] for t in doc.Tables],
            "inline_pictures": doc.InlineShapes.Count, "floating_shapes": doc.Shapes.Count,
            "fonts": fonts, "sizes": sizes, "title": pf(title_p or 1), "body": pf(body_i),
            "footer_text": foot.Range.Text.strip("\r"), "footer_fields": [f.Code.Text.strip() for f in foot.Range.Fields],
            "footer_page_numbers": foot.PageNumbers.Count,
            "footer_shapes": foot.Shapes.Count,
            "page_mm": [round(ps.PageWidth / 72 * 25.4, 1), round(ps.PageHeight / 72 * 25.4, 1)],
            "margins_mm": [round(x / 72 * 25.4, 1) for x in (ps.TopMargin, ps.BottomMargin, ps.LeftMargin, ps.RightMargin)]}


def hwpx_info(path):
    with zipfile.ZipFile(path) as z:
        h = z.read("Contents/header.xml").decode("utf-8")
        s = z.read("Contents/section0.xml").decode("utf-8")
    hangul_faces = re.search(r'<hh:fontface lang="HANGUL"[^>]*>(.*?)</hh:fontface>', h, re.S)
    faces = re.findall(r'face="([^"]+)"', hangul_faces.group(1)) if hangul_faces else []
    char = {m.group(1): {"height": int(m.group(2)), "hangul_font_id": None, "bold": False}
            for m in re.finditer(r'<hh:charPr id="(\d+)" height="(\d+)"', h)}
    for m in re.finditer(r'<hh:charPr id="(\d+)".*?</hh:charPr>', h, re.S):
        blk = m.group(0)
        f = re.search(r'<hh:fontRef hangul="(\d+)"', blk)
        char[m.group(1)]["hangul_font_id"] = int(f.group(1)) if f else None
        char[m.group(1)]["bold"] = "<hh:bold" in blk
    para = {}
    for m in re.finditer(r'<hh:paraPr id="(\d+)".*?</hh:paraPr>', h, re.S):
        ls = re.search(r'<hh:lineSpacing type="(\w+)" value="(\d+)"', m.group(0))
        para[m.group(1)] = (ls.group(1), int(ls.group(2))) if ls else None
    # 글자 조각을 나오는 순서대로 — 그 앞의 가장 가까운 단락 모양(paraPrIDRef) · 글자 모양(charPrIDRef)과 함께
    texts, cur_p, cur_c = [], None, None
    for m in re.finditer(r'<hp:p [^>]*paraPrIDRef="(\d+)"|<hp:run charPrIDRef="(\d+)"|<hp:t>([^<]*)</hp:t>', s):
        if m.group(1):
            cur_p = m.group(1)
        elif m.group(2):
            cur_c = m.group(2)
        elif m.group(3):
            texts.append((m.group(3), cur_p, cur_c))

    def desc(item):
        t, pid, cp = item
        c = char.get(cp, {})
        fid = c.get("hangul_font_id")
        return {"text": t[:18], "font": faces[fid] if fid is not None and fid < len(faces) else None, "size_pt": c.get("height", 0) / 100,
                "bold": c.get("bold"), "line_spacing": para.get(pid)}
    title = next((x for x in texts if T["title"] in x[0]), ("(제목 없음)", None, None))
    head = next((x for x in texts if x[0].strip() == T["head"]), ("(소제목 없음)", None, None))
    body = next(x for x in texts if len(x[0]) > 40)
    tbl = [(int(a), int(b)) for a, b in re.findall(r'<hp:tbl [^>]*rowCnt="(\d+)" colCnt="(\d+)"', s)]
    pg = re.search(r'<hp:pagePr [^>]*width="(\d+)" height="(\d+)"', s)
    mg = re.search(r'<hp:margin [^>]*header="(\d+)" footer="(\d+)"[^>]*left="(\d+)" right="(\d+)" top="(\d+)" bottom="(\d+)"', s) or \
        re.search(r'<hp:margin ([^/]*)/>', s)
    mm = lambda v: round(int(v) / 7200 * 25.4, 1)  # noqa: E731  HWPUNIT 1/7200 인치
    return {"bytes": os.path.getsize(path), "fonts_hangul": faces, "title": desc(title), "head": desc(head), "body": desc(body),
            "text_runs": len(texts), "tables": tbl,
            "pictures": s.count("<hp:pic "), "page_num_ctrl": s.count("<hp:pageNum "), "footer_ctrl": s.count("<hp:footer"),
            "auto_num_page": s.count('numType="PAGE"'),
            "page_mm": [mm(pg.group(1)), mm(pg.group(2))] if pg else None,
            "margin_raw": mg.group(0)[:160] if mg else None}


def pdf_info(p):
    with pymupdf.open(p) as doc:
        fonts = sorted({f[3] for pg in doc for f in pg.get_fonts(full=True)})
        imgs = sum(len(pg.get_images(full=True)) for pg in doc)
        bottoms = []
        for pg in doc:
            ws = pg.get_text("words")
            bottoms.append(" ".join(w[4] for w in ws if w[1] > pg.rect.height * 0.85) or "(없음)")
        return {"pages": doc.page_count, "fonts": fonts, "images": imgs, "page_bottom_text": bottoms, "bytes": os.path.getsize(p)}


# ── 시작 ───────────────────────────────────────────────────
pythoncom.CoInitialize()
th = threading.Thread(target=watch, daemon=True)
th.start()
w = win32.DispatchEx("Word.Application")
hwp = None
try:
    w.Visible = False
    w.DisplayAlerts = 0
    R["word_build"] = w.Build
    R["word_converters"] = [{"name": c.FormatName, "ext": c.Extensions, "class": c.ClassName, "open": c.CanOpen, "save": c.CanSave}
                            for c in w.FileConverters]

    # W1 워드로 연습 문서
    doc = w.Documents.Add()
    WATCH_PIDS[win32process.GetWindowThreadProcessId(doc.ActiveWindow.Hwnd)[1]] = "word"
    sel = w.Selection

    def para(text, size=None, bold=False):
        sel.Font.Size = size or 11
        sel.Font.Bold = bold
        sel.TypeText(text)
        sel.TypeParagraph()
        sel.Font.Size = 11
        sel.Font.Bold = False

    para(T["title"], 16, True)
    for _ in range(N_BEFORE):
        para(T["body"] * 2)
    para(T["head"], 13, True)
    t = doc.Tables.Add(sel.Range, 3, 3, 1, 0)  # wdWord9TableBehavior · 고정 너비
    for i in range(3):
        for j in range(3):
            t.Cell(i + 1, j + 1).Range.Text = T["table"][i][j]
    sel.EndKey(6)
    sel.TypeParagraph()
    for _ in range(N_AFTER):
        para(T["body"] * 2)
    sel.InlineShapes.AddPicture(PIC)
    sel.TypeParagraph()
    para(T["tail"])
    doc.Sections(1).Footers(1).PageNumbers.Add(1, True)  # 가운데, 첫 쪽에도
    for k in ("Author", "Last Author"):
        doc.BuiltinDocumentProperties(k).Value = "CODINGNOW"
    WDOCX = rm(os.path.join(WORK, "practice_word.docx"))
    doc.SaveAs2(WDOCX, 16)
    R["W_word_docx"] = word_info(doc)
    wpdf = rm(os.path.join(WORK, "practice_word.pdf"))
    doc.ExportAsFixedFormat(wpdf, 17)
    R["W_word_pdf"] = pdf_info(wpdf)
    doc.Close(0)

    # 한글
    hwp = win32.Dispatch("HWPFrame.HwpObject")
    hwp.XHwpWindows.Item(0).Visible = False
    R["hangul_version"] = hwp.Version

    def hwp_pids():
        out = []
        win32gui.EnumWindows(lambda h, _: out.append(win32process.GetWindowThreadProcessId(h)[1])
                             if win32gui.GetWindowText(h).endswith("- 한글") or win32gui.GetWindowText(h).endswith("- Hangul") else None or True, 0)
        return set(out)
    for p_ in hwp_pids():
        WATCH_PIDS[p_] = "hangul"

    def save(path, fmt):
        n0 = len(SEEN)
        ok = hwp.SaveAs(rm(path), fmt, "")
        return {"ok": bool(ok), "bytes": os.path.getsize(path) if os.path.exists(path) else None,
                "compat_dialog": [x for x in SEEN[n0:] if x[0] == "hangul"]}

    def ins(text):
        a = hwp.HParameterSet.HInsertText
        hwp.HAction.GetDefault("InsertText", a.HSet)
        a.Text = text
        return hwp.HAction.Execute("InsertText", a.HSet)

    def charshape(height, bold):
        """선택 없이 실행 — 커서 뒤로 칠 글자의 모양"""
        cs = hwp.HParameterSet.HCharShape
        hwp.HAction.GetDefault("CharShape", cs.HSet)
        cs.Height = height
        cs.Bold = 1 if bold else 0
        return hwp.HAction.Execute("CharShape", cs.HSet)

    # W2 워드 문서를 한글로 열기 → HWP · HWPX · PDF
    R["W_open_in_hangul"] = {"open": bool(hwp.Open(WDOCX, "", "")), "pages": hwp.PageCount}
    R["W_save_hwp"] = save(os.path.join(WORK, "docx_to.hwp"), "HWP")
    R["W_save_hwpx"] = save(os.path.join(WORK, "docx_to.hwpx"), "HWPX")
    R["W_hwpx"] = hwpx_info(os.path.join(WORK, "docx_to.hwpx"))
    R["W_save_pdf"] = save(os.path.join(WORK, "docx_to.pdf"), "PDF")
    R["W_hangul_pdf"] = pdf_info(os.path.join(WORK, "docx_to.pdf"))

    # H1 한글로 같은 내용의 연습 문서
    hwp.Clear(1)
    charshape(1600, True)
    ins(T["title"])
    charshape(1000, False)
    ins("\r\n" + "\r\n".join([T["body"] * 2] * N_BEFORE) + "\r\n")
    charshape(1300, True)
    ins(T["head"])
    charshape(1000, False)
    ins("\r\n")
    tc = hwp.HParameterSet.HTableCreation
    hwp.HAction.GetDefault("TableCreate", tc.HSet)
    tc.Rows, tc.Cols, tc.WidthType, tc.HeightType = 3, 3, 2, 0
    R["H_table_create"] = bool(hwp.HAction.Execute("TableCreate", tc.HSet))
    first = hwp.GetPos()[0]
    for k in range(9):
        hwp.SetPos(first + k, 0, 0)
        ins(T["table"][k // 3][k % 3])
    hwp.MovePos(3, 0, 0)  # 문서 끝
    ins("\r\n" + "\r\n".join([T["body"] * 2] * N_AFTER) + "\r\n")
    pic = hwp.InsertPicture(PIC, True, 0, False, False, 0, 0, 0)
    R["H_picture"] = pic is not None
    hwp.MovePos(3, 0, 0)
    ins("\r\n" + T["tail"])
    pn = hwp.HParameterSet.HPageNumPos
    hwp.HAction.GetDefault("PageNumPos", pn.HSet)
    pn.DrawPos = hwp.PageNumPosition("BottomCenter")
    pn.NumberFormat = hwp.NumberFormat("Digit")
    pn.SideChar = ord("-")
    R["H_page_num"] = bool(hwp.HAction.Execute("PageNumPos", pn.HSet))
    HHWP = os.path.join(WORK, "native.hwp")
    R["H_save_hwp"] = save(HHWP, "HWP")
    R["H_pages_in_hangul"] = hwp.PageCount
    R["H_save_hwpx"] = save(os.path.join(WORK, "native.hwpx"), "HWPX")
    R["H_hwpx"] = hwpx_info(os.path.join(WORK, "native.hwpx"))
    R["H_save_pdf"] = save(os.path.join(WORK, "native.pdf"), "PDF")
    R["H_hangul_pdf"] = pdf_info(os.path.join(WORK, "native.pdf"))
    # H2 한글 → DOCX(OOXML) · 같은 파일을 한글이 다시 열면?
    assert hwp.Open(HHWP, "HWP", "")
    HDOCX = os.path.join(WORK, "hwp_to.docx")
    R["H_save_docx"] = save(HDOCX, "OOXML")
    R["H_save_docx_name_DOCX"] = {"ok": False, "note": "SaveAs 형식 이름 'DOCX' 는 False — 'OOXML' 로만 저장됨(11/11 시험)"}
    hwp.Clear(1)
    hwp.Quit()
    hwp = None

    # H3 한글이 만든 DOCX 를 워드로
    d2 = w.Documents.Open(HDOCX, False, True, False)
    R["H_docx_in_word"] = word_info(d2)
    hpdf = rm(os.path.join(WORK, "hwp_to_word.pdf"))
    d2.ExportAsFixedFormat(hpdf, 17)
    R["H_docx_word_pdf"] = pdf_info(hpdf)
    d2.Close(0)

    # 워드로 .hwp · .hwpx 직접 열기
    ara = next((c.OpenFormat for c in w.FileConverters if c.ClassName == "ARA20"), None)
    R["ara20_open_format"] = ara
    for key, path, fmt in (("word_open_hwp", HHWP, 0), ("word_open_hwp_ara20", HHWP, ara),
                           ("word_open_hwpx", os.path.join(WORK, "native.hwpx"), 0)):
        n0 = len(SEEN)
        try:
            d3 = w.Documents.Open(path, False, True, False, "", "", False, "", "", fmt)
            info = {"opened": True, "chars": len(d3.Content.Text), "pages": d3.Content.Information(4),
                    "first_text": d3.Content.Text[:60]}
            d3.Close(0)
        except Exception as e:  # noqa: BLE001
            msg = str(e)
            try:
                msg = e.excepinfo[2] or msg
            except Exception:  # noqa: BLE001
                pass
            info = {"opened": False, "error": msg[:240]}
        time.sleep(1.5)
        info["dialogs"] = [x for x in SEEN[n0:] if x[0] == "word"]
        R[key] = info
finally:
    STOP = True
    if hwp is not None:
        try:
            hwp.Clear(1)
            hwp.Quit()
        except Exception:  # noqa: BLE001
            pass
    w.Quit(0)
R["dialogs_seen"] = SEEN
user = os.environ.get("USERNAME", "")
R["meta"] = {"pc_user_in_native_hwp": user.encode("utf-16-le") in open(os.path.join(WORK, "native.hwp"), "rb").read() if user else None,
             "pc_user_in_docx_to_hwp": user.encode("utf-16-le") in open(os.path.join(WORK, "docx_to.hwp"), "rb").read() if user else None}
with zipfile.ZipFile(os.path.join(WORK, "hwp_to.docx")) as z:
    core = z.read("docProps/core.xml").decode("utf-8")
    R["meta"]["hwp_to_docx_creator"] = re.findall(r"<dc:creator>([^<]*)</dc:creator>", core)
    R["meta"]["hwp_to_docx_lastmod"] = re.findall(r"<cp:lastModifiedBy>([^<]*)</cp:lastModifiedBy>", core)
    R["hwp_to_docx_parts"] = z.namelist()
out = os.path.join(HERE, "measure_en.json" if EN else "measure.json")
json.dump(R, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps({k: v for k, v in R.items() if k != "word_converters"}, ensure_ascii=False, indent=1)[:9000])
