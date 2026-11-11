"""HWP(한글 2024 저장) 문서 요약 정보의 작성자 · 마지막 저장한 사람을 CODINGNOW 로 — 공개 저장소에 이 PC 계정 이름이 남지 않게(11/11).
한글이 저장할 때 '마지막 저장한 사람'에 윈도 계정 이름을 넣고, 자동화로는 이 칸을 바꿀 수 없었다(make_files.py 작업 기록).
HWP 는 OLE 복합 파일이고 요약 정보는 '\\x05HwpSummaryInformation' 스트림(표준 속성 집합 형식)이다 — 윈도 구조화 저장소 API(pywin32 pythoncom)로
그 스트림만 읽어 문자열 속성을 바꾸고 다시 쓴다(길이가 달라져도 저장소 API 가 자리를 다시 잡는다). 나머지 스트림은 건드리지 않는다.
usage: python hwp_author.py <파일.hwp> [...]  (--dry 면 읽기만)"""
import struct, sys
import pythoncom
from win32com import storagecon

sys.stdout.reconfigure(encoding="utf-8")
NAME = "\x05HwpSummaryInformation"
VT_LPWSTR, VT_LPSTR = 0x1F, 0x1E
TARGET = "CODINGNOW"
PIDS = {2: "title", 3: "subject", 4: "author", 5: "keywords", 6: "comments", 7: "template", 8: "last_author", 9: "rev"}


def parse(b):
    bo, ver, sysid, clsid, nsec = struct.unpack_from("<HHI16sI", b, 0)
    fmtid, off = struct.unpack_from("<16sI", b, 28)
    size, cnt = struct.unpack_from("<II", b, off)
    props = []
    for k in range(cnt):
        pid, poff = struct.unpack_from("<II", b, off + 8 + 8 * k)
        props.append([pid, poff])
    # 값의 끝 = 다음 값의 시작(오프셋 순) 또는 구역 끝
    order = sorted(props, key=lambda p: p[1])
    raw = {}
    for i, (pid, poff) in enumerate(order):
        end = order[i + 1][1] if i + 1 < len(order) else size
        raw[pid] = b[off + poff: off + end]
    return b[:off], off, size, [p[0] for p in props], raw


def decode(v):
    vt = struct.unpack_from("<I", v, 0)[0]
    if vt == VT_LPWSTR:
        n = struct.unpack_from("<I", v, 4)[0]
        return vt, v[8:8 + 2 * n].decode("utf-16le").rstrip("\x00")
    if vt == VT_LPSTR:
        n = struct.unpack_from("<I", v, 4)[0]
        return vt, v[8:8 + n].rstrip(b"\x00").decode("cp949", "replace")
    return vt, None


def enc_wstr(s):
    data = (s + "\x00").encode("utf-16le")
    out = struct.pack("<II", VT_LPWSTR, len(s) + 1) + data
    return out + b"\x00" * ((4 - len(out) % 4) % 4)


def rebuild(head, off, pids, raw):
    body = b""
    table = b""
    cur = 8 + 8 * len(pids)
    vals = []
    for pid in pids:
        vals.append((pid, cur, raw[pid]))
        cur += len(raw[pid])
    for pid, poff, v in vals:
        table += struct.pack("<II", pid, poff)
        body += v
    sec = struct.pack("<II", 8 + len(table) + len(body), len(pids)) + table + body
    return head + sec


def run(path, dry):
    flags = storagecon.STGM_READ if dry else storagecon.STGM_READWRITE
    stg = pythoncom.StgOpenStorage(path, None, flags | storagecon.STGM_SHARE_EXCLUSIVE, None, 0)
    st = stg.OpenStream(NAME, None, flags | storagecon.STGM_SHARE_EXCLUSIVE, 0)
    b = b""
    while True:
        chunk = st.Read(65536)
        if not chunk:
            break
        b += chunk
    head, off, size, pids, raw = parse(b)
    before = {PIDS.get(p, p): decode(raw[p])[1] for p in pids if decode(raw[p])[1] is not None}
    changed = []
    for p in pids:
        vt, s = decode(raw[p])
        if p in (4, 8) and s is not None and s != TARGET:
            raw[p] = enc_wstr(TARGET)
            changed.append(PIDS[p])
    print(path, "| 읽음:", before, "| 바꿀 칸:", changed)
    if dry or not changed:
        return
    nb = rebuild(head, off, pids, raw)
    st.Seek(0, 0)
    st.SetSize(len(nb))
    st.Write(nb)
    st.Commit(0)
    stg.Commit(0)
    del st, stg
    # 다시 읽어 확인
    stg = pythoncom.StgOpenStorage(path, None, storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE, None, 0)
    st = stg.OpenStream(NAME, None, storagecon.STGM_READ | storagecon.STGM_SHARE_EXCLUSIVE, 0)
    b2 = st.Read(1 << 20)
    _, _, _, pids2, raw2 = parse(b2)
    print("  → 다시 읽음:", {PIDS.get(p, p): decode(raw2[p])[1] for p in pids2 if decode(raw2[p])[1] is not None})


dry = "--dry" in sys.argv
for f in [a for a in sys.argv[1:] if a != "--dry"]:
    run(f, dry)
