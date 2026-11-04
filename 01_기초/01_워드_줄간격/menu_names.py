"""홈 → 단락 → '줄 간격' 단추의 메뉴 항목 이름을 워드 화면에서 직접 읽는다(접근성 정보 MSAA, ctypes). 결과: menu_names.json
보이는 워드를 따로 띄워 새 문서를 만들고, 단추를 눌러 메뉴를 열고 항목 이름을 읽은 뒤 메뉴를 찍고(shots/s5_menu.png) 다시 눌러 닫는다.
저장하지 않는다. 오피스 리본의 접근성 객체는 IDispatch 호출을 지원하지 않아(E_NOTIMPL) IAccessible 함수표를 직접 부른다.
실행: python menu_names.py"""
import ctypes, ctypes.wintypes as W, json, os, time
import win32com.client as win32
from screens_util import grab, find_windows, pid_of

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "shots"); os.makedirs(OUT, exist_ok=True)
u32, oleacc, oleaut = ctypes.windll.user32, ctypes.windll.oleacc, ctypes.windll.oleaut32
oleaut.SysFreeString.argtypes = [ctypes.c_void_p]
HR = ctypes.c_long


class VARIANT(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort), ("r1", ctypes.c_ushort), ("r2", ctypes.c_ushort), ("r3", ctypes.c_ushort),
                ("val", ctypes.c_void_p), ("val2", ctypes.c_void_p)]


def vi4(n):
    v = VARIANT(); v.vt = 3; v.val = n; return v


class GUID(ctypes.Structure):
    _fields_ = [("a", ctypes.c_ulong), ("b", ctypes.c_ushort), ("c", ctypes.c_ushort), ("d", ctypes.c_ubyte * 8)]


IID_IAcc = GUID(0x618736E0, 0x3C3D, 0x11CF, (ctypes.c_ubyte * 8)(0x81, 0x0C, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71))


def vcall(ptr, idx, *types):
    vtbl = ctypes.cast(ctypes.c_void_p(ptr), ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p)))[0]
    return ctypes.WINFUNCTYPE(HR, ctypes.c_void_p, *types)(vtbl[idx])


class Acc:
    def __init__(self, ptr):
        self.p = ptr

    def count(self):
        n = ctypes.c_long(); vcall(self.p, 8, ctypes.POINTER(ctypes.c_long))(self.p, ctypes.byref(n)); return n.value

    def _bstr(self, idx, child):
        b = ctypes.c_void_p()
        hr = vcall(self.p, idx, VARIANT, ctypes.POINTER(ctypes.c_void_p))(self.p, vi4(child), ctypes.byref(b))
        if hr != 0 or not b.value:
            return None
        s = ctypes.wstring_at(b.value); oleaut.SysFreeString(b); return s

    def name(self, child=0):
        return self._bstr(10, child)

    def do(self, child=0):
        return vcall(self.p, 25, VARIANT)(self.p, vi4(child))

    def children(self):
        n = self.count()
        if n <= 0:
            return []
        arr = (VARIANT * n)(); got = ctypes.c_long()
        oleacc.AccessibleChildren(ctypes.c_void_p(self.p), 0, n, arr, ctypes.byref(got))
        out = []
        for v in arr[:got.value]:
            if v.vt == 9 and v.val:  # VT_DISPATCH → IAccessible
                q = ctypes.c_void_p()
                vcall(v.val, 0, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p))(v.val, ctypes.byref(IID_IAcc), ctypes.byref(q))
                if q.value:
                    out.append((Acc(q.value), 0))
            elif v.vt == 3:
                out.append((self, v.val & 0xFFFFFFFF))
        return out


def acc_from_hwnd(hwnd):
    p = ctypes.c_void_p()
    hr = oleacc.AccessibleObjectFromWindow(W.HWND(hwnd), ctypes.c_long(-4), ctypes.byref(IID_IAcc), ctypes.byref(p))
    return Acc(p.value) if hr == 0 and p.value else None


def walk(acc, depth=0, out=None):
    out = [] if out is None else out
    for a, cid in acc.children():
        try:
            name = a.name(cid)
        except Exception:
            name = None
        out.append((a, cid, (name or "").replace("​", "")))
        if cid == 0 and depth < 16:
            walk(a, depth + 1, out)
    return out


def child_windows(hwnd, cls):
    found = []
    PROC = ctypes.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)

    def cb(h, _):
        b = ctypes.create_unicode_buffer(256); u32.GetClassNameW(h, b, 256)
        if b.value == cls:
            found.append(h)
        return True
    u32.EnumChildWindows(hwnd, PROC(cb), 0)
    return found


R = {}
w = win32.DispatchEx("Word.Application")
try:
    doc = w.Documents.Add()
    w.Visible = True
    w.WindowState = 0
    w.Width, w.Height = 760, 640
    time.sleep(2)
    hwnd = find_windows("OpusApp")[0]
    pid = pid_of(hwnd)
    target, seen = None, []
    for k in child_windows(hwnd, "NetUIHWND"):
        acc = acc_from_hwnd(k)
        if not acc:
            continue
        for a, cid, name in walk(acc):
            seen.append(name)
            if name in ("선 및 단락 간격", "줄 및 단락 간격"):
                target, R["button"] = (a, cid), name
                break
        if target:
            break
    if not target:
        R["seen_count"] = len(seen)  # 이름 목록은 남기지 않는다(계정 이름이 섞여 있다)
        json.dump(R, open(os.path.join(HERE, "menu_names.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        raise SystemExit(f"button not found ({len(seen)} names)")
    before = set(find_windows("Net UI Tool Window", pid))
    target[0].do(target[1])
    time.sleep(1.5)
    pops = [h for h in find_windows("Net UI Tool Window", pid) if h not in before] or find_windows("Net UI Tool Window", pid)
    R["popups"] = len(pops)
    items = []
    for p in pops:
        acc = acc_from_hwnd(p)
        if acc:
            items += [n for _, _, n in walk(acc) if n]
        grab(p, os.path.join(OUT, "s5_menu.png"))
    R["menu_items"] = items
    target[0].do(target[1])  # 같은 단추를 다시 눌러 닫는다(키 입력은 다른 창으로 갈 수 있어 쓰지 않는다)
    time.sleep(0.5)
    doc.Close(0)
finally:
    w.Quit(0)
json.dump(R, open(os.path.join(HERE, "menu_names.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(json.dumps(R, ensure_ascii=False, indent=1))
