"""screens.py · menu_names.py 가 함께 쓰는 창 찍기 도구 — PrintWindow 로 창 하나를 PNG 로 저장, 클래스 이름으로 창 찾기."""
import ctypes, ctypes.wintypes as W, os
from PIL import Image

u32, g32 = ctypes.windll.user32, ctypes.windll.gdi32
u32.SetProcessDPIAware()


def grab(hwnd, out):
    r = W.RECT(); u32.GetWindowRect(hwnd, ctypes.byref(r))
    w, h = r.right - r.left, r.bottom - r.top
    hdc = u32.GetWindowDC(hwnd); mdc = g32.CreateCompatibleDC(hdc); bmp = g32.CreateCompatibleBitmap(hdc, w, h)
    g32.SelectObject(mdc, bmp); u32.PrintWindow(hwnd, mdc, 2)

    class BMI(ctypes.Structure):
        _fields_ = [("a", W.DWORD), ("w", W.LONG), ("h", W.LONG), ("p", W.WORD), ("b", W.WORD), ("c", W.DWORD), ("s", W.DWORD),
                    ("x", W.LONG), ("y", W.LONG), ("u", W.DWORD), ("i", W.DWORD)]
    bi = BMI(ctypes.sizeof(BMI), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)
    buf = ctypes.create_string_buffer(w * h * 4)
    g32.GetDIBits(mdc, bmp, 0, h, buf, ctypes.byref(bi), 0)
    g32.DeleteObject(bmp); g32.DeleteDC(mdc); u32.ReleaseDC(hwnd, hdc)
    Image.frombuffer("RGBA", (w, h), buf, "raw", "BGRA", 0, 1).convert("RGB").save(out)
    print(os.path.basename(out), w, h)


def find_windows(cls, pid=None):
    found = []
    PROC = ctypes.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)

    def cb(h, _):
        b = ctypes.create_unicode_buffer(256); u32.GetClassNameW(h, b, 256)
        if b.value == cls and u32.IsWindowVisible(h):
            p = W.DWORD(); u32.GetWindowThreadProcessId(h, ctypes.byref(p))
            if pid is None or p.value == pid:
                found.append(h)
        return True
    u32.EnumWindows(PROC(cb), 0)
    return found


def pid_of(hwnd):
    p = W.DWORD(); u32.GetWindowThreadProcessId(hwnd, ctypes.byref(p)); return p.value
