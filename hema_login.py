# -*- coding: utf-8 -*-
"""hema_login.py — Homes 自动登录 Python 原子兜底脚本（第二兜底）

适用场景：RUNBOOK「自动登录机制」的 MCP computer 坐标法失效时——
  ① 新建 tab 视口 0x0（整页渲染到视口外，iframe 负坐标，resize 无效）
  ② 多 Chrome 窗口时 computer 工具 surface 绑定漂移（截图/点击落到前台窗口，粘贴落空）
本脚本用 Win32 API 在单进程内一次完成「窗口激活 + 坐标点击 + 剪贴板粘贴 + 回车」，
不经过 MCP computer 工具，天然免疫上述两种失效模式。
2026-09-13 晨实战验证成功（登录页 iframe (868,230.875)、视口原点 (0,221)、
物理点击落点 (1247,553)/(1247,604)，一次登录成功拿到 QUARK_PARAMS）。

用法：
  python hema_login.py <iframe_x> <iframe_y> [--css-vw 1536] [--title HEMAOS] [--dry-run]
  iframe_x/iframe_y：javascript_tool 实测的登录页跨域 iframe（mini_login）左上角 CSS 坐标
  --css-vw：登录页 CSS 视口宽度（window.innerWidth，与 iframe 同一次测量取得），默认 1536
  --title：目标 Chrome 窗口标题关键字，默认 HEMAOS
  --dry-run：只探测窗口/视口并输出将使用的坐标，不点击不粘贴（安全自检用）

流程：EnumWindows 找目标 Chrome 窗口 → EnumChildWindows 取 Chrome_RenderWidgetHostHWND
（top≥100 且面积最大者 = 真实 web 视口）→ Alt 键 trick + SetForegroundWindow 前置
→ SendInput 点击账号框/密码框（视口原点 + CSS 坐标 × 缩放比）→ 剪贴板粘贴凭据
（读同目录 credentials.json，全程零回显）→ Enter 提交 → 清空剪贴板。

纪律：凭据内容严禁出现在任何输出/日志里；同一自然日本脚本 + MCP 坐标法合计尝试 ≤2 次。
"""
import ctypes, json, os, sys, time
from ctypes import wintypes

u32 = ctypes.windll.user32
k32 = ctypes.windll.kernel32

# ── 64 位必须设置 argtypes/restype，否则 GlobalLock 报 OverflowError（首版实测踩坑）──
k32.GlobalAlloc.restype = ctypes.c_void_p
k32.GlobalAlloc.argtypes = [ctypes.c_uint, ctypes.c_size_t]
k32.GlobalLock.restype = ctypes.c_void_p
k32.GlobalLock.argtypes = [ctypes.c_void_p]
k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
u32.SetClipboardData.restype = ctypes.c_void_p
u32.SetClipboardData.argtypes = [ctypes.c_uint, ctypes.c_void_p]
u32.OpenClipboard.argtypes = [ctypes.c_void_p]

try:
    u32.SetProcessDPIAware()  # 物理像素坐标系，与 GetWindowRect 一致
except Exception:
    pass

# ── SendInput 结构 ──
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.c_size_t)]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
                ("dwExtraInfo", ctypes.c_size_t)]

class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]

class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", ctypes.c_ulong), ("u", _INPUT_UNION)]

MOUSEEVENTF_MOVE, MOUSEEVENTF_LEFTDOWN, MOUSEEVENTF_LEFTUP, MOUSEEVENTF_ABSOLUTE = 0x1, 0x2, 0x4, 0x8000
KEYEVENTF_KEYUP = 0x2
VK_MENU, VK_CONTROL, VK_RETURN, VK_V = 0x12, 0x11, 0x0D, 0x56
SM_CXSCREEN, SM_CYSCREEN = 0, 1

def send(*inputs):
    arr = (INPUT * len(inputs))(*inputs)
    u32.SendInput(len(arr), ctypes.byref(arr), ctypes.sizeof(INPUT))

def mouse_click(x, y):
    sw = u32.GetSystemMetrics(SM_CXSCREEN) or 1920
    sh = u32.GetSystemMetrics(SM_CYSCREEN) or 1200
    nx = int(x * 65535 / (sw - 1))
    ny = int(y * 65535 / (sh - 1))
    send(INPUT(type=0, mi=MOUSEINPUT(nx, ny, 0, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, 0, 0)))
    time.sleep(0.1)
    send(INPUT(type=0, mi=MOUSEINPUT(nx, ny, 0, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0)))

def key_tap(vk, up=False):
    send(INPUT(type=1, ki=KEYBDINPUT(vk, 0, KEYEVENTF_KEYUP if up else 0, 0, 0)))

def ctrl_v():
    key_tap(VK_CONTROL); key_tap(VK_V); time.sleep(0.05)
    key_tap(VK_V, up=True); key_tap(VK_CONTROL, up=True)

def set_clipboard(text):
    CF_UNICODETEXT = 13
    if not u32.OpenClipboard(0):
        raise RuntimeError('OpenClipboard failed')
    try:
        u32.EmptyClipboard()
        data = text.encode('utf-16-le') + b'\x00\x00'
        h = k32.GlobalAlloc(0x0042, len(data))
        if not h:
            raise RuntimeError('GlobalAlloc failed')
        p = k32.GlobalLock(h)
        ctypes.memmove(p, data, len(data))
        k32.GlobalUnlock(h)
        if not u32.SetClipboardData(CF_UNICODETEXT, h):
            raise RuntimeError('SetClipboardData failed')
    finally:
        u32.CloseClipboard()

EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
EnumChildProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

def find_chrome_window(keyword):
    """EnumWindows 找标题含 keyword 的 Chrome 顶层窗口；多个时优先前台窗口。"""
    cands = []
    def cb(hwnd, _):
        if not u32.IsWindowVisible(hwnd):
            return True
        n = u32.GetWindowTextLengthW(hwnd)
        if n:
            buf = ctypes.create_unicode_buffer(n + 1)
            u32.GetWindowTextW(hwnd, buf, n + 1)
            if keyword in buf.value:
                cands.append(hwnd)
        return True
    u32.EnumWindows(EnumWindowsProc(cb), 0)
    if not cands:
        return None
    fg = u32.GetForegroundWindow()
    return fg if fg in cands else cands[0]

def find_viewport(hwnd):
    """EnumChildWindows 找 Chrome_RenderWidgetHostHWND：top≥100 且面积最大者 = 真实 web 视口。
    返回物理像素 (left, top, right, bottom)。窗口最大化时 top 即 Chrome UI（标签栏+地址栏）高度。"""
    rects = []
    def cb(h, _):
        cls = ctypes.create_unicode_buffer(256)
        u32.GetClassNameW(h, cls, 256)
        if cls.value == 'Chrome_RenderWidgetHostHWND':
            r = wintypes.RECT()
            u32.GetWindowRect(h, ctypes.byref(r))
            w, hh = r.right - r.left, r.bottom - r.top
            if w > 0 and hh > 0:
                rects.append((r.left, r.top, r.right, r.bottom, w * hh))
        return True
    u32.EnumChildWindows(hwnd, EnumChildProc(cb), 0)
    real = [r for r in rects if r[1] >= 100]  # top<100 的是隐藏/偏移副本，非真实视口
    pool = real or rects
    if not pool:
        return None
    return max(pool, key=lambda r: r[4])[:4]

def foreground(hwnd):
    """Alt 键 trick：先按一下 Alt 解除前台锁定，再 SetForegroundWindow，否则跨进程置前会被系统拒绝。"""
    if u32.IsIconic(hwnd):
        u32.ShowWindow(hwnd, 9)  # SW_RESTORE
    key_tap(VK_MENU); time.sleep(0.05); key_tap(VK_MENU, up=True); time.sleep(0.1)
    u32.SetForegroundWindow(hwnd)
    u32.BringWindowToTop(hwnd)
    time.sleep(0.6)  # 等窗口完成置前与重绘

def main():
    args = sys.argv[1:]
    dry = '--dry-run' in args
    args = [a for a in args if not a.startswith('--dry')]
    kw = 'HEMAOS'; css_vw = 1536.0
    pos = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == '--css-vw':
            css_vw = float(args[i + 1]); i += 2; continue
        if a == '--title':
            kw = args[i + 1]; i += 2; continue
        pos.append(float(a)); i += 1
    if len(pos) != 2:
        print('USAGE: hema_login.py <iframe_x> <iframe_y> [--css-vw 1536] [--title HEMAOS] [--dry-run]')
        sys.exit(2)
    ix, iy = pos

    hwnd = find_chrome_window(kw)
    if not hwnd:
        print('NO_WINDOW title~%s' % kw)
        sys.exit(3)
    # 必须先前台化再枚举：窗口处于后台时渲染子窗口矩形不刷新，真实视口子窗口甚至缺席
    # （2026-09-15 实测：后台枚举只得 (0,0) 幻影矩形，前台化后才出现 top=221 的真实视口）
    foreground(hwnd)
    vp = find_viewport(hwnd)
    if vp is None or vp[1] < 100:
        time.sleep(1.0)
        vp2 = find_viewport(hwnd)
        if vp2 is not None and (vp is None or vp2[1] >= 100):
            vp = vp2
    if not vp:
        print('NO_VIEWPORT hwnd=%d' % hwnd)
        sys.exit(4)
    vx, vy, vr, vb = vp
    vw, vh = vr - vx, vb - vy
    scale = vw / css_vw  # 物理像素/CSS 像素（本机实测 1920/1536=1.25）
    # RUNBOOK 实测 iframe 内布局：输入框距 iframe 左边 20px、宽 220、高 36；账号框顶边距 17、密码框顶边距 58
    acc = (vx + (ix + 130) * scale, vy + (iy + 35) * scale)
    pwd = (vx + (ix + 130) * scale, vy + (iy + 76) * scale)
    print('WINDOW %d' % hwnd)
    print('VIEWPORT_PX %d %d %d %d' % (vx, vy, vr, vb))
    print('SCALE %.4f' % scale)
    print('POINTS (%d, %d) (%d, %d)' % (round(acc[0]), round(acc[1]), round(pwd[0]), round(pwd[1])))
    if dry:
        print('DRY_RUN_OK no action taken')
        return

    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, 'credentials.json'), encoding='utf-8') as f:
        cred = json.load(f)
    user = str(cred.get('username', '')).strip()
    pwdv = str(cred.get('password', '')).strip()
    if not user or not pwdv or '请填写' in user or '请填写' in pwdv:
        print('CRED_NOT_READY')
        sys.exit(5)

    foreground(hwnd)
    set_clipboard(user); time.sleep(0.15)
    mouse_click(*acc); time.sleep(0.4)
    ctrl_v(); time.sleep(0.4)
    set_clipboard(pwdv); time.sleep(0.15)
    mouse_click(*pwd); time.sleep(0.4)
    ctrl_v(); time.sleep(0.4)
    set_clipboard('')  # 提交前先清剪贴板
    key_tap(VK_RETURN); time.sleep(0.2)
    print('LOGIN_SUBMITTED')

if __name__ == '__main__':
    main()
