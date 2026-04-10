import ctypes
import threading
from ctypes import wintypes

from .constants import (
    APP_ID,
    TRAY_COMMAND_DISABLE,
    TRAY_COMMAND_ENABLE,
    TRAY_COMMAND_EXIT,
    TRAY_COMMAND_REBOOT,
    TRAY_COMMAND_SHOW,
)


WM_APP = 0x8000
WM_COMMAND = 0x0111
WM_CLOSE = 0x0010
WM_DESTROY = 0x0002
WM_CONTEXTMENU = 0x007B
WM_NULL = 0x0000
WM_LBUTTONUP = 0x0202
WM_LBUTTONDBLCLK = 0x0203
WM_RBUTTONUP = 0x0205

NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002
NIM_SETVERSION = 0x00000004
NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
NOTIFYICON_VERSION_4 = 4

IMAGE_ICON = 1
LR_LOADFROMFILE = 0x00000010
LR_DEFAULTSIZE = 0x00000040

MF_STRING = 0x00000000
MF_SEPARATOR = 0x00000800
TPM_LEFTALIGN = 0x0000
TPM_RIGHTBUTTON = 0x0002

IDI_APPLICATION = 32512

TRAY_MESSAGE = WM_APP + 1

MENU_SHOW = 1001
MENU_DISABLE = 1002
MENU_ENABLE = 1003
MENU_REBOOT = 1004
MENU_EXIT = 1005

MIM_STYLE = 0x00000010
MNS_NOCHECK = 0x80000000

ATOM = getattr(wintypes, "ATOM", wintypes.WORD)
HBRUSH = getattr(wintypes, "HBRUSH", wintypes.HANDLE)
HCURSOR = getattr(wintypes, "HCURSOR", wintypes.HANDLE)
HICON = getattr(wintypes, "HICON", wintypes.HANDLE)
HMENU = getattr(wintypes, "HMENU", wintypes.HANDLE)
LRESULT = getattr(wintypes, "LRESULT", ctypes.c_ssize_t)
LPCRECT = getattr(wintypes, "LPCRECT", ctypes.POINTER(wintypes.RECT))
UINT_PTR = getattr(wintypes, "UINT_PTR", ctypes.c_size_t)


def loword(value):
    return value & 0xFFFF


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", wintypes.HWND),
        ("message", wintypes.UINT),
        ("wParam", wintypes.WPARAM),
        ("lParam", wintypes.LPARAM),
        ("time", wintypes.DWORD),
        ("pt", POINT),
        ("lPrivate", wintypes.DWORD),
    ]


WNDPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
)


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", HICON),
        ("hCursor", HCURSOR),
        ("hbrBackground", HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", HICON),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", HICON),
    ]


class MENUINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("fMask", wintypes.DWORD),
        ("dwStyle", wintypes.DWORD),
        ("cyMax", wintypes.UINT),
        ("hbrBack", HBRUSH),
        ("dwContextHelpID", wintypes.DWORD),
        ("dwMenuData", ctypes.c_size_t),
    ]


user32 = ctypes.WinDLL("user32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
user32.RegisterClassW.restype = ATOM
user32.UnregisterClassW.argtypes = [wintypes.LPCWSTR, wintypes.HINSTANCE]
user32.UnregisterClassW.restype = wintypes.BOOL
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD,
    wintypes.LPCWSTR,
    wintypes.LPCWSTR,
    wintypes.DWORD,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HWND,
    HMENU,
    wintypes.HINSTANCE,
    wintypes.LPVOID,
]
user32.CreateWindowExW.restype = wintypes.HWND
user32.DefWindowProcW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.DefWindowProcW.restype = LRESULT
user32.DestroyWindow.argtypes = [wintypes.HWND]
user32.DestroyWindow.restype = wintypes.BOOL
user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.PostQuitMessage.restype = None
user32.GetMessageW.argtypes = [ctypes.POINTER(MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = wintypes.BOOL
user32.TranslateMessage.argtypes = [ctypes.POINTER(MSG)]
user32.TranslateMessage.restype = wintypes.BOOL
user32.DispatchMessageW.argtypes = [ctypes.POINTER(MSG)]
user32.DispatchMessageW.restype = LRESULT
user32.CreatePopupMenu.argtypes = []
user32.CreatePopupMenu.restype = HMENU
user32.SetMenuInfo.argtypes = [HMENU, ctypes.POINTER(MENUINFO)]
user32.SetMenuInfo.restype = wintypes.BOOL
user32.AppendMenuW.argtypes = [HMENU, wintypes.UINT, UINT_PTR, wintypes.LPCWSTR]
user32.AppendMenuW.restype = wintypes.BOOL
user32.TrackPopupMenu.argtypes = [
    HMENU,
    wintypes.UINT,
    ctypes.c_int,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.HWND,
    LPCRECT,
]
user32.TrackPopupMenu.restype = wintypes.BOOL
user32.DestroyMenu.argtypes = [HMENU]
user32.DestroyMenu.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [wintypes.HWND]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.PostMessageW.restype = wintypes.BOOL
user32.GetCursorPos.argtypes = [ctypes.POINTER(POINT)]
user32.GetCursorPos.restype = wintypes.BOOL
user32.LoadImageW.argtypes = [
    wintypes.HINSTANCE,
    wintypes.LPCWSTR,
    wintypes.UINT,
    ctypes.c_int,
    ctypes.c_int,
    wintypes.UINT,
]
user32.LoadImageW.restype = wintypes.HANDLE
user32.LoadIconW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
user32.LoadIconW.restype = HICON
user32.DestroyIcon.argtypes = [HICON]
user32.DestroyIcon.restype = wintypes.BOOL
shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.POINTER(NOTIFYICONDATAW)]
shell32.Shell_NotifyIconW.restype = wintypes.BOOL
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE


class TrayIcon:
    def __init__(self, tooltip, command_queue, labels, icon_path=None):
        self.tooltip = tooltip[:127]
        self.command_queue = command_queue
        self.labels = labels
        self.icon_path = icon_path
        self.hwnd = None
        self.hicon = None
        self._class_name = f"{APP_ID}.TrayWindow"
        self._window_proc = WNDPROC(self._wnd_proc)
        self._thread = None
        self._ready = threading.Event()
        self._error = None
        self._command_ids = {
            MENU_SHOW: TRAY_COMMAND_SHOW,
            MENU_DISABLE: TRAY_COMMAND_DISABLE,
            MENU_ENABLE: TRAY_COMMAND_ENABLE,
            MENU_REBOOT: TRAY_COMMAND_REBOOT,
            MENU_EXIT: TRAY_COMMAND_EXIT,
        }

    def start(self):
        if self._thread and self._thread.is_alive():
            return

        self._thread = threading.Thread(target=self._run, name="tray-icon", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=5)

        if self._error:
            raise RuntimeError(self._error)

    def stop(self):
        if self.hwnd:
            user32.PostMessageW(self.hwnd, WM_CLOSE, 0, 0)
        if self._thread:
            self._thread.join(timeout=5)

    def _load_icon(self):
        if self.icon_path:
            handle = user32.LoadImageW(
                None,
                self.icon_path,
                IMAGE_ICON,
                0,
                0,
                LR_LOADFROMFILE | LR_DEFAULTSIZE,
            )
            if handle:
                return HICON(handle)

        return user32.LoadIconW(None, ctypes.cast(ctypes.c_void_p(IDI_APPLICATION), wintypes.LPCWSTR))

    def _build_notify_data(self):
        data = NOTIFYICONDATAW()
        data.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        data.hWnd = self.hwnd
        data.uID = 1
        data.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
        data.uCallbackMessage = TRAY_MESSAGE
        data.hIcon = self.hicon
        data.szTip = self.tooltip
        data.uVersion = NOTIFYICON_VERSION_4
        return data

    def _show_menu(self):
        menu = user32.CreatePopupMenu()
        if not menu:
            return

        try:
            menu_info = MENUINFO()
            menu_info.cbSize = ctypes.sizeof(MENUINFO)
            menu_info.fMask = MIM_STYLE
            menu_info.dwStyle = MNS_NOCHECK
            user32.SetMenuInfo(menu, ctypes.byref(menu_info))

            user32.AppendMenuW(menu, MF_STRING, MENU_SHOW, self.labels["show"])
            user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
            user32.AppendMenuW(menu, MF_STRING, MENU_DISABLE, self.labels["disable"])
            user32.AppendMenuW(menu, MF_STRING, MENU_ENABLE, self.labels["enable"])
            user32.AppendMenuW(menu, MF_STRING, MENU_REBOOT, self.labels["reboot"])
            user32.AppendMenuW(menu, MF_SEPARATOR, 0, None)
            user32.AppendMenuW(menu, MF_STRING, MENU_EXIT, self.labels["exit"])

            point = POINT()
            user32.GetCursorPos(ctypes.byref(point))
            user32.SetForegroundWindow(self.hwnd)
            user32.TrackPopupMenu(
                menu,
                TPM_LEFTALIGN | TPM_RIGHTBUTTON,
                point.x,
                point.y,
                0,
                self.hwnd,
                None,
            )
            user32.PostMessageW(self.hwnd, WM_NULL, 0, 0)
        finally:
            user32.DestroyMenu(menu)

    def _wnd_proc(self, hwnd, message, w_param, l_param):
        if message == TRAY_MESSAGE:
            event_code = loword(l_param)

            if event_code in (WM_LBUTTONUP, WM_LBUTTONDBLCLK):
                self.command_queue.put(TRAY_COMMAND_SHOW)
                return 0

            if event_code in (WM_RBUTTONUP, WM_CONTEXTMENU):
                self._show_menu()
                return 0

        if message == WM_COMMAND:
            menu_id = loword(w_param)
            command = self._command_ids.get(menu_id)
            if command:
                self.command_queue.put(command)
                return 0

        if message == WM_CLOSE:
            data = self._build_notify_data()
            shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(data))
            user32.DestroyWindow(hwnd)
            return 0

        if message == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0

        return user32.DefWindowProcW(hwnd, message, w_param, l_param)

    def _run(self):
        h_instance = kernel32.GetModuleHandleW(None)
        class_info = WNDCLASSW()
        class_info.lpfnWndProc = self._window_proc
        class_info.hInstance = h_instance
        class_info.lpszClassName = self._class_name

        atom = 0
        try:
            atom = user32.RegisterClassW(ctypes.byref(class_info))
            self.hwnd = user32.CreateWindowExW(
                0,
                self._class_name,
                self._class_name,
                0,
                0,
                0,
                0,
                0,
                None,
                None,
                h_instance,
                None,
            )
            if not self.hwnd:
                raise RuntimeError("could not create tray window")

            self.hicon = self._load_icon()
            data = self._build_notify_data()
            if not shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(data)):
                raise RuntimeError("could not create tray icon")
            shell32.Shell_NotifyIconW(NIM_SETVERSION, ctypes.byref(data))
        except Exception as exc:
            self._error = str(exc)
            if self.hwnd:
                user32.DestroyWindow(self.hwnd)
                self.hwnd = None
            if self.hicon:
                user32.DestroyIcon(self.hicon)
                self.hicon = None
            if atom:
                user32.UnregisterClassW(self._class_name, h_instance)
            self._ready.set()
            return

        self._ready.set()

        message = MSG()
        while user32.GetMessageW(ctypes.byref(message), None, 0, 0) > 0:
            user32.TranslateMessage(ctypes.byref(message))
            user32.DispatchMessageW(ctypes.byref(message))

        if self.hicon:
            user32.DestroyIcon(self.hicon)
            self.hicon = None
        self.hwnd = None
        if atom:
            user32.UnregisterClassW(self._class_name, h_instance)
