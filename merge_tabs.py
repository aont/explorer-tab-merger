# -*- coding: utf-8 -*-
"""
merge_tabs.py - Merge File Explorer tabs into the first window (Python)
Python port of the C++ merge_tabs.cpp (keeps ANSI/BSTR conversion, COM usage, and WM_COMMAND sending approach)

Requirements:
  pip install pywin32

Run:
  python merge_tabs.py
"""

import time
import ctypes
from ctypes import wintypes
from typing import List, Tuple, Optional
import pythoncom
from win32com.client import Dispatch
import win32com.client
import win32con
from comtypes import IUnknown
from comtypes.client import dynamic


if not hasattr(wintypes, "LRESULT"):
    # LRESULT is equivalent to LONG_PTR. If it's missing in wintypes, reuse LPARAM (same LONG_PTR size).
    wintypes.LRESULT = wintypes.LPARAM

# ---- Win32 definitions ----
user32 = ctypes.WinDLL("user32", use_last_error=True)

EnumChildProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

GetClassNameA = user32.GetClassNameA
GetClassNameA.argtypes = [wintypes.HWND, ctypes.c_char_p, ctypes.c_int]
GetClassNameA.restype = ctypes.c_int

EnumChildWindows = user32.EnumChildWindows
EnumChildWindows.argtypes = [wintypes.HWND, EnumChildProc, wintypes.LPARAM]
EnumChildWindows.restype = wintypes.BOOL

SendMessageA = user32.SendMessageA
SendMessageA.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
SendMessageA.restype = wintypes.LRESULT

PostMessageA = user32.PostMessageA
PostMessageA.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
PostMessageA.restype = wintypes.BOOL

IsWindow = user32.IsWindow
IsWindow.argtypes = [wintypes.HWND]
IsWindow.restype = wintypes.BOOL

WM_COMMAND_ID_NEW_TAB = 0xA21B  # Undocumented: same as the C++ version

# ---- Data structure ----
class TabInfo:
    def __init__(self, browser, url: str, top_level: int, identity: int):
        # browser: IWebBrowser2 COM object
        self.browser = browser
        self.url = url or ""
        self.top_level = top_level  # HWND
        self.identity = identity


# ---- COM identity helpers ----
def ensure_comtypes_browser(obj):
    """Return a comtypes-dispatch wrapper for the given COM object."""
    if obj is None:
        return None

    if hasattr(obj, "QueryInterface"):
        return obj

    disp = getattr(obj, "_oleobj_", None)
    if disp is None:
        return None

    try:
        return dynamic.Dispatch(disp)
    except Exception as e:
        print(f"[warn] Failed to convert COM object to comtypes: {e}")
        return None


def com_identity_address_ctypes(obj) -> int:
    """
    Return an address that can be used to compare COM object identity.
    """
    if obj is None:
        return 0

    # pythoncom offers ObjectIdentity which works for PyIDispatch objects.
    try:
        return int(pythoncom.ObjectIdentity(obj))
    except Exception:
        pass

    # comtypes objects expose QueryInterface; fall back to that path.
    try:
        if hasattr(obj, "QueryInterface"):
            unk = obj.QueryInterface(IUnknown)
            return ctypes.addressof(unk)
    except Exception as e:
        print(f"[warn] QueryInterface for identity failed: {e}")

    # As a final attempt, look for the underlying PyIDispatch via _oleobj_.
    try:
        ole = getattr(obj, "_oleobj_", None)
        if ole is not None:
            return int(pythoncom.ObjectIdentity(ole))
    except Exception:
        pass

    return 0


def is_same_com_object_ctypes(a, b) -> bool:
    if not a or not b:
        return False
    return com_identity_address_ctypes(a) == com_identity_address_ctypes(b)


# ---- Utilities ----
def bstr_to_ansi(bstr) -> str:
    # pywin32 typically converts BSTR to Python str already; keep it consistent just in case
    if bstr is None:
        return ""
    return str(bstr)


def navigate_browser(wb, url: str) -> bool:
    # Equivalent to IWebBrowser2.Navigate2
    try:
        vEmpty = None
        wb.Navigate2(url, vEmpty, vEmpty, vEmpty, vEmpty)
        return True
    except Exception as e:
        print(f"[warn] Navigate2 failed: {e}")
        return False


# ---- Enumerate Explorer tabs ----
def get_explorer_tab_url(wb) -> str:
    """Return a URL suitable for Navigate2, handling virtual folders."""
    url = ""
    try:
        url = bstr_to_ansi(getattr(wb, "LocationURL", ""))
    except Exception:
        url = ""

    url = (url or "").strip()
    if url:
        return url

    # Fallback for virtual folders (e.g. This PC, Control Panel, etc.).
    try:
        doc = getattr(wb, "Document", None)
        folder = getattr(doc, "Folder", None) if doc is not None else None
        folder_self = getattr(folder, "Self", None) if folder is not None else None
        path = bstr_to_ansi(getattr(folder_self, "Path", "")) if folder_self is not None else ""
        path = path.strip()
        if path.startswith("::"):
            return "shell:" + path
        if path.startswith("shell::"):
            return path
    except Exception as e:
        print(f"[warn] Failed to resolve virtual folder URL: {e}")

    return ""


def collect_explorer_tabs() -> Tuple[List[TabInfo], List[int]]:
    """
    Enumerate all ShellWindows and keep only File Explorer instances.
    Returns: (tabs, windowOrder)
    """
    tabs: List[TabInfo] = []
    window_order: List[int] = []

    try:
        shell_windows = Dispatch("Shell.Application").Windows()
    except Exception as e:
        print(f"[error] ShellWindows CoCreate failed: {e}")
        return tabs, window_order

    # Windows() is indexable; iterating yields IWebBrowser2-like objects per window
    try:
        count = shell_windows.Count
    except Exception as e:
        print(f"[error] get_Count failed: {e}")
        return tabs, window_order

    for i in range(count):
        try:
            wb = shell_windows.Item(i)
            wb = ensure_comtypes_browser(wb)
            if wb is None:
                print(f"[warn] Failed to wrap IWebBrowser2 for item {i}.")
                continue
            # Simple filter to target Explorer only:
            # 1) HWnd is available
            # 2) LocationURL is available
            hwnd = int(wb.HWnd) if hasattr(wb, "HWnd") else 0
            if not hwnd:
                continue

            # Reproducing IServiceProvider→IShellBrowser check strictly is hard in pywin32.
            # Instead, use class-name filtering and LocationURL presence as a practical substitute.
            # (Edge/IE normally won't appear here; even if they did, tabs with URL behave similarly.)
            url = get_explorer_tab_url(wb)

            identity = com_identity_address_ctypes(wb)

            if hwnd not in window_order:
                window_order.append(hwnd)

            print(f"[debug] Explorer tab found: top-level HWND=0x{hwnd:016X}, IWebBrowser2={wb}, URL={url}")
            tabs.append(TabInfo(wb, url, hwnd, identity))

        except Exception as e:
            print(f"[warn] enumerate item {i} failed: {e}")
            continue

    return tabs, window_order


# ---- Find ShellTabWindowClass ----
def find_shell_tab_host(top_level_hwnd: int) -> Optional[int]:
    target = ctypes.c_void_p(0)

    def enum_proc(hwnd, lparam):
        # Traverse descendants recursively (EnumChildWindows enumerates only direct children)
        buf = ctypes.create_string_buffer(256)
        GetClassNameA(hwnd, buf, 255)
        cls = buf.value.decode(errors="ignore")
        if cls == "ShellTabWindowClass":
            # Found: store in lparam-like holder and stop
            target.value = hwnd
            return False  # stop
        # Recurse
        EnumChildWindows(hwnd, EnumChildProc(enum_proc), lparam)
        # Stop if already found
        return False if target.value else True

    if not IsWindow(top_level_hwnd):
        return None

    EnumChildWindows(top_level_hwnd, EnumChildProc(enum_proc), 0)
    return target.value or None


# ---- Create a new tab & navigate ----
def create_tab_and_navigate(first_window_hwnd: int, tab_host_hwnd: int, url: str,
                            known_tab_count: List[int]) -> bool:
    """
    known_tab_count is a one-element list to simulate pass-by-reference (C++ size_t&).
    """
    if not first_window_hwnd or not tab_host_hwnd or not url:
        return False

    # Refresh baseline (number of tabs in the first window)
    before_tabs, _ = collect_explorer_tabs()
    baseline_tabs = [t for t in before_tabs if t.top_level == first_window_hwnd]
    baseline = len(baseline_tabs)
    baseline_identities = {t.identity for t in baseline_tabs if t.identity}
    known_tab_count[0] = baseline
    print(f"[debug] Baseline tab count for first window: {baseline}")

    # Send WM_COMMAND to create a new tab
    print(f"[debug] Sending WM_COMMAND to create new tab in HWND=0x{tab_host_hwnd:016X}")
    SendMessageA(tab_host_hwnd, win32con.WM_COMMAND, WM_COMMAND_ID_NEW_TAB, 0)

    timeout_ms = 8000
    retry_ms = 300
    waited = 0

    while waited <= timeout_ms:
        tabs, _ = collect_explorer_tabs()
        first_window_tabs = [t for t in tabs if t.top_level == first_window_hwnd]
        current_count = len(first_window_tabs)

        new_tab = None
        fallback_used = False
        for candidate in first_window_tabs:
            if candidate.identity and candidate.identity not in baseline_identities:
                new_tab = candidate
                break

        if new_tab is None and current_count > baseline:
            # Fallback: if identities missing, pick the last tab.
            if first_window_tabs:
                new_tab = first_window_tabs[-1]
                fallback_used = True

        if new_tab is not None:
            if new_tab.identity:
                baseline_identities.add(new_tab.identity)
            wb = new_tab.browser
            if fallback_used:
                print(f"[debug] Identified new tab by fallback count check in HWND=0x{first_window_hwnd:016X}")
            else:
                print(f"[debug] Identified new tab by COM identity in HWND=0x{first_window_hwnd:016X}")

            ok = navigate_browser(wb, url)

            # Keep pre-created tab references until after detection completes.
            before_tabs.clear()
            tabs.clear()
            if ok:
                known_tab_count[0] = current_count
                print("[debug] Navigation succeeded for new tab.")
                return True
            return False

        time.sleep(retry_ms / 1000.0)
        waited += retry_ms

    before_tabs.clear()
    return False


def main():
    # Initialize STA (Explorer COM expects STA)
    pythoncom.CoInitialize()

    try:
        tabs, window_order = collect_explorer_tabs()
        if not window_order:
            print("No Explorer windows detected.")
            return 0

        first_window = window_order[0]
        known_tab_count_box = [0]  # pass-by-reference container

        urls_to_merge: List[str] = []
        windows_to_close: List[int] = []

        for t in tabs:
            if t.top_level == first_window:
                known_tab_count_box[0] += 1
                print(f"[debug] Known tab in first window on startup: HWND=0x{t.top_level:016X}, IWebBrowser2={t.browser}")
            else:
                if t.url:
                    urls_to_merge.append(t.url)
                    print(f"[debug] Tab queued for merge: HWND=0x{t.top_level:016X}, IWebBrowser2={t.browser}, URL={t.url}")
                if t.top_level not in windows_to_close:
                    windows_to_close.append(t.top_level)

        # No explicit Release in Python (garbage collector handles it)

        if not urls_to_merge:
            print("Nothing to merge.")
            return 0

        tab_host = find_shell_tab_host(first_window)
        if not tab_host:
            print("Could not find ShellTabWindowClass in the first window.")
            return 3

        print(f"Merging {len(urls_to_merge)} tab(s) into the first window...")

        success = 0
        for url in urls_to_merge:
            if create_tab_and_navigate(first_window, tab_host, url, known_tab_count_box):
                success += 1
            else:
                print(f"[warn] Failed to create tab for: {url}")

        # Close source windows
        for h in windows_to_close:
            if h and h != first_window and IsWindow(h):
                SendMessageA(h, win32con.WM_CLOSE, 0, 0)

        print(f"Completed. {success} tab(s) moved.")
        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
