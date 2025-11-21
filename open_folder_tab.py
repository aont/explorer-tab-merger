# -*- coding: utf-8 -*-
"""
open_folder_tab.py - Open a folder in a new tab of the first Explorer window, or launch the folder if no window exists.

Usage:
    python open_folder_tab.py "C:\\Path\\To\\Folder"
"""

import sys
import time
import ctypes
from ctypes import wintypes
from typing import List, Tuple, Optional
import pythoncom
from win32com.client import Dispatch
import win32con

if not hasattr(wintypes, "LRESULT"):
    wintypes.LRESULT = wintypes.LPARAM

# ---- Win32 definitions ----
user32 = ctypes.WinDLL("user32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)

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

IsWindow = user32.IsWindow
IsWindow.argtypes = [wintypes.HWND]
IsWindow.restype = wintypes.BOOL

ShellExecuteW = shell32.ShellExecuteW
ShellExecuteW.argtypes = [wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_int]
ShellExecuteW.restype = wintypes.HINSTANCE

WM_COMMAND_ID_NEW_TAB = 0xA21B


class TabInfo:
    def __init__(self, browser, url: str, top_level: int):
        self.browser = browser
        self.url = url or ""
        self.top_level = top_level


def bstr_to_ansi(bstr) -> str:
    if bstr is None:
        return ""
    return str(bstr)


def normalize_folder_path(path: str) -> str:
    return path.replace("/", "\\")


def navigate_browser(wb, url: str) -> bool:
    try:
        vEmpty = None
        wb.Navigate2(url, vEmpty, vEmpty, vEmpty, vEmpty)
        return True
    except Exception as e:
        print(f"[warn] Navigate2 failed: {e}")
        return False


def get_explorer_tab_url(wb) -> str:
    url = ""
    try:
        url = bstr_to_ansi(getattr(wb, "LocationURL", ""))
    except Exception:
        url = ""

    url = (url or "").strip()
    if url:
        return url

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
    tabs: List[TabInfo] = []
    window_order: List[int] = []

    try:
        shell_windows = Dispatch("Shell.Application").Windows()
    except Exception as e:
        print(f"[error] ShellWindows CoCreate failed: {e}")
        return tabs, window_order

    try:
        count = shell_windows.Count
    except Exception as e:
        print(f"[error] get_Count failed: {e}")
        return tabs, window_order

    for i in range(count):
        try:
            wb = shell_windows.Item(i)
            hwnd = int(wb.HWnd) if hasattr(wb, "HWnd") else 0
            if not hwnd:
                continue

            url = get_explorer_tab_url(wb)

            if hwnd not in window_order:
                window_order.append(hwnd)

            tabs.append(TabInfo(wb, url, hwnd))
        except Exception as e:
            print(f"[warn] enumerate item {i} failed: {e}")
            continue

    return tabs, window_order


def find_shell_tab_host(top_level_hwnd: int) -> Optional[int]:
    target = ctypes.c_void_p(0)

    def enum_proc(hwnd, lparam):
        buf = ctypes.create_string_buffer(256)
        GetClassNameA(hwnd, buf, 255)
        cls = buf.value.decode(errors="ignore")
        if cls == "ShellTabWindowClass":
            target.value = hwnd
            return False
        EnumChildWindows(hwnd, EnumChildProc(enum_proc), lparam)
        return False if target.value else True

    if not IsWindow(top_level_hwnd):
        return None

    EnumChildWindows(top_level_hwnd, EnumChildProc(enum_proc), 0)
    return target.value or None


def create_tab_and_navigate(first_window_hwnd: int, tab_host_hwnd: int, url: str) -> bool:
    if not first_window_hwnd or not tab_host_hwnd or not url:
        return False

    before_tabs, _ = collect_explorer_tabs()
    baseline_browsers = [t.browser for t in before_tabs if t.top_level == first_window_hwnd]
    baseline = len(baseline_browsers)

    SendMessageA(tab_host_hwnd, win32con.WM_COMMAND, WM_COMMAND_ID_NEW_TAB, 0)

    timeout_ms = 8000
    retry_ms = 300
    waited = 0

    while waited <= timeout_ms:
        tabs, _ = collect_explorer_tabs()
        first_window_tabs = [t for t in tabs if t.top_level == first_window_hwnd]
        current_count = len(first_window_tabs)

        new_tab_info = None
        for candidate in first_window_tabs:
            is_new = True
            for baseline_browser in baseline_browsers:
                try:
                    if candidate.browser == baseline_browser:
                        is_new = False
                        break
                except Exception as e:
                    print(f"[warn] Comparison with baseline browser failed: {e}")
            if is_new:
                new_tab_info = candidate
                break

        if new_tab_info is not None and current_count > baseline:
            ok = navigate_browser(new_tab_info.browser, url)
            del tabs[:]
            baseline_browsers.clear()
            if ok:
                return True
            return False

        time.sleep(retry_ms / 1000.0)
        waited += retry_ms

    baseline_browsers.clear()
    return False


def launch_with_shell_execute(path: str) -> int:
    result = ShellExecuteW(None, "open", path, None, None, win32con.SW_SHOWNORMAL)
    return 0 if result > 32 else 2


def main(argv: List[str]) -> int:
    if len(argv) < 2:
        print("Usage: python open_folder_tab.py <folder path>")
        return 1

    target_path = normalize_folder_path(argv[1])
    if not target_path:
        print("Empty folder path provided.")
        return 1

    pythoncom.CoInitialize()
    try:
        tabs, window_order = collect_explorer_tabs()

        if not window_order:
            print("No Explorer window found; launching folder via ShellExecute.")
            return launch_with_shell_execute(target_path)

        first_window = window_order[0]
        tab_host = find_shell_tab_host(first_window)
        if not tab_host:
            print("Could not find ShellTabWindowClass in the first window.")
            return 3

        if not create_tab_and_navigate(first_window, tab_host, target_path):
            print("Failed to create or navigate new tab; falling back to ShellExecute.")
            return launch_with_shell_execute(target_path)

        return 0
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

