"""
===============================================================================
Script: Debug Save As Dialog - UIA & Win32 Control Inspector
===============================================================================

Description:
Diagnostic tool to inspect the exact window geometry, DPI scaling, Win32 child
HWNDs, and UI Automation (UIA) control tree of Power BI Desktop's 'Save As'
dialog. Writes all findings to 'debug_save_as_tree.txt' and captures a
diagnostic screenshot ('debug_save_as_screen.bmp') so we can verify exact
coordinates and control types without guessing.

Usage:
1. Open the 'Save As' dialog in Power BI Desktop ('...' -> 'Export data'),
   then run:
       py debug_save_as.py
   (Or run `py debug_save_as.py --trigger-export` to open Save As automatically.)
2. Paste the terminal output (or 'debug_save_as_tree.txt') so we can lock in
   the exact control IDs and coordinates.
===============================================================================
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import sys
import time
from pathlib import Path

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from pywinauto import Desktop

OUT_DIR = Path(r"C:\Users\jd0855\Desktop\VCI_Automation\bulk-send-emails-ljso")
REPORT_TXT = OUT_DIR / "debug_save_as_tree.txt"
SCREEN_BMP = OUT_DIR / "debug_save_as_screen.bmp"


def _get_hwnd_text(hwnd: int) -> str:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value.strip()


def _get_hwnd_class(hwnd: int) -> str:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    buf = ctypes.create_unicode_buffer(128)
    user32.GetClassNameW(hwnd, buf, 128)
    return buf.value


def _get_hwnd_rect(hwnd: int) -> tuple[int, int, int, int]:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    r = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(r))
    return int(r.left), int(r.top), int(r.right), int(r.bottom)


def _find_all_save_as_windows() -> list[int]:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    IsWindowVisible = user32.IsWindowVisible

    found: list[int] = []

    @EnumWindowsProc
    def cb(hwnd, lparam):
        if not IsWindowVisible(hwnd):
            return True
        title = _get_hwnd_text(hwnd)
        cls = _get_hwnd_class(hwnd)
        if any(k in title.casefold() for k in ("save as", "export as", "export data")) or cls == "#32770":
            found.append(int(hwnd))
        return True

    EnumWindows(cb, 0)
    return found


def _save_region_bmp(left: int, top: int, width: int, height: int, out_path: Path) -> bool:
    if width <= 0 or height <= 0:
        return False
    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    hdc_screen = user32.GetDC(0)
    if not hdc_screen:
        return False
    hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
    hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, width, height)
    old_obj = gdi32.SelectObject(hdc_mem, hbmp)

    try:
        SRCCOPY = 0x00CC0020
        CAPTUREBLT = 0x40000000
        if not gdi32.BitBlt(hdc_mem, 0, 0, width, height, hdc_screen, left, top, SRCCOPY | CAPTUREBLT):
            return False

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [
                ("biSize", wintypes.DWORD),
                ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG),
                ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD),
                ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD),
                ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG),
                ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD),
            ]

        row_bytes = ((width * 24 + 31) // 32) * 4
        img_size = row_bytes * height
        bmi = BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.biWidth = width
        bmi.biHeight = height  # bottom-up BMP
        bmi.biPlanes = 1
        bmi.biBitCount = 24
        bmi.biCompression = 0
        bmi.biSizeImage = img_size

        buf = ctypes.create_string_buffer(img_size)
        if gdi32.GetDIBits(hdc_mem, hbmp, 0, height, buf, ctypes.byref(bmi), 0) != height:
            return False

        file_header = (
            b"BM"
            + (54 + img_size).to_bytes(4, "little")
            + (0).to_bytes(2, "little")
            + (0).to_bytes(2, "little")
            + (54).to_bytes(4, "little")
        )
        out_path.write_bytes(file_header + bytes(bmi) + buf.raw)
        return True
    finally:
        gdi32.SelectObject(hdc_mem, old_obj)
        gdi32.DeleteObject(hbmp)
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(0, hdc_screen)


def inspect_save_as(trigger_export: bool = False) -> None:
    lines: list[str] = []
    user32 = ctypes.WinDLL("user32", use_last_error=True)

    # Virtual desktop bounds
    v_left = user32.GetSystemMetrics(76)    # SM_XVIRTUALSCREEN
    v_top = user32.GetSystemMetrics(77)     # SM_YVIRTUALSCREEN
    v_width = user32.GetSystemMetrics(78)   # SM_CXVIRTUALSCREEN
    v_height = user32.GetSystemMetrics(79)  # SM_CYVIRTUALSCREEN
    lines.append(f"Virtual Desktop: left={v_left}, top={v_top}, width={v_width}, height={v_height}")

    if trigger_export:
        from export_managers_info import (
            OPEN_TIMEOUT_SECONDS,
            PBIX_PATH,
            WINDOW_TITLE_KEYWORD,
            _click_more_options_and_export,
            open_or_attach_pbix,
        )
        win = open_or_attach_pbix(PBIX_PATH, WINDOW_TITLE_KEYWORD, OPEN_TIMEOUT_SECONDS)
        _click_more_options_and_export(win)
        time.sleep(1.5)

    hwnds = _find_all_save_as_windows()
    if not hwnds:
        raise RuntimeError(
            "No visible 'Save As' window found. Either open 'Save As' in Power BI first "
            "or run: py debug_save_as.py --trigger-export"
        )

    lines.append(f"\nFound {len(hwnds)} matching Save As / #32770 window(s):")
    for h in hwnds:
        lines.append(f"  HWND={hex(h)} | class='{_get_hwnd_class(h)}' | title='{_get_hwnd_text(h)}' | rect={_get_hwnd_rect(h)}")

    target_hwnd = hwnds[0]
    dlg_rect = _get_hwnd_rect(target_hwnd)
    try:
        dpi = user32.GetDpiForWindow(target_hwnd)
    except Exception:
        dpi = 96
    lines.append(f"\nInspecting HWND={hex(target_hwnd)} | rect={dlg_rect} | DPI={dpi} ({dpi/96.0:.2f}x)")

    # Save screenshot of the Save As dialog
    pad = 20
    cap_l = dlg_rect[0] - pad
    cap_t = dlg_rect[1] - pad
    cap_w = (dlg_rect[2] - dlg_rect[0]) + pad * 2
    cap_h = (dlg_rect[3] - dlg_rect[1]) + pad * 2
    if _save_region_bmp(cap_l, cap_t, cap_w, cap_h, SCREEN_BMP):
        lines.append(f"Saved dialog screenshot to: {SCREEN_BMP}")

    # 1. Enumerate all Win32 child HWNDs
    lines.append("\n================ WIN32 CHILD HWNDS ================")
    EnumChildWindows = user32.EnumChildWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    GetDlgCtrlID = user32.GetDlgCtrlID
    IsWindowVisible = user32.IsWindowVisible

    win32_children: list[str] = []

    @EnumWindowsProc
    def child_cb(chwnd, lparam):
        vis = bool(IsWindowVisible(chwnd))
        cls = _get_hwnd_class(chwnd)
        cid = GetDlgCtrlID(chwnd)
        crect = _get_hwnd_rect(chwnd)
        ctxt = _get_hwnd_text(chwnd)
        win32_children.append(
            f"  hwnd={hex(chwnd)} | vis={vis} | id={cid} | class='{cls}' | rect={crect} | text='{ctxt}'"
        )
        return True

    EnumChildWindows(target_hwnd, child_cb, 0)
    lines.extend(win32_children)

    # 2. Enumerate all UIA descendants
    lines.append("\n================ UIA DESCENDANTS ================")
    uia_win = Desktop(backend="uia").window(handle=target_hwnd)
    for idx, c in enumerate(uia_win.descendants()):
        try:
            ctype = str(c.element_info.control_type or "")
            aid = str(c.element_info.automation_id or "")
            cname = str(c.element_info.name or "")
            crect = c.rectangle()
            r_tuple = (crect.left, crect.top, crect.right, crect.bottom)
            vis = c.is_visible()
            val = ""
            try:
                val = str(c.get_value() or "")
            except Exception:
                try:
                    val = str(c.window_text() or "")
                except Exception:
                    val = ""
            lines.append(
                f"  [{idx:03d}] vis={vis} | type={ctype} | id='{aid}' | name='{cname}' | rect={r_tuple} | val='{val}'"
            )
        except Exception as exc:
            lines.append(f"  [{idx:03d}] <error reading control: {exc}>")

    output = "\n".join(lines)
    REPORT_TXT.write_text(output, encoding="utf-8")
    print(output)
    print(f"\nDiagnostic report saved to: {REPORT_TXT}")


if __name__ == "__main__":
    trigger = "--trigger-export" in sys.argv
    inspect_save_as(trigger_export=trigger)