"""
===============================================================================
Script: Managers Info PBIX - Data Refresh, Save & Table Visual CSV Export
===============================================================================

Description:
Automates exporting the 'Manager Contact Information' table visual from Power BI
Desktop ('W:/Dashboards VCI/LJSO/Managers/Managers Info.pbix') into a dated CSV
file in the same directory. Includes a SKIP_DATA_REFRESH toggle to bypass data
refresh and PBIX saving when you only need a fast table export, while keeping
the DPI-aware pixel detection for the top-right More Options (...) menu.

Workflow Outline:
1. DPI & Window Acquisition: Enables Per-Monitor DPI awareness, attaches to
   an open 'Managers Info' Power BI Desktop window (or launches the PBIX),
   and maximizes it for consistent canvas geometry.
2. Data-Only Refresh & Save (Optional - controlled by SKIP_DATA_REFRESH):
   When enabled, triggers 'Data' only refresh, closes any query warning modal,
   and saves the PBIX file. Skipped when SKIP_DATA_REFRESH = True.
3. Table Visual Export: Locates the dark blue 'Managers Information' header
   banner via native Windows GDI capture, clicks inside the table visual,
   scans the top-right header to lock onto the exact '...' ('More options')
   dots, hovers and clicks '...', and selects 'Export data'.
4. File Save & Overwrite Verification: Inspects the UIA 'Address:' toolbar
   in 'Save As', navigates to 'W:/Dashboards VCI/LJSO/Managers' if needed,
   waits for the network folder view to finish loading, clicks the UIA
   'File name:' Edit box (id='1001'), clears and pastes
   'managers-contact-info-YYYY-MM-DD' via real keyboard input so Windows
   IFileSaveDialog registers the filename change, clicks 'Save' (id='1'),
   handles any overwrite prompt, and verifies the CSV file on disk.
===============================================================================
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import logging
import os
import shutil
import time
from datetime import date
from pathlib import Path

# Enable Per-Monitor DPI Awareness BEFORE importing pywinauto so window
# rectangles, GDI screen captures, and mouse coordinates match 1:1 on
# high-DPI / multi-monitor setups.
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from pywinauto import Desktop, keyboard, mouse

# ==========================================
# CONFIGURATION & PATHS
# ==========================================
SKIP_DATA_REFRESH = True  # Set to False to run 'Data' only refresh + PBIX save before exporting

PBIX_PATH = Path(r"W:\Dashboards VCI\LJSO\Managers\Managers Info.pbix")
EXPORT_DIR = Path(r"W:\Dashboards VCI\LJSO\Managers")
WINDOW_TITLE_KEYWORD = "Managers Info"
DEFAULT_VISUAL_EXPORT_NAME = "Manager Contact Information.csv"
EXPORT_PREFIX = "managers-contact-info"

OPEN_TIMEOUT_SECONDS = 180
REFRESH_TIMEOUT_SECONDS = 1200
SAVE_PBIX_TIMEOUT_SECONDS = 120
RENDER_TIMEOUT_SECONDS = 45
SAVE_CSV_TIMEOUT_SECONDS = 90

LOG = logging.getLogger("export_managers_info")


class WorkflowError(RuntimeError):
    pass


ExportWorkflowError = WorkflowError


def text(control) -> str:
    try:
        return str(control.window_text() or "").strip()
    except Exception:
        return ""


def visible(control) -> bool:
    try:
        r = control.rectangle()
        return control.is_visible() and r.right > r.left and r.bottom > r.top
    except Exception:
        return False


def _set_clipboard_text(value: str) -> None:
    """Copies a Unicode string onto the Windows clipboard (64-bit pointer safe)."""
    CF_UNICODETEXT = 13
    GMEM_MOVEABLE = 0x0002

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL

    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL

    if not user32.OpenClipboard(None):
        time.sleep(0.1)
        if not user32.OpenClipboard(None):
            raise WorkflowError("Unable to open Windows clipboard.")
    try:
        user32.EmptyClipboard()
        data = value.encode("utf-16-le") + b"\x00\x00"
        h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not h_mem:
            raise WorkflowError("GlobalAlloc failed while setting clipboard text.")
        ptr = kernel32.GlobalLock(h_mem)
        if not ptr:
            raise WorkflowError("GlobalLock returned NULL while setting clipboard text.")
        try:
            ctypes.memmove(ptr, data, len(data))
        finally:
            kernel32.GlobalUnlock(h_mem)
        if not user32.SetClipboardData(CF_UNICODETEXT, h_mem):
            raise WorkflowError("SetClipboardData failed.")
    finally:
        user32.CloseClipboard()


def _get_dpi_scale(window) -> float:
    """Returns the DPI scale factor (e.g., 1.0, 1.25, 1.5) for the monitor hosting `window`."""
    try:
        hwnd = window.handle
        dpi = ctypes.windll.user32.GetDpiForWindow(hwnd)
        if dpi and dpi > 0:
            return dpi / 96.0
    except Exception:
        pass
    return 1.0


def _grab_screen_region_rgb(
    left: int, top: int, width: int, height: int
) -> list[list[tuple[int, int, int]]] | None:
    """
    Captures a rectangular region of the virtual desktop across primary and
    secondary monitors in 1:1 physical pixels using Windows GDI BitBlt.
    """
    if width <= 0 or height <= 0:
        return None

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32

    hdc_screen = user32.GetDC(0)
    if not hdc_screen:
        return None

    hdc_mem = gdi32.CreateCompatibleDC(hdc_screen)
    hbmp = gdi32.CreateCompatibleBitmap(hdc_screen, width, height)
    old_obj = gdi32.SelectObject(hdc_mem, hbmp)

    try:
        SRCCOPY = 0x00CC0020
        CAPTUREBLT = 0x40000000
        ok = gdi32.BitBlt(
            hdc_mem, 0, 0, width, height, hdc_screen, left, top, SRCCOPY | CAPTUREBLT
        )
        if not ok:
            return None

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

        bmi = BITMAPINFOHEADER()
        bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.biWidth = width
        bmi.biHeight = -height  # top-down DIB
        bmi.biPlanes = 1
        bmi.biBitCount = 32
        bmi.biCompression = 0  # BI_RGB

        buf_size = width * height * 4
        buf = ctypes.create_string_buffer(buf_size)
        DIB_RGB_COLORS = 0
        lines = gdi32.GetDIBits(
            hdc_mem, hbmp, 0, height, buf, ctypes.byref(bmi), DIB_RGB_COLORS
        )
        if lines != height:
            return None

        raw = buf.raw
        rows: list[list[tuple[int, int, int]]] = []
        stride = width * 4
        for y in range(height):
            row_offset = y * stride
            row: list[tuple[int, int, int]] = []
            for x in range(width):
                idx = row_offset + x * 4
                b = raw[idx]
                g = raw[idx + 1]
                r = raw[idx + 2]
                row.append((r, g, b))
            rows.append(row)
        return rows
    except Exception as exc:
        LOG.debug("GDI screen capture failed: %s", exc)
        return None
    finally:
        gdi32.SelectObject(hdc_mem, old_obj)
        gdi32.DeleteObject(hbmp)
        gdi32.DeleteDC(hdc_mem)
        user32.ReleaseDC(0, hdc_screen)


def _find_power_bi_window(title_keyword: str):
    matches = [
        w for w in Desktop(backend="uia").windows()
        if title_keyword.casefold() in text(w).casefold()
        and "power bi" in text(w).casefold()
    ]
    if not matches:
        matches = [
            w for w in Desktop(backend="uia").windows()
            if title_keyword.casefold() in text(w).casefold()
            and not text(w).lower().endswith(".py")
        ]
    if not matches:
        raise WorkflowError(f"Power BI window containing '{title_keyword}' is not open.")
    return matches[0]


def open_or_attach_pbix(pbix_path: Path, title_keyword: str, timeout_seconds: int = 180):
    opened_fresh = False
    try:
        window = _find_power_bi_window(title_keyword)
        LOG.info("Attached to open Power BI window: '%s'", text(window))
    except Exception:
        if not pbix_path.exists():
            raise WorkflowError(f"PBIX file not found: {pbix_path}")
        LOG.info("Launching Power BI Desktop for: %s", pbix_path)
        os.startfile(str(pbix_path))  # type: ignore[attr-defined]
        opened_fresh = True

        deadline = time.time() + timeout_seconds
        last_err = None
        while time.time() < deadline:
            time.sleep(3)
            try:
                window = _find_power_bi_window(title_keyword)
                break
            except Exception as err:
                last_err = err
        else:
            raise WorkflowError(
                f"Power BI Desktop did not load within {timeout_seconds}s"
            ) from last_err

    window.set_focus()
    try:
        window.maximize()
    except Exception:
        pass
    time.sleep(8.0 if opened_fresh else 1.5)
    return window


def locate_refresh_button(window, timeout_seconds: int = 15):
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        candidates = []
        for c in window.descendants():
            if not visible(c):
                continue
            c_type = str(c.element_info.control_type or "")
            if c_type not in ("Button", "SplitButton", "MenuItem"):
                continue

            auto_id = str(c.element_info.automation_id or "").casefold()
            name = str(c.element_info.name or "").casefold()
            t = text(c).casefold()

            if auto_id in ("refreshqueries", "refresh", "refreshbtn", "refreshbutton"):
                return c

            if t == "refresh" or name == "refresh":
                candidates.append(c)
            elif t.startswith("refresh") or name.startswith("refresh"):
                candidates.append(c)

        if candidates:
            return min(candidates, key=lambda x: x.rectangle().top)

        try:
            home_tabs = [
                c for c in window.descendants(control_type="TabItem")
                if visible(c) and text(c).casefold() == "home"
            ]
            if home_tabs:
                home_tabs[0].click_input()
                time.sleep(1)
        except Exception:
            pass

        time.sleep(1)

    return None


def trigger_data_only_refresh(window) -> None:
    """
    Opens the Refresh dropdown in the Power BI ribbon and selects 'Data'
    so only data rows refresh without the 20-minute schema evaluation.
    """
    LOG.info("Initiating 'Data' only refresh sequence...")
    window.set_focus()
    time.sleep(0.5)

    try:
        home_tabs = [
            c for c in window.descendants(control_type="TabItem")
            if visible(c) and text(c).casefold() == "home"
        ]
        if home_tabs:
            home_tabs[0].click_input()
            time.sleep(0.8)
    except Exception:
        pass

    refresh_button = locate_refresh_button(window, timeout_seconds=15)
    if not refresh_button:
        raise WorkflowError("Power BI 'Refresh' ribbon button was not found")

    r = refresh_button.rectangle()
    chevron_x = r.left + 22
    chevron_y = r.bottom + 15

    LOG.info("Clicking the Refresh downward arrow at (%d, %d)...", chevron_x, chevron_y)
    mouse.click(button="left", coords=(chevron_x, chevron_y))
    time.sleep(0.6)

    data_x = chevron_x
    data_y = chevron_y + 100

    LOG.info("Clicking 'Data' option directly below arrow at (%d, %d)...", data_x, data_y)
    mouse.click(button="left", coords=(data_x, data_y))
    time.sleep(1.0)


def find_and_click_refresh_close_button(window) -> bool:
    """
    Detects if the Power BI Refresh modal dialog has finished loading with
    query warnings/errors ('X of the loaded queries contained errors') and
    clicks its 'Close' button safely without ever touching the window title bar.
    """
    win_rect = window.rectangle()
    roots = [window] + [
        w for w in Desktop(backend="uia").windows()
        if w != window and visible(w)
    ]

    dialog_markers = []
    for root in roots:
        try:
            for c in root.descendants():
                if not visible(c):
                    continue
                t = text(c).casefold()
                n = str(c.element_info.name or "").casefold()
                if any(phrase in t or phrase in n for phrase in ("contained errors", "view errors", "rows loaded")):
                    dialog_markers.append(c)
        except Exception:
            continue

    if not dialog_markers:
        return False

    for marker in dialog_markers:
        parent = marker
        for _ in range(6):
            try:
                parent = parent.parent()
                if parent is None:
                    break
                close_btns = [
                    b for b in parent.descendants(control_type="Button")
                    if visible(b)
                    and (text(b).casefold() == "close" or str(b.element_info.name or "").casefold() == "close")
                    and b.rectangle().top > win_rect.top + 100
                    and b.rectangle().right < win_rect.right - 60
                ]
                if close_btns:
                    target_btn = max(close_btns, key=lambda b: b.rectangle().top)
                    LOG.info("Refresh completed with query warnings. Clicking 'Close' button at %s...", target_btn.rectangle())
                    target_btn.click_input()
                    return True
            except Exception:
                break

    candidate_closes = []
    for root in roots:
        try:
            for b in root.descendants(control_type="Button"):
                if not visible(b):
                    continue
                t = text(b).casefold()
                n = str(b.element_info.name or "").casefold()
                if t == "close" or n == "close":
                    r = b.rectangle()
                    if (
                        r.top > win_rect.top + 120
                        and r.bottom < win_rect.bottom - 40
                        and r.left > win_rect.left + 150
                        and r.right < win_rect.right - 80
                    ):
                        candidate_closes.append(b)
        except Exception:
            continue

    if candidate_closes:
        target_btn = max(candidate_closes, key=lambda b: b.rectangle().top)
        LOG.info("Refresh completed with query warnings. Clicking 'Close' button at %s...", target_btn.rectangle())
        target_btn.click_input()
        return True

    return False


def refresh_report(window, timeout_seconds: int = 1200, mode: str = "data_only") -> dict:
    started = time.time()
    if mode.casefold() in ("data_only", "data"):
        trigger_data_only_refresh(window)
    else:
        LOG.info("Locating Power BI ribbon 'Refresh' button for full refresh...")
        refresh_controls = [
            c for c in window.descendants()
            if visible(c) and "refresh" in str(c.element_info.automation_id or "").casefold()
        ]
        if not refresh_controls:
            raise WorkflowError("Power BI 'Refresh' button was not found in the ribbon")
        refresh_controls[0].click_input()

    LOG.info("Refresh initiated; polling query progress...")
    time.sleep(5)

    stable_since = None
    while time.time() - started < timeout_seconds:
        time.sleep(2)

        fatal_dialogs = [
            c for c in Desktop(backend="uia").windows()
            if visible(c) and any(err in text(c).casefold() for err in ("unexpected error", "cannot connect to the server"))
        ]
        if fatal_dialogs:
            raise WorkflowError(f"Power BI fatal error popup: {text(fatal_dialogs[0])}")

        if find_and_click_refresh_close_button(window):
            time.sleep(2)
            if find_and_click_refresh_close_button(window):
                time.sleep(1.5)
            elapsed = round(time.time() - started, 1)
            LOG.info("Data refresh complete and 'Close' button clicked in %s seconds", elapsed)
            return {"completed": True, "elapsedSeconds": elapsed, "mode": mode, "dismissed_query_errors": True}

        win_rect = window.rectangle()
        roots = [window] + [
            w for w in Desktop(backend="uia").windows()
            if w != window and visible(w)
        ]

        busy = False
        for root in roots:
            try:
                for c in root.descendants():
                    if not visible(c):
                        continue
                    t = text(c).casefold()
                    n = str(c.element_info.name or "").casefold()
                    c_type = str(c.element_info.control_type or "")

                    if any(
                        phrase in t or phrase in n
                        for phrase in (
                            "cancel refresh",
                            "refreshing",
                            "loading data",
                            "creating connection",
                            "waiting for other queries",
                            "applying changes",
                            "evaluating",
                            "syncing schema",
                            "rows loaded",
                            "contained errors",
                            "view errors",
                        )
                    ):
                        busy = True
                        break

                    if c_type == "Button" and (t == "cancel" or n == "cancel"):
                        r = c.rectangle()
                        if r.top > win_rect.top + 120 and r.right < win_rect.right - 80:
                            busy = True
                            break
            except Exception:
                continue
            if busy:
                break

        if not busy:
            stable_since = stable_since or time.time()
            if time.time() - stable_since >= 6:
                elapsed = round(time.time() - started, 1)
                LOG.info("Data refresh completed successfully in %s seconds", elapsed)
                return {"completed": True, "elapsedSeconds": elapsed, "mode": mode, "dismissed_query_errors": False}
        else:
            stable_since = None

    raise WorkflowError(f"Power BI refresh did not complete within {timeout_seconds} seconds")


def save_report(window, pbix_path: Path, timeout_seconds: int = 120) -> dict:
    LOG.info("Saving Power BI report: %s", pbix_path.name)
    before_mtime = pbix_path.stat().st_mtime if pbix_path.exists() else None

    window.set_focus()
    time.sleep(0.5)
    keyboard.send_keys("^s")

    deadline = time.time() + timeout_seconds
    after_mtime = before_mtime
    while time.time() < deadline:
        time.sleep(2)
        if pbix_path.exists():
            after_mtime = pbix_path.stat().st_mtime
            if before_mtime is None or after_mtime > before_mtime:
                LOG.info("Save verified. PBIX file timestamp updated on disk")
                return {"saved": True, "path": str(pbix_path)}

    raise WorkflowError("Save was triggered, but the PBIX file modified timestamp did not change on disk")


def _detect_blue_banner_bounds(window) -> tuple[int, int] | None:
    """
    Scans the upper report canvas of the Power BI window using 1:1 DPI-aware
    GDI capture to locate the dark blue 'Managers Information' header banner
    (RGB ~ [9, 71, 128]). Returns (banner_right_x, banner_bottom_y) in physical
    screen coordinates.
    """
    r = window.rectangle()
    win_w = r.right - r.left
    win_h = r.bottom - r.top
    if win_w <= 0 or win_h <= 0:
        return None

    scan_left = r.left + int(win_w * 0.03)
    scan_top = r.top + int(win_h * 0.14)
    scan_w = int(win_w * 0.94)
    scan_h = int(win_h * 0.28)

    rows = _grab_screen_region_rgb(scan_left, scan_top, scan_w, scan_h)
    if not rows:
        return None

    xs: list[int] = []
    ys: list[int] = []
    for dy in range(0, len(rows), 2):
        row = rows[dy]
        for dx in range(0, len(row), 2):
            red, green, blue = row[dx]
            if red < 35 and 45 <= green <= 105 and blue >= 100:
                xs.append(scan_left + dx)
                ys.append(scan_top + dy)

    if len(xs) < 100:
        return None

    xs.sort()
    ys.sort()
    banner_right = xs[int(len(xs) * 0.995)]
    banner_bottom = ys[int(len(ys) * 0.995)]
    return banner_right, banner_bottom


def _find_three_dots_in_patch(
    patch_rgb: list[list[tuple[int, int, int]]],
    patch_left: int,
    patch_top: int,
    min_y: int,
    scale: float = 1.0,
) -> tuple[int, int] | None:
    """
    Finds the exact center (x, y) of the horizontal '...' ('More options') icon
    inside a small RGB patch around the top-right of the table visual.
    """
    h = len(patch_rgb)
    w = len(patch_rgb[0]) if h else 0
    best_match: tuple[int, int] | None = None
    pad = max(4, int(round(4 * scale)))
    max_dot_w = max(4, int(round(5 * scale)))
    min_gap = 2.5 * scale
    max_gap = 9.5 * scale
    max_gap_diff = 2.5 * scale

    for y in range(pad, h - pad):
        screen_y = patch_top + y
        if screen_y < min_y:
            continue

        dark_xs = [
            x for x in range(w)
            if max(patch_rgb[y][x]) < 185
        ]
        if len(dark_xs) < 3:
            continue

        runs: list[tuple[int, int]] = []
        start = prev = dark_xs[0]
        for x in dark_xs[1:]:
            if x == prev + 1:
                prev = x
            else:
                runs.append((start, prev))
                start = prev = x
        runs.append((start, prev))

        dot_runs = [run for run in runs if 1 <= (run[1] - run[0] + 1) <= max_dot_w]
        for i in range(len(dot_runs) - 2):
            d1, d2, d3 = dot_runs[i], dot_runs[i + 1], dot_runs[i + 2]
            c1 = (d1[0] + d1[1]) / 2.0
            c2 = (d2[0] + d2[1]) / 2.0
            c3 = (d3[0] + d3[1]) / 2.0
            gap1 = c2 - c1
            gap2 = c3 - c2
            if min_gap <= gap1 <= max_gap and min_gap <= gap2 <= max_gap and abs(gap1 - gap2) <= max_gap_diff:
                cx = int(round(c2))
                above = patch_rgb[y - pad][cx]
                below = patch_rgb[y + pad][cx]
                if min(above) > 220 and min(below) > 220:
                    candidate = (patch_left + cx, screen_y)
                    if best_match is None or candidate[0] >= best_match[0]:
                        best_match = candidate

    return best_match


def _wait_for_canvas_and_find_more_options(
    window, timeout_seconds: int = 45
) -> tuple[int, int]:
    """
    Waits for the report canvas to render, clicks inside the table visual to
    reveal its top-right action header, and locates the exact '...' icon center.
    """
    window.set_focus()
    scale = _get_dpi_scale(window)
    banner_bounds = None
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        banner_bounds = _detect_blue_banner_bounds(window)
        if banner_bounds is not None:
            break
        time.sleep(1.5)

    r = window.rectangle()
    win_w = r.right - r.left
    win_h = r.bottom - r.top

    table_focus_x = r.left + int(win_w * 0.40)
    table_focus_y = r.top + int(win_h * 0.52)
    LOG.info("Focusing table visual at (%d, %d) [DPI scale=%.2fx]...", table_focus_x, table_focus_y, scale)
    mouse.move(coords=(table_focus_x, table_focus_y))
    time.sleep(0.4)
    mouse.click(button="left", coords=(table_focus_x, table_focus_y))
    time.sleep(1.0)

    if banner_bounds is not None:
        banner_right, banner_bottom = banner_bounds
        approx_x = banner_right - int(win_w * 0.0345)
        approx_y = banner_bottom + int(win_h * 0.0453)
        LOG.info(
            "Detected blue report banner (right=%d, bottom=%d); scanning for '...' around (%d, %d)...",
            banner_right, banner_bottom, approx_x, approx_y
        )

        half_w = int(90 * scale)
        half_h = int(45 * scale)
        patch_left = max(r.left, approx_x - half_w)
        patch_top = max(r.top, approx_y - half_h)
        patch_rgb = _grab_screen_region_rgb(patch_left, patch_top, half_w * 2, half_h * 2)
        if patch_rgb:
            exact_dots = _find_three_dots_in_patch(
                patch_rgb,
                patch_left,
                patch_top,
                min_y=banner_bottom + int(12 * scale),
                scale=scale,
            )
            if exact_dots is not None:
                LOG.info("Locked onto exact '...' center pixel at %s", exact_dots)
                return exact_dots

        return approx_x, approx_y

    filters_left = r.left + int(win_w * 0.929)
    for c in window.descendants(control_type="Text"):
        if visible(c) and text(c).casefold() == "filters":
            cr = c.rectangle()
            if cr.left > r.left + int(win_w * 0.65):
                filters_left = cr.left
                break

    more_x = filters_left - int(win_w * 0.037)
    more_y = r.top + int(win_h * 0.385)
    LOG.info("Using Filters-pane anchored '...' target=(%d, %d)", more_x, more_y)
    return more_x, more_y


def _click_more_options_and_export(window) -> None:
    more_x, more_y = _wait_for_canvas_and_find_more_options(window, RENDER_TIMEOUT_SECONDS)

    LOG.info("Hovering and clicking '...' ('More options') at (%d, %d)...", more_x, more_y)
    mouse.move(coords=(more_x, more_y))
    time.sleep(0.5)
    mouse.click(button="left", coords=(more_x, more_y))
    time.sleep(1.0)

    clicked = False
    deadline = time.time() + 3.0
    while time.time() < deadline:
        matches = []
        for root in list(Desktop(backend="uia").windows()) + [window]:
            try:
                for c in root.descendants(control_type="MenuItem"):
                    if visible(c) and text(c).casefold() == "export data":
                        matches.append(c)
            except Exception:
                continue

        if matches:
            LOG.info("Clicking 'Export data' via UIA at %s...", matches[0].rectangle())
            matches[0].click_input()
            clicked = True
            break
        time.sleep(0.25)

    if not clicked:
        r = window.rectangle()
        win_w = r.right - r.left
        win_h = r.bottom - r.top
        export_x = more_x + max(45, int(win_w * 0.042))
        export_y = more_y + max(25, int(win_h * 0.036))
        LOG.info("Clicking 'Export data' by dropdown offset at (%d, %d)...", export_x, export_y)
        mouse.click(button="left", coords=(export_x, export_y))

    time.sleep(1.0)


def _find_save_as_hwnd() -> int | None:
    """Finds the visible Win32 '#32770' Save As dialog HWND directly."""
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    GetWindowTextW = user32.GetWindowTextW
    GetWindowTextLengthW = user32.GetWindowTextLengthW
    GetClassNameW = user32.GetClassNameW
    IsWindowVisible = user32.IsWindowVisible

    found: list[int] = []

    @EnumWindowsProc
    def callback(hwnd, lparam):
        if not IsWindowVisible(hwnd):
            return True
        cls_buf = ctypes.create_unicode_buffer(64)
        GetClassNameW(hwnd, cls_buf, 64)
        if cls_buf.value == "#32770":
            length = GetWindowTextLengthW(hwnd)
            if length > 0:
                txt_buf = ctypes.create_unicode_buffer(length + 1)
                GetWindowTextW(hwnd, txt_buf, length + 1)
                if txt_buf.value.strip().casefold() in ("save as", "export as", "export data"):
                    found.append(int(hwnd))
                    return False
        return True

    EnumWindows(callback, 0)
    return found[0] if found else None


def _get_hwnd_text(hwnd: int) -> str:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value.strip()


def _safe_replace_file(source: Path, destination: Path, max_attempts: int = 5) -> None:
    """Moves/copies `source` to `destination`, always overwriting any existing file."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == destination.resolve():
        return

    last_err = None
    for _ in range(max_attempts):
        try:
            if destination.exists():
                try:
                    destination.unlink()
                except OSError:
                    pass
            shutil.copy2(str(source), str(destination))
            if destination.exists() and destination.stat().st_size > 0:
                try:
                    source.unlink()
                except OSError:
                    pass
                return
        except Exception as err:
            last_err = err
            time.sleep(1.0)

    raise WorkflowError(f"Failed to overwrite '{destination}' from '{source}': {last_err}")


def _wait_for_save_as_ready(save_dlg, timeout_seconds: float = 5.0) -> None:
    """
    Waits briefly if the Save As address bar ProgressBar ('Loading') is still
    binding a network/DFS folder so IFileSaveDialog has a resolved IShellItem.
    """
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        loading = False
        try:
            for p in save_dlg.descendants(control_type="ProgressBar"):
                if visible(p) and "loading" in str(p.element_info.name or "").casefold():
                    loading = True
                    break
        except Exception:
            pass
        if not loading:
            break
        time.sleep(0.3)
    time.sleep(0.6)


def _ensure_save_as_folder(save_dlg, target_folder: str) -> None:
    """
    Checks the UIA 'Address: ...' toolbar (id='1001'). If the Save As dialog is
    already in `target_folder`, skips address-bar navigation completely; otherwise
    clicks the right side of the Address toolbar to enter `target_folder`.
    """
    addr_bars = [
        c for c in save_dlg.descendants(control_type="ToolBar")
        if visible(c) and str(c.element_info.automation_id or "") == "1001"
        and str(c.element_info.name or "").casefold().startswith("address:")
    ]
    if addr_bars:
        current_addr = str(addr_bars[0].element_info.name or "")
        if target_folder.casefold() in current_addr.casefold():
            LOG.info("Save As already open in target folder (%s).", current_addr)
            return

        ar = addr_bars[0].rectangle()
        click_x = ar.right - 20
        click_y = (ar.top + ar.bottom) // 2
        LOG.info("Navigating Save As address bar to: %s", target_folder)
        _set_clipboard_text(target_folder)
        mouse.click(button="left", coords=(click_x, click_y))
        time.sleep(0.3)
        keyboard.send_keys("^v{ENTER}")
        time.sleep(1.5)


def _set_filename_box_uia(save_dlg, file_stem: str) -> None:
    """
    Locates the exact UIA Edit control id='1001' ('File name:') inside Save As,
    clicks it, and pastes `file_stem` via real keyboard input (Ctrl+A, Backspace,
    Ctrl+V) so Windows IFileSaveDialog receives genuine edit-change events
    and commits the new filename instead of keeping 'Manager Contact Information'.
    """
    fname_edits = [
        c for c in save_dlg.descendants(control_type="Edit")
        if visible(c)
        and str(c.element_info.automation_id or "") == "1001"
        and "file name" in str(c.element_info.name or "").casefold()
    ]
    if not fname_edits:
        raise WorkflowError("Could not locate UIA 'File name:' Edit control (id='1001') in Save As.")

    edit_ctrl = fname_edits[0]
    er = edit_ctrl.rectangle()
    click_x = er.left + 80
    click_y = (er.top + er.bottom) // 2

    LOG.info("Focusing 'File name:' Edit (id='1001') at (%d, %d) and pasting: %s", click_x, click_y, file_stem)
    _set_clipboard_text(file_stem)
    mouse.click(button="left", coords=(click_x, click_y))
    time.sleep(0.25)
    keyboard.send_keys("^a{BACKSPACE}")
    time.sleep(0.15)
    keyboard.send_keys("^v")
    time.sleep(0.35)

    try:
        current_val = str(edit_ctrl.get_value() or "").strip()
    except Exception:
        current_val = text(edit_ctrl)

    if file_stem.casefold() not in current_val.casefold():
        _set_clipboard_text(file_stem)
        mouse.click(button="left", coords=(click_x, click_y))
        time.sleep(0.2)
        keyboard.send_keys("^a{BACKSPACE}^v")
        time.sleep(0.35)


def _dismiss_error_popup_if_present() -> bool:
    """Checks for a 'File not found' Save As error popup and dismisses it."""
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    EnumWindows = user32.EnumWindows
    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    found_err: list[int] = []

    @EnumWindowsProc
    def cb(hwnd, lparam):
        if user32.IsWindowVisible(hwnd) and _get_hwnd_text(hwnd).casefold() == "save as":
            if not user32.GetDlgItem(hwnd, 40965):
                found_err.append(int(hwnd))
                return False
        return True

    EnumWindows(cb, 0)
    if found_err:
        user32.SetForegroundWindow(found_err[0])
        time.sleep(0.2)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.4)
        return True
    return False


def _locate_exported_file_on_disk(target_csv: Path, started_ts: float) -> Path | None:
    """
    Forces a fresh SMB directory listing of `target_csv.parent` (bypassing
    Windows negative file-attribute cache on network drives). Checks for
    `target_csv` first, and also checks if Power BI saved the default
    'Manager Contact Information.csv' after `started_ts`—renaming/overwriting
    it cleanly to `target_csv`.
    """
    try:
        if target_csv.exists() and target_csv.stat().st_size > 0:
            return target_csv
    except OSError:
        pass

    try:
        entries = os.listdir(str(target_csv.parent))
        stem_cf = target_csv.stem.casefold()
        default_cf = DEFAULT_VISUAL_EXPORT_NAME.casefold()

        for entry in entries:
            if entry.casefold() == target_csv.name.casefold() or entry.casefold() == stem_cf:
                candidate = target_csv.parent / entry
                if candidate.is_file() and candidate.stat().st_size > 0:
                    if candidate.suffix.lower() != ".csv":
                        renamed = candidate.with_suffix(".csv")
                        _safe_replace_file(candidate, renamed)
                        return renamed
                    return candidate

        for entry in entries:
            if entry.casefold() == default_cf:
                candidate = target_csv.parent / entry
                if candidate.is_file() and candidate.stat().st_size > 0:
                    if candidate.stat().st_mtime >= (started_ts - 2.0):
                        LOG.info(
                            "Found fresh export '%s'; renaming/overwriting to '%s'...",
                            candidate.name, target_csv.name
                        )
                        _safe_replace_file(candidate, target_csv)
                        return target_csv
    except OSError:
        pass

    return None


def _handle_save_as_dialog(window, target_csv: Path, timeout_seconds: int = 90) -> Path:
    """
    Waits for the Windows 'Save As' dialog, ensures the target folder is loaded,
    pastes `target_csv.stem` into the UIA 'File name:' Edit box (id='1001'),
    clicks 'Save' (id='1'), handles overwrite prompts, and verifies/renames the
    exported CSV on disk.
    """
    export_started_ts = time.time()
    folder_path = str(target_csv.parent)
    file_stem = target_csv.stem  # e.g. 'managers-contact-info-2026-11-12'
    downloads_dir = Path.home() / "Downloads"
    downloads_csv = downloads_dir / target_csv.name
    default_target_csv = target_csv.parent / DEFAULT_VISUAL_EXPORT_NAME
    default_downloads_csv = downloads_dir / DEFAULT_VISUAL_EXPORT_NAME

    target_csv.parent.mkdir(parents=True, exist_ok=True)
    for existing in (target_csv, downloads_csv, default_target_csv, default_downloads_csv):
        if existing.exists():
            try:
                existing.unlink()
            except OSError:
                pass

    deadline = time.time() + timeout_seconds
    dialog_hwnd = None
    while time.time() < deadline:
        dialog_hwnd = _find_save_as_hwnd()
        if dialog_hwnd is not None:
            break
        time.sleep(0.4)

    if not dialog_hwnd:
        raise WorkflowError("Windows 'Save As' dialog did not appear after clicking Export data.")

    save_dlg = Desktop(backend="uia").window(handle=dialog_hwnd)
    save_dlg.set_focus()
    _wait_for_save_as_ready(save_dlg, timeout_seconds=5.0)

    # Step 1: Ensure folder is W:\Dashboards VCI\LJSO\Managers (skips if already there)
    _ensure_save_as_folder(save_dlg, folder_path)
    _wait_for_save_as_ready(save_dlg, timeout_seconds=4.0)

    # Step 2: Paste ONLY the filename stem into 'File name:' (id='1001') via keyboard
    _set_filename_box_uia(save_dlg, file_stem)

    # Step 3: Click 'Save' (id='1')
    save_btns = [
        b for b in save_dlg.descendants(control_type="Button")
        if visible(b) and str(b.element_info.automation_id or "") == "1"
    ]
    if save_btns:
        save_btns[0].click_input()
    else:
        keyboard.send_keys("%s")

    time.sleep(1.0)

    # Fallback: If the DFS share popped up 'File not found', save to Downloads and move/overwrite
    if _dismiss_error_popup_if_present():
        LOG.warning("DFS share rejected direct Save As; saving to Downloads and overwriting into %s...", target_csv)
        _ensure_save_as_folder(save_dlg, str(downloads_dir))
        _wait_for_save_as_ready(save_dlg, timeout_seconds=3.0)
        _set_filename_box_uia(save_dlg, file_stem)
        if save_btns:
            save_btns[0].click_input()
        else:
            keyboard.send_keys("%s")
        time.sleep(1.0)

    # Step 4: Handle 'Confirm Save As' overwrite prompt if it appears
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    confirm_hwnd = None
    for _ in range(4):
        EnumWindows = user32.EnumWindows
        EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        found_confirm: list[int] = []

        @EnumWindowsProc
        def confirm_cb(hwnd, lparam):
            if user32.IsWindowVisible(hwnd) and "confirm save as" in _get_hwnd_text(hwnd).casefold():
                found_confirm.append(int(hwnd))
                return False
            return True

        EnumWindows(confirm_cb, 0)
        if found_confirm:
            confirm_hwnd = found_confirm[0]
            break
        time.sleep(0.3)

    if confirm_hwnd:
        LOG.info("Confirming overwrite prompt ('Yes')...")
        user32.SetForegroundWindow(confirm_hwnd)
        time.sleep(0.2)
        keyboard.send_keys("%y")

    # Step 5: Poll disk with fresh SMB directory listings until Power BI finishes writing the CSV
    file_deadline = time.time() + 60
    while time.time() < file_deadline:
        found_path = _locate_exported_file_on_disk(target_csv, export_started_ts)
        if found_path is not None:
            LOG.info("Export verified: %s (%d bytes)", found_path, found_path.stat().st_size)
            return found_path

        for dl_candidate in (downloads_csv, default_downloads_csv):
            if dl_candidate.exists() and dl_candidate.stat().st_size > 0:
                _safe_replace_file(dl_candidate, target_csv)
                LOG.info("Export verified (moved & overwritten): %s (%d bytes)", target_csv, target_csv.stat().st_size)
                return target_csv

        time.sleep(0.5)

    raise WorkflowError(f"Exported file was not created on disk: {target_csv}")


def export_managers_contact_csv(
    pbix_path: Path = PBIX_PATH,
    export_dir: Path = EXPORT_DIR,
    export_date: date | None = None,
    skip_refresh: bool = SKIP_DATA_REFRESH,
) -> Path:
    """Main entry point callable by email scripts or standalone execution."""
    run_date = export_date or date.today()
    date_str = run_date.strftime("%Y-%m-%d")
    target_csv = export_dir / f"{EXPORT_PREFIX}-{date_str}.csv"

    window = open_or_attach_pbix(pbix_path, WINDOW_TITLE_KEYWORD, OPEN_TIMEOUT_SECONDS)

    if skip_refresh:
        LOG.info("SKIP_DATA_REFRESH is True -> skipping data refresh and PBIX save.")
    else:
        refresh_report(window, timeout_seconds=REFRESH_TIMEOUT_SECONDS, mode="data_only")
        save_report(window, pbix_path, timeout_seconds=SAVE_PBIX_TIMEOUT_SECONDS)

    _click_more_options_and_export(window)
    return _handle_save_as_dialog(window, target_csv, SAVE_CSV_TIMEOUT_SECONDS)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    exported_file = export_managers_contact_csv()
    print(f"Successfully exported: {exported_file}")