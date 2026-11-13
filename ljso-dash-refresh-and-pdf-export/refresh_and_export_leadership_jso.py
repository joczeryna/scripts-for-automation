"""
Leadership JSO Dashboard: Automated Refresh, Save & Dual-Path PDF Export
Project Directory: C:/Users/jd0855/Desktop/VCI_Automation/ljso-dash-refresh-and-pdf-export
File: refresh_and_export_leadership_jso.py

Workflow:
1. Connect to or launch 'Leadership JSO.pbix' located in 'W:/Dashboards VCI/LJSO/'.
2. Execute data refresh ('Data' only) and poll until all background queries complete.
3. If the Refresh summary modal remains open due to query warnings/errors, click 'Close'.
4. Save the updated PBIX file and verify the disk timestamp updates.
5. Export the dashboard to PDF via File -> Export -> Export to PDF.
6. Locate the exported PDF on disk (including Windows Temp / Acrobat viewer target) and close the viewer.
7. Distribute PDF outputs:
   - Primary: 'W:/Dashboards VCI/LJSO/Main/Leadership JSO.pdf' (always overwriting)
   - Archive: 'W:/Dashboards VCI/LJSO/Backup-Archive/Leadership JSO - YYYY-MM-DD.pdf'
"""

from __future__ import annotations

import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

from pywinauto import Desktop, keyboard, mouse


DEFAULT_CONFIG = {
    "pbix_file": r"W:\Dashboards VCI\LJSO\Leadership JSO.pbix",
    "window_title_keyword": "Leadership JSO",
    "refresh_mode": "data_only",
    "open_timeout_seconds": 180,
    "refresh_timeout_seconds": 1200,
    "save_timeout_seconds": 120,
    "pdf_export_timeout_seconds": 300,
    "close_power_bi_when_finished": False,
    "destinations": {
        "main_folder": r"W:\Dashboards VCI\LJSO\Main",
        "main_filename": "Leadership JSO.pdf",
        "backup_folder": r"W:\Dashboards VCI\LJSO\Backup-Archive",
        "backup_filename_prefix": "Leadership JSO",
        "date_format": "%Y-%m-%d"
    }
}

LOG = logging.getLogger("ljso_refresh_export")


class WorkflowError(RuntimeError):
    pass


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


def exact(root, title: str, control_type: str):
    matches = [
        c for c in root.descendants(control_type=control_type)
        if visible(c) and (
            text(c).casefold() == title.casefold()
            or str(c.element_info.name or "").strip().casefold() == title.casefold()
        )
    ]
    if not matches:
        raise WorkflowError(f"Visible {control_type} not found: '{title}'")
    return matches[0]


def find_power_bi_window(title_keyword: str):
    matches = [
        w for w in Desktop(backend="uia").windows()
        if title_keyword.casefold() in text(w).casefold()
        and ".pdf" not in text(w).casefold()
    ]
    if not matches:
        raise WorkflowError(f"Power BI window containing '{title_keyword}' is not open")
    return matches[0]


def open_and_wait_for_power_bi(pbix_path: Path, title_keyword: str, timeout_seconds: int = 180) -> tuple[object, bool]:
    opened_by_runner = False
    try:
        window = find_power_bi_window(title_keyword)
        LOG.info("Attached to already open Power BI Desktop window: '%s'", text(window))
    except Exception:
        if not pbix_path.exists():
            raise WorkflowError(f"Target PBIX file does not exist: {pbix_path}")
        if os.name != "nt":
            raise WorkflowError("Power BI Desktop automation requires a Windows environment")

        LOG.info("Launching Power BI Desktop for: %s", pbix_path)
        os.startfile(str(pbix_path))  # type: ignore[attr-defined]
        opened_by_runner = True

        deadline = time.time() + timeout_seconds
        last_error: Exception | None = None
        while time.time() < deadline:
            time.sleep(3)
            try:
                window = find_power_bi_window(title_keyword)
                break
            except Exception as error:
                last_error = error
        else:
            raise WorkflowError(
                f"Power BI Desktop did not become ready within {timeout_seconds} seconds"
            ) from last_error

    window.set_focus()
    time.sleep(1)

    try:
        if hasattr(window, "maximize"):
            window.maximize()
            time.sleep(1)
    except Exception as ex:
        LOG.warning("Could not auto-maximize window (%s); proceeding with current geometry", ex)

    try:
        exact(window, "Report view", "TabItem").click_input()
        time.sleep(2)
    except Exception:
        LOG.warning("Could not select 'Report view' tab item; proceeding in current workspace view")

    return window, opened_by_runner


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

    # 1. Ensure Home tab is active
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

    # 2. Locate Refresh ribbon control
    refresh_button = locate_refresh_button(window, timeout_seconds=15)
    if not refresh_button:
        raise WorkflowError("Power BI 'Refresh' ribbon button was not found")

    r = refresh_button.rectangle()

    # Step A: Click the downward arrow (chevron)
    chevron_x = r.left + 22
    chevron_y = r.bottom + 15

    LOG.info("Clicking the Refresh downward arrow at (%d, %d)...", chevron_x, chevron_y)
    mouse.click(button="left", coords=(chevron_x, chevron_y))
    time.sleep(0.6)

    # Step B: Click directly on 'Data' in the dropdown menu
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

    # Strategy 1: Walk up from the error/status text to the enclosing dialog container
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

    # Strategy 2: Canvas-bounded search for 'Close' button (strictly excluding title-bar X)
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

        # 1. Check for fatal crash popups
        fatal_dialogs = [
            c for c in Desktop(backend="uia").windows()
            if visible(c) and any(err in text(c).casefold() for err in ("unexpected error", "cannot connect to the server"))
        ]
        if fatal_dialogs:
            raise WorkflowError(f"Power BI fatal error popup: {text(fatal_dialogs[0])}")

        # 2. Check if Refresh modal finished with query warnings and has a 'Close' button
        if find_and_click_refresh_close_button(window):
            time.sleep(2)
            if find_and_click_refresh_close_button(window):
                time.sleep(1.5)
            elapsed = round(time.time() - started, 1)
            LOG.info("Data refresh complete and 'Close' button clicked in %s seconds", elapsed)
            return {"completed": True, "elapsedSeconds": elapsed, "mode": mode, "dismissed_query_errors": True}

        # 3. Check if Refresh modal or background queries are still running
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

        # 4. Clean completion (0 errors, modal auto-dismissed)
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


def trigger_pdf_export(window) -> None:
    """
    Exports the dashboard to PDF using the backstage path:
    File -> Export -> Export to PDF
    """
    LOG.info("Triggering Power BI Export to PDF (File -> Export -> Export to PDF)...")
    window.set_focus()
    time.sleep(1.0)
    win_rect = window.rectangle()

    client_left = win_rect.left + (8 if win_rect.left < 0 else 0)
    client_top = win_rect.top + (8 if win_rect.top < 0 else 0)

    save_rect = None
    home_tab = None
    try:
        save_btns = [
            c for c in window.descendants(control_type="Button")
            if visible(c) and str(c.element_info.automation_id or "") == "save"
        ]
        if save_btns:
            save_rect = save_btns[0].rectangle()
    except Exception:
        pass

    try:
        for c in window.descendants(control_type="TabItem"):
            if visible(c) and (text(c).casefold() == "home" or str(c.element_info.name or "").casefold() == "home"):
                r = c.rectangle()
                if (r.top - client_top) < 200 and (r.left - client_left) < 300:
                    home_tab = c
                    break
    except Exception:
        pass

    home_rect = home_tab.rectangle() if home_tab else None

    if save_rect and home_rect and (home_rect.top > save_rect.top):
        dy = (home_rect.top + home_rect.height() // 2) - (save_rect.top + save_rect.height() // 2)
        dpi_scale = max(1.0, dy / 33.0)
    else:
        dpi_scale = 1.0
    LOG.info("Detected UI scale factor: %.2fx (home_rect=%s)", dpi_scale, home_rect)

    # Step 1: Click 'File' ONCE (strictly to the left of 'Home' tab)
    if home_rect:
        file_x = max(client_left + int(20 * dpi_scale), home_rect.left - int(55 * dpi_scale))
        file_y = home_rect.top + (home_rect.height() // 2)
    elif save_rect:
        file_x = save_rect.left + int(15 * dpi_scale)
        file_y = save_rect.top + (save_rect.height() // 2) + int(33 * dpi_scale)
    else:
        file_x = client_left + int(37 * dpi_scale)
        file_y = client_top + int(52 * dpi_scale)

    tight_file_candidates = []
    for c in window.descendants():
        if not visible(c):
            continue
        t = text(c).casefold()
        n = str(c.element_info.name or "").casefold()
        auto_id = str(c.element_info.automation_id or "").casefold()
        if auto_id == "filemenubutton" or t == "file" or n == "file":
            r = c.rectangle()
            max_right = home_rect.left if home_rect else (client_left + int(75 * dpi_scale))
            if (r.top - client_top) < int(120 * dpi_scale) and r.right <= max_right and r.width() < int(75 * dpi_scale):
                tight_file_candidates.append(c)

    if tight_file_candidates:
        r = tight_file_candidates[0].rectangle()
        file_x = r.left + (r.width() // 2)
        file_y = r.top + (r.height() // 2)
        LOG.info("Clicking 'File' via tight UIA box at (%d, %d)...", file_x, file_y)
    else:
        LOG.info("Clicking 'File' to the left of 'Home' at (%d, %d)...", file_x, file_y)

    mouse.click(button="left", coords=(file_x, file_y))
    time.sleep(2.0)

    # Step 2: Click 'Export' on the File backstage left sidebar
    sidebar_max_x = client_left + int(215 * dpi_scale)
    export_items = []
    save_as_rect = None

    for c in window.descendants():
        if not visible(c):
            continue
        t = text(c).casefold()
        n = str(c.element_info.name or "").casefold()
        r = c.rectangle()
        if r.left < sidebar_max_x and r.width() < int(230 * dpi_scale):
            if t == "export" or n == "export":
                export_items.append(c)
            elif t == "save as" or n == "save as":
                save_as_rect = r

    if export_items:
        export_items.sort(
            key=lambda c: 0 if str(c.element_info.control_type or "") in ("Button", "MenuItem", "TabItem", "ListItem") else 1
        )
        r = export_items[0].rectangle()
        export_x = r.left + (r.width() // 2)
        export_y = r.top + (r.height() // 2)
        LOG.info("Clicking 'Export' in backstage sidebar via UIA at (%d, %d)...", export_x, export_y)
    elif save_as_rect:
        export_x = save_as_rect.left + (save_as_rect.width() // 2)
        export_y = save_as_rect.top + (save_as_rect.height() // 2) + int(176 * dpi_scale)
        LOG.info("Clicking 'Export' below 'Save as' at (%d, %d)...", export_x, export_y)
    else:
        export_x = client_left + int(90 * dpi_scale)
        export_y = client_top + int(412 * dpi_scale)
        LOG.info("Clicking 'Export' in sidebar at (%d, %d)...", export_x, export_y)

    mouse.click(button="left", coords=(export_x, export_y))
    time.sleep(2.0)

    # Step 3: Click 'Export to PDF' on the right pane
    pdf_items = []
    for c in window.descendants():
        if not visible(c):
            continue
        t = text(c).casefold()
        n = str(c.element_info.name or "").casefold()
        r = c.rectangle()
        if r.left >= sidebar_max_x and r.width() < int(400 * dpi_scale) and ("export to pdf" in t or "export to pdf" in n):
            pdf_items.append(c)

    if pdf_items:
        pdf_items.sort(
            key=lambda c: 0 if str(c.element_info.control_type or "") in ("Button", "MenuItem", "TabItem", "ListItem") else 1
        )
        r = pdf_items[0].rectangle()
        pdf_x = r.left + (r.width() // 2)
        pdf_y = r.top + (r.height() // 2)
        LOG.info("Clicking 'Export to PDF' via UIA at (%d, %d)...", pdf_x, pdf_y)
    elif save_as_rect:
        pdf_x = save_as_rect.left + int(230 * dpi_scale)
        pdf_y = save_as_rect.top + (save_as_rect.height() // 2)
        LOG.info("Clicking 'Export to PDF' to the right of 'Save as' at (%d, %d)...", pdf_x, pdf_y)
    else:
        pdf_x = client_left + int(290 * dpi_scale)
        pdf_y = client_top + int(228 * dpi_scale)
        LOG.info("Clicking 'Export to PDF' at (%d, %d)...", pdf_x, pdf_y)

    mouse.click(button="left", coords=(pdf_x, pdf_y))
    time.sleep(2.0)


def get_pdf_search_roots(pbix_path: Path) -> list[Path]:
    roots: list[Path] = [
        pbix_path.parent,
        Path.home() / "Documents",
        Path.home() / "Downloads",
        Path.home() / "Desktop",
        Path(tempfile.gettempdir()),
    ]
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        local_temp = Path(local_app_data) / "Temp"
        if local_temp not in roots:
            roots.append(local_temp)
        pbi_dir = Path(local_app_data) / "Microsoft" / "Power BI Desktop"
        if pbi_dir.exists():
            roots.append(pbi_dir)
    return [r for r in roots if r.exists()]


def find_pdf_from_viewer_cmdline(stem: str) -> Path | None:
    """
    When Power BI exports a PDF, it writes the file to a temp path and opens it in
    Adobe Acrobat / default PDF viewer. Inspects viewer command lines to grab that
    exact file path directly regardless of which monitor the viewer opened on.
    """
    if os.name != "nt":
        return None
    try:
        ps_cmd = (
            f"Get-CimInstance Win32_Process | "
            f"Where-Object {{ $_.CommandLine -like '*{stem}*.pdf*' }} | "
            f"Select-Object -ExpandProperty CommandLine"
        )
        output = subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            stderr=subprocess.DEVNULL,
            text=True,
            timeout=6,
        )
        for line in output.splitlines():
            matches = re.findall(r'([A-Za-z]:\\[^"*?<>|\r\n]+?\.pdf)', line, flags=re.IGNORECASE)
            for m in matches:
                p = Path(m.strip().strip('"'))
                if p.exists() and stem.casefold() in p.name.casefold():
                    return p
    except Exception:
        pass
    return None


def scan_recent_pdfs(roots: list[Path], stem: str, min_mtime: float, max_depth: int = 2) -> list[Path]:
    found: list[Path] = []
    stem_cf = stem.casefold()

    def _walk(folder: Path, depth: int) -> None:
        if depth > max_depth:
            return
        try:
            with os.scandir(folder) as it:
                for entry in it:
                    try:
                        if entry.is_file(follow_symlinks=False):
                            name_cf = entry.name.casefold()
                            if name_cf.endswith(".pdf") and stem_cf in name_cf:
                                st = entry.stat(follow_symlinks=False)
                                if st.st_mtime >= min_mtime and st.st_size > 0:
                                    found.append(Path(entry.path))
                        elif entry.is_dir(follow_symlinks=False) and depth < max_depth:
                            _walk(Path(entry.path), depth + 1)
                    except OSError:
                        continue
        except OSError:
            pass

    for r in roots:
        _walk(r, 0)

    found.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return found


def wait_for_pdf_and_stabilize(pbix_path: Path, timeout_seconds: int = 300) -> Path:
    search_roots = get_pdf_search_roots(pbix_path)
    default_save_target = pbix_path.parent / f"{pbix_path.stem}.pdf"
    start_time = time.time()
    deadline = start_time + timeout_seconds
    LOG.info("Waiting for PDF rendering to finalize on disk...")

    time.sleep(3)
    save_dialogs = [
        c for c in Desktop(backend="uia").windows()
        if visible(c) and text(c).casefold() in ("save as", "export as")
    ]
    if save_dialogs:
        dlg = save_dialogs[0]
        dlg.set_focus()
        fname_box = exact(dlg, "File name:", "Edit")
        fname_box.set_edit_text(str(default_save_target))
        time.sleep(0.5)
        exact(dlg, "Save", "Button").click_input()
        LOG.info("Handled explicit Save As dialog; targeting: %s", default_save_target)

    last_size = -1
    stable_count = 0
    detected_path: Path | None = None

    while time.time() < deadline:
        time.sleep(2)

        progress_dialogs = [
            w for w in Desktop(backend="uia").windows()
            if "export to pdf" in text(w).casefold()
        ]
        modal_active = any(visible(d) for d in progress_dialogs)

        # Check disk locations (including Windows Temp where Power BI writes before launching Acrobat)
        candidates = scan_recent_pdfs(search_roots, pbix_path.stem, start_time - 10, max_depth=2)
        if not candidates:
            viewer_pdf = find_pdf_from_viewer_cmdline(pbix_path.stem)
            if viewer_pdf and viewer_pdf.exists() and viewer_pdf.stat().st_mtime >= start_time - 15:
                candidates = [viewer_pdf]

        if candidates:
            candidate = candidates[0]
            if detected_path != candidate:
                detected_path = candidate
                LOG.info("Detected exported PDF file: %s", candidate)

            current_size = candidate.stat().st_size
            if current_size > 0 and current_size == last_size and not modal_active:
                stable_count += 1
                if stable_count >= 2:
                    LOG.info("PDF rendered and stabilized: %s (%s bytes)", candidate, current_size)
                    return candidate
            else:
                stable_count = 0
            last_size = current_size

    raise WorkflowError(f"Exported PDF was not generated or stabilized within {timeout_seconds} seconds")


def close_pdf_viewer_if_opened(pdf_name: str) -> None:
    """
    Closes any PDF viewer window (such as Adobe Acrobat Reader) across all monitors
    displaying the exported PDF so it does not hold a lock or clutter the screen.
    """
    try:
        viewers = [
            w for w in Desktop(backend="uia").windows()
            if pdf_name.casefold() in text(w).casefold()
            and "power bi" not in text(w).casefold()
            and "visual studio code" not in text(w).casefold()
        ]
        for v in viewers:
            title = text(v)
            v.close()
            LOG.info("Closed PDF viewer window: %s", title)
    except Exception:
        pass


def safe_copy_replace(source: Path, destination: Path, max_attempts: int = 5, retry_delay: float = 2.0) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    last_error = None

    for attempt in range(1, max_attempts + 1):
        try:
            if destination.exists():
                try:
                    destination.unlink()
                except OSError:
                    pass

            shutil.copy2(str(source), str(destination))

            if destination.exists() and destination.stat().st_size > 0:
                return

        except PermissionError as err:
            last_error = err
            LOG.warning(
                "Destination '%s' is locked by another process (attempt %d/%d). Retrying in %.1fs...",
                destination.name, attempt, max_attempts, retry_delay
            )
            time.sleep(retry_delay)
        except Exception as err:
            last_error = err
            time.sleep(retry_delay)

    raise WorkflowError(
        f"Failed to overwrite '{destination}' after {max_attempts} attempts. "
        f"Ensure the file is not currently open/locked by another user. Error: {last_error}"
    )


def distribute_pdf(
    source_pdf: Path,
    main_dir: Path,
    main_filename: str,
    backup_dir: Path,
    backup_prefix: str,
    date_format: str,
) -> dict:
    LOG.info("Distributing PDF outputs to production shares...")

    close_pdf_viewer_if_opened(main_filename)
    close_pdf_viewer_if_opened(source_pdf.name)
    time.sleep(1.0)

    today = date.today()
    date_str = today.strftime(date_format)

    # 1. Main Destination: Always replaces W:\Dashboards VCI\LJSO\Main\Leadership JSO.pdf
    main_dest = main_dir / main_filename
    LOG.info("Saving primary Main PDF (overwriting if present): %s", main_dest)
    safe_copy_replace(source=source_pdf, destination=main_dest)

    # 2. Backup Destination: W:\Dashboards VCI\LJSO\Backup-Archive\Leadership JSO - YYYY-MM-DD.pdf
    backup_filename = f"{backup_prefix} - {date_str}.pdf"
    backup_dest = backup_dir / backup_filename
    LOG.info("Saving dated Backup-Archive PDF: %s", backup_dest)
    safe_copy_replace(source=source_pdf, destination=backup_dest)

    LOG.info("File verification passed. Main: %s bytes | Backup: %s bytes",
             main_dest.stat().st_size, backup_dest.stat().st_size)

    return {
        "main_file": str(main_dest),
        "main_size_bytes": main_dest.stat().st_size,
        "backup_file": str(backup_dest),
        "backup_size_bytes": backup_dest.stat().st_size,
        "main_replaced": True,
    }


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.StreamHandler(sys.stdout)],
    )

    script_dir = Path(__file__).resolve().parent
    settings_path = Path(sys.argv[1] if len(sys.argv) > 1 else script_dir / "settings.json").resolve()

    if settings_path.exists():
        LOG.info("Loading configuration from: %s", settings_path)
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
    else:
        LOG.info("settings.json not found; utilizing built-in enterprise defaults")
        settings = DEFAULT_CONFIG

    try:
        pbix_path = Path(settings.get("pbix_file", DEFAULT_CONFIG["pbix_file"])).resolve()
        keyword = settings.get("window_title_keyword", pbix_path.stem)
        refresh_mode = settings.get("refresh_mode", "data_only")

        open_timeout = int(settings.get("open_timeout_seconds", 180))
        refresh_timeout = int(settings.get("refresh_timeout_seconds", 1200))
        save_timeout = int(settings.get("save_timeout_seconds", 120))
        export_timeout = int(settings.get("pdf_export_timeout_seconds", 300))

        dest_cfg = settings.get("destinations", DEFAULT_CONFIG["destinations"])
        main_dir = Path(dest_cfg.get("main_folder", r"W:\Dashboards VCI\LJSO\Main"))
        main_filename = dest_cfg.get("main_filename", f"{pbix_path.stem}.pdf")
        backup_dir = Path(dest_cfg.get("backup_folder", r"W:\Dashboards VCI\LJSO\Backup-Archive"))
        backup_prefix = dest_cfg.get("backup_filename_prefix", pbix_path.stem)
        date_format = dest_cfg.get("date_format", "%Y-%m-%d")

        window, opened_by_runner = open_and_wait_for_power_bi(pbix_path, keyword, open_timeout)
        refresh_result = refresh_report(window, refresh_timeout, mode=refresh_mode)
        save_result = save_report(window, pbix_path, save_timeout)

        trigger_pdf_export(window)
        generated_pdf = wait_for_pdf_and_stabilize(pbix_path, export_timeout)
        close_pdf_viewer_if_opened(generated_pdf.name)

        distribution_result = distribute_pdf(
            source_pdf=generated_pdf,
            main_dir=main_dir,
            main_filename=main_filename,
            backup_dir=backup_dir,
            backup_prefix=backup_prefix,
            date_format=date_format,
        )

        if settings.get("close_power_bi_when_finished", False):
            LOG.info("Closing Power BI Desktop instance as configured...")
            window.close()

        summary = {
            "workflow": "LEADERSHIP JSO AUTOMATED REFRESH AND PDF EXPORT",
            "executionDate": date.today().isoformat(),
            "pbixPath": str(pbix_path),
            "openedByRunner": opened_by_runner,
            "refresh": refresh_result,
            "save": save_result,
            "distribution": distribution_result,
            "status": "SUCCESS",
        }
        LOG.info("WORKFLOW_COMPLETED=%s", json.dumps(summary))
        return 0

    except Exception as ex:
        LOG.exception("Leadership JSO workflow failed: %s", ex)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())