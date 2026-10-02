from __future__ import annotations

import json
import logging
import sys
import time
from datetime import date
from pathlib import Path

from pywinauto import Desktop, keyboard, mouse


LOG = logging.getLogger("leadership_jso_filter_update")
MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
ROLLING_PAGES = ["JSO - Deviation Statuses", "JSO - All Audits"]


class FilterUpdateError(RuntimeError):
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


def exact(window, title: str, control_type: str):
    matches = [
        c for c in window.descendants(control_type=control_type)
        if visible(c) and text(c).casefold() == title.casefold()
    ]
    if not matches:
        raise FilterUpdateError(f"Visible {control_type} not found: {title}")
    return matches[0]


def power_bi_window():
    matches = [
        w for w in Desktop(backend="uia").windows()
        if "safety reports" in text(w).casefold()
    ]
    if not matches:
        raise FilterUpdateError("Safety Reports is not open in Power BI Desktop")
    return matches[0]


def rolling_months(today: date) -> list[str]:
    values = []
    year, month = today.year, today.month
    for offset in (2, 1, 0):
        index = year * 12 + month - 1 - offset
        values.append(MONTHS[index % 12])
    return values


def select_page(window, page_name: str) -> None:
    exact(window, page_name, "TabItem").click_input()
    time.sleep(4)


def select_table_visual(window) -> dict:
    deadline = time.time() + 75
    last_headers = []

    while time.time() < deadline:
        all_cells = [
            c for c in window.descendants(control_type="DataItem")
            if visible(c)
        ]
        pivot_cells = [
            c for c in all_cells
            if "pivotTableCellWrap" in str(c.element_info.class_name or "")
        ]
        cells = pivot_cells or all_cells
        last_headers = sorted({text(c) for c in cells if text(c)})

        # CMonth Name is the preferred table anchor. CY Name is in the same
        # visual and is a safe fallback when Power BI has not exposed the
        # second header through UI Automation after a refresh.
        header = None
        header_name = ""
        for candidate_name in ("cmonth name", "cy name"):
            matches = [
                c for c in cells
                if " ".join(text(c).casefold().split()) == candidate_name
            ]
            if matches:
                header = matches[0]
                header_name = candidate_name
                break

        if header is not None:
            hr = header.rectangle()
            body = [
                c for c in cells
                if c.rectangle().top > hr.bottom
                and abs(c.rectangle().left - hr.left) <= 8
            ]
            target = min(body, key=lambda c: c.rectangle().top) if body else header
            r = target.rectangle()
            mouse.click(
                button="left",
                coords=(r.left + r.width() // 2, r.top + r.height() // 2),
            )
            time.sleep(3)
            if any(
                visible(c)
                and " ".join(text(c).casefold().split()) == "filters on this visual"
                for c in window.descendants(control_type="Text")
            ):
                return {
                    "cellRectangle": [r.left, r.top, r.right, r.bottom],
                    "anchorHeader": header_name,
                }

        time.sleep(3)

    diagnostics = Path(__file__).resolve().parent / "diagnostics"
    diagnostics.mkdir(parents=True, exist_ok=True)
    screenshot = diagnostics / (
        "table_visual_selection_failed_" + time.strftime("%Y%m%d_%H%M%S") + ".png"
    )
    try:
        window.capture_as_image().save(screenshot)
    except Exception:
        screenshot = None
    raise FilterUpdateError(
        "The report table could not be selected after 75 seconds. "
        "Visible table values included: "
        + ", ".join(last_headers[:20])
        + (f". Screenshot: {screenshot}" if screenshot else "")
    )


def month_card_button(window):
    matches = [
        c for c in window.descendants(control_type="Button")
        if visible(c)
        and text(c).casefold().startswith("cmonth name ")
        and "title" in str(c.element_info.class_name or "")
    ]
    if not matches:
        raise FilterUpdateError("CMonth Name filter card was not found")
    return min(matches, key=lambda c: c.rectangle().top)


def expand_month_card(window) -> None:
    # A visible searchInput is the most reliable indication that the
    # CMonth Name card is already expanded. Power BI does not always expose
    # the Filter type ComboBox after returning from an Export dialog.
    if any(
        visible(c) and str(c.element_info.class_name or "") == "searchInput"
        for c in window.descendants(control_type="Edit")
    ):
        return
    expanders = [
        c for c in window.descendants(control_type="Button")
        if visible(c)
        and "cmonth name" in text(c).casefold()
        and "expand or collapse" in text(c).casefold()
    ]
    if expanders:
        expanders[0].click_input()
    else:
        # Some Power BI builds expose only the filter-card title button.
        # Clicking the title expands the same card without changing filters.
        month_card_button(window).click_input()
    time.sleep(3)
    if not any(
        visible(c) and str(c.element_info.class_name or "") == "searchInput"
        for c in window.descendants(control_type="Edit")
    ):
        raise FilterUpdateError("CMonth Name filter card did not expand")


def filter_search(window):
    edits = [
        c for c in window.descendants(control_type="Edit")
        if visible(c) and str(c.element_info.class_name or "") == "searchInput"
    ]
    if not edits:
        raise FilterUpdateError("CMonth Name filter search box was not found")
    return edits[0]


def search_month(window, month: str):
    search = filter_search(window)
    search.click_input()
    keyboard.send_keys("^a{BACKSPACE}")
    keyboard.send_keys(month, with_spaces=True)
    time.sleep(2)
    matches = [
        c for c in window.descendants(control_type="CheckBox")
        if visible(c) and text(c).casefold() == month.casefold()
    ]
    if not matches:
        raise FilterUpdateError(f"Month checkbox was not found after search: {month}")
    return matches[0]


def toggle_state(control) -> int:
    try:
        return int(control.get_toggle_state())
    except Exception:
        return -1


def set_month(window, month: str, required: bool) -> dict:
    box = search_month(window, month)
    before = toggle_state(box)
    if before not in (0, 1):
        raise FilterUpdateError(f"Could not determine checkbox state for {month}")
    desired = 1 if required else 0
    if before != desired:
        box.click_input()
        time.sleep(3)
        # Reacquire after Power BI rerenders the filter list.
        box = search_month(window, month)
        after = toggle_state(box)
        if after != desired:
            raise FilterUpdateError(
                f"{month} checkbox verification failed; expected {desired}, got {after}"
            )
    else:
        after = before
    return {"month": month, "required": required, "before": before, "after": after}


def clear_search(window) -> None:
    search = filter_search(window)
    search.click_input()
    keyboard.send_keys("^a{BACKSPACE}")
    time.sleep(2)


def summary_matches(summary: str, required: list[str]) -> bool:
    folded = summary.casefold()
    return all(month.casefold() in folded for month in required) and all(
        month.casefold() not in folded for month in MONTHS if month not in required
    )


def update_page(window, page_name: str, required: list[str]) -> dict:
    select_page(window, page_name)
    anchor = select_table_visual(window)
    before = text(month_card_button(window))
    if summary_matches(before, required):
        return {"page": page_name, "before": before, "after": before, "changed": False, "anchor": anchor}

    expand_month_card(window)
    checkbox_changes = []
    for month in MONTHS:
        checkbox_changes.append(set_month(window, month, month in required))
    clear_search(window)
    after = text(month_card_button(window))
    if not summary_matches(after, required):
        raise FilterUpdateError(
            f"{page_name} verification failed. Expected {required}; Power BI reports: {after}"
        )
    return {
        "page": page_name,
        "before": before,
        "after": after,
        "changed": True,
        "checkboxes": checkbox_changes,
        "anchor": anchor,
    }


def refresh_report(window, timeout_seconds: int) -> dict:
    buttons = [
        c for c in window.descendants(control_type="Button")
        if visible(c) and str(c.element_info.automation_id or "") == "refreshQueries"
    ]
    if not buttons:
        raise FilterUpdateError("Power BI Refresh button was not found")
    started = time.time()
    buttons[0].click_input()
    time.sleep(5)

    stable_since = None
    while time.time() - started < timeout_seconds:
        time.sleep(2)
        refresh_buttons = [
            c for c in window.descendants(control_type="Button")
            if visible(c) and str(c.element_info.automation_id or "") == "refreshQueries"
        ]
        busy_text = [
            text(c).casefold() for c in window.descendants()
            if visible(c) and text(c)
        ]
        busy = any(
            phrase in value
            for value in busy_text
            for phrase in ("cancel refresh", "refreshing", "loading data", "applying changes")
        )
        enabled = bool(refresh_buttons) and refresh_buttons[0].is_enabled()
        if enabled and not busy:
            stable_since = stable_since or time.time()
            if time.time() - stable_since >= 8:
                return {"completed": True, "elapsedSeconds": round(time.time() - started, 1)}
        else:
            stable_since = None
    raise FilterUpdateError(
        f"Power BI refresh did not complete within {timeout_seconds} seconds"
    )


def validate_leadership(window, current_year: int) -> dict:
    select_page(window, "Leadership JSOs")
    anchor = select_table_visual(window)
    month_summary = text(month_card_button(window))
    all_cells = [
        c for c in window.descendants(control_type="DataItem")
        if visible(c)
    ]
    pivot_cells = [
        c for c in all_cells
        if "pivotTableCellWrap" in str(c.element_info.class_name or "")
    ]
    cells = pivot_cells or all_cells
    year_headers = [
        c for c in cells
        if " ".join(text(c).casefold().split()) == "cy name"
    ]
    if not year_headers:
        raise FilterUpdateError("Leadership JSOs CY Name column was not found")
    yr = year_headers[0].rectangle()
    visible_years = sorted({
        text(c) for c in cells
        if c.rectangle().top > yr.bottom
        and abs(c.rectangle().left - yr.left) <= 8
        and text(c)
    })
    if "(all)" not in month_summary.casefold():
        raise FilterUpdateError(
            f"Leadership JSOs must include all months; Power BI reports: {month_summary}"
        )
    if visible_years != [str(current_year)]:
        raise FilterUpdateError(
            f"Leadership JSOs must show only {current_year}; visible CY Name values: {visible_years}"
        )
    return {
        "page": "Leadership JSOs",
        "monthFilter": month_summary,
        "visibleYearValues": visible_years,
        "changed": False,
        "anchor": anchor,
    }


def save_report(window, pbix_path: Path) -> dict:
    before = pbix_path.stat().st_mtime if pbix_path.exists() else None
    save_buttons = [
        c for c in window.descendants(control_type="Button")
        if visible(c) and str(c.element_info.automation_id or "") == "save"
    ]
    if save_buttons:
        save_buttons[0].click_input()
    else:
        keyboard.send_keys("^s")
    deadline = time.time() + 120
    after = before
    while time.time() < deadline:
        time.sleep(2)
        if pbix_path.exists():
            after = pbix_path.stat().st_mtime
            if before is None or after > before:
                break
    if before is not None and (after is None or after <= before):
        raise FilterUpdateError("Power BI Save was requested, but the PBIX modified time did not change")
    return {"path": str(pbix_path), "modifiedTimeChanged": before is None or after > before}


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    settings_path = Path(sys.argv[1] if len(sys.argv) > 1 else "settings.json").resolve()
    try:
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
        pbix_path = Path(settings.get("power_bi_file", r"C:\Users\jd0855\Downloads\Safety Reports.pbix"))
        today = date.today()
        required = rolling_months(today)
        window = power_bi_window()
        window.set_focus()
        exact(window, "Report view", "TabItem").click_input()
        time.sleep(3)

        refresh = refresh_report(
            window,
            int(settings.get("refresh_timeout_seconds", 1800)),
        )
        refresh_save = save_report(window, pbix_path)

        leadership = validate_leadership(window, today.year)
        pages = [update_page(window, page, required) for page in ROLLING_PAGES]
        changed = any(item["changed"] for item in pages)
        filter_save = save_report(window, pbix_path) if changed else {
            "path": str(pbix_path), "modifiedTimeChanged": False, "notRequired": True
        }
        result = {
            "mode": "CONTROLLED POWER BI MONTH FILTER UPDATE",
            "runDate": today.isoformat(),
            "rollingMonths": required,
            "leadership": leadership,
            "rollingPages": pages,
            "filtersChanged": changed,
            "refreshPerformed": True,
            "refresh": refresh,
            "saveAfterRefresh": refresh_save,
            "filterSavePerformed": changed,
            "filterSaveVerification": filter_save,
            "exportsCreated": 0,
            "googleSheetsWrites": 0,
        }
        LOG.info("LEADERSHIP_JSO_MONTH_FILTER_UPDATE=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Power BI month filter update failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
