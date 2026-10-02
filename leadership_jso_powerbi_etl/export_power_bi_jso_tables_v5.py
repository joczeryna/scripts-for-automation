from __future__ import annotations

import csv
import json
import logging
import shutil
import sys
import time
from pathlib import Path

from pywinauto import Desktop, keyboard, mouse

import update_power_bi_month_filters_v2 as month_filters


LOG = logging.getLogger("leadership_jso_exports")
EXPORT_NAMES = {
    "Leadership JSOs": "Leadership_JSOs",
    "JSO - Deviation Statuses": "JSO_Deviation_Statuses",
    "JSO - All Audits": "JSO_All_Audits",
}


class ExportError(RuntimeError):
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
        if visible(c) and text(c).casefold() == title.casefold()
    ]
    if not matches:
        raise ExportError(f"Visible {control_type} not found: {title}")
    return matches[0]


def power_bi_window():
    matches = [
        w for w in Desktop(backend="uia").windows()
        if "safety reports" in text(w).casefold()
    ]
    if not matches:
        raise ExportError("Safety Reports is not open in Power BI Desktop")
    return matches[0]


def select_page(window, page_name: str) -> None:
    exact(window, page_name, "TabItem").click_input()
    time.sleep(4)


def table_menu(window):
    cells = [
        c for c in window.descendants(control_type="DataItem")
        if visible(c) and "pivotTableCellWrap" in str(c.element_info.class_name or "")
    ]
    headers = [c for c in cells if text(c).casefold() == "cmonth name"]
    if not headers:
        raise ExportError("CMonth Name table column was not found")
    h = headers[0].rectangle()
    body = [
        c for c in cells
        if c.rectangle().top > h.bottom and abs(c.rectangle().left - h.left) <= 8
    ]
    target = min(body, key=lambda c: c.rectangle().top) if body else headers[0]
    r = target.rectangle()
    mouse.move(coords=(r.left + r.width() // 2, r.top + r.height() // 2))
    time.sleep(2)
    menus = [
        c for c in window.descendants(control_type="Button")
        if visible(c)
        and text(c) == "More options"
        and str(c.element_info.class_name or "").strip() == "vcMenuBtn"
    ]
    if len(menus) != 1:
        raise ExportError(f"Expected one table More options button; found {len(menus)}")
    return menus[0]


def wait_for_save_dialog(window, timeout: int = 30):
    deadline = time.time() + timeout
    while time.time() < deadline:
        dialogs = [
            c for c in window.descendants(control_type="Window")
            if visible(c) and text(c) == "Save As"
        ]
        if dialogs:
            return dialogs[0]
        time.sleep(0.5)
    raise ExportError("Power BI Save As dialog did not appear")


def wait_for_file(path: Path, timeout: int = 180) -> None:
    deadline = time.time() + timeout
    prior = -1
    stable = 0
    while time.time() < deadline:
        if path.exists():
            size = path.stat().st_size
            if size > 0 and size == prior:
                stable += 1
                if stable >= 3:
                    return
            else:
                stable = 0
            prior = size
        time.sleep(1)
    raise ExportError(f"Export file was not created or stabilized: {path}")


def settle_export(output: Path, started_at: float, timeout: int = 180) -> dict:
    downloads_dir = Path.home() / "Downloads"
    deadline = time.time() + timeout
    tracked: dict[Path, tuple[int, int]] = {}
    while time.time() < deadline:
        candidates = []
        for folder in (output.parent, downloads_dir):
            if not folder.exists():
                continue
            candidates.extend(
                path for path in folder.glob("*.csv")
                if path.stat().st_mtime >= started_at - 2
            )
        if output.exists():
            candidates.append(output)
        for candidate in sorted(set(candidates), key=lambda p: p.stat().st_mtime, reverse=True):
            prior, stable = tracked.get(candidate, (-1, 0))
            size = candidate.stat().st_size
            stable = stable + 1 if size > 0 and size == prior else 0
            tracked[candidate] = (size, stable)
            if stable >= 3:
                if candidate != output:
                    output.parent.mkdir(parents=True, exist_ok=True)
                    candidate.replace(output)
                    return {
                        "powerBiOriginalName": candidate.name,
                        "renamedTo": output.name,
                        "movedFrom": str(candidate.parent),
                    }
                return {
                    "powerBiOriginalName": output.name,
                    "renamedTo": output.name,
                    "movedFrom": "",
                }
        time.sleep(1)
    raise ExportError(
        f"No new Power BI CSV stabilized in {output.parent} or {downloads_dir}"
    )


def export_page(window, page_name: str, output: Path) -> dict:
    select_page(window, page_name)
    menu = table_menu(window)
    menu.click_input()
    time.sleep(1)
    exact(window, "Export data", "MenuItem").click_input()
    dialog = wait_for_save_dialog(window)
    dialog.set_focus()
    keyboard.send_keys("^l")
    time.sleep(0.5)
    keyboard.send_keys(str(output.parent), with_spaces=True)
    keyboard.send_keys("{ENTER}")
    time.sleep(2)
    filename = exact(dialog, "File name:", "Edit")
    filename.set_edit_text(output.name)
    started_at = time.time()
    exact(dialog, "Save", "Button").click_input()
    placement = settle_export(output, started_at)
    return {
        "page": page_name,
        "file": str(output),
        "bytes": output.stat().st_size,
        **placement,
    }


def read_csv(path: Path) -> tuple[list[str], list[list[str]], str]:
    errors = []
    for encoding in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                rows = list(csv.reader(handle))
            if not rows:
                raise ExportError(f"Export is empty: {path}")
            return rows[0], rows[1:], encoding
        except UnicodeError as error:
            errors.append(f"{encoding}: {error}")
    raise ExportError(f"Could not decode {path}: {'; '.join(errors)}")


def remove_grand_total(path: Path) -> int:
    headers, rows, encoding = read_csv(path)
    kept = []
    removed = 0
    for row in rows:
        values = {str(value).strip().casefold() for value in row}
        if "grand total" in values:
            removed += 1
        else:
            kept.append(row)
    if removed:
        with path.open("w", encoding=encoding, newline="") as handle:
            writer = csv.writer(handle)
            writer.writerow(headers)
            writer.writerows(kept)
    return removed


def validate_export(path: Path, page_name: str, allow_empty: bool = False) -> dict:
    grand_totals_removed = remove_grand_total(path) if page_name == "Leadership JSOs" else 0
    headers, rows, encoding = read_csv(path)
    clean_headers = [str(value).strip() for value in headers]
    if len(clean_headers) < 3 or not all(clean_headers):
        raise ExportError(f"Invalid headers in {path.name}: {clean_headers}")
    if not rows and not allow_empty:
        raise ExportError(f"No data rows were exported for {page_name}")
    header_keys = {value.casefold().replace(" ", "") for value in clean_headers}
    if "cmonthname" not in header_keys:
        raise ExportError(f"CMonth Name header is missing from {path.name}")
    if page_name != "Leadership JSOs" and "auditid" not in header_keys:
        raise ExportError(f"Audit Id header is missing from {path.name}")
    return {
        "page": page_name,
        "file": str(path),
        "encoding": encoding,
        "columns": len(clean_headers),
        "rows": len(rows),
        "headers": clean_headers,
        "grandTotalRowsRemoved": grand_totals_removed,
        "bytes": path.stat().st_size,
    }


def set_all_audits_months(window, required: list[str]) -> dict:
    last_error = None
    for attempt in range(1, 4):
        try:
            select_page(window, "JSO - All Audits")
            month_filters.select_table_visual(window)
            before = text(month_filters.month_card_button(window))
            if month_filters.summary_matches(before, required):
                return {
                    "required": required,
                    "summary": before,
                    "checkboxes": [],
                    "changed": False,
                    "recoveryAttempt": attempt,
                }
            month_filters.expand_month_card(window)
            break
        except Exception as error:
            last_error = error
            if attempt == 3:
                raise ExportError(
                    f"Could not open the All Audits CMonth Name filter after 3 attempts: {error}"
                ) from error
            # Escape any menu/dialog state left behind by Export data, then
            # reacquire the page and table visual from a clean UI state.
            keyboard.send_keys("{ESC}")
            time.sleep(3)
    else:
        raise ExportError(f"Could not prepare All Audits month filter: {last_error}")
    states = [
        month_filters.set_month(window, month, month in required)
        for month in month_filters.MONTHS
    ]
    month_filters.clear_search(window)
    summary = text(month_filters.month_card_button(window))
    if not month_filters.summary_matches(summary, required):
        raise ExportError(
            f"All Audits month-filter verification failed. Expected {required}; got {summary}"
        )
    return {
        "required": required,
        "summary": summary,
        "checkboxes": states,
        "changed": True,
        "recoveryAttempt": attempt,
    }


def combine_monthly_exports(parts: list[Path], output: Path) -> dict:
    combined_headers = None
    combined_rows = []
    part_counts = []
    for part in parts:
        headers, rows, _ = read_csv(part)
        if combined_headers is None:
            combined_headers = headers
        elif headers != combined_headers:
            raise ExportError(
                f"All Audits monthly headers do not match: {part.name}"
            )
        if len(rows) >= 30000:
            raise ExportError(
                f"{part.name} reached the 30,000-row Power BI export limit; "
                "the month must be split further before loading Google Sheets."
            )
        combined_rows.extend(rows)
        part_counts.append({"file": str(part), "rows": len(rows)})
    if not combined_headers or not combined_rows:
        raise ExportError("All Audits monthly exports produced no combined data")
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(combined_headers)
        writer.writerows(combined_rows)
    return {"parts": part_counts, "combinedRows": len(combined_rows)}


def export_all_audits_by_month(window, export_dir: Path, stamp: str, months: list[str]) -> tuple[dict, dict]:
    parts_dir = export_dir / "_all_audits_parts"
    parts_dir.mkdir(parents=True, exist_ok=False)
    parts = []
    filters = []
    try:
        for month in months:
            filters.append(set_all_audits_months(window, [month]))
            part = parts_dir / f"JSO_All_Audits_{month}_{stamp}.csv"
            export_page(window, "JSO - All Audits", part)
            # Validate headers and enforce the Power BI per-export row cap now.
            # On the first day(s) of a month, the refreshed source may not yet
            # contain current-month rows. Prior months must never be empty.
            validation = validate_export(
                part,
                "JSO - All Audits",
                allow_empty=(month == months[-1]),
            )
            if validation["rows"] >= 30000:
                raise ExportError(
                    f"All Audits {month} export reached 30,000 rows and is incomplete"
                )
            parts.append(part)

        combined = export_dir / f"JSO_All_Audits_{stamp}.csv"
        merge = combine_monthly_exports(parts, combined)
        merge["parts"] = [
            {"month": month, "rows": detail["rows"]}
            for month, detail in zip(months, merge["parts"])
        ]
        restored = set_all_audits_months(window, months)
    finally:
        # Keep only the final combined file; monthly files are temporary.
        shutil.rmtree(parts_dir, ignore_errors=True)
    return (
        validate_export(combined, "JSO - All Audits"),
        {
            "monthlyFilters": filters,
            "restoredFilter": restored,
            **merge,
            "temporaryMonthlyFilesRemoved": True,
        },
    )


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    settings_path = Path(sys.argv[1] if len(sys.argv) > 1 else "settings.json").resolve()
    try:
        settings = json.loads(settings_path.read_text(encoding="utf-8"))
        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        export_dir = settings_path.parent / "exports" / stamp
        export_dir.mkdir(parents=True, exist_ok=False)
        window = power_bi_window()
        window.set_focus()
        exact(window, "Report view", "TabItem").click_input()
        time.sleep(3)

        created = []
        for page_name in settings["report_pages"]:
            if page_name == "JSO - All Audits":
                continue
            base = EXPORT_NAMES.get(page_name)
            if not base:
                raise ExportError(f"No controlled export filename is configured for {page_name}")
            output = export_dir / f"{base}_{stamp}.csv"
            created.append(export_page(window, page_name, output))

        validations = [
            validate_export(Path(item["file"]), item["page"])
            for item in created
        ]
        # All Audits always uses one simultaneous three-month filter and one
        # export: current month plus the two preceding months.
        rolling_months = month_filters.rolling_months(
            __import__("datetime").date.today()
        )
        all_audits_filter = set_all_audits_months(window, rolling_months)
        all_audits_output = export_dir / f"JSO_All_Audits_{stamp}.csv"
        all_audits_export = export_page(
            window,
            "JSO - All Audits",
            all_audits_output,
        )
        all_audits_validation = validate_export(
            Path(all_audits_export["file"]),
            "JSO - All Audits",
        )
        validations.append(all_audits_validation)
        pbix_path = Path(settings.get("power_bi_file", r"C:\Users\jd0855\Downloads\Safety Reports.pbix"))
        restore_save = month_filters.save_report(window, pbix_path)
        result = {
            "mode": "CONTROLLED POWER BI TABLE EXPORT - NO GOOGLE SHEETS WRITES",
            "runTimestamp": stamp,
            "exportDirectory": str(export_dir),
            "exports": validations,
            "allAuditsFilter": {
                "requiredMonths": rolling_months,
                "summary": all_audits_filter["summary"],
                "singleExport": True,
                "exportedRows": all_audits_validation["rows"],
                "rowCountAcceptedWithoutLimit": True,
            },
            "exportsCreated": len(validations),
            "allExportsValid": len(validations) == len(settings["report_pages"]),
            "powerBiRefreshPerformed": False,
            "powerBiSavePerformed": True,
            "powerBiSaveAfterFilterRestore": restore_save,
            "googleSheetsWrites": 0,
        }
        LOG.info("LEADERSHIP_JSO_EXPORT_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Power BI JSO table export failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
