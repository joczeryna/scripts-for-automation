from __future__ import annotations

import logging
import re
import time
import zipfile
from datetime import date
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page, TimeoutError as PlaywrightTimeout


def _visible(locator: Locator) -> Locator | None:
    for index in range(locator.count()):
        item = locator.nth(index)
        if item.is_visible():
            return item
    return None


def _find_across_surfaces(page: Page, build_locator) -> tuple[Locator | None, Any]:
    for surface in [page, *page.frames]:
        try:
            item = _visible(build_locator(surface))
            if item:
                return item, surface
        except Exception:
            continue
    return None, page


def _wait_for_chrome_download(staging_dir: Path, started_at: float, timeout_seconds: int = 180) -> Path:
    deadline = time.time() + timeout_seconds
    last_size: dict[Path, int] = {}
    stable_count: dict[Path, int] = {}

    while time.time() < deadline:
        completed = list(staging_dir.glob("*.xlsx"))
        partial_named = list(staging_dir.glob("*.xlsx.crdownload"))
        candidates = [
            path
            for path in [*completed, *partial_named]
            if path.stat().st_mtime >= started_at - 2
        ]
        for candidate in candidates:
            size = candidate.stat().st_size
            if size > 0 and last_size.get(candidate) == size:
                stable_count[candidate] = stable_count.get(candidate, 0) + 1
                # Chrome can close after writing a valid XLSX but before removing
                # the .crdownload suffix. ZIP integrity proves the file finished.
                if stable_count[candidate] >= 3 and zipfile.is_zipfile(candidate):
                    return candidate
            else:
                stable_count[candidate] = 0
            last_size[candidate] = size
        time.sleep(1)

    raise RuntimeError(f"Chrome download did not complete within {timeout_seconds} seconds")


def wait_for_authenticated_page(page: Page, expected_text: str, seconds: int) -> None:
    deadline = time.time() + seconds
    while time.time() < deadline:
        if expected_text.casefold() in page.locator("body").inner_text().casefold():
            return
        time.sleep(2)
    raise RuntimeError(
        f"Timed out waiting for '{expected_text}'. Complete sign-in/MFA in Chrome and run again."
    )


def _choose_parameter(
    page: Page,
    label: str,
    values: list[str],
    clear_default_select_all: bool = True,
) -> None:
    logging.info("Setting filter %s = %s", label, values)
    label_rx = re.compile(re.escape(label), re.I)

    label_node = None
    surface = page
    deadline = time.time() + 180
    while time.time() < deadline and not label_node:
        surfaces = [page, *page.frames]
        for candidate_surface in surfaces:
            try:
                candidate = _visible(candidate_surface.get_by_text(label_rx, exact=False))
                if candidate:
                    label_node = candidate
                    surface = candidate_surface
                    logging.info("Located filter '%s' in frame: %s", label, candidate_surface.url)
                    break
            except Exception:
                continue
        if not label_node:
            page.wait_for_timeout(2000)

    if not label_node:
        debug_dir = Path(__file__).resolve().parents[1] / "logs"
        debug_dir.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(debug_dir / "powerbi_debug.png"), full_page=True)
        debug_sections = []
        for index, frame in enumerate(page.frames):
            try:
                body_text = frame.locator("body").inner_text(timeout=10000)
            except Exception as exc:
                body_text = f"Unable to read frame: {exc}"
            debug_sections.append(
                f"\n===== FRAME {index} =====\nURL: {frame.url}\n\n{body_text}\n"
            )
        debug_file = debug_dir / "powerbi_debug.txt"
        debug_file.write_text("\n".join(debug_sections), encoding="utf-8")
        raise RuntimeError(
            f"Power BI filter label not found: {label}. Debug files created at {debug_file}"
        )

    container = label_node.locator("xpath=ancestor::*[self::div or self::td or self::tr][1]")
    control = _visible(container.locator("input, [role='combobox'], select"))
    if not control:
        control = _visible(label_node.locator("xpath=following::*[self::input or @role='combobox' or self::select][1]"))
    if not control:
        raise RuntimeError(f"Input control not found for filter: {label}")

    tag = control.evaluate("element => element.tagName.toLowerCase()")
    if tag == "select":
        for value in values:
            try:
                control.select_option(label=value)
            except Exception:
                control.select_option(value=value)
        return

    control.click()
    page.wait_for_timeout(400)

    # Paginated-report multiselect parameters open with Select All checked.
    # Clear that default before choosing the specific required value(s).
    select_all = _visible(
        surface.get_by_text(re.compile(r"^Select All$", re.I), exact=False)
    )
    if select_all and clear_default_select_all:
        select_all.click()
        page.wait_for_timeout(300)

    for value in values:
        option = _visible(
            surface.get_by_role(
                "option",
                name=re.compile(f"^{re.escape(value)}$", re.I),
                exact=False,
            )
        )
        if not option:
            option = _visible(
                surface.get_by_text(
                    re.compile(f"^{re.escape(value)}$", re.I),
                    exact=False,
                )
            )
        if option:
            option.click()
        else:
            control.fill(value)
            control.press("Enter")
        page.wait_for_timeout(250)
    page.keyboard.press("Tab")
    page.wait_for_timeout(300)


def _click_view_report(page: Page, report_name: str, wait_seconds: int = 15) -> None:
    view_button, _ = _find_across_surfaces(
        page, lambda surface: surface.get_by_role("button", name=re.compile("View report", re.I))
    )
    if not view_button:
        view_button, _ = _find_across_surfaces(
            page, lambda surface: surface.get_by_text(re.compile("View report", re.I), exact=False)
        )
    if not view_button:
        raise RuntimeError(f"View report button not found for {report_name}")

    deadline = time.time() + 60
    while time.time() < deadline and view_button.is_disabled():
        page.wait_for_timeout(1000)
    if view_button.is_disabled():
        raise RuntimeError(f"View report button remained disabled for {report_name}")
    view_button.click()
    page.wait_for_timeout(wait_seconds * 1000)


def export_power_bi_report(page: Page, report: dict[str, Any], run_date: date, output_dir: Path) -> Path:
    page.goto(report["url"], wait_until="domcontentloaded")
    wait_for_authenticated_page(page, report["name"].replace("_", " ").split()[0], 300)

    initialization = report.get("two_stage_initialization")
    if initialization:
        logging.info("Starting first-stage initialization for %s", report["name"])
        _choose_parameter(
            page,
            initialization["filter_label"],
            [initialization["filter_value"]],
            clear_default_select_all=False,
        )
        _click_view_report(
            page,
            report["name"],
            initialization.get("load_wait_seconds", 20),
        )
        logging.info("JSO second initialization completed; proceeding to export for %s", report["name"])

    for item in report["filters"]:
        if item.get("use_default"):
            logging.info("Leaving filter %s at its configured default", item["label"])
            continue
        if "value_by_day" in item:
            day_map = item["value_by_day"]
            value = day_map.get(run_date.strftime("%A"), day_map["default"])
            values = [value]
        else:
            values = item.get("values", [item["value"]])

        if len(values) == 1 and values[0].strip().casefold() == "select all":
            logging.info("Leaving filter %s at its default Select All", item["label"])
            continue

        _choose_parameter(
            page,
            item["label"],
            values,
            clear_default_select_all=not item.get("starts_empty", False),
        )

    report_surface = page
    if not report.get("export_after_initialization"):
        view_button, report_surface = _find_across_surfaces(
            page, lambda surface: surface.get_by_role("button", name=re.compile("View report", re.I))
        )
        if not view_button:
            view_button, report_surface = _find_across_surfaces(
                page, lambda surface: surface.get_by_text(re.compile("View report", re.I), exact=False)
            )
        if not view_button:
            raise RuntimeError(f"View report button not found for {report['name']}")

        enabled_deadline = time.time() + 60
        while time.time() < enabled_deadline and view_button.is_disabled():
            page.wait_for_timeout(1000)
        if view_button.is_disabled():
            debug_dir = Path(__file__).resolve().parents[1] / "logs"
            debug_dir.mkdir(parents=True, exist_ok=True)
            page.screenshot(
                path=str(debug_dir / "view_report_disabled.png"),
                full_page=True,
            )
            raise RuntimeError(f"View report button remained disabled for {report['name']}")
        view_button.click()
        page.wait_for_timeout(15000)
    else:
        logging.info("Exporting %s immediately after JSO initialization", report["name"])

    export_button, report_surface = _find_across_surfaces(
        page,
        lambda surface: surface.get_by_text(
            re.compile("^Export$", re.I),
            exact=False,
        ),
    )
    if not export_button:
        export_button, report_surface = _find_across_surfaces(
            page, lambda surface: surface.get_by_role("button", name=re.compile("Export", re.I))
        )
    if not export_button:
        raise RuntimeError(f"Export menu not found for {report['name']}")
    export_button.click()
    excel = _visible(report_surface.get_by_text(re.compile(r"Microsoft Excel.*xlsx", re.I)))
    if not excel:
        excel, report_surface = _find_across_surfaces(
            page, lambda surface: surface.get_by_text(re.compile(r"Microsoft Excel.*xlsx", re.I))
        )
    if not excel:
        raise RuntimeError("Microsoft Excel (.xlsx) export option was not found")

    staging_dir = output_dir / "_browser_downloads"
    staging_dir.mkdir(parents=True, exist_ok=True)
    cdp = page.context.new_cdp_session(page)
    cdp.send(
        "Browser.setDownloadBehavior",
        {
            "behavior": "allow",
            "downloadPath": str(staging_dir.resolve()),
            "eventsEnabled": True,
        },
    )
    started_at = time.time()
    excel.click()
    downloaded_file = _wait_for_chrome_download(staging_dir, started_at)
    destination = output_dir / f"{report['name']}_{run_date:%Y-%m-%d}.xlsx"
    downloaded_file.replace(destination)
    logging.info("Downloaded %s", destination)
    return destination


def upload_to_drive(page: Page, drive: dict[str, Any], files: list[Path]) -> None:
    page.goto(drive["folder_url"], wait_until="domcontentloaded")
    wait_for_authenticated_page(page, drive["folder_name"], 300)

    new_button = _visible(page.get_by_role("button", name=re.compile("New", re.I)))
    if not new_button:
        new_button = _visible(page.get_by_text(re.compile("^New$", re.I)))
    if not new_button:
        raise RuntimeError("Google Drive New button not found")
    new_button.click()
    upload_item = _visible(
        page.get_by_role("menuitem", name=re.compile(r"^File upload$", re.I))
    )
    if not upload_item:
        upload_text = _visible(
            page.get_by_text(re.compile(r"^File upload$", re.I), exact=False)
        )
        if upload_text:
            upload_item = upload_text.locator("xpath=ancestor::*[@role='menuitem'][1]")
    if not upload_item:
        raise RuntimeError("Google Drive File upload menu item not found")

    with page.expect_file_chooser() as chooser_info:
        upload_item.click()
    chooser_info.value.set_files([str(path.resolve()) for path in files])

    page.get_by_text(re.compile(r"upload complete|uploads complete", re.I)).wait_for(
        state="visible", timeout=180000
    )
    logging.info("Uploaded %d files to %s", len(files), drive["folder_name"])


def _run_sheet_menu(page: Page, menu: str, item: str, wait_seconds: int) -> None:
    menu_button = _visible(page.get_by_role("button", name=menu, exact=True))
    if not menu_button:
        raise RuntimeError(f"Google Sheets menu button not found: {menu}")
    menu_button.click()

    menu_item = _visible(page.get_by_role("menuitem", name=item, exact=True))
    if not menu_item:
        menu_item = _visible(page.get_by_text(item, exact=True))
    if not menu_item:
        raise RuntimeError(f"Google Sheets menu item not found: {item}")
    menu_item.click()
    logging.info("Started Google Sheets menu action: %s", item)
    page.wait_for_timeout(wait_seconds * 1000)


def update_executive_dashboard(page: Page, dashboard: dict[str, Any], run_date: date) -> None:
    page.goto(dashboard["url"], wait_until="domcontentloaded")
    wait_for_authenticated_page(page, dashboard["name"], 300)
    page.wait_for_timeout(5000)

    _run_sheet_menu(
        page, dashboard["updates_menu"], dashboard["compliance_item"],
        dashboard.get("post_click_wait_seconds", 2)
    )
    logging.info("Safety Compliance Report triggered; dashboard automation is complete")
