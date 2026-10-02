from __future__ import annotations

import logging
import re
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from playwright.sync_api import Locator, Page


class AuthenticationError(RuntimeError):
    pass


def _visible(locator: Locator) -> Locator | None:
    for index in range(locator.count()):
        candidate = locator.nth(index)
        try:
            if candidate.is_visible():
                return candidate
        except Exception:
            continue
    return None


def _surfaces(page: Page):
    return [page, *page.frames]


def _find_across_surfaces(page: Page, builder) -> tuple[Locator | None, Any]:
    for surface in _surfaces(page):
        try:
            found = _visible(builder(surface))
            if found:
                return found, surface
        except Exception:
            continue
    return None, page


def _wait_for_text(page: Page, expected: str, seconds: int) -> None:
    deadline = time.time() + seconds
    while time.time() < deadline:
        for surface in _surfaces(page):
            try:
                if expected.casefold() in surface.locator("body").inner_text(timeout=5_000).casefold():
                    return
            except Exception:
                continue
        time.sleep(2)
    if "login" in page.url.casefold() or "microsoftonline" in page.url.casefold():
        raise AuthenticationError(
            "Power BI authentication is required. Complete sign-in/MFA in the opened Chrome window."
        )
    raise RuntimeError(f"Timed out waiting for page content containing '{expected}'.")


def _parameter_control(page: Page, label: str) -> tuple[Locator, Any]:
    label_rx = re.compile(re.escape(label), re.I)
    deadline = time.time() + 180
    while time.time() < deadline:
        for surface in _surfaces(page):
            try:
                label_node = _visible(surface.get_by_text(label_rx, exact=False))
                if not label_node:
                    continue
                container = label_node.locator(
                    "xpath=ancestor::*[self::div or self::td or self::tr][1]"
                )
                control = _visible(container.locator("input, [role='combobox'], select"))
                if not control:
                    control = _visible(
                        label_node.locator(
                            "xpath=following::*[self::input or @role='combobox' or self::select][1]"
                        )
                    )
                if control:
                    return control, surface
            except Exception:
                continue
        page.wait_for_timeout(2_000)
    raise RuntimeError(f"Power BI parameter control not found: {label}")


def _choose_parameter(
    page: Page,
    label: str,
    value: str,
    clear_select_all: bool,
) -> None:
    logging.info("Setting %s = %s", label, value)
    control, surface = _parameter_control(page, label)
    control.click()
    page.wait_for_timeout(500)

    if clear_select_all:
        # Do not blindly click Select All: when it starts unchecked that would
        # select Blank() as well. Explicitly clear Blank() instead.
        blank_text = _visible(
            surface.get_by_text(re.compile(r"^\s*Blank\(\)\s*$", re.I), exact=False)
        )
        if blank_text:
            blank_row = blank_text.locator(
                "xpath=ancestor::*[self::label or @role='option' or "
                "@role='menuitemcheckbox' or contains(@class,'item')][1]"
            )
            blank_checkbox = _visible(
                blank_row.locator("input[type='checkbox'], [role='checkbox']")
            )
            blank_checked = False
            if blank_checkbox:
                try:
                    blank_checked = blank_checkbox.is_checked()
                except Exception:
                    blank_checked = blank_checkbox.get_attribute("aria-checked") in ("true", "mixed")
            if blank_checked:
                blank_text.click()
                page.wait_for_timeout(500)

    value_rx = re.compile(
        r"^\s*" + r"\s+".join(re.escape(part) for part in value.split()) + r"\s*$",
        re.I,
    )
    option = _visible(surface.get_by_role("option", name=value_rx, exact=False))
    if not option:
        option, option_surface = _find_across_surfaces(
            page,
            lambda candidate: candidate.get_by_text(value_rx, exact=False),
        )
        if option:
            surface = option_surface
    if not option and value.casefold().startswith("vci construction"):
        # Power BI can render the company caption without the LLC suffix or
        # with irregular whitespace even though the selected value is the same.
        company_rx = re.compile(r"VCI\s+CONSTRUCTION(?:\s*,?\s*LLC)?", re.I)
        option, option_surface = _find_across_surfaces(
            page,
            lambda candidate: candidate.get_by_text(company_rx, exact=False),
        )
        if option:
            surface = option_surface
    if not option:
        debug_dir = Path(__file__).resolve().parents[1] / "logs"
        debug_dir.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(debug_dir / "dvir_parameter_debug.png"), full_page=True)
        debug_text = []
        for index, frame in enumerate(page.frames):
            try:
                body = frame.locator("body").inner_text(timeout=10_000)
            except Exception as exc:
                body = f"Unable to read frame: {exc}"
            debug_text.append(f"===== FRAME {index} =====\nURL: {frame.url}\n\n{body}")
        (debug_dir / "dvir_parameter_debug.txt").write_text(
            "\n\n".join(debug_text), encoding="utf-8"
        )
        raise RuntimeError(f"Value '{value}' was not found for parameter '{label}'.")
    option_row = option.locator(
        "xpath=ancestor::*[self::label or @role='option' or "
        "@role='menuitemcheckbox' or contains(@class,'item')][1]"
    )
    option_checkbox = _visible(
        option_row.locator("input[type='checkbox'], [role='checkbox']")
    )
    already_checked = False
    if option_checkbox:
        try:
            already_checked = option_checkbox.is_checked()
        except Exception:
            already_checked = option_checkbox.get_attribute("aria-checked") == "true"
    if not already_checked:
        option.click()
    page.keyboard.press("Tab")
    page.wait_for_timeout(800)

    try:
        selected_text = control.input_value().strip()
    except Exception:
        selected_text = (control.inner_text() or control.get_attribute("value") or "").strip()
    if not selected_text:
        selected_text = (control.get_attribute("title") or "").strip()
    logging.info("Confirmed %s displays: %s", label, selected_text or "<unreadable>")
    normalized_selected = re.sub(r"\s+", " ", selected_text).casefold()
    if value.casefold().startswith("vci construction"):
        if "vci construction" not in normalized_selected or "blank" in normalized_selected:
            raise RuntimeError(
                f"{label} retained the wrong selection: '{selected_text}'. "
                "Expected only VCI CONSTRUCTION, LLC."
            )
    elif value.casefold() == "select all":
        selected_items = [
            item.strip()
            for item in re.split(r",\s{2,}|\n", selected_text)
            if item.strip()
        ]
        explicitly_all = "select all" in normalized_selected
        expanded_all = len(selected_items) > 2 and any(
            office in normalized_selected
            for office in ("anaheim", "baldwin", "corona", "upland")
        )
        if "required" in normalized_selected or not (explicitly_all or expanded_all):
            raise RuntimeError(
                f"{label} retained the wrong selection: '{selected_text}'. Expected Select All."
            )
        logging.info("Confirmed %s has all available locations selected", label)


def _fill_date_parameter(page: Page, label: str, value: date) -> None:
    logging.info("Setting %s = %s", label, value)
    control, _ = _parameter_control(page, label)
    control.click()
    date_text = value.strftime("%m/%d/%Y")
    control.press("Control+A")
    control.fill(date_text)
    control.press("Enter")
    control.press("Tab")
    page.wait_for_timeout(800)

    displayed = ""
    try:
        displayed = control.input_value().strip()
    except Exception:
        displayed = (control.get_attribute("value") or "").strip()
    retained_date = None
    for date_format in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            retained_date = datetime.strptime(displayed, date_format).date()
            break
        except ValueError:
            continue
    if retained_date != value:
        raise RuntimeError(
            f"{label} did not retain the requested date. "
            f"Expected {date_text}; Power BI displayed '{displayed}'."
        )
    logging.info("Confirmed %s displayed as %s", label, displayed)


def _capture_dvir_render_state(page: Page) -> None:
    debug_dir = Path(__file__).resolve().parents[1] / "logs"
    debug_dir.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(debug_dir / "dvir_rendered_report.png"), full_page=True)

    body_texts = []
    for index, frame in enumerate(page.frames):
        try:
            body = frame.locator("body").inner_text(timeout=10_000)
        except Exception as exc:
            body = f"Unable to read frame: {exc}"
        body_texts.append(f"===== FRAME {index} =====\nURL: {frame.url}\n\n{body}")
    combined = "\n\n".join(body_texts)
    (debug_dir / "dvir_rendered_report.txt").write_text(combined, encoding="utf-8")
    if re.search(r"\bno data\b|no rows|no results", combined, re.I):
        raise RuntimeError(
            "Power BI rendered no DVIR rows for the selected parameters. "
            "Review logs\\dvir_rendered_report.png before retrying."
        )


def _click_view_report(page: Page, wait_seconds: int = 20) -> None:
    button, _ = _find_across_surfaces(
        page, lambda surface: surface.get_by_role("button", name=re.compile("View report", re.I))
    )
    if not button:
        button, _ = _find_across_surfaces(
            page, lambda surface: surface.get_by_text(re.compile("View report", re.I), exact=False)
        )
    if not button:
        raise RuntimeError("View report button was not found.")
    deadline = time.time() + 60
    while time.time() < deadline and button.is_disabled():
        page.wait_for_timeout(1_000)
    if button.is_disabled():
        raise RuntimeError("View report button remained disabled.")
    button.click()
    page.wait_for_timeout(wait_seconds * 1_000)


def _configure_downloads(page: Page, staging_dir: Path) -> None:
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


def _wait_for_download(staging_dir: Path, started_at: float, seconds: int = 900) -> Path:
    deadline = time.time() + seconds
    sizes: dict[Path, int] = {}
    stable: dict[Path, int] = {}
    while time.time() < deadline:
        for candidate in [*staging_dir.glob("*.xlsx"), *staging_dir.glob("*.crdownload")]:
            if candidate.stat().st_mtime < started_at - 2:
                continue
            size = candidate.stat().st_size
            if size > 0 and sizes.get(candidate) == size:
                stable[candidate] = stable.get(candidate, 0) + 1
                if stable[candidate] >= 3 and zipfile.is_zipfile(candidate):
                    return candidate
            else:
                stable[candidate] = 0
            sizes[candidate] = size
        time.sleep(1)
    raise RuntimeError(
        f"Power BI Excel download did not complete within {seconds} seconds."
    )


def _save_download(
    page: Page,
    click_target: Locator,
    output_path: Path,
    timeout_seconds: int = 900,
) -> Path:
    staging_dir = output_path.parent / "_browser_downloads"
    _configure_downloads(page, staging_dir)
    started_at = time.time()
    click_target.click()
    downloaded = _wait_for_download(staging_dir, started_at, timeout_seconds)
    if output_path.exists():
        output_path.unlink()
    downloaded.replace(output_path)
    logging.info("Downloaded %s", output_path)
    return output_path


def export_dvir_details(
    page: Page,
    settings: dict[str, Any],
    run_date: date,
    output_dir: Path,
    authentication_wait_seconds: int,
) -> Path:
    page.goto(settings["url"], wait_until="domcontentloaded", timeout=180_000)
    _wait_for_text(page, "Company Code", authentication_wait_seconds)

    # Stage 1: Company initializes the dependent report parameters.
    _choose_parameter(
        page,
        settings["company_filter"],
        settings["company_value"],
        clear_select_all=True,
    )
    _click_view_report(page, wait_seconds=15)

    # Stage 2: set the dependent filters and render the final report.
    _choose_parameter(
        page,
        settings["office_filter"],
        settings["office_value"],
        clear_select_all=False,
    )
    _fill_date_parameter(
        page,
        settings["start_date_filter"],
        run_date - timedelta(days=int(settings.get("lookback_days", 7))),
    )
    _fill_date_parameter(page, settings["end_date_filter"], run_date)
    _click_view_report(page, wait_seconds=20)
    _capture_dvir_render_state(page)

    export_button, surface = _find_across_surfaces(
        page, lambda item: item.get_by_text(re.compile(r"^Export$", re.I), exact=False)
    )
    if not export_button:
        raise RuntimeError("DVIR Export menu was not found.")
    export_button.click()
    excel = _visible(surface.get_by_text(re.compile(r"Microsoft Excel.*xlsx", re.I)))
    if not excel:
        excel, _ = _find_across_surfaces(
            page, lambda item: item.get_by_text(re.compile(r"Microsoft Excel.*xlsx", re.I))
        )
    if not excel:
        raise RuntimeError("DVIR Microsoft Excel (.xlsx) export option was not found.")
    return _save_download(
        page,
        excel,
        output_dir / settings["output_filename"],
        int(settings.get("download_timeout_seconds", 900)),
    )


def _move_audit_date_to_latest(page: Page, label: str) -> None:
    logging.info("Moving %s end date to the latest available date", label)
    # The updated report renamed this slicer from "Audit Date" to
    # "AuditDate". Ignore spaces and punctuation so either caption works.
    compact_label = re.sub(r"[^a-z0-9]", "", label.casefold())
    label_pattern = r"^\s*" + r"\s*".join(
        re.escape(character) for character in compact_label
    ) + r"\s*$"
    label_node, surface = _find_across_surfaces(
        page,
        lambda item: item.get_by_text(re.compile(label_pattern, re.I), exact=False),
    )
    if not label_node:
        debug_dir = Path(__file__).resolve().parents[1] / "logs"
        debug_dir.mkdir(parents=True, exist_ok=True)
        page.screenshot(
            path=str(debug_dir / "safety_deviations_date_filter_not_found.png"),
            full_page=True,
        )
        raise RuntimeError(
            f"Safety Deviations date filter was not found: {label}. "
            "A diagnostic screenshot was saved in the logs folder."
        )

    # The report update separated the title and date controls into nested
    # elements. Search the whole slicer visual before falling back globally.
    container = label_node.locator(
        "xpath=ancestor::*[contains(@class,'visual-container') or "
        "contains(@class,'visualContainer')][1]"
    )
    if container.count() == 0:
        container = label_node.locator("xpath=ancestor::*[self::div][1]")
    sliders = container.get_by_role("slider")
    if sliders.count() < 2:
        sliders = surface.get_by_role("slider")
    visible_sliders = [sliders.nth(i) for i in range(sliders.count()) if sliders.nth(i).is_visible()]
    if len(visible_sliders) < 2:
        raise RuntimeError("The two Audit Date slider handles were not found.")
    end_slider = visible_sliders[-1]
    end_slider.focus()
    end_slider.press("End")
    page.wait_for_timeout(2_000)
    displayed = (
        end_slider.get_attribute("aria-valuetext")
        or end_slider.get_attribute("aria-label")
        or end_slider.get_attribute("aria-valuenow")
        or "<unreadable>"
    )
    logging.info("Confirmed %s latest endpoint displays: %s", label, displayed)


def _capture_safety_state(page: Page) -> None:
    debug_dir = Path(__file__).resolve().parents[1] / "logs"
    debug_dir.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(debug_dir / "safety_deviations_export_state.png"), full_page=True)
    body_texts = []
    for index, frame in enumerate(page.frames):
        try:
            body = frame.locator("body").inner_text(timeout=10_000)
        except Exception as exc:
            body = f"Unable to read frame: {exc}"
        body_texts.append(f"===== FRAME {index} =====\nURL: {frame.url}\n\n{body}")
    (debug_dir / "safety_deviations_export_state.txt").write_text(
        "\n\n".join(body_texts), encoding="utf-8"
    )


def _safety_detail_more_options(page: Page) -> tuple[Locator, Any]:
    """Locate More options inside the detail table, not a summary visual."""
    detail_headers = ("Deviation Question", "Auditid", "Audit ID")
    for surface in _surfaces(page):
        for header in detail_headers:
            try:
                header_node = _visible(
                    surface.get_by_text(re.compile(f"^{re.escape(header)}$", re.I))
                )
                if not header_node:
                    continue
                visual = header_node.locator(
                    "xpath=ancestor::*[contains(@class,'visual-container') or "
                    "contains(@class,'visualContainer')][1]"
                )
                if visual.count() == 0:
                    continue
                visual.hover()
                button = _visible(
                    visual.locator(
                        "button[aria-label*='More options'], "
                        "[role='button'][aria-label*='More options'], "
                        "button[title*='More options']"
                    )
                )
                if button:
                    return button, surface
            except Exception:
                continue

    # Fallback: after hovering the bottom of the report, Power BI exposes the
    # detail visual's header icons. The detail table is the last large visual.
    viewport = page.viewport_size or {"width": 1280, "height": 720}
    page.mouse.move(max(viewport["width"] - 80, 1), max(viewport["height"] - 250, 1))
    page.wait_for_timeout(1_000)
    for surface in _surfaces(page):
        try:
            buttons = surface.locator(
                "button[aria-label*='More options'], "
                "[role='button'][aria-label*='More options'], "
                "button[title*='More options']"
            )
            visible_buttons = [
                buttons.nth(i) for i in range(buttons.count()) if buttons.nth(i).is_visible()
            ]
            if visible_buttons:
                return visible_buttons[-1], surface
        except Exception:
            continue
    raise RuntimeError("Safety Deviations detail table More options button was not found.")


def export_safety_deviations(
    page: Page,
    settings: dict[str, Any],
    run_date: date,
    output_dir: Path,
    authentication_wait_seconds: int,
) -> Path:
    page.goto(settings["url"], wait_until="domcontentloaded", timeout=180_000)
    _wait_for_text(page, "Safety Deviations", authentication_wait_seconds)
    page.wait_for_timeout(8_000)
    _move_audit_date_to_latest(page, settings["date_filter"])
    page.wait_for_timeout(8_000)
    _capture_safety_state(page)

    more_options, surface = _safety_detail_more_options(page)
    more_options.click()

    export_data = _visible(surface.get_by_text(re.compile(r"^Export data$", re.I)))
    if not export_data:
        export_data, _ = _find_across_surfaces(
            page, lambda item: item.get_by_text(re.compile(r"^Export data$", re.I))
        )
    if not export_data:
        raise RuntimeError("Safety Deviations Export data option was not found.")
    export_data.click()

    current_layout, _ = _find_across_surfaces(
        page,
        lambda item: item.get_by_text(re.compile(r"Data with current layout", re.I)),
    )
    if not current_layout:
        raise RuntimeError("Data with current layout option was not found.")
    current_layout.click()

    export_button, _ = _find_across_surfaces(
        page, lambda item: item.get_by_role("button", name=re.compile(r"^Export$", re.I))
    )
    if not export_button:
        raise RuntimeError("Safety Deviations export confirmation button was not found.")

    date_text = run_date.strftime(settings.get("output_date_format", "%m-%d-%Y"))
    filename = settings["output_filename_pattern"].format(run_date=date_text)
    return _save_download(
        page,
        export_button,
        output_dir / filename,
        int(settings.get("download_timeout_seconds", 300)),
    )


def upload_to_drive(page: Page, drive: dict[str, Any], files: list[Path]) -> None:
    page.goto(drive["folder_url"], wait_until="domcontentloaded", timeout=180_000)
    _wait_for_text(page, drive["folder_name"], 300)
    page.wait_for_timeout(3_000)

    # Drive commonly hides .xlsx extensions and permits duplicate filenames.
    # Inspect the current folder first and upload only genuinely missing files.
    folder_text = page.locator("body").inner_text(timeout=30_000).casefold()
    missing_files = [path for path in files if path.stem.casefold() not in folder_text]
    existing_files = [path for path in files if path not in missing_files]
    for path in existing_files:
        logging.info("Skipping Drive upload because the file already exists: %s", path.name)
    if not missing_files:
        logging.info("Both validated files already exist in %s; no upload was performed", drive["folder_name"])
        return

    new_button = _visible(page.get_by_role("button", name=re.compile(r"^New$", re.I)))
    if not new_button:
        new_button = _visible(page.get_by_text(re.compile(r"^New$", re.I)))
    if not new_button:
        raise RuntimeError("Google Drive New button was not found.")
    new_button.click()
    upload_item = _visible(page.get_by_role("menuitem", name=re.compile(r"^File upload$", re.I)))
    if not upload_item:
        upload_text = _visible(
            page.get_by_text(re.compile(r"^File upload$", re.I), exact=False)
        )
        if upload_text:
            upload_item = upload_text.locator("xpath=ancestor::*[@role='menuitem'][1]")
    if not upload_item:
        raise RuntimeError("Google Drive File upload menu item was not found.")
    with page.expect_file_chooser() as chooser:
        upload_item.click()
    chooser.value.set_files([str(path.resolve()) for path in missing_files])

    # If the intake folder already contains either exact filename, replace it.
    # Keeping both would make the downstream exact-name lookup ambiguous.
    conflict_deadline = time.time() + 8
    while time.time() < conflict_deadline:
        conflict = _visible(
            page.get_by_text(
                re.compile(r"already exists|same name|replace existing", re.I),
                exact=False,
            )
        )
        if conflict:
            replace_button = _visible(
                page.get_by_role(
                    "button", name=re.compile(r"replace(?: all)?|update", re.I)
                )
            )
            if not replace_button:
                replace_button = _visible(
                    page.get_by_text(
                        re.compile(r"^replace(?: all)?$|^update$", re.I), exact=False
                    )
                )
            if not replace_button:
                raise RuntimeError(
                    "Google Drive reported an existing filename, but its Replace option "
                    "could not be found. No completion marker was written."
                )
            replace_button.click()
            logging.info("Replacing existing Drive file(s) with the validated export pair")
            break
        page.wait_for_timeout(500)

    page.get_by_text(re.compile(r"upload complete|uploads complete", re.I)).wait_for(
        state="visible", timeout=240_000
    )
    logging.info(
        "Google Drive confirmed upload complete for %d validated files in %s",
        len(missing_files),
        drive["folder_name"],
    )
