"""
===============================================================================
Script: Leadership JSO Report - Bulk Email Automation
===============================================================================

Description:
Automates extracting manager email addresses from 'Managers Info.pbix' (or
reading the most recent exported 'managers-contact-info-*.csv' file) and
distributing the 'Leadership JSO.pdf' report via Google Workspace Gmail in Chrome.
Embeds a clean, descriptive hyperlink to the VCI Portal Safety page along with
an inline visual reference screenshot directly inside the email body.

Workflow Outline:
1. PBIX Table Export (Optional): When RUN_PBIX_EXPORT = True, runs
   export_managers_contact_csv() to export 'Manager Contact Information' from
   'W:/Dashboards VCI/LJSO/Managers/Managers Info.pbix' into
   'W:/Dashboards VCI/LJSO/Managers/managers-contact-info-YYYY-MM-DD.csv'.
   When False, loads the most recently modified matching CSV in that folder.
2. Data Ingestion: Reads the 'Email Address' column (handling Power BI CSV
   encodings) and preserves all manager entries.
3. File Verification: Confirms 'W:/Dashboards VCI/LJSO/Main/Leadership JSO.pdf'
   exists on disk before launching the browser.
4. Browser Automation: Launches Google Chrome using a persistent local profile
   (retaining Gmail login sessions) and navigates to mail.google.com.
5. Rich Email Composition & Dispatch:
   - TEST MODE (TEST_MODE = True): Generates distinct '+alias' BCC addresses
     ('jo.deleon+1@vcicom.com' .. '+N') so Gmail creates N valid BCC chips
     while routing solely to Jo's inbox, keeps leadership CCs (Tom, Christian,
     Jo) so the team can inspect the test email, and prefixes '[TEST]'.
   - LIVE MODE (TEST_MODE = False): Places all unique eligible manager emails
     in BCC and leadership (Tom, Christian, Jo) in CC.
   - Enters each recipient one-by-one and presses Enter after every email so
     every address converts cleanly into a Gmail chip.
   - Builds the message body using Trusted-Types-safe DOM nodes with a
     descriptive hyperlink on 'VCI Portal Safety Page' (pointing to the Safety
     page under 'Safety Dashboard > Leadership JSOs'), embeds the visual
     reference screenshot inline, attaches 'Leadership JSO.pdf', sends, and
     confirms the "Message sent" toast.
===============================================================================
"""
from __future__ import annotations

import base64
import ctypes
from ctypes import wintypes
import glob
import logging
import mimetypes
import os
import time
from datetime import datetime
from pathlib import Path
import pandas as pd
from playwright.sync_api import Page, sync_playwright

from export_managers_info import export_managers_contact_csv

# ==========================================
# CONFIGURATION & PATHS
# ==========================================
TEST_MODE = True         # True = send 1 test email with '+alias' BCC chips to TEST_RECIPIENT + team CCs; False = live send
RUN_PBIX_EXPORT = True   # True = export fresh CSV from Power BI; False = use latest existing CSV on disk
TEST_RECIPIENT = "jo.deleon@vcicom.com"

BASE_DIR = r"C:\Users\jd0855\Desktop\VCI_Automation\bulk-send-emails-ljso"
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profiles", "gmail")

MANAGERS_FOLDER = r"W:\Dashboards VCI\LJSO\Managers"
PDF_PATH = r"W:\Dashboards VCI\LJSO\Main\Leadership JSO.pdf"

PORTAL_SAFETY_URL = "https://sites.google.com/vcicom.com/vci/safety?authuser=0"
PORTAL_REF_IMAGE_PATH = os.path.join(BASE_DIR, "vci_safety_portal_ref.jpg")

STATIC_CCS = [
    "tom.thai@vcicom.com",
    "christian.amaya@vcicom.com",
    "jo.deleon@vcicom.com",
]


# ==========================================
# DATA RETRIEVAL
# ==========================================
def get_latest_managers_file(folder_path: str) -> str:
    search_pattern = os.path.join(folder_path, "managers-contact-info-*")
    matching_files = [
        f for f in glob.glob(search_pattern)
        if f.lower().endswith((".csv", ".xlsx", ".xls"))
    ]
    if not matching_files:
        raise FileNotFoundError(
            f"No file matching 'managers-contact-info-*' found in {folder_path}"
        )
    return max(matching_files, key=os.path.getmtime)


def load_manager_emails(file_path: str) -> list[str]:
    print(f"Loading manager contact info from: {file_path}")

    if file_path.lower().endswith(".csv"):
        try:
            df = pd.read_csv(file_path, encoding="utf-8-sig")
        except (UnicodeDecodeError, UnicodeError):
            df = pd.read_csv(file_path, encoding="utf-16le")
    else:
        df = pd.read_excel(file_path)

    df.columns = df.columns.str.strip()

    email_col = "Email Address"
    if email_col not in df.columns:
        raise KeyError(
            f"Column '{email_col}' not found. Available columns: {list(df.columns)}"
        )

    df = df.dropna(subset=[email_col]).copy()
    df[email_col] = df[email_col].astype(str).str.strip()
    df = df[df[email_col].str.contains("@", na=False)]

    # Deduplicate by (Manager Name, Email Address) if a name column exists so
    # all manager rows are preserved even if test rows share an address.
    name_col = "Manager Name" if "Manager Name" in df.columns else (
        "Dycom Leader Name" if "Dycom Leader Name" in df.columns else None
    )
    if name_col:
        df[name_col] = df[name_col].fillna("").astype(str).str.strip()
        df = df.drop_duplicates(subset=[name_col, email_col])
    else:
        df = df.drop_duplicates()

    return df[email_col].tolist()


def build_test_bcc_aliases(base_email: str, count: int) -> list[str]:
    """
    Expands 'jo.deleon@vcicom.com' into ['jo.deleon+1@vcicom.com', ..., '+N']
    so Gmail creates N distinct valid BCC chips that all deliver to `base_email`.
    """
    total = max(1, count)
    local, domain = base_email.split("@", 1)
    return [f"{local}+{i}@{domain}" for i in range(1, total + 1)]


# ==========================================
# INLINE IMAGE & GMAIL COMPOSE HELPERS
# ==========================================
def _resolve_portal_image_path(configured_path: str) -> str | None:
    """Finds the VCI Safety portal reference image even if saved as .jpg, .jpeg, or .png."""
    p = Path(configured_path)
    if p.exists() and p.is_file():
        return str(p)
    for ext in (".jpg", ".jpeg", ".png"):
        candidate = p.with_suffix(ext)
        if candidate.exists() and candidate.is_file():
            return str(candidate)
    return None


def _copy_image_to_win32_clipboard(image_path: str) -> bool:
    """
    Loads a JPG/PNG/BMP image via built-in Windows GDI+ and places it onto
    the Windows clipboard as CF_DIB (64-bit pointer safe) for Ctrl+V fallback.
    """
    try:
        gdiplus = ctypes.WinDLL("gdiplus")
        gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

        class GdiplusStartupInput(ctypes.Structure):
            _fields_ = [
                ("GdiplusVersion", wintypes.UINT),
                ("DebugEventCallback", ctypes.c_void_p),
                ("SuppressBackgroundThread", wintypes.BOOL),
                ("SuppressExternalCodecs", wintypes.BOOL),
            ]

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

        class BITMAP(ctypes.Structure):
            _fields_ = [
                ("bmType", wintypes.LONG),
                ("bmWidth", wintypes.LONG),
                ("bmHeight", wintypes.LONG),
                ("bmWidthBytes", wintypes.LONG),
                ("bmPlanes", wintypes.WORD),
                ("bmBitsPixel", wintypes.WORD),
                ("bmBits", ctypes.c_void_p),
            ]

        token = ctypes.c_ulong()
        startup_input = GdiplusStartupInput(1, None, False, False)
        if gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(startup_input), None) != 0:
            return False

        try:
            gp_bitmap = ctypes.c_void_p()
            if gdiplus.GdipCreateBitmapFromFile(ctypes.c_wchar_p(image_path), ctypes.byref(gp_bitmap)) != 0:
                return False

            hbitmap = wintypes.HBITMAP()
            try:
                if gdiplus.GdipCreateHBITMAPFromBitmap(gp_bitmap, ctypes.byref(hbitmap), 0xFFFFFFFF) != 0:
                    return False
            finally:
                gdiplus.GdipDisposeImage(gp_bitmap)

            try:
                bmp = BITMAP()
                gdi32.GetObjectW.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p]
                gdi32.GetObjectW.restype = ctypes.c_int
                if not gdi32.GetObjectW(hbitmap, ctypes.sizeof(BITMAP), ctypes.byref(bmp)):
                    return False

                width, height = bmp.bmWidth, bmp.bmHeight
                row_bytes = ((width * 24 + 31) // 32) * 4
                img_bytes = row_bytes * height

                bmi = BITMAPINFOHEADER()
                bmi.biSize = ctypes.sizeof(BITMAPINFOHEADER)
                bmi.biWidth = width
                bmi.biHeight = height
                bmi.biPlanes = 1
                bmi.biBitCount = 24
                bmi.biCompression = 0  # BI_RGB
                bmi.biSizeImage = img_bytes

                total_size = ctypes.sizeof(BITMAPINFOHEADER) + img_bytes
                GMEM_MOVEABLE = 0x0002
                CF_DIB = 8

                kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
                kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
                kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
                kernel32.GlobalLock.restype = wintypes.LPVOID
                kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
                kernel32.GlobalUnlock.restype = wintypes.BOOL

                h_mem = kernel32.GlobalAlloc(GMEM_MOVEABLE, total_size)
                if not h_mem:
                    return False

                ptr = kernel32.GlobalLock(h_mem)
                if not ptr:
                    return False

                try:
                    ctypes.memmove(ptr, ctypes.byref(bmi), ctypes.sizeof(BITMAPINFOHEADER))
                    hdc = user32.GetDC(0)
                    try:
                        gdi32.GetDIBits.argtypes = [
                            wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT,
                            ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT
                        ]
                        gdi32.GetDIBits.restype = ctypes.c_int
                        lines = gdi32.GetDIBits(
                            hdc, hbitmap, 0, height,
                            ctypes.c_void_p(ptr + ctypes.sizeof(BITMAPINFOHEADER)),
                            ctypes.byref(bmi), 0
                        )
                        if lines != height:
                            return False
                    finally:
                        user32.ReleaseDC(0, hdc)
                finally:
                    kernel32.GlobalUnlock(h_mem)

                user32.OpenClipboard.argtypes = [wintypes.HWND]
                user32.OpenClipboard.restype = wintypes.BOOL
                user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
                user32.SetClipboardData.restype = wintypes.HANDLE

                if not user32.OpenClipboard(None):
                    return False
                try:
                    user32.EmptyClipboard()
                    if not user32.SetClipboardData(CF_DIB, h_mem):
                        return False
                finally:
                    user32.CloseClipboard()

                return True
            finally:
                gdi32.DeleteObject(hbitmap)
        finally:
            gdiplus.GdiplusShutdown(token)
    except Exception:
        return False


def compose_email_body_with_portal_ref(
    page: Page,
    body_box,
    greeting_line: str,
    today_str: str,
    portal_image_path: str | None,
) -> None:
    """
    Populates Gmail's rich-text Message Body using Trusted-Types-safe DOM
    creation (TrustedHTML-compliant):
      1. Greeting & report intro text
      2. Descriptive inline hyperlink on 'VCI Portal Safety Page'
         (under Safety Dashboard > Leadership JSOs)
      3. Inline visual reference image (when present on disk)
      4. Closing signature
    """
    body_box.click()
    time.sleep(0.2)

    body_box.evaluate(
        """(el, cfg) => {
            el.focus();
            el.replaceChildren();

            const addLine = (nodes) => {
                const div = document.createElement('div');
                if (!nodes || nodes.length === 0) {
                    div.appendChild(document.createElement('br'));
                } else {
                    for (const n of nodes) {
                        div.appendChild(n);
                    }
                }
                el.appendChild(div);
            };

            addLine([document.createTextNode(cfg.greeting)]);
            addLine([]);
            addLine([document.createTextNode(`Please see attached Leadership JSO report as of ${cfg.today}.`)]);
            addLine([]);

            const linkBold = document.createElement('b');
            linkBold.textContent = 'VCI Portal Safety Page';

            const link = document.createElement('a');
            link.href = cfg.portalUrl;
            link.appendChild(linkBold);

            const navPath = document.createElement('b');
            navPath.textContent = 'Safety Dashboard > Leadership JSOs';

            addLine([
                document.createTextNode('You can also access and download this report anytime via the '),
                link,
                document.createTextNode(' (under '),
                navPath,
                document.createTextNode('):'),
            ]);
            addLine([]);

            const range = document.createRange();
            range.selectNodeContents(el);
            range.collapse(false);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }""",
        {
            "greeting": greeting_line,
            "today": today_str,
            "portalUrl": PORTAL_SAFETY_URL,
        },
    )
    time.sleep(0.3)

    resolved_img = _resolve_portal_image_path(portal_image_path) if portal_image_path else None
    if resolved_img:
        print(f"Embedding inline portal visual reference: {resolved_img}")
        mime_type = mimetypes.guess_type(resolved_img)[0] or "image/jpeg"
        b64_data = base64.b64encode(Path(resolved_img).read_bytes()).decode("ascii")
        img_filename = os.path.basename(resolved_img)

        # Method 1: Dispatch synthetic paste event with image File blob directly into Message Body
        body_box.evaluate(
            """(el, payload) => {
                el.focus();
                const binStr = atob(payload.b64);
                const bytes = new Uint8Array(binStr.length);
                for (let i = 0; i < binStr.length; i++) {
                    bytes[i] = binStr.charCodeAt(i);
                }
                const file = new File([bytes], payload.filename, { type: payload.mime });
                const dt = new DataTransfer();
                dt.items.add(file);
                const pasteEvt = new ClipboardEvent('paste', {
                    clipboardData: dt,
                    bubbles: true,
                    cancelable: true
                });
                el.dispatchEvent(pasteEvt);
            }""",
            {"b64": b64_data, "filename": img_filename, "mime": mime_type},
        )
        time.sleep(1.5)

        # Method 2 Fallback: Copy CF_DIB to Windows clipboard & press Ctrl+V
        if body_box.locator("img").count() == 0:
            if _copy_image_to_win32_clipboard(resolved_img):
                body_box.click()
                page.keyboard.press("Control+End")
                page.keyboard.press("Control+v")
                time.sleep(1.5)

        inline_img = body_box.locator("img").first
        inline_img.wait_for(state="visible", timeout=20000)
        inline_img.evaluate(
            """img => {
                img.style.maxWidth = '650px';
                img.style.height = 'auto';
                img.style.border = '1px solid #d0d7de';
            }"""
        )
        time.sleep(0.5)
    else:
        print(
            f"[WARNING] Portal reference image not found at '{portal_image_path}'. "
            f"Save the screenshot there to embed it inline."
        )

    # Append closing signature using Trusted-Types-safe DOM creation
    body_box.evaluate(
        """el => {
            el.focus();

            const spacer = document.createElement('div');
            spacer.appendChild(document.createElement('br'));
            el.appendChild(spacer);

            const sig = document.createElement('div');
            sig.appendChild(document.createTextNode('Best regards,'));
            sig.appendChild(document.createElement('br'));
            sig.appendChild(document.createTextNode('Leadership JSO - Automation'));
            el.appendChild(sig);
        }"""
    )
    time.sleep(0.3)


def _find_visible_recipient_input(scope, field_label: str):
    """Returns a visible input/combobox/textbox matching To, Cc, or Bcc inside `scope`."""
    selectors = [
        f'input[aria-label="{field_label}"]:visible',
        f'input[aria-label="{field_label} recipients"]:visible',
        f'input[aria-label^="{field_label}"]:visible',
        f'input[aria-label*="{field_label}"]:visible',
        f'div[aria-label="{field_label}"][role="textbox"]:visible',
        f'div[aria-label^="{field_label}"][role="textbox"]:visible',
    ]
    for sel in selectors:
        loc = scope.locator(sel).first
        try:
            if loc.count() > 0 and loc.is_visible():
                return loc
        except Exception:
            continue
    return None


def expand_and_fill_gmail_field(page: Page, field_label: str, recipients: list[str]) -> None:
    """
    Expands To/Cc/Bcc in Gmail's Compose header, deduplicates addresses, and
    enters each email one-by-one followed by Enter so Gmail tokenizes every
    address into its own recipient chip cleanly.
    """
    if not recipients:
        return

    seen: set[str] = set()
    unique_recipients: list[str] = []
    for raw_email in recipients:
        cleaned = raw_email.strip()
        key = cleaned.casefold()
        if cleaned and key not in seen:
            seen.add(key)
            unique_recipients.append(cleaned)

    if not unique_recipients:
        return

    compose_dialogs = page.locator('div[role="dialog"]:visible')
    scope = compose_dialogs.last if compose_dialogs.count() > 0 else page

    collapsed = scope.locator(':text-is("Recipients"):visible').first
    try:
        if collapsed.count() > 0 and collapsed.is_visible():
            collapsed.click()
            time.sleep(0.3)
    except Exception:
        pass

    target_input = _find_visible_recipient_input(scope, field_label)

    if field_label in ("Cc", "Bcc") and target_input is None:
        to_input = _find_visible_recipient_input(scope, "To")
        if to_input is not None:
            to_input.click()
            time.sleep(0.3)

        toggle = scope.locator(
            f'span:text-is("{field_label}"):visible, '
            f'button:text-is("{field_label}"):visible, '
            f'a:text-is("{field_label}"):visible'
        ).first
        if toggle.count() > 0 and toggle.is_visible():
            toggle.click()
            time.sleep(0.4)
            target_input = _find_visible_recipient_input(scope, field_label)

    if field_label in ("Cc", "Bcc") and target_input is None:
        if field_label == "Bcc":
            page.keyboard.press("Control+Shift+b")
            time.sleep(0.4)
        else:
            cc_link = scope.locator('span:text-is("Cc"), button:text-is("Cc")').last
            if cc_link.count() > 0:
                cc_link.evaluate("el => el.click()")
                time.sleep(0.4)
        target_input = _find_visible_recipient_input(scope, field_label)

    if target_input is not None:
        target_input.click()
        time.sleep(0.2)
        for email in unique_recipients:
            target_input.fill(email)
            target_input.press("Enter")
            time.sleep(0.06)
    else:
        for email in unique_recipients:
            page.keyboard.insert_text(email)
            page.keyboard.press("Enter")
            time.sleep(0.06)

    time.sleep(0.3)


# ==========================================
# MAIN AUTOMATION
# ==========================================
def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    if not os.path.exists(PDF_PATH):
        raise FileNotFoundError(f"Attachment not found at: {PDF_PATH}")

    if RUN_PBIX_EXPORT:
        exported_path = str(export_managers_contact_csv())
    else:
        exported_path = get_latest_managers_file(MANAGERS_FOLDER)

    os.makedirs(PROFILE_DIR, exist_ok=True)
    manager_emails = load_manager_emails(exported_path)
    today_str = datetime.now().strftime("%B %d, %Y")

    if TEST_MODE:
        print(f"[TEST MODE] Found {len(manager_emails)} manager entries in export.")
        bcc_list = build_test_bcc_aliases(TEST_RECIPIENT, len(manager_emails))
        print(
            f"[TEST MODE] Generated {len(bcc_list)} distinct '+alias' BCC addresses "
            f"(e.g., {bcc_list[0]} .. {bcc_list[-1]}) routing to {TEST_RECIPIENT}, plus team CCs."
        )
        cc_list = STATIC_CCS
        subject = f"[TEST] Leadership JSO Report – {today_str}"
    else:
        cc_list = STATIC_CCS
        bcc_list = sorted(set(manager_emails))
        subject = f"Leadership JSO Report – {today_str}"

    if not bcc_list:
        print("No eligible recipients found.")
        return

    print(f"Launching Chrome. Preparing to send with {len(cc_list)} CC and {len(bcc_list)} BCC recipient(s)...")

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            channel="chrome",
            headless=False,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--start-maximized"],
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://mail.google.com")

        print("Waiting for Gmail... (Log in if prompted)")
        compose_btn = page.locator(
            'div[role="button"][gh="cm"], div[role="button"]:has-text("Compose")'
        )
        compose_btn.first.wait_for(state="visible", timeout=180000)

        # Open Compose window
        compose_btn.first.click()
        page.wait_for_selector('input[name="subjectbox"]', timeout=15000)
        time.sleep(0.8)

        # Fill CC and BCC (pressing Enter after every single email)
        expand_and_fill_gmail_field(page, "Cc", cc_list)
        expand_and_fill_gmail_field(page, "Bcc", bcc_list)

        # Fill Subject
        page.locator('input[name="subjectbox"]').fill(subject)

        # Fill Body with descriptive VCI Portal Safety hyperlink & inline visual reference
        body_box = page.locator('div[role="textbox"][aria-label="Message Body"]').first
        compose_email_body_with_portal_ref(
            page=page,
            body_box=body_box,
            greeting_line="Dear Team,",
            today_str=today_str,
            portal_image_path=PORTAL_REF_IMAGE_PATH,
        )

        # Attach PDF and wait for upload chip to finish
        print("Attaching report...")
        pdf_name = os.path.basename(PDF_PATH)
        page.locator('input[type="file"]').first.set_input_files(PDF_PATH)
        page.locator(f'div[aria-label*="{pdf_name}"], a[aria-label*="{pdf_name}"]').first.wait_for(
            state="visible", timeout=30000
        )
        time.sleep(1.5)

        # Send (Ctrl + Enter)
        print("Sending email...")
        body_box.click()
        page.keyboard.press("Control+Enter")

        # Check immediately if Gmail raised an address-validation modal
        time.sleep(1.0)
        alert_dlg = page.locator('div[role="alertdialog"]:visible').first
        if alert_dlg.count() > 0 and alert_dlg.is_visible():
            raise RuntimeError(f"Gmail rejected sending: {alert_dlg.inner_text()}")

        # Wait for confirmation toast
        page.locator('span:has-text("Message sent")').first.wait_for(
            state="visible", timeout=20000
        )
        print("Email sent successfully!")

        time.sleep(2)
        context.close()


if __name__ == "__main__":
    main()