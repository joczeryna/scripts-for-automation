from __future__ import annotations

import base64
import ctypes
import json
import logging
import os
import socket
import struct
import time
from pathlib import Path
from urllib.parse import urlparse


GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"

LOGIN_TIMEOUT_SECONDS = 15 * 60
LOGIN_CHECK_INTERVAL_SECONDS = 2.0


class AuthenticationError(RuntimeError):
    pass


def _bring_chrome_to_front() -> None:
    """Bring an open Google Chrome window to the foreground on Windows so prompts are visible."""
    if os.name != "nt":
        return
    try:
        user32 = ctypes.windll.user32
        matches: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def enum_cb(hwnd, _lparam):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buf = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buf, length + 1)
                    if "Google Chrome" in buf.value:
                        matches.append(hwnd)
            return True

        user32.EnumWindows(enum_cb, 0)
        if matches:
            hwnd = matches[0]
            if user32.IsIconic(hwnd):
                user32.ShowWindow(hwnd, 9)
            user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def _get_devtools_ws_url(cdp_endpoint: str | None = None) -> str:
    if cdp_endpoint and cdp_endpoint != "auto":
        return cdp_endpoint

    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise AuthenticationError("LOCALAPPDATA environment variable not found.")

    port_file = (
        Path(local_app_data)
        / "Google"
        / "Chrome"
        / "User Data"
        / "DevToolsActivePort"
    )
    if not port_file.exists():
        raise AuthenticationError(
            "Chrome remote debugging is not active. In Chrome, open "
            "chrome://inspect/#remote-debugging and check "
            "'Allow remote debugging for this browser instance'."
        )

    lines = port_file.read_text(encoding="utf-8").strip().splitlines()
    if len(lines) < 2:
        raise AuthenticationError("Chrome DevToolsActivePort file is incomplete.")

    port, ws_path = lines[0].strip(), lines[1].strip()
    return f"ws://127.0.0.1:{port}{ws_path}"


class _DirectCdpClient:
    """
    Lightweight local CDP WebSocket client that attaches ONLY to the single
    tab it creates—never stalling on sleeping tabs or Chrome extensions.
    """

    def __init__(self, ws_url: str, timeout: float = 45.0) -> None:
        parsed = urlparse(ws_url)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 9222
        path = parsed.path or "/"
        if parsed.query:
            path = f"{path}?{parsed.query}"

        self._next_id = 0
        self._buf = bytearray()
        self._sock = socket.create_connection((host, port), timeout=timeout)
        self._sock.settimeout(timeout)

        ws_key = base64.b64encode(os.urandom(16)).decode("ascii")
        handshake = (
            f"GET {path} HTTP/1.1\r\n"
            f"Host: {host}:{port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {ws_key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        ).encode("ascii")
        self._sock.sendall(handshake)

        while b"\r\n\r\n" not in self._buf:
            chunk = self._sock.recv(4096)
            if not chunk:
                raise AuthenticationError("Chrome closed the connection during handshake.")
            self._buf.extend(chunk)

        header_end = self._buf.index(b"\r\n\r\n") + 4
        header_text = bytes(self._buf[:header_end]).decode("latin-1")
        del self._buf[:header_end]

        if "101" not in header_text.splitlines()[0]:
            raise AuthenticationError(
                f"Unexpected CDP handshake response: {header_text.splitlines()[0]}"
            )

    def _recv_exact(self, n: int) -> bytes:
        while len(self._buf) < n:
            chunk = self._sock.recv(65536)
            if not chunk:
                raise ConnectionError("Chrome closed the CDP WebSocket.")
            self._buf.extend(chunk)
        out = bytes(self._buf[:n])
        del self._buf[:n]
        return out

    def _send_frame(self, payload: bytes, opcode: int = 0x1) -> None:
        length = len(payload)
        first_byte = 0x80 | opcode
        if length < 126:
            header = struct.pack("!BB", first_byte, 0x80 | length)
        elif length <= 65535:
            header = struct.pack("!BBH", first_byte, 0x80 | 126, length)
        else:
            header = struct.pack("!BBQ", first_byte, 0x80 | 127, length)

        mask = os.urandom(4)
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
        self._sock.sendall(header + mask + masked)

    def _recv_text(self) -> str:
        fragments: list[bytes] = []
        while True:
            b1, b2 = self._recv_exact(2)
            fin = (b1 & 0x80) != 0
            opcode = b1 & 0x0F
            masked = (b2 & 0x80) != 0
            length = b2 & 0x7F

            if length == 126:
                length = struct.unpack("!H", self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack("!Q", self._recv_exact(8))[0]

            mask_key = self._recv_exact(4) if masked else b""
            data = self._recv_exact(length)
            if masked:
                data = bytes(b ^ mask_key[i % 4] for i, b in enumerate(data))

            if opcode == 0x8:
                raise ConnectionError("Chrome closed the CDP WebSocket.")
            if opcode == 0x9:
                self._send_frame(data, opcode=0xA)
                continue
            if opcode == 0xA:
                continue

            fragments.append(data)
            if fin:
                return b"".join(fragments).decode("utf-8")

    def send(
        self,
        method: str,
        params: dict | None = None,
        session_id: str | None = None,
    ) -> dict:
        self._next_id += 1
        cmd_id = self._next_id
        msg: dict = {"id": cmd_id, "method": method}
        if params:
            msg["params"] = params
        if session_id:
            msg["sessionId"] = session_id

        self._send_frame(json.dumps(msg).encode("utf-8"))

        while True:
            raw = self._recv_text()
            reply = json.loads(raw)
            if reply.get("id") == cmd_id:
                if "error" in reply:
                    raise RuntimeError(f"CDP {method} failed: {reply['error']}")
                return reply.get("result", {})

    def close(self) -> None:
        try:
            self._send_frame(b"", opcode=0x8)
        except Exception:
            pass
        try:
            self._sock.close()
        except Exception:
            pass


def _build_xlsx_export_url(resolved_url: str, spreadsheet_id: str) -> str:
    parsed = urlparse(resolved_url)
    prefix = (
        parsed.path.split("/d/")[0]
        if "/spreadsheets/u/" in parsed.path
        else "/spreadsheets"
    )
    return f"https://docs.google.com{prefix}/d/{spreadsheet_id}/export?format=xlsx"


def _get_tab_state(cdp: _DirectCdpClient, session_id: str) -> tuple[str, str]:
    eval_res = cdp.send(
        "Runtime.evaluate",
        {
            "expression": "JSON.stringify({href: window.location.href, ready: document.readyState})",
            "returnByValue": True,
        },
        session_id=session_id,
    )
    raw_state = eval_res.get("result", {}).get("value")
    if not raw_state:
        return "", ""
    state = json.loads(raw_state)
    return state.get("href", ""), state.get("ready", "")


def setup_google_login(
    spreadsheet_id: str,
    cdp_endpoint: str | None = None,
) -> None:
    """
    Open a single tab in the user's running Chrome so they can sign in to Google.
    Closes only that tab once access is confirmed.
    """
    ws_url = _get_devtools_ws_url(cdp_endpoint)
    sheet_url = GOOGLE_SHEET_URL.format(spreadsheet_id=spreadsheet_id)

    _bring_chrome_to_front()
    print("Attaching to Chrome — click 'Allow' in Chrome if prompted...")
    try:
        cdp = _DirectCdpClient(ws_url, timeout=60.0)
    except Exception as exc:
        raise AuthenticationError(
            f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
            "debugging' is checked at chrome://inspect/#remote-debugging and "
            "click Allow if Chrome prompts you."
        ) from exc

    target_id = None
    try:
        try:
            created = cdp.send("Target.createTarget", {"url": "about:blank"})
        except Exception as exc:
            raise AuthenticationError(
                f"Chrome did not approve remote debugging ({exc}). Click 'Allow' in Chrome and retry."
            ) from exc

        target_id = created["targetId"]
        attached = cdp.send(
            "Target.attachToTarget",
            {"targetId": target_id, "flatten": True},
        )
        session_id = attached["sessionId"]

        cdp.send("Page.enable", session_id=session_id)
        cdp.send("Page.bringToFront", session_id=session_id)
        cdp.send("Page.navigate", {"url": sheet_url}, session_id=session_id)

        print()
        print("A new tab is now open in your Chrome window.")
        print(
            "Sign in with the authorized VCI Google account and "
            "confirm that the spreadsheet opens."
        )
        print(
            "This tab will close automatically after Google "
            "confirms access to the spreadsheet."
        )

        deadline = time.monotonic() + LOGIN_TIMEOUT_SECONDS
        while time.monotonic() < deadline:
            try:
                current_url, ready = _get_tab_state(cdp, session_id)
                if (
                    spreadsheet_id in current_url
                    and "accounts.google.com" not in current_url
                    and ready in ("interactive", "complete")
                ):
                    print("Google sign-in confirmed. Closing setup tab...")
                    return
            except Exception:
                # Tab may be mid-redirect during Okta/Google SSO; keep polling.
                pass

            time.sleep(LOGIN_CHECK_INTERVAL_SECONDS)

        raise TimeoutError(
            "Google sign-in was not confirmed within 15 minutes. "
            "Run --setup-login again and complete the sign-in."
        )
    finally:
        if target_id:
            try:
                cdp.send("Target.closeTarget", {"targetId": target_id})
            except Exception:
                pass
        cdp.close()


def download_reports(
    reports: list[dict],
    download_folder: Path,
    cdp_endpoint: str | None = None,
) -> list[dict]:
    """
    Export configured Google Sheets as Excel workbooks through a single tab
    in the user's running Chrome browser session.
    """
    download_folder.mkdir(parents=True, exist_ok=True)
    resolved_download_dir = str(download_folder.resolve())
    ws_url = _get_devtools_ws_url(cdp_endpoint)

    _bring_chrome_to_front()
    print("Attaching to Chrome — click 'Allow' in Chrome if prompted...")
    try:
        cdp = _DirectCdpClient(ws_url, timeout=45.0)
    except Exception as exc:
        raise AuthenticationError(
            f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
            "debugging' is checked at chrome://inspect/#remote-debugging and "
            "click Allow if Chrome prompts you."
        ) from exc

    download_results = []
    target_id = None

    try:
        try:
            created = cdp.send("Target.createTarget", {"url": "about:blank"})
        except Exception as exc:
            raise AuthenticationError(
                f"Chrome did not approve remote debugging ({exc}). Click 'Allow' in Chrome and retry."
            ) from exc

        target_id = created["targetId"]
        attached = cdp.send(
            "Target.attachToTarget",
            {"targetId": target_id, "flatten": True},
        )
        session_id = attached["sessionId"]

        cdp.send("Page.enable", session_id=session_id)
        cdp.send("Page.bringToFront", session_id=session_id)

        # Direct downloads from this Chrome session into the temporary download folder
        try:
            cdp.send(
                "Browser.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": resolved_download_dir,
                },
            )
        except Exception:
            cdp.send(
                "Page.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": resolved_download_dir,
                },
                session_id=session_id,
            )

        # Open the first spreadsheet to confirm active Google/Okta SSO and detect account prefix
        first_sheet_id = reports[0]["spreadsheet_id"]
        first_sheet_url = GOOGLE_SHEET_URL.format(spreadsheet_id=first_sheet_id)
        logging.info("Opening %s in active Chrome window.", first_sheet_url)
        cdp.send("Page.navigate", {"url": first_sheet_url}, session_id=session_id)

        deadline = time.monotonic() + 25.0
        current_url = ""
        while time.monotonic() < deadline:
            current_url, ready = _get_tab_state(cdp, session_id)
            if first_sheet_id in current_url and ready in ("interactive", "complete"):
                break
            time.sleep(0.5)

        if "accounts.google.com" in current_url or first_sheet_id not in current_url:
            raise AuthenticationError(
                "Your Chrome Google session has expired. Sign in via Chrome and retry."
            )

        for report in reports:
            report_name = report["name"]
            spreadsheet_id = report["spreadsheet_id"]
            temporary_filename = (
                report_name.lower().replace(" ", "_") + "_download.xlsx"
            )
            destination_path = download_folder / temporary_filename
            if destination_path.exists():
                destination_path.unlink()

            export_url = _build_xlsx_export_url(current_url, spreadsheet_id)
            existing_files = {p.name for p in download_folder.iterdir()}

            print(f"Downloading {report_name}...")
            cdp.send(
                "Runtime.evaluate",
                {
                    "expression": f"window.location.assign({json.dumps(export_url)});",
                    "returnByValue": True,
                },
                session_id=session_id,
            )

            dl_deadline = time.monotonic() + 180.0
            downloaded_candidate: Path | None = None

            while time.monotonic() < dl_deadline:
                new_completed = [
                    p
                    for p in download_folder.iterdir()
                    if p.is_file()
                    and p.name not in existing_files
                    and not p.name.endswith((".crdownload", ".tmp"))
                    and not p.name.startswith(".")
                ]
                if new_completed:
                    candidate = max(
                        new_completed, key=lambda p: p.stat().st_mtime
                    )
                    size_1 = candidate.stat().st_size
                    if size_1 > 0:
                        time.sleep(0.3)
                        if (
                            candidate.exists()
                            and candidate.stat().st_size == size_1
                        ):
                            downloaded_candidate = candidate
                            break

                try:
                    tab_href, _ = _get_tab_state(cdp, session_id)
                    if "accounts.google.com" in tab_href:
                        raise AuthenticationError(
                            f"Your Chrome Google session expired while exporting {report_name}."
                        )
                except AuthenticationError:
                    raise
                except Exception:
                    pass

                time.sleep(0.5)

            if downloaded_candidate is None:
                raise RuntimeError(
                    f"Google did not return the export for {report_name} within 180 seconds."
                )

            downloaded_candidate.replace(destination_path)

            if not destination_path.exists() or destination_path.stat().st_size == 0:
                raise ValueError(
                    f"The downloaded file for {report_name} is empty."
                )

            header_bytes = destination_path.read_bytes()[:4]
            if not header_bytes.startswith(b"PK"):
                raise AuthenticationError(
                    f"Google did not return a valid Excel workbook for {report_name}. "
                    "The Google session may have expired."
                )

            download_results.append(
                {
                    "report": report,
                    "download_path": destination_path,
                }
            )
            print(f"Downloaded {report_name} successfully.")

    finally:
        if target_id:
            try:
                cdp.send("Target.closeTarget", {"targetId": target_id})
            except Exception:
                pass
        cdp.close()

    return download_results