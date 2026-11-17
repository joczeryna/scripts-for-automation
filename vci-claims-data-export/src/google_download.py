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
from urllib.parse import quote, urlparse


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


def _get_devtools_ws_url() -> str:
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
            raise AuthenticationError(f"Unexpected CDP handshake response: {header_text.splitlines()[0]}")

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

    def send(self, method: str, params: dict | None = None, session_id: str | None = None) -> dict:
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


def _build_gviz_export_url(resolved_url: str, spreadsheet_id: str, source_tab: str, gid: str | None = None) -> str:
    parsed = urlparse(resolved_url)
    prefix = parsed.path.split("/d/")[0] if "/spreadsheets/u/" in parsed.path else "/spreadsheets"
    base = f"https://docs.google.com{prefix}/d/{spreadsheet_id}/gviz/tq?tqx=out:csv"
    
    # Target the exact tab by ID. Because it's the correct spreadsheet with 1 tab, 
    # gviz will naturally pull the exact bounds of the data (Columns A-O).
    if gid:
        return f"{base}&gid={gid}"
        
    encoded_tab = quote(source_tab, safe="")
    return f"{base}&sheet={encoded_tab}"


def setup_login(spreadsheet_id: str, cdp_endpoint: str | None = "auto") -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    ws_url = _get_devtools_ws_url() if not cdp_endpoint or cdp_endpoint == "auto" else cdp_endpoint

    _bring_chrome_to_front()
    print("Attaching to Chrome — click 'Allow' in Chrome if prompted...")
    cdp = _DirectCdpClient(ws_url, timeout=60.0)
    target_id = None
    try:
        created = cdp.send("Target.createTarget", {"url": sheet_url})
        target_id = created["targetId"]
        attached = cdp.send("Target.attachToTarget", {"targetId": target_id, "flatten": True})
        session_id = attached["sessionId"]
        cdp.send("Page.bringToFront", session_id=session_id)

        print("Sign in to Google in the opened Chrome tab and confirm that the spreadsheet opens.")
        input("When the CURRENT tab is visible, return here and press Enter...")

        eval_res = cdp.send(
            "Runtime.evaluate",
            {"expression": "window.location.href", "returnByValue": True},
            session_id=session_id,
        )
        current_url = eval_res.get("result", {}).get("value", "")
        if "accounts.google.com" in current_url or spreadsheet_id not in current_url:
            raise AuthenticationError(
                "Google sign-in was not completed or the spreadsheet did not open."
            )
        print("Google spreadsheet access confirmed.")
    finally:
        if target_id:
            try:
                cdp.send("Target.closeTarget", {"targetId": target_id})
            except Exception:
                pass
        cdp.close()


def download_tab_values(
    spreadsheet_id: str,
    source_tab: str,
    output_path: Path,
    cdp_endpoint: str | None = "auto",
    gid: str | None = None,
) -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    ws_url = _get_devtools_ws_url() if not cdp_endpoint or cdp_endpoint == "auto" else cdp_endpoint

    _bring_chrome_to_front()
    print("Attaching to Chrome — click 'Allow' in Chrome if prompted...")
    try:
        cdp = _DirectCdpClient(ws_url, timeout=45.0)
    except Exception as exc:
        raise AuthenticationError(
            f"Could not connect to Chrome ({exc}). Ensure Chrome is open and 'Allow remote "
            "debugging' is checked at chrome://inspect/#remote-debugging."
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
        attached = cdp.send("Target.attachToTarget", {"targetId": target_id, "flatten": True})
        session_id = attached["sessionId"]

        cdp.send("Page.enable", session_id=session_id)
        cdp.send("Page.bringToFront", session_id=session_id)
        logging.info("Opening %s in active Chrome window.", sheet_url)
        cdp.send("Page.navigate", {"url": sheet_url}, session_id=session_id)

        # Wait up to 25 seconds for sheet navigation + any silent Okta/CompanyView SSO redirect
        deadline = time.monotonic() + 25.0
        current_url = ""
        while time.monotonic() < deadline:
            eval_res = cdp.send(
                "Runtime.evaluate",
                {
                    "expression": "JSON.stringify({href: window.location.href, ready: document.readyState})",
                    "returnByValue": True,
                },
                session_id=session_id,
            )
            raw_state = eval_res.get("result", {}).get("value")
            if raw_state:
                state = json.loads(raw_state)
                current_url = state.get("href", "")
                ready = state.get("ready", "")
                if spreadsheet_id in current_url and ready in ("interactive", "complete"):
                    break
            time.sleep(0.5)

        if "accounts.google.com" in current_url or spreadsheet_id not in current_url:
            raise AuthenticationError(
                "Your Chrome Google/Okta session has expired. Open Chrome, sign in to Google, and retry."
            )

        # Pass the explicit GID to strictly target the correct tab
        export_url = _build_gviz_export_url(current_url, spreadsheet_id, source_tab, gid)
        js_fetch = (
            "(async () => {"
            f"  const resp = await fetch({json.dumps(export_url)}, {{ credentials: 'include' }});"
            "  const text = await resp.text();"
            "  return JSON.stringify({ status: resp.status, ok: resp.ok, text });"
            "})()"
        )
        fetch_eval = cdp.send(
            "Runtime.evaluate",
            {
                "expression": js_fetch,
                "awaitPromise": True,
                "returnByValue": True,
            },
            session_id=session_id,
        )
        if "exceptionDetails" in fetch_eval:
            raise RuntimeError(f"Browser CSV fetch failed: {fetch_eval['exceptionDetails']}")

        payload = json.loads(fetch_eval["result"]["value"])
        status = int(payload["status"])
        if status in (401, 403):
            raise AuthenticationError(
                "Google rejected the active Chrome session. Refresh your login in Chrome and retry."
            )
        if not payload["ok"]:
            raise RuntimeError(f"Google CSV export failed with HTTP {status}.")

        output_path.write_bytes(payload["text"].encode("utf-8"))
    finally:
        if target_id:
            try:
                cdp.send("Target.closeTarget", {"targetId": target_id})
            except Exception:
                pass
        cdp.close()

    body_prefix = output_path.read_bytes()[:2] if output_path.exists() else b""
    if body_prefix in (b"<!", b"<h"):
        if output_path.exists():
            output_path.unlink()
        raise AuthenticationError(
            "Google did not return the tab values. Refresh your login in Chrome and retry."
        )
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError("Google returned an empty CSV export.")