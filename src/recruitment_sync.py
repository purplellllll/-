"""Local Gmail-to-Feishu candidate synchronizer.

The program reads mail from one Gmail label, extracts four candidate fields,
appends one row to a Feishu spreadsheet, and stores Gmail message IDs in
SQLite for idempotency.
"""

from __future__ import annotations

import argparse
import base64
import contextlib
from datetime import datetime, timedelta
import hashlib
import html
import io
import json
import logging
import mimetypes
import msvcrt
import os
import re
import secrets
import sqlite3
import sys
import tempfile
import time
import webbrowser
import zipfile
from dataclasses import dataclass
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any
from urllib import error, parse, request
from xml.etree import ElementTree as ET


GMAIL_SCOPES = (
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
)
GMAIL_SCOPE = " ".join(GMAIL_SCOPES)
GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
FEISHU_API = "https://open.feishu.cn/open-apis"
DOCX_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
# Increment this whenever parsing policy changes so previously rejected Gmail
# messages are reconsidered by the scheduled synchronizer.
PARSER_VERSION = 9
EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
CANDIDATE_FIELD_KEYS = ("name", "student_id", "email", "phone", "major", "group")
ALLOWED_APPLICATION_GROUPS = ("开发组", "测试组", "运营组")
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png", ".webp", ".bmp"}


class SyncError(RuntimeError):
    """An expected, actionable synchronization error."""


def json_request(
    url: str,
    *,
    method: str = "GET",
    headers: dict[str, str] | None = None,
    payload: dict[str, Any] | None = None,
    form: dict[str, str] | None = None,
    timeout: int = 30,
) -> dict[str, Any]:
    if payload is not None and form is not None:
        raise ValueError("payload and form are mutually exclusive")
    body: bytes | None = None
    request_headers = dict(headers or {})
    if payload is not None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        request_headers["Content-Type"] = "application/json; charset=utf-8"
    if form is not None:
        body = parse.urlencode(form).encode("utf-8")
        request_headers["Content-Type"] = "application/x-www-form-urlencoded"
    req = request.Request(url, data=body, headers=request_headers, method=method)
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:800]
        raise SyncError(f"HTTP {exc.code} calling {parse.urlparse(url).path}: {detail}") from exc
    except error.URLError as exc:
        raise SyncError(f"Network error calling {parse.urlparse(url).netloc}: {exc.reason}") from exc
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SyncError(f"Non-JSON response from {parse.urlparse(url).netloc}") from exc


def multipart_file_request(
    url: str,
    *,
    headers: dict[str, str] | None,
    fields: dict[str, str],
    file_field: str,
    filename: str,
    content: bytes,
    timeout: int = 30,
) -> dict[str, Any]:
    """Send one binary file using ``multipart/form-data`` and decode JSON.

    The Feishu IM file endpoint requires a multipart request, whereas the
    rest of this small client intentionally uses JSON requests.  Keeping this
    transport primitive here lets the interview module reuse the same error
    handling without adding another HTTP dependency.
    """
    boundary = "----recruitment-" + secrets.token_hex(16)
    safe_filename = filename.replace("\\", "_").replace('"', "_").replace("\r", "_").replace("\n", "_")
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.extend(
            (
                f"--{boundary}\r\n".encode("ascii"),
                f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"),
                str(value).encode("utf-8"),
                b"\r\n",
            )
        )
    chunks.extend(
        (
            f"--{boundary}\r\n".encode("ascii"),
            (
                f'Content-Disposition: form-data; name="{file_field}"; '
                f'filename="{safe_filename}"\r\n'
            ).encode("utf-8"),
            b"Content-Type: application/octet-stream\r\n\r\n",
            content,
            b"\r\n",
            f"--{boundary}--\r\n".encode("ascii"),
        )
    )
    request_headers = dict(headers or {})
    request_headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
    req = request.Request(url, data=b"".join(chunks), headers=request_headers, method="POST")
    try:
        with request.urlopen(req, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")[:800]
        raise SyncError(f"HTTP {exc.code} calling {parse.urlparse(url).path}: {detail}") from exc
    except error.URLError as exc:
        raise SyncError(f"Network error calling {parse.urlparse(url).netloc}: {exc.reason}") from exc
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise SyncError(f"Non-JSON response from {parse.urlparse(url).netloc}") from exc


def b64url_decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def get_nested(data: dict[str, Any], path: str) -> Any:
    node: Any = data
    for part in path.split("."):
        if not isinstance(node, dict) or part not in node:
            return None
        node = node[part]
    return node


def resolve_path(project_root: Path, value: str) -> Path:
    candidate = Path(value)
    return candidate if candidate.is_absolute() else project_root / candidate


def load_config(config_file: Path) -> tuple[dict[str, Any], Path]:
    if not config_file.exists():
        raise SyncError(
            f"Missing {config_file.name}. Copy config.example.json to config.json and fill the local values."
        )
    try:
        config = json.loads(config_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SyncError(f"Invalid JSON in {config_file.name}: {exc}") from exc
    for section in ("gmail", "feishu", "runtime"):
        if not isinstance(config.get(section), dict):
            raise SyncError(f"config.json is missing the '{section}' section")
    return config, config_file.parent.resolve()


def configure_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
    )


def write_private_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, path)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    result: dict[str, str] = {}

    def do_GET(self) -> None:  # noqa: N802 - HTTP server callback name
        parameters = parse.parse_qs(parse.urlparse(self.path).query)
        self.__class__.result = {key: values[0] for key, values in parameters.items() if values}
        content = "<h2>Gmail 授权完成</h2><p>你可以关闭此页面并回到终端。</p>"
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content.encode("utf-8"))))
        self.end_headers()
        self.wfile.write(content.encode("utf-8"))

    def log_message(self, _format: str, *args: Any) -> None:
        return


class GmailAuth:
    def __init__(self, config: dict[str, Any], project_root: Path):
        self.client_file = resolve_path(project_root, str(config["gmail"].get("client_secret_file", "")))
        self.token_file = resolve_path(project_root, str(config["gmail"].get("token_file", "")))
        self.account_email = compact_value(str(config["gmail"].get("account_email", "")))

    def _client(self) -> dict[str, str]:
        if not self.client_file.exists():
            raise SyncError(
                f"Gmail client JSON was not found at {self.client_file}. Put the downloaded Desktop app JSON there."
            )
        data = json.loads(self.client_file.read_text(encoding="utf-8"))
        client = data.get("installed") or data.get("web")
        if not isinstance(client, dict) or not client.get("client_id"):
            raise SyncError("The Gmail client JSON has no installed/web OAuth client definition")
        return {"client_id": str(client["client_id"]), "client_secret": str(client.get("client_secret", ""))}

    def authorize(self) -> None:
        client = self._client()
        OAuthCallbackHandler.result = {}
        server = HTTPServer(("127.0.0.1", 0), OAuthCallbackHandler)
        redirect_uri = f"http://127.0.0.1:{server.server_port}/"
        verifier = secrets.token_urlsafe(64)
        challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode("ascii")).digest()).rstrip(b"=").decode("ascii")
        state = secrets.token_urlsafe(24)
        parameters = {
            "client_id": client["client_id"],
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": GMAIL_SCOPE,
            "access_type": "offline",
            "prompt": "consent",
            "state": state,
            "code_challenge": challenge,
            "code_challenge_method": "S256",
        }
        if EMAIL_PATTERN.fullmatch(self.account_email):
            # Skip Google's account chooser when the operator has configured
            # exactly one recruiting mailbox.  The user still must approve
            # the requested Gmail scopes on Google's consent page.
            parameters["login_hint"] = self.account_email
        url = GOOGLE_AUTH_URL + "?" + parse.urlencode(parameters)
        auth_url_file = self.token_file.parent / "pending-gmail-auth-url.txt"
        auth_url_file.parent.mkdir(parents=True, exist_ok=True)
        auth_url_file.write_text(url, encoding="utf-8")
        print("A browser window will open. Sign in to the Gmail account used for recruiting and approve Gmail read/send access.", flush=True)
        print("If no window appears, copy this URL into your browser:", flush=True)
        print(url, flush=True)
        webbrowser.open(url)
        server.timeout = 1
        deadline = time.time() + 300
        while time.time() < deadline and not OAuthCallbackHandler.result:
            server.handle_request()
        server.server_close()
        result = OAuthCallbackHandler.result
        if not result:
            raise SyncError("Gmail authorization timed out after five minutes")
        if result.get("state") != state:
            raise SyncError("Gmail OAuth state validation failed")
        if result.get("error"):
            raise SyncError(f"Gmail authorization was declined: {result['error']}")
        if not result.get("code"):
            raise SyncError("Gmail authorization returned no authorization code")
        token = json_request(
            GOOGLE_TOKEN_URL,
            method="POST",
            form={
                "code": result["code"],
                "client_id": client["client_id"],
                "client_secret": client["client_secret"],
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
                "code_verifier": verifier,
            },
        )
        if "access_token" not in token or "refresh_token" not in token:
            raise SyncError("Gmail OAuth did not return a reusable access and refresh token")
        token["expires_at"] = int(time.time()) + int(token.get("expires_in", 3600))
        write_private_json(self.token_file, token)
        auth_url_file.unlink(missing_ok=True)
        print(f"Gmail authorization saved locally in {self.token_file}")

    def access_token(self) -> str:
        client = self._client()
        if not self.token_file.exists():
            raise SyncError("Gmail is not authorized yet. Run: .\\run.ps1 authorize-gmail")
        token = json.loads(self.token_file.read_text(encoding="utf-8"))
        if token.get("access_token") and int(token.get("expires_at", 0)) > time.time() + 60:
            return str(token["access_token"])
        refresh_token = token.get("refresh_token")
        if not refresh_token:
            raise SyncError("The saved Gmail token has no refresh token. Run authorize-gmail again.")
        refreshed = json_request(
            GOOGLE_TOKEN_URL,
            method="POST",
            form={
                "client_id": client["client_id"],
                "client_secret": client["client_secret"],
                "refresh_token": str(refresh_token),
                "grant_type": "refresh_token",
            },
        )
        if "access_token" not in refreshed:
            raise SyncError("Gmail refresh failed. Run authorize-gmail again.")
        token.update(refreshed)
        token["expires_at"] = int(time.time()) + int(token.get("expires_in", 3600))
        write_private_json(self.token_file, token)
        return str(token["access_token"])


class GmailClient:
    def __init__(self, auth: GmailAuth):
        self.auth = auth
        self._mailbox_address: str | None = None

    def call(
        self,
        method: str,
        path: str,
        *,
        query: dict[str, Any] | None = None,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        url = "https://gmail.googleapis.com/gmail/v1/users/me" + path
        if query:
            url += "?" + parse.urlencode({key: value for key, value in query.items() if value is not None})
        return json_request(
            url,
            method=method,
            headers={"Authorization": f"Bearer {self.auth.access_token()}"},
            payload=payload,
        )

    def label_id(self, label_name: str) -> str:
        labels = self.call("GET", "/labels").get("labels", [])
        for label in labels:
            if label.get("name") == label_name:
                return str(label["id"])
        raise SyncError(f"Gmail label '{label_name}' was not found. Create it in Gmail before the first sync.")

    def messages(self, label_id: str, limit: int) -> list[dict[str, Any]]:
        response = self.call("GET", "/messages", query={"labelIds": label_id, "maxResults": limit})
        return list(response.get("messages", []))

    def message(self, message_id: str) -> dict[str, Any]:
        return self.call("GET", f"/messages/{parse.quote(message_id)}", query={"format": "full"})

    def attachment(self, message_id: str, attachment_id: str) -> bytes:
        response = self.call(
            "GET",
            f"/messages/{parse.quote(message_id)}/attachments/{parse.quote(attachment_id)}",
        )
        data = response.get("data")
        if not isinstance(data, str):
            raise SyncError("Gmail attachment response contains no data")
        return b64url_decode(data)

    def mailbox_address(self) -> str:
        if self._mailbox_address:
            return self._mailbox_address
        address = compact_value(str(self.call("GET", "/profile").get("emailAddress", "")))
        if not EMAIL_PATTERN.fullmatch(address):
            raise SyncError("Gmail did not return a valid mailbox address for the authorized account")
        self._mailbox_address = address
        return address

    def send_plain_message(self, recipient: str, subject: str, content: str) -> str:
        """Send a UTF-8 plain-text notice from the authorized Gmail account."""
        message = EmailMessage()
        message["From"] = self.mailbox_address()
        message["To"] = recipient
        message["Subject"] = subject
        message.set_content(content)
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
        try:
            response = self.call("POST", "/messages/send", payload={"raw": raw})
        except SyncError as exc:
            if "HTTP 403" in str(exc):
                raise SyncError(
                    "Gmail has not granted send permission yet. Run .\\run.ps1 authorize-gmail and approve sending email."
                ) from exc
            raise
        message_id = str(response.get("id", "")).strip()
        if not message_id:
            raise SyncError("Gmail did not return a message ID after sending the interview notice")
        return message_id


def decoded_part(part: dict[str, Any]) -> str:
    data = get_nested(part, "body.data")
    if not isinstance(data, str):
        return ""
    raw = b64url_decode(data)
    charset = "utf-8"
    for header in part.get("headers", []):
        if str(header.get("name", "")).lower() == "content-type":
            match = re.search(r"charset=[\"']?([^;\"'\s]+)", str(header.get("value", "")), flags=re.I)
            if match:
                charset = match.group(1)
    try:
        return raw.decode(charset, errors="replace")
    except LookupError:
        return raw.decode("utf-8", errors="replace")


def strip_html(value: str) -> str:
    value = re.sub(r"(?is)<(script|style).*?>.*?</\1>", " ", value)
    value = re.sub(r"(?s)<[^>]+>", " ", value)
    return html.unescape(re.sub(r"[ \t]+", " ", value))


def walk_message_parts(part: dict[str, Any]) -> list[dict[str, Any]]:
    result = [part]
    for child in part.get("parts", []) or []:
        result.extend(walk_message_parts(child))
    return result


def message_content(gmail: GmailClient, message: dict[str, Any]) -> tuple[str, list[tuple[str, bytes]]]:
    text_parts: list[str] = []
    html_parts: list[str] = []
    attachments: list[tuple[str, bytes]] = []
    for part in walk_message_parts(message.get("payload", {})):
        mime_type = str(part.get("mimeType", "")).lower()
        filename = str(part.get("filename", "")).strip()
        if mime_type == "text/plain":
            text_parts.append(decoded_part(part))
        elif mime_type == "text/html":
            html_parts.append(strip_html(decoded_part(part)))
        if filename:
            attachment_id = get_nested(part, "body.attachmentId")
            inline_data = get_nested(part, "body.data")
            if isinstance(attachment_id, str):
                attachments.append((filename, gmail.attachment(str(message["id"]), attachment_id)))
            elif isinstance(inline_data, str):
                attachments.append((filename, b64url_decode(inline_data)))
    return "\n".join(text_parts or html_parts), attachments


def normalize_label(value: str) -> str:
    return re.sub(r"[\s:：_-]+", "", value).lower()


def compact_value(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip(" \t\r\n:：|，,;；")


def column_name(column_number: int) -> str:
    """Convert a one-based column number into an A1-style column name."""
    result = ""
    while column_number:
        column_number, remainder = divmod(column_number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_text(value: Any) -> str:
    """Normalize a Sheets cell, including Feishu's hyperlink return format."""
    if isinstance(value, list):
        text_parts: list[str] = []
        link_parts: list[str] = []
        for item in value:
            if not isinstance(item, dict):
                continue
            text = compact_value(str(item.get("text", "")))
            link = compact_value(str(item.get("link", "")))
            if text:
                text_parts.append(text)
            if link:
                link_parts.append(link)
        # Feishu may return a hyperlink as an empty text run followed by a
        # URL run.  Prefer an HTTP(S) link whenever one exists so link cells
        # are not mistaken for blank values or visible labels.
        for link in link_parts:
            if parse.urlparse(link).scheme in {"http", "https"} and parse.urlparse(link).netloc:
                return link
        if text_parts:
            return compact_value("".join(text_parts))
        if link_parts:
            return link_parts[0]
        return ""
    return compact_value("" if value is None else str(value))


def normalize_scheduled_at(value: Any) -> str:
    """Convert Feishu/Excel date serials into the datetime format used by Offers.

    Feishu may return a date-time cell as either an ISO-like string or a
    spreadsheet serial (for example ``46268.666...``).  Keeping the raw
    serial would make the Offer template fall back to ``XX:XX``.
    """
    normalized = cell_text(value)
    if not normalized:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}[ T]+\d{1,2}:\d{2}", normalized):
        try:
            parsed = datetime.fromisoformat(normalized.replace("T", " "))
            return parsed.strftime("%Y-%m-%d %H:%M")
        except ValueError:
            return normalized
    try:
        serial = float(normalized)
    except ValueError:
        return normalized
    # Feishu date serials for ordinary calendar dates fall in this range.
    if not 20000 <= serial <= 100000:
        return normalized
    try:
        parsed = datetime(1899, 12, 30) + timedelta(days=serial)
    except (OverflowError, ValueError):
        return normalized
    return parsed.strftime("%Y-%m-%d %H:%M")


FIELD_LABELS = {
    "name": {"姓名", "name"},
    "student_id": {"学号", "学生编号", "studentid"},
    "email": {"email", "e-mail", "邮箱", "电子邮箱", "电子邮件"},
    "phone": {"电话", "联系电话", "手机", "手机号", "mobile", "phone"},
    "major": {"专业", "所学专业", "所在专业", "主修专业", "专业名称", "major", "specialty"},
    "group": {"应聘组别", "应聘岗位", "申请组别", "申请岗位", "意向组别", "意向岗位", "求职意向", "岗位", "职位", "position", "role", "job"},
}

# Labels that often sit next to résumé fields in the same PDF table row. They
# delimit a blank value, instead of being treated as that value.
TABLE_STOP_LABELS = {
    "出生年月",
    "性别",
    "民族",
    "政治面貌",
    "籍贯",
    "现所在地",
    "学历",
    "学位",
    "学校",
    "院校",
    "教育经历",
    "电子邮件",
    "联系方式",
    "专业技能",
    "专业领域",
    "在校经历",
    "工作经历",
    "自我评价",
}

NAME_HEADING_BLACKLIST = {
    "个人简历",
    "简历模板",
    "个人信息",
    "教育经历",
    "技术栈",
    "专业技能",
    "项目经历",
    "校园经历",
    "在校经历",
    "工作经历",
    "工作经验",
    "培训经历",
    "自我评价",
    "求职意向",
    "竞赛与荣誉",
    "个人项目",
    "项目经验",
    "社会实践",
    "活动组织",
    "才艺特长",
    "专业领域",
}

# A blank form can place adjacent labels in the PDF text stream.  Those labels
# must never be mistaken for a candidate's name or major just because their
# cells happen to follow a ``姓名`` / ``专业`` heading.
NAME_VALUE_NOISE = {
    "性别",
    "照片",
    "学校",
    "籍贯",
    "专业班级",
    "年龄",
    "政治面貌",
    "外语能力",
    "邮箱",
    "电话",
    "教育背景",
    "技能证书",
    "获得荣誉",
    "自我评价",
}
MAJOR_VALUE_NOISE = NAME_VALUE_NOISE | {
    "班级",
    "学历",
    "学位",
    "联系方式",
    "电子邮箱",
}


def extract_docx_rows(raw: bytes) -> tuple[list[list[str]], str]:
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            document = archive.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise SyncError("The .docx attachment cannot be read") from exc
    root = ET.fromstring(document)
    rows: list[list[str]] = []
    for row in root.findall(".//w:tr", DOCX_NS):
        cells: list[str] = []
        for cell in row.findall("./w:tc", DOCX_NS):
            text = "".join(node.text or "" for node in cell.findall(".//w:t", DOCX_NS))
            cells.append(compact_value(text))
        if cells:
            rows.append(cells)
    paragraphs = [compact_value("".join(node.text or "" for node in paragraph.findall(".//w:t", DOCX_NS))) for paragraph in root.findall(".//w:p", DOCX_NS)]
    joined = "\n".join([" | ".join(row) for row in rows] + [line for line in paragraphs if line])
    return rows, joined


def extract_pdf_text(raw: bytes) -> str:
    """Extract selectable text from a PDF attachment.

    Image-only scans contain no text layer and need OCR, so they are reported
    for manual review instead of guessing at personal information.
    """
    # PyMuPDF keeps the reading order of many merged-cell résumé tables much
    # better than pypdf.  In particular it avoids splitting "E-mail" and its
    # value across unrelated table cells.  Keep pypdf as the compatibility
    # fallback when PyMuPDF is unavailable or cannot read a particular file.
    try:
        import fitz

        with fitz.open(stream=raw, filetype="pdf") as document:
            text = "\n".join(page.get_text("text") for page in document)
        if compact_value(text):
            return compact_value(text)
    except Exception as exc:
        logging.debug("PyMuPDF text extraction failed; using pypdf fallback: %s", exc)
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise SyncError("PDF parsing requires pypdf. Run: python -m pip install -r requirements.txt") from exc
    try:
        reader = PdfReader(io.BytesIO(raw))
        if reader.is_encrypted and reader.decrypt("") == 0:
            raise SyncError("The PDF attachment is encrypted and cannot be read")
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
    except SyncError:
        raise
    except Exception as exc:
        raise SyncError("The PDF attachment cannot be read") from exc
    text = compact_value(text)
    if not text:
        raise SyncError("The PDF has no selectable text (it may be a scanned image and needs OCR)")
    return text


def extract_pdf_rows(raw: bytes) -> list[list[str]]:
    """Extract grid-table cells from a selectable-text PDF when available.

    PDF text streams can be ordered by drawing position rather than by table
    cell. PyMuPDF preserves the label/value relationships in table-style
    résumés; pypdf text extraction remains the fallback for ordinary layouts.
    """
    try:
        import fitz
    except ImportError:
        return []
    rows: list[list[str]] = []
    try:
        with fitz.open(stream=raw, filetype="pdf") as document:
            for page in document:
                # Some PyMuPDF versions print an optional layout suggestion;
                # suppress it because synchronization runs in the background.
                with contextlib.redirect_stdout(io.StringIO()):
                    tables = page.find_tables()
                for table in tables.tables:
                    for row in table.extract():
                        values = [compact_value("" if cell is None else str(cell)) for cell in row]
                        if any(values):
                            rows.append(values)
    except Exception as exc:
        logging.debug("PDF table detection failed; using text fallback: %s", exc)
    return rows


def extract_pdf_content(raw: bytes) -> tuple[list[list[str]], str]:
    """Return both structural PDF table rows and regular extracted text."""
    return extract_pdf_rows(raw), extract_pdf_text(raw)


def extract_pdf_top_name(raw: bytes) -> str:
    """Infer a Chinese name only from a prominent title at the top of a PDF.

    This deliberately has a narrow confidence boundary: it is a fallback for
    visually designed resumes that omit a ``姓名`` label, not a free-form text
    guesser.  Ambiguous titles are left blank.
    """
    try:
        import fitz
    except ImportError:
        return ""
    try:
        with fitz.open(stream=raw, filetype="pdf") as document:
            if not document.page_count:
                return ""
            page = document[0]
            page_width = float(page.rect.width)
            top_limit = float(page.rect.height) * 0.25
            choices: list[tuple[float, str]] = []
            for block in page.get_text("dict").get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        candidate = re.sub(r"\s+", "", str(span.get("text", "")))
                        if not re.fullmatch(r"[\u4e00-\u9fff]{2,4}", candidate):
                            continue
                        if candidate in NAME_HEADING_BLACKLIST:
                            continue
                        bbox = span.get("bbox", [0, 0, 0, 0])
                        top = float(bbox[1])
                        if top > top_limit:
                            continue
                        font_size = float(span.get("size", 0))
                        center = (float(bbox[0]) + float(bbox[2])) / 2
                        centeredness = max(0.0, 1.0 - abs(center - page_width / 2) / (page_width / 2))
                        choices.append((font_size * 100 + centeredness * 20 - top / 10, candidate))
    except Exception as exc:
        logging.debug("PDF top-name detection failed: %s", exc)
        return ""
    return max(choices, default=(0.0, ""))[1]


def extract_doc_content(raw: bytes) -> tuple[list[list[str]], str]:
    """Read text and table cells from a legacy .doc through Microsoft Word."""
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise SyncError("Legacy .doc parsing requires Microsoft Word and pywin32 on this computer") from exc

    temporary_path: Path | None = None
    word: Any = None
    document: Any = None
    initialized = False
    try:
        with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as temporary:
            temporary.write(raw)
            temporary_path = Path(temporary.name)
        pythoncom.CoInitialize()
        initialized = True
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            FileName=str(temporary_path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        rows: list[list[str]] = []
        for table in document.Tables:
            # Word cannot enumerate Rows when a legacy table has vertically
            # merged cells. Its full range still has each cell delimited by
            # CR+BEL, which is enough for our adjacent label/value parser.
            table_text = str(table.Range.Text)
            cells = [
                compact_value(cell.replace("\x07", " ").replace("\r", " "))
                for cell in table_text.split("\r\x07")
            ]
            cells = [cell for cell in cells if cell]
            if cells:
                rows.append(cells)
        text = str(document.Content.Text)
    except Exception as exc:
        raise SyncError("The legacy .doc attachment cannot be read by Microsoft Word") from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(SaveChanges=0)
            except Exception:
                pass
        if initialized:
            pythoncom.CoUninitialize()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    text = text.replace("\r\x07", "\n").replace("\x07", " ").replace("\r", "\n")
    joined = "\n".join([" | ".join(row) for row in rows] + [compact_value(text)])
    if not compact_value(joined):
        raise SyncError("The legacy .doc attachment contains no readable text")
    return rows, joined


def extract_doc_text(raw: bytes) -> str:
    """Backward-compatible text-only helper for legacy .doc attachments."""
    return extract_doc_content(raw)[1]


def fields_from_rows(rows: list[list[str]]) -> dict[str, str]:
    values: dict[str, str] = {}
    aliases = {normalize_label(alias): key for key, aliases in FIELD_LABELS.items() for alias in aliases}
    stop_labels = set(aliases) | {normalize_label(label) for label in TABLE_STOP_LABELS}
    for row in rows:
        for index, cell in enumerate(row[:-1]):
            key = aliases.get(normalize_label(cell))
            if not key or values.get(key):
                continue
            for following_cell in row[index + 1 :]:
                candidate = compact_value(following_cell)
                if not candidate:
                    continue
                if normalize_label(candidate) in stop_labels:
                    break
                values[key] = candidate
                break
    return values


def first_match(patterns: list[str], value: str) -> str:
    for pattern in patterns:
        match = re.search(pattern, value, flags=re.I | re.M)
        if match:
            return compact_value(match.group(1))
    return ""


def compact_chinese_value(value: str) -> str:
    """Remove artificial spaces introduced inside a CJK value by PDF text extraction."""
    return re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", compact_value(value))


def fields_from_text(value: str) -> dict[str, str]:
    fields: dict[str, str] = {}
    fields["name"] = first_match(
        [
            r"(?:姓\s*名|\bname)\s*[:：]?\s*([^\r\n|，,;；]{1,40}?)"
            r"(?=\s*(?:学\s*号|学生编号|student\s*id|邮箱|电子邮箱|电子邮件|e-?mail|联系电话|手机(?:号)?|电话|专业|所学专业|主修专业|major|mobile|phone)\s*[:：]?|[\r\n|，,;；]|\s*$)"
        ],
        value,
    )
    fields["student_id"] = first_match(
        [r"(?:学\s*号|学生编号|student\s*id)\s*[:：]?\s*([A-Za-z0-9_-]{5,32})"], value
    )
    fields["email"] = first_match(
        [r"(?:e-?mail|邮箱|电子邮箱|电子邮件)\s*[:：]?\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})"], value
    )
    if not fields["email"]:
        generic_email = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", value)
        fields["email"] = generic_email.group(0) if generic_email else ""
    fields["phone"] = first_match(
        [r"(?:联系电话|手机号|手机|电话|mobile|phone)\s*[:：]?\s*((?:\+?86[-\s]?)?1[3-9]\d[-\s]?\d{4}[-\s]?\d{4})"], value
    )
    if not fields["phone"]:
        generic_phone = re.search(r"(?<!\d)(?:\+?86[-\s]?)?1[3-9]\d[-\s]?\d{4}[-\s]?\d{4}(?!\d)", value)
        fields["phone"] = generic_phone.group(0) if generic_phone else ""
    fields["major"] = first_match(
        [
            r"(?:所(?:学|在)?\s*专\s*业|主修\s*(?:专\s*业)?|专\s*业(?!\s*(?:领域|技能|能力|课程))(?:名称)?|\bmajor)\s*[:：]?\s*"
            r"([^\r\n|，,;；]{2,80}?)(?=\s*(?:姓名|学号|学生编号|student\s*id|邮箱|电子邮箱|电子邮件|e-?mail|联系电话|手机(?:号)?|电话|学历|学位|学校|院校|教育经历|毕业院校|mobile|phone)\s*[:：]?|[\r\n|，,;；]|\s*$)"
        ],
        value,
    )
    if fields["major"]:
        fields["major"] = compact_chinese_value(fields["major"])
    fields["group"] = first_match(
        [
            r"(?:应聘(?:组别|岗位)?|申请(?:组别|岗位)?|意向(?:组别|岗位)?|求职意向|岗位|职位|\bposition|\brole|\bjob)\s*[:：]?\s*"
            r"([^\r\n|，,;；]{2,80}?)(?=\s*(?:姓名|学号|学生编号|student\s*id|邮箱|电子邮箱|电子邮件|e-?mail|联系电话|手机(?:号)?|电话|专业|所学专业|主修专业|学历|学位|学校|院校|教育经历|毕业院校|mobile|phone)\s*[:：]?|[\r\n|，,;；]|\s*$)"
        ],
        value,
    )
    return {key: compact_value(item) for key, item in fields.items() if item}


def message_headers(message: dict[str, Any]) -> dict[str, str]:
    return {
        str(header.get("name", "")).lower(): str(header.get("value", ""))
        for header in message.get("payload", {}).get("headers", [])
    }


def preferred_email(attachment_value: str, body_value: str) -> str:
    """Use only an address explicitly found in the résumé or email body."""
    return next((value for value in (attachment_value, body_value) if compact_value(value)), "")


def normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", compact_value(value))
    if len(digits) == 13 and digits.startswith("86"):
        digits = digits[2:]
    return digits if re.fullmatch(r"1[3-9]\d{9}", digits) else ""


def normalize_application_group(value: str) -> str:
    """Accept only the three recruiting groups used by this workflow."""
    normalized = compact_value(value)
    return next((group for group in ALLOWED_APPLICATION_GROUPS if group in normalized), "")


def clean_candidate_fields(fields: dict[str, str]) -> dict[str, str]:
    """Keep only safely formatted values; unknown or malformed fields stay blank."""
    cleaned = {key: compact_value(str(fields.get(key, ""))) for key in CANDIDATE_FIELD_KEYS}
    name = compact_chinese_value(cleaned["name"])
    is_chinese_name = bool(re.fullmatch(r"[\u4e00-\u9fff]{2,8}", name))
    is_latin_name = bool(re.fullmatch(r"[A-Za-z][A-Za-z .'-]{1,79}", name))
    if (
        not (is_chinese_name or is_latin_name)
        or name in NAME_HEADING_BLACKLIST
        or any(label in name for label in NAME_VALUE_NOISE)
    ):
        cleaned["name"] = ""
    else:
        cleaned["name"] = name
    email = cleaned["email"]
    cleaned["email"] = email if re.fullmatch(EMAIL_PATTERN, email) else ""
    cleaned["phone"] = normalize_phone(cleaned["phone"])
    major = compact_chinese_value(cleaned["major"])
    if (
        major in {"学校", "学类", "专业", "专业技能"}
        or major.startswith("学校")
        or len(major) > 50
        or any(label in major for label in MAJOR_VALUE_NOISE)
    ):
        cleaned["major"] = ""
    else:
        cleaned["major"] = major
    cleaned["group"] = normalize_application_group(cleaned["group"])
    return cleaned


class CandidateLLM:
    """Optional, validated second-pass extraction through an OpenAI-compatible API."""

    def __init__(self, config: dict[str, Any]):
        options = config.get("llm", {}) if isinstance(config.get("llm", {}), dict) else {}
        self.enabled = bool(options.get("enabled", False))
        self.base_url = str(options.get("base_url", "")).strip().rstrip("/")
        self.model = str(options.get("model", "")).strip()
        self.vision_model = str(options.get("vision_model", "qwen-vl-plus")).strip()
        self.key_env = str(options.get("api_key_env", "")).strip()
        self.timeout_seconds = max(5, int(options.get("timeout_seconds", 45)))
        self.max_input_characters = max(1000, int(options.get("max_input_characters", 12000)))

    def should_enrich(self, candidate: dict[str, str], source_text: str) -> bool:
        return self.enabled and bool(compact_value(source_text)) and any(not candidate.get(key) for key in CANDIDATE_FIELD_KEYS)

    @staticmethod
    def _content_as_text(response: dict[str, Any]) -> str:
        choices = response.get("choices", [])
        if not isinstance(choices, list) or not choices or not isinstance(choices[0], dict):
            return ""
        message = choices[0].get("message", {})
        if not isinstance(message, dict):
            return ""
        content = message.get("content", "")
        if isinstance(content, str):
            return content
        if isinstance(content, list):
            return "".join(str(part.get("text", "")) for part in content if isinstance(part, dict))
        return ""

    def enrich(self, candidate: dict[str, str], source_text: str) -> dict[str, str]:
        if not self.should_enrich(candidate, source_text):
            return {}
        if not self.base_url or not self.model:
            logging.warning("LLM second pass is enabled but base_url or model is missing; continuing without it.")
            return {}
        api_key = os.environ.get(self.key_env, "").strip() if self.key_env else ""
        headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
        prompt = (
            "Extract only explicitly stated candidate details from this Chinese résumé/email text. "
            "Return one JSON object and nothing else, with exactly these string keys: "
            "name, student_id, email, phone, major, group. "
            "Use an empty string when a field is absent or uncertain. Never infer values, never use an email sender address, "
            "and keep phone as an 11-digit mainland-China mobile number when present.\n\n"
            f"Rule-pass result (only fill blanks or malformed values): {json.dumps(candidate, ensure_ascii=False)}\n\n"
            f"Source text:\n{source_text[:self.max_input_characters]}"
        )
        try:
            response = json_request(
                f"{self.base_url}/chat/completions",
                method="POST",
                headers=headers,
                payload={
                    "model": self.model,
                    "temperature": 0,
                    "messages": [
                        {"role": "system", "content": "You are a precise résumé information extractor. Output valid JSON only."},
                        {"role": "user", "content": prompt},
                    ],
                },
                timeout=self.timeout_seconds,
            )
            content = self._content_as_text(response).strip()
            start, end = content.find("{"), content.rfind("}")
            if start < 0 or end < start:
                raise SyncError("LLM did not return a JSON object")
            decoded = json.loads(content[start : end + 1])
            if not isinstance(decoded, dict):
                raise SyncError("LLM returned a non-object result")
            return clean_candidate_fields({key: str(decoded.get(key, "")) for key in CANDIDATE_FIELD_KEYS})
        except (SyncError, ValueError, TypeError, json.JSONDecodeError) as exc:
            logging.warning("LLM second-pass extraction failed; continuing with rule results: %s", exc)
            return {}

    def extract_image(self, raw: bytes, filename: str) -> dict[str, str]:
        """Extract candidate fields from an image résumé with Qwen-VL."""
        if not self.enabled or not self.base_url or not self.vision_model:
            return {}
        if len(raw) > 10 * 1024 * 1024:
            raise SyncError(f"The image résumé '{filename}' exceeds the 10 MB vision limit")
        api_key = os.environ.get(self.key_env, "").strip() if self.key_env else ""
        headers = {"Authorization": f"Bearer {api_key}"} if api_key else {}
        mime_type = mimetypes.guess_type(filename)[0] or "image/jpeg"
        image_url = f"data:{mime_type};base64,{base64.b64encode(raw).decode('ascii')}"
        prompt = (
            "Read this Chinese résumé image and return one JSON object with exactly these string keys: "
            "name, student_id, email, phone, major, group. Use an empty string when a field is absent or uncertain. "
            "Never infer values, never use the sender address, and group must be only 开发组, 测试组, or 运营组."
        )
        try:
            response = json_request(
                f"{self.base_url}/chat/completions",
                method="POST",
                headers=headers,
                payload={
                    "model": self.vision_model,
                    "temperature": 0,
                    "messages": [
                        {"role": "system", "content": "You are a precise résumé information extractor. Output valid JSON only."},
                        {
                            "role": "user",
                            "content": [
                                {"type": "image_url", "image_url": {"url": image_url}},
                                {"type": "text", "text": prompt},
                            ],
                        },
                    ],
                },
                timeout=self.timeout_seconds,
            )
            content = self._content_as_text(response).strip()
            start, end = content.find("{"), content.rfind("}")
            if start < 0 or end < start:
                raise SyncError("Vision model did not return a JSON object")
            decoded = json.loads(content[start : end + 1])
            if not isinstance(decoded, dict):
                raise SyncError("Vision model returned a non-object result")
            return clean_candidate_fields({key: str(decoded.get(key, "")) for key in CANDIDATE_FIELD_KEYS})
        except (SyncError, ValueError, TypeError, json.JSONDecodeError) as exc:
            logging.warning("Image résumé extraction failed; continuing with blank fields: %s", exc)
            return {}


def parse_candidate(
    gmail: GmailClient, message: dict[str, Any], llm: CandidateLLM | None = None
) -> tuple[dict[str, str], list[str]]:
    body, attachments = message_content(gmail, message)
    fields = fields_from_text(body)
    body_email = fields.get("email", "")
    attachment_email = ""
    attachment_fields: dict[str, str] = {}
    attachment_errors: list[str] = []
    source_texts = [body]
    for filename, raw in attachments:
        suffix = Path(filename).suffix.lower()
        try:
            if suffix == ".docx":
                rows, document_text = extract_docx_rows(raw)
                document_fields = fields_from_rows(rows)
                for key, value in fields_from_text(document_text).items():
                    document_fields.setdefault(key, value)
            elif suffix == ".doc":
                rows, document_text = extract_doc_content(raw)
                document_fields = fields_from_rows(rows)
                for key, value in fields_from_text(document_text).items():
                    document_fields.setdefault(key, value)
            elif suffix == ".pdf":
                rows, document_text = extract_pdf_content(raw)
                document_fields = fields_from_rows(rows)
                for key, value in fields_from_text(document_text).items():
                    # PDF table extraction can merge labels with adjacent
                    # values.  Prefer explicitly labelled text for fields
                    # whose formats can be verified independently.
                    if key in {"email", "phone", "major", "group"} or not document_fields.get(key):
                        document_fields[key] = value
                if not document_fields.get("name"):
                    document_fields["name"] = extract_pdf_top_name(raw)
            elif suffix in IMAGE_SUFFIXES:
                if not llm:
                    raise SyncError("Image résumé extraction requires the Qwen-VL second pass")
                document_text = ""
                document_fields = llm.extract_image(raw, filename)
            else:
                continue
            source_texts.append(document_text)
            for key, value in document_fields.items():
                if key == "email":
                    attachment_email = attachment_email or value
                else:
                    attachment_fields.setdefault(key, value)
        except SyncError as exc:
            attachment_errors.append(f"{filename}: {exc}")
    # A resume is the source of record. Its explicitly extracted values take
    # precedence over loose values found in a free-form email body.
    fields.update(attachment_fields)
    # Only information explicitly found in the résumé or email body is stored.
    # The Gmail sender address is not assumed to be the candidate's address.
    fields["email"] = preferred_email(attachment_email, body_email)
    fields = clean_candidate_fields(fields)
    if llm:
        llm_fields = llm.enrich(fields, "\n\n".join(part for part in source_texts if compact_value(part)))
        for key, value in llm_fields.items():
            if not fields.get(key):
                fields[key] = value
    # Missing applicant fields are intentionally represented by an empty cell
    # in Feishu.  A resume should not be rejected merely because a particular
    # field (for example, a phone number) cannot be extracted.
    return fields, attachment_errors


class SyncRunLock:
    """A process-wide Windows lock that prevents overlapping sync runs."""

    def __init__(self, lock_file: Path):
        self.lock_file = lock_file
        self.handle: Any | None = None

    def acquire(self) -> bool:
        self.lock_file.parent.mkdir(parents=True, exist_ok=True)
        handle = self.lock_file.open("a+", encoding="utf-8")
        try:
            handle.seek(0)
            if not handle.read(1):
                handle.seek(0)
                handle.write("\n")
                handle.flush()
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            handle.close()
            return False
        self.handle = handle
        return True

    def release(self) -> None:
        if not self.handle:
            return
        try:
            self.handle.seek(0)
            msvcrt.locking(self.handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass
        finally:
            self.handle.close()
            self.handle = None


class SyncState:
    def __init__(self, database_file: Path):
        database_file.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(database_file)
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS message_state (
                message_id TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                feishu_record_id TEXT,
                error_reason TEXT,
                updated_at INTEGER NOT NULL
            )
            """
        )
        columns = {row[1] for row in self.connection.execute("PRAGMA table_info(message_state)")}
        if "parser_version" not in columns:
            self.connection.execute("ALTER TABLE message_state ADD COLUMN parser_version INTEGER NOT NULL DEFAULT 1")
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS interview_link_email_state (
                record_range TEXT PRIMARY KEY,
                recipient_email TEXT NOT NULL,
                interview_link TEXT NOT NULL,
                gmail_message_id TEXT NOT NULL,
                sent_at INTEGER NOT NULL
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS interview_offer_review_state (
                review_id TEXT PRIMARY KEY,
                record_range TEXT NOT NULL,
                recipient_email TEXT NOT NULL,
                interview_link TEXT NOT NULL,
                candidate_name TEXT NOT NULL,
                scheduled_at TEXT NOT NULL,
                offer_subject TEXT NOT NULL,
                offer_content TEXT NOT NULL,
                reviewer_email TEXT NOT NULL,
                review_message_id TEXT NOT NULL,
                status TEXT NOT NULL,
                decided_message_id TEXT,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL,
                UNIQUE(record_range, recipient_email, interview_link)
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS candidate_backup_state (
                message_id TEXT PRIMARY KEY,
                feishu_record_id TEXT NOT NULL,
                candidate_json TEXT NOT NULL,
                status TEXT NOT NULL,
                backup_path TEXT,
                error_reason TEXT,
                updated_at INTEGER NOT NULL
            )
            """
        )
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS candidate_resume_upload_state (
                message_id TEXT PRIMARY KEY,
                feishu_record_id TEXT NOT NULL,
                status TEXT NOT NULL,
                files_json TEXT NOT NULL,
                error_reason TEXT,
                updated_at INTEGER NOT NULL
            )
            """
        )
        self.connection.commit()

    @staticmethod
    def _record_row(record_range: str) -> str:
        """Return the sheet row number, independent of the range's last column."""
        match = re.search(r"![A-Z]+(\d+):", str(record_range))
        return match.group(1) if match else ""

    @classmethod
    def _same_record_range(cls, first: str, second: str) -> bool:
        return bool(first == second or (cls._record_row(first) and cls._record_row(first) == cls._record_row(second)))

    def handled(self, message_id: str) -> bool:
        row = self.connection.execute(
            "SELECT status, parser_version FROM message_state WHERE message_id = ?", (message_id,)
        ).fetchone()
        return bool(row and (row[0] == "success" or (row[0] == "invalid" and row[1] >= PARSER_VERSION)))

    def save(self, message_id: str, status: str, *, record_id: str | None = None, reason: str | None = None) -> None:
        self.connection.execute(
            """
            INSERT INTO message_state(message_id, status, feishu_record_id, error_reason, updated_at, parser_version)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(message_id) DO UPDATE SET
                status=excluded.status,
                feishu_record_id=excluded.feishu_record_id,
                error_reason=excluded.error_reason,
                updated_at=excluded.updated_at,
                parser_version=excluded.parser_version
            """,
            (message_id, status, record_id, reason, int(time.time()), PARSER_VERSION),
        )
        self.connection.commit()

    def summary(self) -> list[tuple[str, int]]:
        return list(self.connection.execute("SELECT status, COUNT(*) FROM message_state GROUP BY status ORDER BY status"))

    def queue_candidate_backup(self, message_id: str, record_id: str, candidate: dict[str, str]) -> None:
        self.connection.execute(
            """
            INSERT INTO candidate_backup_state(
                message_id, feishu_record_id, candidate_json, status, backup_path, error_reason, updated_at
            ) VALUES (?, ?, ?, 'pending', NULL, NULL, ?)
            ON CONFLICT(message_id) DO UPDATE SET
                feishu_record_id=excluded.feishu_record_id,
                candidate_json=excluded.candidate_json,
                status=CASE WHEN candidate_backup_state.status='completed' THEN 'completed' ELSE 'pending' END,
                error_reason=CASE WHEN candidate_backup_state.status='completed' THEN candidate_backup_state.error_reason ELSE NULL END,
                updated_at=excluded.updated_at
            """,
            (message_id, record_id, json.dumps(candidate, ensure_ascii=False, sort_keys=True), int(time.time())),
        )
        self.connection.commit()

    def pending_candidate_backups(self) -> list[tuple[str, str, dict[str, str]]]:
        rows = self.connection.execute(
            """
            SELECT message_id, feishu_record_id, candidate_json
            FROM candidate_backup_state
            WHERE status != 'completed'
            ORDER BY updated_at
            """
        ).fetchall()
        result: list[tuple[str, str, dict[str, str]]] = []
        for message_id, record_id, candidate_json in rows:
            try:
                decoded = json.loads(candidate_json)
            except json.JSONDecodeError:
                decoded = {}
            candidate = {str(key): str(value) for key, value in decoded.items()} if isinstance(decoded, dict) else {}
            result.append((str(message_id), str(record_id), candidate))
        return result

    def complete_candidate_backup(self, message_id: str, backup_path: Path) -> None:
        self.connection.execute(
            """
            UPDATE candidate_backup_state
            SET status='completed', backup_path=?, error_reason=NULL, updated_at=?
            WHERE message_id=?
            """,
            (str(backup_path), int(time.time()), message_id),
        )
        self.connection.commit()

    def fail_candidate_backup(self, message_id: str, reason: str) -> None:
        self.connection.execute(
            """
            UPDATE candidate_backup_state
            SET status='pending', error_reason=?, updated_at=?
            WHERE message_id=?
            """,
            (reason[:1000], int(time.time()), message_id),
        )
        self.connection.commit()

    def queue_candidate_resume_upload(self, message_id: str, record_id: str) -> None:
        """Queue a Feishu Drive copy without making row insertion dependent on it."""
        self.connection.execute(
            """
            INSERT INTO candidate_resume_upload_state(
                message_id, feishu_record_id, status, files_json, error_reason, updated_at
            ) VALUES (?, ?, 'pending', '{}', NULL, ?)
            ON CONFLICT(message_id) DO UPDATE SET
                feishu_record_id=excluded.feishu_record_id,
                status=CASE
                    WHEN candidate_resume_upload_state.status='completed'
                    THEN 'completed' ELSE 'pending' END,
                error_reason=CASE
                    WHEN candidate_resume_upload_state.status='completed'
                    THEN candidate_resume_upload_state.error_reason ELSE NULL END,
                updated_at=excluded.updated_at
            """,
            (message_id, record_id, int(time.time())),
        )
        self.connection.commit()

    def queue_completed_backups_for_resume_upload(self) -> int:
        """Backfill old locally archived résumés after this feature is enabled."""
        rows = self.connection.execute(
            """
            SELECT message_id, feishu_record_id
            FROM candidate_backup_state
            WHERE status='completed' AND feishu_record_id != ''
            """
        ).fetchall()
        for message_id, record_id in rows:
            self.queue_candidate_resume_upload(str(message_id), str(record_id))
        return len(rows)

    def pending_candidate_resume_uploads(self) -> list[tuple[str, str, dict[str, dict[str, str]], str, str]]:
        """Return pending Drive uploads together with their local backup state."""
        rows = self.connection.execute(
            """
            SELECT uploads.message_id, uploads.feishu_record_id, uploads.files_json,
                   COALESCE(backups.status, ''), COALESCE(backups.backup_path, '')
            FROM candidate_resume_upload_state AS uploads
            LEFT JOIN candidate_backup_state AS backups ON backups.message_id=uploads.message_id
            WHERE uploads.status != 'completed'
            ORDER BY uploads.updated_at
            """
        ).fetchall()
        result: list[tuple[str, str, dict[str, dict[str, str]], str, str]] = []
        for message_id, record_id, files_json, backup_status, backup_path in rows:
            try:
                decoded = json.loads(files_json)
            except json.JSONDecodeError:
                decoded = {}
            uploaded_files: dict[str, dict[str, str]] = {}
            if isinstance(decoded, dict):
                for key, value in decoded.items():
                    if isinstance(value, dict):
                        uploaded_files[str(key)] = {str(k): str(v) for k, v in value.items()}
            result.append((str(message_id), str(record_id), uploaded_files, str(backup_status), str(backup_path)))
        return result

    def save_candidate_resume_upload_progress(
        self,
        message_id: str,
        uploaded_files: dict[str, dict[str, str]],
    ) -> None:
        self.connection.execute(
            """
            UPDATE candidate_resume_upload_state
            SET files_json=?, status='pending', error_reason=NULL, updated_at=?
            WHERE message_id=?
            """,
            (json.dumps(uploaded_files, ensure_ascii=False, sort_keys=True), int(time.time()), message_id),
        )
        self.connection.commit()

    def complete_candidate_resume_upload(self, message_id: str) -> None:
        self.connection.execute(
            """
            UPDATE candidate_resume_upload_state
            SET status='completed', error_reason=NULL, updated_at=?
            WHERE message_id=?
            """,
            (int(time.time()), message_id),
        )
        self.connection.commit()

    def fail_candidate_resume_upload(self, message_id: str, reason: str) -> None:
        self.connection.execute(
            """
            UPDATE candidate_resume_upload_state
            SET status='pending', error_reason=?, updated_at=?
            WHERE message_id=?
            """,
            (reason[:1000], int(time.time()), message_id),
        )
        self.connection.commit()

    def interview_notice_sent(self, record_range: str, recipient_email: str, interview_link: str) -> bool:
        rows = self.connection.execute(
            "SELECT record_range, recipient_email, interview_link FROM interview_link_email_state"
        ).fetchall()
        return any(
            self._same_record_range(str(existing_range), record_range)
            and str(existing_email) == recipient_email
            and str(existing_link) == interview_link
            for existing_range, existing_email, existing_link in rows
        )

    def save_interview_notice(
        self,
        record_range: str,
        recipient_email: str,
        interview_link: str,
        gmail_message_id: str,
    ) -> None:
        existing = self.connection.execute(
            "SELECT record_range FROM interview_link_email_state"
        ).fetchall()
        previous_range = next(
            (str(item[0]) for item in existing if self._same_record_range(str(item[0]), record_range)),
            "",
        )
        if previous_range and previous_range != record_range:
            self.connection.execute(
                """
                UPDATE interview_link_email_state
                SET record_range=?, recipient_email=?, interview_link=?, gmail_message_id=?, sent_at=?
                WHERE record_range=?
                """,
                (record_range, recipient_email, interview_link, gmail_message_id, int(time.time()), previous_range),
            )
            self.connection.commit()
            return
        self.connection.execute(
            """
            INSERT INTO interview_link_email_state(
                record_range, recipient_email, interview_link, gmail_message_id, sent_at
            ) VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(record_range) DO UPDATE SET
                recipient_email=excluded.recipient_email,
                interview_link=excluded.interview_link,
                gmail_message_id=excluded.gmail_message_id,
                sent_at=excluded.sent_at
            """,
            (record_range, recipient_email, interview_link, gmail_message_id, int(time.time())),
        )
        self.connection.commit()

    def offer_review(
        self,
        record_range: str,
        recipient_email: str,
        interview_link: str,
    ) -> dict[str, str] | None:
        row = self.connection.execute(
            """
            SELECT review_id, candidate_name, scheduled_at, offer_subject, offer_content,
                   reviewer_email, review_message_id, status
            FROM interview_offer_review_state
            WHERE record_range=? AND recipient_email=? AND interview_link=?
            """,
            (record_range, recipient_email, interview_link),
        ).fetchone()
        if not row:
            candidates = self.connection.execute(
                """
                SELECT review_id, record_range, candidate_name, scheduled_at, offer_subject, offer_content,
                       reviewer_email, review_message_id, status
                FROM interview_offer_review_state
                WHERE recipient_email=? AND interview_link=?
                ORDER BY updated_at DESC
                """,
                (recipient_email, interview_link),
            ).fetchall()
            row = next(
                (
                    (item[0], item[2], item[3], item[4], item[5], item[6], item[7], item[8])
                    for item in candidates
                    if self._same_record_range(str(item[1]), record_range)
                ),
                None,
            )
        if not row:
            return None
        keys = (
            "review_id",
            "candidate_name",
            "scheduled_at",
            "offer_subject",
            "offer_content",
            "reviewer_open_id",
            "review_message_id",
            "status",
        )
        return {key: str(value or "") for key, value in zip(keys, row)}

    def create_offer_review(
        self,
        *,
        review_id: str | None = None,
        record_range: str,
        recipient_email: str,
        interview_link: str,
        candidate_name: str,
        scheduled_at: str,
        offer_subject: str,
        offer_content: str,
        reviewer_open_id: str,
        review_message_id: str,
    ) -> str:
        review_id = review_id or "OFFER-" + secrets.token_hex(6).upper()
        now = int(time.time())
        self.connection.execute(
            """
            INSERT INTO interview_offer_review_state(
                review_id, record_range, recipient_email, interview_link,
                candidate_name, scheduled_at, offer_subject, offer_content,
                reviewer_email, review_message_id, status, decided_message_id,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', NULL, ?, ?)
            """,
            (
                review_id,
                record_range,
                recipient_email,
                interview_link,
                candidate_name,
                scheduled_at,
                offer_subject,
                offer_content,
                reviewer_open_id,
                review_message_id,
                now,
                now,
            ),
        )
        self.connection.commit()
        return review_id

    def pending_offer_reviews(self) -> list[dict[str, str]]:
        rows = self.connection.execute(
            """
            SELECT review_id, reviewer_email, review_message_id
            FROM interview_offer_review_state
            WHERE status='pending'
            ORDER BY created_at
            """
        ).fetchall()
        return [
            {"review_id": str(review_id), "reviewer_email": str(email), "review_message_id": str(message_id)}
            for review_id, email, message_id in rows
        ]

    def offer_review_by_id(self, review_id: str) -> dict[str, str] | None:
        row = self.connection.execute(
            """
            SELECT review_id, record_range, recipient_email, interview_link,
                   candidate_name, reviewer_email, status
            FROM interview_offer_review_state
            WHERE review_id=?
            """,
            (review_id,),
        ).fetchone()
        if not row:
            return None
        keys = (
            "review_id",
            "record_range",
            "recipient_email",
            "interview_link",
            "candidate_name",
            "reviewer_open_id",
            "status",
        )
        return {key: str(value or "") for key, value in zip(keys, row)}

    def decide_offer_review(self, review_id: str, status: str, decision_message_id: str) -> None:
        if status not in {"approved", "rejected"}:
            raise ValueError("Offer review status must be approved or rejected")
        self.connection.execute(
            """
            UPDATE interview_offer_review_state
            SET status=?, decided_message_id=?, updated_at=?
            WHERE review_id=? AND status='pending'
            """,
            (status, decision_message_id, int(time.time()), review_id),
        )
        self.connection.commit()


class CandidateBackup:
    """Maintain one local, atomic archive folder per successfully stored résumé."""

    # Keep only files that can be delivered and linked as original résumés.
    supported_suffixes = {".pdf", ".docx"} | IMAGE_SUFFIXES

    def __init__(self, config: dict[str, Any], project_root: Path):
        settings = config.get("backup", {}) if isinstance(config.get("backup", {}), dict) else {}
        self.enabled = bool(settings.get("enabled", True))
        directory = str(settings.get("directory", "data/candidate-backups"))
        self.root = resolve_path(project_root, directory)

    @staticmethod
    def _safe_component(value: str, fallback: str) -> str:
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", Path(value).name).strip(" .")
        return cleaned[:160] or fallback

    @staticmethod
    def _write_bytes(path: Path, content: bytes) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_name(path.name + ".tmp")
        temporary.write_bytes(content)
        os.replace(temporary, path)

    def archive(
        self,
        message_id: str,
        record_id: str,
        candidate: dict[str, str],
        gmail: GmailClient,
        message: dict[str, Any],
    ) -> Path:
        if not self.enabled:
            return self.root
        directory = self.root / self._safe_component(message_id, "gmail-message")
        _body, attachments = message_content(gmail, message)
        saved_attachments: list[dict[str, str]] = []
        for index, (filename, raw) in enumerate(attachments, start=1):
            suffix = Path(filename).suffix.lower()
            if suffix not in self.supported_suffixes:
                continue
            safe_name = self._safe_component(filename, f"resume-{index}{suffix}")
            saved_name = f"{index:02d}_{safe_name}"
            self._write_bytes(directory / saved_name, raw)
            saved_attachments.append({"original_name": filename, "backup_name": saved_name})
        manifest = {
            "gmail_message_id": message_id,
            "feishu_record_id": record_id,
            "backed_up_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
            "candidate": candidate,
            "resume_attachments": saved_attachments,
        }
        write_private_json(directory / "candidate.json", manifest)
        return directory


def flush_candidate_backups(
    state: SyncState,
    backup: CandidateBackup,
    gmail: GmailClient,
    message_cache: dict[str, dict[str, Any]],
) -> dict[str, int]:
    """Complete pending archives after Feishu insertion without duplicating rows."""
    result = {"completed": 0, "failed": 0, "disabled": 0}
    if not backup.enabled:
        result["disabled"] = len(state.pending_candidate_backups())
        return result
    for message_id, record_id, candidate in state.pending_candidate_backups():
        try:
            message = message_cache.get(message_id) or gmail.message(message_id)
            destination = backup.archive(message_id, record_id, candidate, gmail, message)
            state.complete_candidate_backup(message_id, destination)
            result["completed"] += 1
            logging.info("Candidate résumé backup completed.")
        except (OSError, SyncError) as exc:
            state.fail_candidate_backup(message_id, str(exc))
            result["failed"] += 1
            logging.error("Candidate résumé backup failed and will retry: %s", exc)
    return result


class FeishuClient:
    def __init__(self, config: dict[str, Any], config_file: Path | None = None):
        self.full_config = config
        self.config = config["feishu"]
        self.config_file = config_file
        self._token: str | None = None
        self._expires_at = 0.0
        self._sheet_context: dict[str, Any] | None = None

    def token(self) -> str:
        if self._token and self._expires_at > time.time() + 60:
            return self._token
        required = ("app_id", "app_secret")
        if any(not str(self.config.get(key, "")).strip() or "replace-with" in str(self.config.get(key, "")) for key in required):
            raise SyncError("Feishu app_id or app_secret is missing in config.json")
        response = json_request(
            FEISHU_API + "/auth/v3/tenant_access_token/internal/",
            method="POST",
            payload={"app_id": self.config["app_id"], "app_secret": self.config["app_secret"]},
        )
        if response.get("code") != 0 or not get_nested(response, "tenant_access_token"):
            raise SyncError(f"Feishu token request failed: {response.get('msg', 'unknown error')}")
        self._token = str(response["tenant_access_token"])
        self._expires_at = time.time() + int(response.get("expire", 7200))
        return self._token

    def call(self, method: str, path: str, *, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        response = json_request(
            FEISHU_API + path,
            method=method,
            headers={"Authorization": f"Bearer {self.token()}"},
            payload=payload,
        )
        if response.get("code") != 0:
            raise SyncError(f"Feishu API failed: {response.get('msg', 'unknown error')} (code {response.get('code')})")
        return response

    def send_bot_text(self, open_id: str, text: str) -> str:
        """Send a private, plain-text bot message to one Feishu user."""
        response = json_request(
            FEISHU_API + "/im/v1/messages?receive_id_type=open_id",
            method="POST",
            headers={"Authorization": f"Bearer {self.token()}"},
            payload={
                "receive_id": open_id,
                "msg_type": "text",
                "content": json.dumps({"text": text}, ensure_ascii=True),
            },
        )
        if response.get("code") != 0 or not get_nested(response, "data.message_id"):
            raise SyncError(f"Feishu private message failed: {response.get('msg', 'unknown error')}")
        return str(response["data"]["message_id"])

    def send_bot_card(self, open_id: str, card: dict[str, Any]) -> str:
        """Send a private interactive card to one Feishu user."""
        response = json_request(
            FEISHU_API + "/im/v1/messages?receive_id_type=open_id",
            method="POST",
            headers={"Authorization": f"Bearer {self.token()}"},
            payload={
                "receive_id": open_id,
                "msg_type": "interactive",
                "content": json.dumps(card, ensure_ascii=True),
            },
        )
        if response.get("code") != 0 or not get_nested(response, "data.message_id"):
            raise SyncError(f"Feishu private card failed: {response.get('msg', 'unknown error')}")
        return str(response["data"]["message_id"])

    def sheet_context(self) -> dict[str, Any]:
        if self._sheet_context is not None:
            return self._sheet_context
        spreadsheet_token = str(self.config.get("spreadsheet_token", "")).strip()
        if not spreadsheet_token or spreadsheet_token.startswith("replace-with"):
            raise SyncError("Feishu spreadsheet_token is missing in config.json")
        sheets_response = self.call(
            "GET", f"/sheets/v3/spreadsheets/{parse.quote(spreadsheet_token)}/sheets/query"
        )
        sheets = [item for item in sheets_response.get("data", {}).get("sheets", []) if not item.get("hidden")]
        configured_sheet_id = str(self.config.get("sheet_id", "")).strip()
        if configured_sheet_id:
            selected = next((item for item in sheets if str(item.get("sheet_id")) == configured_sheet_id), None)
            if selected is None:
                raise SyncError("Configured Feishu sheet_id does not exist or is hidden")
        elif len(sheets) == 1:
            selected = sheets[0]
        elif not sheets:
            raise SyncError("The Feishu spreadsheet has no visible worksheet")
        else:
            available = ", ".join(f"{item.get('title')} ({item.get('sheet_id')})" for item in sheets)
            raise SyncError("Multiple visible worksheets found. Set feishu.sheet_id in config.json. Available: " + available)
        sheet_id = str(selected["sheet_id"])
        encoded_range = parse.quote(f"{sheet_id}!A1:Z1", safe="!")
        header_response = self.call(
            "GET", f"/sheets/v2/spreadsheets/{parse.quote(spreadsheet_token)}/values/{encoded_range}"
        )
        values = header_response.get("data", {}).get("valueRange", {}).get("values", [])
        if not values or not isinstance(values[0], list):
            raise SyncError("The first row of the selected Feishu worksheet must contain column headers")
        headers = [compact_value("" if value is None else str(value)) for value in values[0]]
        while headers and not headers[-1]:
            headers.pop()
        if not any(headers):
            raise SyncError("The first row of the selected Feishu worksheet has no column headers")
        self._sheet_context = {
            "spreadsheet_token": spreadsheet_token,
            "sheet_id": sheet_id,
            "sheet_title": str(selected.get("title", "")),
            "headers": headers,
            "row_count": int((selected.get("grid_properties") or {}).get("row_count") or 0),
            "column_count": int((selected.get("grid_properties") or {}).get("column_count") or 0),
        }
        return self._sheet_context

    def next_serial(self, context: dict[str, Any], serial_header: str) -> str:
        headers = list(context["headers"])
        if serial_header not in headers:
            raise SyncError(f"The serial-number header '{serial_header}' is missing from the Feishu spreadsheet")
        column = column_name(headers.index(serial_header) + 1)
        value_range = f"{context['sheet_id']}!{column}2:{column}5000"
        response = self.call(
            "GET",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values/"
            f"{parse.quote(value_range, safe='!')}",
        )
        values = response.get("data", {}).get("valueRange", {}).get("values", [])
        highest = 0
        for row in values:
            if not isinstance(row, list) or not row:
                continue
            match = re.fullmatch(r"\d+", cell_text(row[0]))
            if match:
                highest = max(highest, int(match.group(0)))
        return str(highest + 1)

    def first_empty_candidate_row(self, context: dict[str, Any]) -> int:
        """Return the first fully blank row below the spreadsheet headers."""
        headers = list(context["headers"])
        right_column = column_name(len(headers))
        grid_rows = max(int(context.get("row_count") or 0), 2)
        value_range = f"{context['sheet_id']}!A2:{right_column}{grid_rows}"
        response = self.call(
            "GET",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values/"
            f"{parse.quote(value_range, safe='!')}",
        )
        values = response.get("data", {}).get("valueRange", {}).get("values", [])
        for row_number, current_row in enumerate(values, start=2):
            if not isinstance(current_row, list) or not any(cell_text(cell).strip() for cell in current_row):
                return row_number
        return max(2, grid_rows + 1)

    def create_candidate(self, candidate: dict[str, str]) -> str:
        context = self.sheet_context()
        configured_fields = self.config.get("fields", {})
        source_keys = ("name", "student_id", "email", "phone", "major", "group", "interview_link")
        field_headers = {key: str(configured_fields.get(key, "")).strip() for key in source_keys}
        serial_header = str(configured_fields.get("serial", "")).strip()
        if not serial_header or any(not header for header in field_headers.values()):
            raise SyncError("Feishu field mapping is incomplete in config.json")
        source_by_header = {
            # A missing résumé field is a valid empty cell, never a reason to
            # reject a candidate or withhold their interview invitation.
            header: candidate.get(source_key, "").strip()
            for source_key, header in field_headers.items()
        }
        source_by_header[serial_header] = self.next_serial(context, serial_header)
        headers = list(context["headers"])
        missing_headers = set(source_by_header) - set(headers)
        if missing_headers:
            raise SyncError("These Feishu spreadsheet headers are missing: " + ", ".join(sorted(missing_headers)))
        row = [source_by_header.get(header, "") for header in headers]
        target_row = self.first_empty_candidate_row(context)
        record_range = f"{context['sheet_id']}!A{target_row}:{column_name(len(headers))}{target_row}"
        self.call(
            "PUT",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values",
            payload={"valueRange": {"range": record_range, "values": [row]}},
        )
        return record_range

    def ensure_resume_original_column(self) -> str:
        """Create the single column used for original résumé links when needed."""
        configured_fields = self.config.get("fields", {})
        header = str(configured_fields.get("resume_original", "简历原件")).strip()
        if not header:
            raise SyncError("Feishu resume_original field mapping is empty")
        context = self.sheet_context()
        if header in context["headers"]:
            return header
        column = column_name(len(context["headers"]) + 1)
        record_range = f"{context['sheet_id']}!{column}1:{column}1"
        self.call(
            "PUT",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values",
            payload={"valueRange": {"range": record_range, "values": [[header]]}},
        )
        # Header changes must be visible to subsequent candidate row writes.
        self._sheet_context = None
        return header

    def write_candidate_resume_links(self, record_range: str, links: list[tuple[str, str]]) -> None:
        """Write clickable raw Drive URLs for the original file(s) into one row."""
        match = re.search(r"![A-Z]+(\d+):", record_range)
        if not match:
            raise SyncError("Stored Feishu range has no row number")
        header = self.ensure_resume_original_column()
        context = self.sheet_context()
        if header not in context["headers"]:
            raise SyncError("The Feishu résumé-original column could not be created")
        column = column_name(context["headers"].index(header) + 1)
        row_number = match.group(1)
        # Keeping the original filename beside each URL makes multiple
        # attachments unambiguous while Feishu recognizes the URL as a link.
        value = "\n".join(f"{name}: {url}" for name, url in links)
        value_range = f"{context['sheet_id']}!{column}{row_number}:{column}{row_number}"
        self.call(
            "PUT",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values",
            payload={"valueRange": {"range": value_range, "values": [[value]]}},
        )

    def candidate_record_matches_backup(
        self,
        record_range: str,
        candidate: dict[str, Any],
        attachment_names: list[str],
    ) -> bool:
        """Prevent an old local backup from being linked into a reused row.

        The operator can clear and reuse a spreadsheet while preserving local
        sync history.  Before backfilling a historical attachment, require a
        strong identifier match (student number, email, phone) or a matching
        name/filename in the current row.
        """
        match = re.search(r"![A-Z]+(\d+):", record_range)
        if not match:
            return False
        context = self.sheet_context()
        row_number = match.group(1)
        right_column = column_name(len(context["headers"]))
        value_range = f"{context['sheet_id']}!A{row_number}:{right_column}{row_number}"
        response = self.call(
            "GET",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values/"
            f"{parse.quote(value_range, safe='!')}",
        )
        rows = response.get("data", {}).get("valueRange", {}).get("values", [])
        if not rows or not isinstance(rows[0], list):
            return False
        row = rows[0]
        configured_fields = self.config.get("fields", {})
        actual: dict[str, str] = {}
        for key in ("name", "student_id", "email", "phone"):
            header = str(configured_fields.get(key, "")).strip()
            if header and header in context["headers"]:
                index = context["headers"].index(header)
                actual[key] = cell_text(row[index]).strip() if index < len(row) else ""
        for key in ("student_id", "email", "phone"):
            expected = compact_value(str(candidate.get(key, "")))
            if expected and expected == actual.get(key, ""):
                return True
        expected_name = compact_value(str(candidate.get("name", "")))
        actual_name = actual.get("name", "")
        if expected_name and expected_name == actual_name:
            return True
        return bool(actual_name and any(actual_name in filename for filename in attachment_names))

    @staticmethod
    def _folder_token(value: str) -> str:
        """Accept either a Drive folder token or a complete folder URL."""
        raw = value.strip()
        if not raw:
            return ""
        parsed = parse.urlparse(raw)
        if parsed.scheme and parsed.netloc:
            parts = [part for part in parsed.path.split("/") if part]
            if "folder" in parts:
                index = parts.index("folder")
                if len(parts) > index + 1:
                    return parts[index + 1]
            return ""
        return raw

    def resume_archive_settings(self) -> tuple[bool, str]:
        settings = self.config.get("resume_archive", {})
        if not isinstance(settings, dict):
            return False, ""
        enabled = bool(settings.get("enabled", False))
        folder_value = str(settings.get("folder_token", "") or settings.get("folder_url", ""))
        return enabled, self._folder_token(folder_value)

    def ensure_resume_archive_folder(self) -> str:
        """Create the configured Drive destination once and persist its token."""
        enabled, folder_token = self.resume_archive_settings()
        if not enabled:
            return ""
        if folder_token:
            return folder_token
        settings = self.config.setdefault("resume_archive", {})
        if not isinstance(settings, dict):
            raise SyncError("Feishu resume_archive settings must be an object")
        folder_name = str(settings.get("folder_name", "招新简历原件")).strip() or "招新简历原件"
        response = self.call(
            "POST",
            "/drive/v1/files/create_folder",
            payload={"name": folder_name, "folder_token": ""},
        )
        folder_token = str(get_nested(response, "data.token") or "").strip()
        if not folder_token:
            raise SyncError("Feishu Drive folder creation returned no folder token")
        settings["folder_token"] = folder_token
        if self.config_file:
            write_private_json(self.config_file, self.full_config)
        else:
            logging.warning("Feishu résumé folder token was created but could not be persisted to config.json.")
        return folder_token

    def upload_resume_original(self, source: Path, folder_token: str, *, filename: str | None = None) -> tuple[str, str]:
        """Upload a local, unmodified résumé into a shared Feishu Drive folder."""
        if not source.is_file():
            raise SyncError(f"Local résumé backup is missing: {source}")
        original_name = (filename or source.name).replace("\\", "_").replace("/", "_").strip() or source.name
        content = source.read_bytes()
        if not content:
            raise SyncError(f"Local résumé backup is empty: {source.name}")
        if len(content) > 20 * 1024 * 1024:
            raise SyncError(f"Résumé exceeds Feishu single-upload limit (20MB): {source.name}")
        response = multipart_file_request(
            FEISHU_API + "/drive/v1/files/upload_all",
            headers={"Authorization": f"Bearer {self.token()}"},
            fields={
                "file_name": original_name,
                "parent_type": "explorer",
                "parent_node": folder_token,
                "size": str(len(content)),
            },
            file_field="file",
            filename=original_name,
            content=content,
            timeout=60,
        )
        if response.get("code") != 0:
            raise SyncError(f"Feishu Drive upload failed: {response.get('msg', 'unknown error')}")
        file_token = str(get_nested(response, "data.file_token") or "").strip()
        if not file_token:
            raise SyncError("Feishu Drive upload returned no file token")
        settings = self.config.get("resume_archive", {})
        base_url = str(settings.get("file_url_base", "https://feishu.cn")).strip().rstrip("/")
        if not base_url.startswith(("https://", "http://")):
            raise SyncError("Feishu resume_archive.file_url_base must be an https URL")
        return file_token, f"{base_url}/file/{parse.quote(file_token, safe='')}"

    def update_candidate(self, record_range: str, candidate: dict[str, str]) -> str:
        """Rewrite extracted candidate fields in an existing spreadsheet row.

        This is used when a parser improvement corrects a résumé that was
        already synchronized.  The serial number and interview fields are left
        untouched.
        """
        match = re.search(r"![A-Z]+(\d+):", record_range)
        if not match:
            raise SyncError("Stored Feishu range has no row number")
        context = self.sheet_context()
        configured_fields = self.config.get("fields", {})
        source_keys = ("name", "student_id", "email", "phone", "major", "group")
        field_headers = {key: str(configured_fields.get(key, "")).strip() for key in source_keys}
        if any(not header for header in field_headers.values()):
            raise SyncError("Feishu field mapping is incomplete in config.json")
        headers = list(context["headers"])
        missing_headers = set(field_headers.values()) - set(headers)
        if missing_headers:
            raise SyncError("These Feishu spreadsheet headers are missing: " + ", ".join(sorted(missing_headers)))
        row_number = match.group(1)
        for source_key, header in field_headers.items():
            column = column_name(headers.index(header) + 1)
            value_range = f"{context['sheet_id']}!{column}{row_number}:{column}{row_number}"
            self.call(
                "PUT",
                f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values",
                payload={"valueRange": {"range": value_range, "values": [[candidate.get(source_key, "").strip()]]}},
            )
        return record_range

    @staticmethod
    def _display_units(value: str) -> int:
        """Approximate rendered width: CJK glyphs occupy about two ASCII units."""
        return sum(1 if ord(character) < 128 else 2 for character in value)

    def fit_sheet_dimensions(self, row_limit: int = 500) -> dict[str, int]:
        """Resize recruiting-sheet rows and columns to fit current cell content.

        The Feishu Sheets API exposes fixed pixel sizes rather than its desktop
        client's live "auto fit" switch. Recalculating those sizes after each
        new candidate gives the sheet equivalent content-aware sizing.
        """
        context = self.sheet_context()
        headers = list(context["headers"])
        if not headers:
            return {"columns": 0, "rows": 0}
        right_column = column_name(len(headers))
        read_range = f"{context['sheet_id']}!A1:{right_column}{row_limit}"
        response = self.call(
            "GET",
            f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values/"
            f"{parse.quote(read_range, safe='!')}",
        )
        values = response.get("data", {}).get("valueRange", {}).get("values", [])
        configured = self.config.get("fields", {})
        interview = self.config.get("interview", {}) if isinstance(self.config.get("interview", {}), dict) else {}
        interview_fields = interview.get("fields", {}) if isinstance(interview.get("fields", {}), dict) else {}
        minimum_widths = {
            str(configured.get("serial", "")): 72,
            str(configured.get("name", "")): 96,
            str(configured.get("student_id", "")): 132,
            str(configured.get("email", "")): 230,
            str(configured.get("phone", "")): 150,
            str(configured.get("major", "")): 220,
            str(configured.get("group", "")): 220,
            str(configured.get("interview_link", "")): 300,
            str(interview_fields.get("scheduled_at", "")): 190,
            str(interview_fields.get("status", "")): 145,
            str(interview_fields.get("interviewer", "")): 125,
            "简历原件": 460,
            "备注": 360,
        }
        spreadsheet_token = parse.quote(str(context["spreadsheet_token"]))
        for index, header in enumerate(headers):
            widest = self._display_units(header)
            for row in values:
                if isinstance(row, list) and index < len(row):
                    widest = max(widest, self._display_units(cell_text(row[index])))
            width = max(minimum_widths.get(header, 110), min(460, 36 + widest * 9))
            self.call(
                "PUT",
                f"/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range",
                payload={
                    "dimension": {
                        "sheetId": context["sheet_id"],
                        "majorDimension": "COLUMNS",
                        # The Sheets dimension API uses one-based, inclusive
                        # indexes (unlike most Python collections).
                        "startIndex": index + 1,
                        "endIndex": index + 1,
                    },
                    "dimensionProperties": {"fixedSize": width},
                },
            )
        existing_column_count = int(context.get("column_count") or 0)
        if existing_column_count > len(headers):
            self.call(
                "PUT",
                f"/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range",
                payload={
                    "dimension": {
                        "sheetId": context["sheet_id"],
                        "majorDimension": "COLUMNS",
                        "startIndex": len(headers) + 1,
                        "endIndex": existing_column_count,
                    },
                    "dimensionProperties": {"fixedSize": 20},
                },
            )
        existing_row_count = int(context.get("row_count") or 0)
        row_count = max(len(values), 2)
        if existing_row_count:
            row_count = min(row_count, existing_row_count)
        self.call(
            "PUT",
            f"/sheets/v2/spreadsheets/{spreadsheet_token}/dimension_range",
            payload={
                "dimension": {
                    "sheetId": context["sheet_id"],
                    "majorDimension": "ROWS",
                    "startIndex": 1,
                    "endIndex": row_count,
                },
                "dimensionProperties": {"fixedSize": 36},
            },
        )
        return {"columns": len(headers), "rows": row_count}


def flush_feishu_resume_uploads(state: SyncState, feishu: FeishuClient) -> dict[str, int]:
    """Upload original local backups and add their Drive links to sheet rows.

    This deliberately runs after ordinary row insertion and local archival.
    A missing Drive scope, inaccessible folder, or one oversized attachment
    must never duplicate a candidate or prevent the interview workflow.
    """
    pending = state.pending_candidate_resume_uploads()
    result = {
        "queued": len(pending),
        "uploaded": 0,
        "completed": 0,
        "waiting": 0,
        "missing": 0,
        "stale": 0,
        "failed": 0,
    }
    enabled, folder_token = feishu.resume_archive_settings()
    if not enabled:
        result["waiting"] = len(pending)
        if pending:
            logging.info("Feishu original-resume sync is waiting for a shared Drive folder token.")
        return result
    if not folder_token:
        try:
            folder_token = feishu.ensure_resume_archive_folder()
        except SyncError as exc:
            result["waiting"] = len(pending)
            logging.error("Feishu original-resume sync is waiting for Drive folder creation: %s", exc)
            return result
    for message_id, record_id, uploaded_files, backup_status, backup_path in pending:
        if backup_status != "completed" or not backup_path:
            result["waiting"] += 1
            continue
        directory = Path(backup_path)
        manifest_path = directory / "candidate.json"
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            state.fail_candidate_resume_upload(message_id, f"Could not read local résumé manifest: {exc}")
            result["failed"] += 1
            continue
        attachments = manifest.get("resume_attachments", []) if isinstance(manifest, dict) else []
        if not isinstance(attachments, list):
            attachments = []
        # Only original PDF and DOCX files are eligible for Feishu Drive.
        # A legacy .doc or a cloud-mail link intentionally leaves the cell blank.
        attachments = [
            item
            for item in attachments
            if isinstance(item, dict)
            and Path(str(item.get("original_name", ""))).suffix.lower() in {".pdf", ".docx"}
        ]
        if not attachments:
            # The source e-mail carried no real PDF/DOC/DOCX attachment (for
            # example a mailbox 'super attachment' link).  It has no original
            # file that can be copied into Drive, so leave the sheet cell blank.
            state.complete_candidate_resume_upload(message_id)
            result["missing"] += 1
            continue
        candidate = manifest.get("candidate", {}) if isinstance(manifest, dict) else {}
        candidate = candidate if isinstance(candidate, dict) else {}
        attachment_names = [
            str(item.get("original_name", "")) for item in attachments if isinstance(item, dict)
        ]
        try:
            if not feishu.candidate_record_matches_backup(record_id, candidate, attachment_names):
                # The row was likely cleared and reused after the local backup
                # was made.  Never put a former applicant's résumé on a newer
                # applicant's line.
                state.complete_candidate_resume_upload(message_id)
                result["stale"] += 1
                logging.warning("Skipped a historical résumé because its stored sheet row is now different.")
                continue
        except SyncError as exc:
            state.fail_candidate_resume_upload(message_id, str(exc))
            result["failed"] += 1
            continue
        try:
            links: list[tuple[str, str]] = []
            for attachment in attachments:
                if not isinstance(attachment, dict):
                    continue
                backup_name = str(attachment.get("backup_name", "")).strip()
                original_name = str(attachment.get("original_name", "")).strip()
                if not backup_name or not original_name:
                    continue
                uploaded = uploaded_files.get(backup_name, {})
                url = str(uploaded.get("url", "")).strip()
                if not url:
                    file_token, url = feishu.upload_resume_original(
                        directory / backup_name,
                        folder_token,
                        filename=original_name,
                    )
                    uploaded_files[backup_name] = {"file_token": file_token, "name": original_name, "url": url}
                    state.save_candidate_resume_upload_progress(message_id, uploaded_files)
                    result["uploaded"] += 1
                links.append((original_name, url))
            if not links:
                state.complete_candidate_resume_upload(message_id)
                result["missing"] += 1
                continue
            feishu.write_candidate_resume_links(record_id, links)
            state.complete_candidate_resume_upload(message_id)
            result["completed"] += 1
            logging.info("Original résumé link synchronized to Feishu sheet.")
        except (OSError, SyncError) as exc:
            state.fail_candidate_resume_upload(message_id, str(exc))
            result["failed"] += 1
            logging.error("Original résumé upload will retry: %s", exc)
    return result


def valid_interview_link(value: str) -> bool:
    """Accept only ordinary web meeting links, never arbitrary cell text."""
    parsed = parse.urlparse(value)
    return parsed.scheme in {"http", "https"} and bool(parsed.netloc)


def interview_notice_content(
    config: dict[str, Any],
    candidate_name: str,
    scheduled_at: str,
    interview_link: str,
) -> tuple[str, str]:
    interview = config.get("interview", {}) if isinstance(config.get("interview", {}), dict) else {}
    configured_notice = interview.get("candidate_notice", {})
    notice = configured_notice if isinstance(configured_notice, dict) else {}
    organization = compact_value(str(notice.get("organization", "重庆邮电大学2026数学与统计学院110实验室")))
    contact_name = compact_value(str(notice.get("contact_name", "张天齐")))
    contact_qq = compact_value(str(notice.get("contact_qq", "3882313752")))
    contact_email = compact_value(str(notice.get("contact_email", "zhangtianqi_oms@outlook.com")))
    meeting_mode = compact_value(str(notice.get("meeting_mode", "视频面试")))
    meeting_location = compact_value(str(notice.get("meeting_location", "线上面试")))
    meeting_hint = compact_value(str(notice.get("meeting_hint", "请使用腾讯会议电脑客户端进行面试/笔试")))
    date_value, time_value = "2026-XX-XX", "XX:XX"
    scheduled_match = re.fullmatch(r"(\d{4}-\d{2}-\d{2})[ T]+(\d{1,2}:\d{2})", scheduled_at)
    if scheduled_match:
        date_value, time_value = scheduled_match.group(1), scheduled_match.group(2).zfill(5)
    subject = f"【{organization}】面试/笔试通知"
    recipient = candidate_name or "同学"
    lines = [
        f"{recipient}，你好：",
        "",
        f"您有一场{organization}的面试/笔试，请您安排好时间准时参加，有任何变动请随时告知。",
        "",
        f"面试日期：{date_value}",
        f"面试时间：{time_value}",
        f"面试方式：{meeting_mode}",
        f"面试地点：{meeting_location}",
        "",
        f"面试链接：{interview_link}",
        meeting_hint,
        "面试要求：面试时需打开摄像头。",
        "",
        f"联系人：{contact_name}",
        f"QQ：{contact_qq}",
        f"联系邮箱：{contact_email}",
    ]
    return subject, "\n".join(lines)


def offer_review_required(config: dict[str, Any]) -> bool:
    interview = config.get("interview", {}) if isinstance(config.get("interview", {}), dict) else {}
    notice = interview.get("candidate_notice", {}) if isinstance(interview.get("candidate_notice", {}), dict) else {}
    return bool(notice.get("review_required", True))


def offer_reviewer_open_id(config: dict[str, Any]) -> str:
    """Use the existing private resume-audit recipient for Offer approval."""
    interview = config.get("interview", {}) if isinstance(config.get("interview", {}), dict) else {}
    notice = interview.get("candidate_notice", {}) if isinstance(interview.get("candidate_notice", {}), dict) else {}
    configured = compact_value(str(notice.get("reviewer_open_id", "")))
    if configured:
        return configured
    reviewer = interview.get("resume_reviewer", {})
    return compact_value(str(reviewer.get("open_id", ""))) if isinstance(reviewer, dict) else ""


def offer_review_message(
    review_id: str,
    candidate_name: str,
    recipient_email: str,
    subject: str,
    content: str,
) -> tuple[str, str]:
    preview_subject = f"【待审核 Offer {review_id}】{candidate_name or '候选人'}"
    preview = "\n".join(
        [
            "以下 Offer 尚未发送给候选人，请先核对。",
            "",
            f"候选人：{candidate_name or '未识别'}",
            f"拟发送至：{recipient_email}",
            f"审核编号：{review_id}",
            "",
            "拟发送主题：",
            subject,
            "",
            "拟发送正文：",
            content,
            "",
            "确认无误后，请在与机器人的单聊中发送：",
            f"/批准Offer {review_id}",
            "如不发送，请发送：",
            f"/拒绝Offer {review_id}",
            "",
            "未经批准，系统绝不会向候选人发送。",
        ]
    )
    return preview_subject, preview


def offer_review_card(
    review_id: str,
    candidate_name: str,
    recipient_email: str,
    subject: str,
    content: str,
) -> dict[str, Any]:
    """Build a two-button Offer review card for the configured reviewer."""
    body = (
        f"**候选人**：{candidate_name or '未识别'}\n"
        f"**收件邮箱**：{recipient_email}\n"
        f"**审核编号**：{review_id}\n\n"
        f"**拟发送主题**：{subject}\n\n"
        f"**拟发送正文**：\n{content}"
    )
    return {
        "schema": "2.0",
        "config": {"enable_forward": False, "update_multi": True, "width_mode": "fill"},
        "header": {
            "title": {"tag": "plain_text", "content": "待审核 Offer"},
            "subtitle": {"tag": "plain_text", "content": candidate_name or "候选人"},
            "template": "blue",
        },
        "body": {
            "elements": [
                {"tag": "markdown", "content": body},
                {"tag": "hr"},
                {
                    # Card JSON 2.0 does not support the legacy ``action``
                    # container; place each button in its own column instead.
                    "tag": "column_set",
                    "horizontal_spacing": "8px",
                    "columns": [
                        {
                            "tag": "column",
                            "width": "auto",
                            "elements": [
                                {
                                    "tag": "button",
                                    "text": {"tag": "plain_text", "content": "通过"},
                                    "type": "primary_filled",
                                    "behaviors": [
                                        {"type": "callback", "value": {"action": "approve_offer", "review_id": review_id}}
                                    ],
                                }
                            ],
                        },
                        {
                            "tag": "column",
                            "width": "auto",
                            "elements": [
                                {
                                    "tag": "button",
                                    "text": {"tag": "plain_text", "content": "不通过"},
                                    "type": "danger_filled",
                                    "behaviors": [
                                        {"type": "callback", "value": {"action": "reject_offer", "review_id": review_id}}
                                    ],
                                }
                            ],
                        },
                    ],
                },
            ]
        },
    }


def send_interview_link_notices(
    config: dict[str, Any],
    state: SyncState,
    gmail: GmailClient,
    feishu: FeishuClient,
) -> dict[str, int]:
    """Stage every Offer in a private Feishu bot message before email delivery."""
    result = {
        "sent": 0,
        "already_sent": 0,
        "review_requested": 0,
        "review_pending": 0,
        "review_rejected": 0,
        "missing_email": 0,
        "invalid_link": 0,
        "failed": 0,
    }
    context = feishu.sheet_context()
    headers = list(context["headers"])
    configured_fields = feishu.config.get("fields", {})
    source_keys = ("name", "email", "group", "interview_link")
    field_headers = {key: str(configured_fields.get(key, "")).strip() for key in source_keys}
    if any(not header for header in field_headers.values()):
        raise SyncError("Feishu field mapping is incomplete in config.json")
    missing_headers = set(field_headers.values()) - set(headers)
    if missing_headers:
        raise SyncError("These Feishu spreadsheet headers are missing: " + ", ".join(sorted(missing_headers)))
    interview = config.get("interview", {}) if isinstance(config.get("interview", {}), dict) else {}
    interview_fields = interview.get("fields", {}) if isinstance(interview.get("fields", {}), dict) else {}
    requires_review = offer_review_required(config)
    reviewer_open_id = offer_reviewer_open_id(config)
    if requires_review and not reviewer_open_id:
        raise SyncError("Offer review recipient is not set; send /设置简历审核 to the bot in a private chat")
    scheduled_header = str(interview_fields.get("scheduled_at", "")).strip()
    right_column = column_name(len(headers))
    last_row = max(int(context.get("row_count") or 0), 2)
    read_range = f"{context['sheet_id']}!A2:{right_column}{last_row}"
    response = feishu.call(
        "GET",
        f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values/"
        f"{parse.quote(read_range, safe='!')}",
    )
    rows = response.get("data", {}).get("valueRange", {}).get("values", [])
    for row_number, row in enumerate(rows, start=2):
        if not isinstance(row, list):
            continue

        def row_value(header: str) -> str:
            index = headers.index(header)
            return cell_text(row[index]) if index < len(row) else ""

        interview_link = row_value(field_headers["interview_link"])
        if not interview_link:
            continue
        if not valid_interview_link(interview_link):
            result["invalid_link"] += 1
            logging.warning("Interview link in Feishu row %s is not a valid HTTP(S) link.", row_number)
            continue
        recipient_email = row_value(field_headers["email"])
        if not EMAIL_PATTERN.fullmatch(recipient_email):
            result["missing_email"] += 1
            logging.warning("Feishu row %s has an interview link but no valid candidate email.", row_number)
            continue
        record_range = f"{context['sheet_id']}!A{row_number}:{right_column}{row_number}"
        if state.interview_notice_sent(record_range, recipient_email, interview_link):
            result["already_sent"] += 1
            continue
        scheduled_at = (
            normalize_scheduled_at(row_value(scheduled_header))
            if scheduled_header and scheduled_header in headers
            else ""
        )
        candidate_name = row_value(field_headers["name"])
        subject, content = interview_notice_content(
            config,
            candidate_name,
            scheduled_at,
            interview_link,
        )
        if requires_review:
            review = state.offer_review(record_range, recipient_email, interview_link)
            if review is None:
                review_id = "OFFER-" + secrets.token_hex(6).upper()
                review_card = offer_review_card(
                    review_id,
                    candidate_name,
                    recipient_email,
                    subject,
                    content,
                )
                try:
                    review_message_id = feishu.send_bot_card(reviewer_open_id, review_card)
                    state.create_offer_review(
                        review_id=review_id,
                        record_range=record_range,
                        recipient_email=recipient_email,
                        interview_link=interview_link,
                        candidate_name=candidate_name,
                        scheduled_at=scheduled_at,
                        offer_subject=subject,
                        offer_content=content,
                        reviewer_open_id=reviewer_open_id,
                        review_message_id=review_message_id,
                    )
                    result["review_requested"] += 1
                    logging.info("Offer review requested for Feishu row %s.", row_number)
                except SyncError as exc:
                    result["failed"] += 1
                    logging.error("Offer review request failed for Feishu row %s: %s", row_number, exc)
                continue
            if review["status"] == "pending":
                result["review_pending"] += 1
                continue
            if review["status"] == "rejected":
                result["review_rejected"] += 1
                continue
            if review["status"] != "approved":
                result["review_pending"] += 1
                continue
        try:
            message_id = gmail.send_plain_message(recipient_email, subject, content)
            state.save_interview_notice(record_range, recipient_email, interview_link, message_id)
            result["sent"] += 1
            logging.info("Interview notice email sent for Feishu row %s.", row_number)
        except SyncError as exc:
            result["failed"] += 1
            logging.error("Interview notice email failed for Feishu row %s: %s", row_number, exc)
    return result


def run_interview_link_notice_scan(
    config: dict[str, Any],
    root: Path,
    state: SyncState,
    gmail: GmailClient,
    feishu: FeishuClient,
) -> dict[str, int]:
    """Serialize all interview-link scans so one link can trigger one email."""
    lock = SyncRunLock(root / "data" / "interview-link-notice.lock")
    if not lock.acquire():
        logging.info("An interview-link notification scan is already in progress.")
        return {
            "sent": 0,
            "already_sent": 0,
            "missing_email": 0,
            "invalid_link": 0,
            "failed": 0,
            "skipped": 1,
        }
    try:
        result = send_interview_link_notices(config, state, gmail, feishu)
        result["skipped"] = 0
        return result
    finally:
        lock.release()


def validate_remote(config: dict[str, Any], root: Path) -> None:
    gmail = GmailClient(GmailAuth(config, root))
    label_id = gmail.label_id(str(config["gmail"].get("source_label", "")))
    feishu = FeishuClient(config)
    context = feishu.sheet_context()
    available_fields = set(context["headers"])
    expected = {str(value) for value in config["feishu"].get("fields", {}).values()}
    interview_fields = config.get("interview", {}).get("fields", {}) if isinstance(config.get("interview", {}), dict) else {}
    expected.update(str(value) for value in interview_fields.values())
    missing = expected - available_fields
    if missing:
        raise SyncError("These Feishu spreadsheet headers are missing: " + ", ".join(sorted(missing)))
    logging.info(
        "Connection check passed for Gmail label and Feishu spreadsheet '%s' (%s columns).",
        context["sheet_title"],
        len(available_fields),
    )


def run_sync(config: dict[str, Any], root: Path, config_file: Path | None = None) -> int:
    lock = SyncRunLock(root / "data" / "sync-run.lock")
    if not lock.acquire():
        logging.info("A Gmail-to-Feishu synchronization run is already in progress.")
        print(json.dumps({"skipped": True, "reason": "another_sync_is_running"}, ensure_ascii=False))
        return 0
    try:
        return _run_sync_locked(config, root, config_file)
    finally:
        lock.release()


def _run_sync_locked(config: dict[str, Any], root: Path, config_file: Path | None = None) -> int:
    from interview_scheduler import InterviewCoordinator

    database_file = resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
    state = SyncState(database_file)
    gmail = GmailClient(GmailAuth(config, root))
    feishu = FeishuClient(config, config_file)
    interviews = InterviewCoordinator(config, root)
    backup = CandidateBackup(config, root)
    llm = CandidateLLM(config)
    source_label = str(config["gmail"].get("source_label", "")).strip()
    label_id = gmail.label_id(source_label)
    limit = int(config["gmail"].get("max_messages_per_run", 25))
    messages = gmail.messages(label_id, limit)
    synchronized = invalid = skipped = failed = invitations_queued = 0
    message_cache: dict[str, dict[str, Any]] = {}
    for item in messages:
        message_id = str(item.get("id", ""))
        if not message_id:
            continue
        if state.handled(message_id):
            skipped += 1
            continue
        try:
            message = gmail.message(message_id)
            message_cache[message_id] = message
            candidate, attachment_errors = parse_candidate(gmail, message, llm)
            if attachment_errors:
                logging.warning(
                    "Some resume attachments could not be read; storing detected fields anyway: %s",
                    "; ".join(attachment_errors),
                )
            record_id = feishu.create_candidate(candidate)
            state.save(message_id, "success", record_id=record_id)
            try:
                state.queue_candidate_backup(message_id, record_id, candidate)
                state.queue_candidate_resume_upload(message_id, record_id)
            except sqlite3.Error as exc:
                # Feishu insertion has already succeeded. Never mark it as an
                # error (which would cause a duplicate row) merely because the
                # local backup queue needs attention.
                logging.error("Candidate backup could not be queued: %s", exc)
            interviews.queue_candidate(message_id, record_id, candidate)
            invitations_queued += 1
            synchronized += 1
            logging.info("Candidate record created in Feishu.")
        except SyncError as exc:
            state.save(message_id, "error", reason=str(exc))
            failed += 1
            logging.error("Sync failed for one Gmail message: %s", exc)
    backup_result = flush_candidate_backups(state, backup, gmail, message_cache)
    resume_archive_result = flush_feishu_resume_uploads(state, feishu)
    # Archive first so an enabled interview invitation can send the same local
    # original file without downloading the Gmail attachment a second time.
    invitation_result = interviews.dispatch_pending()
    try:
        interview_email_result = run_interview_link_notice_scan(config, root, state, gmail, feishu)
    except SyncError as exc:
        # A mail notice is an independent follow-up action. Do not roll back
        # candidate synchronization if Gmail sending has not yet been enabled.
        logging.error("Interview notice scan failed: %s", exc)
        interview_email_result = {
            "sent": 0,
            "already_sent": 0,
            "missing_email": 0,
            "invalid_link": 0,
            "failed": 1,
            "skipped": 0,
        }
    layout_result: dict[str, int] | None = None
    if synchronized or interview_email_result["sent"] or resume_archive_result["completed"]:
        try:
            layout_result = feishu.fit_sheet_dimensions()
        except SyncError as exc:
            # Sizing is cosmetic; a successful candidate sync must not be
            # retried or duplicated just because a layout request failed.
            logging.warning("Candidate rows were synchronized but sheet auto-fit failed: %s", exc)
    print(
        json.dumps(
            {
                "found": len(messages),
                "synchronized": synchronized,
                "invalid": invalid,
                "skipped": skipped,
                "failed": failed,
                "invitations_queued": invitations_queued,
                "invitations": invitation_result,
                "backups": backup_result,
                "resume_archive": resume_archive_result,
                "interview_email": interview_email_result,
                "layout": layout_result,
            },
            ensure_ascii=False,
        )
    )
    return 1 if failed else 0


def dispatch_interviews(config: dict[str, Any], root: Path) -> int:
    from interview_scheduler import InterviewCoordinator

    coordinator = InterviewCoordinator(config, root)
    result = coordinator.dispatch_pending()
    coordinator.close()
    print(json.dumps(result, ensure_ascii=False))
    return 1 if result["failed"] else 0


def dispatch_interview_notices(config: dict[str, Any], root: Path) -> int:
    database_file = resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
    state = SyncState(database_file)
    gmail = GmailClient(GmailAuth(config, root))
    feishu = FeishuClient(config)
    result = run_interview_link_notice_scan(config, root, state, gmail, feishu)
    print(json.dumps(result, ensure_ascii=False))
    return 1 if result["failed"] else 0


def sync_original_resumes(config: dict[str, Any], root: Path, config_file: Path | None = None) -> int:
    """Backfill Drive links for all previously archived candidate résumés."""
    database_file = resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
    state = SyncState(database_file)
    queued = state.queue_completed_backups_for_resume_upload()
    feishu = FeishuClient(config, config_file)
    result = flush_feishu_resume_uploads(state, feishu)
    if result["completed"]:
        try:
            feishu.fit_sheet_dimensions()
        except SyncError as exc:
            logging.warning("Original résumé links were synchronized but sheet auto-fit failed: %s", exc)
    print(json.dumps({"backfill_queued": queued, "resume_archive": result}, ensure_ascii=False))
    return 1 if result["failed"] else 0


def queue_existing_interviews(config: dict[str, Any], root: Path) -> int:
    from interview_scheduler import InterviewCoordinator

    database_file = resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
    state = SyncState(database_file)
    gmail = GmailClient(GmailAuth(config, root))
    coordinator = InterviewCoordinator(config, root)
    llm = CandidateLLM(config)
    queued = 0
    rows = state.connection.execute(
        "SELECT message_id, feishu_record_id FROM message_state WHERE status='success'"
    ).fetchall()
    for message_id, record_range in rows:
        if not record_range:
            continue
        candidate, _attachment_errors = parse_candidate(gmail, gmail.message(str(message_id)), llm)
        coordinator.queue_candidate(str(message_id), str(record_range), candidate)
        queued += 1
    result = coordinator.dispatch_pending()
    coordinator.close()
    print(json.dumps({"queued": queued, "invitations": result}, ensure_ascii=False))
    return 1 if result["failed"] else 0


def self_test() -> None:
    sample = "姓名 张三\n学号 2026123456\n邮箱 zhangsan@example.com\n联系电话 138 0013 8000\n专业：计算机科学与技术\n应聘组别：开发组"
    parsed = fields_from_text(sample)
    required = {
        "name": "张三",
        "student_id": "2026123456",
        "email": "zhangsan@example.com",
        "phone": "138 0013 8000",
        "major": "计算机科学与技术",
        "group": "开发组",
    }
    if parsed != required:
        raise SyncError(f"Parser self-test failed: {parsed}")
    if preferred_email("resume@example.com", "body@example.com") != "resume@example.com":
        raise SyncError("Attachment email priority self-test failed")
    if preferred_email("", "") != "":
        raise SyncError("Missing email must remain blank")
    table_fields = fields_from_rows([["学号", "2026123456"]])
    for key, value in fields_from_text("202412345678").items():
        table_fields.setdefault(key, value)
    if table_fields.get("student_id") != "2026123456":
        raise SyncError("Explicit table fields must override unstructured fallback values")
    subject, content = interview_notice_content(
        {},
        "张三",
        "2026-08-10 9:05",
        "https://meeting.tencent.com/example",
    )
    if (
        "2026-08-10" not in content
        or "09:05" not in content
        or "面试时需打开摄像头" not in content
        or "110实验室" not in subject
    ):
        raise SyncError("Interview notice template self-test failed")
    if normalize_scheduled_at("46268.666666666664") != "2026-09-03 16:00":
        raise SyncError("Feishu date-serial normalization self-test failed")
    print("Parser self-test passed.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Synchronize recruiting email candidates to a Feishu spreadsheet")
    parser.add_argument(
        "command",
        choices=(
            "authorize-gmail",
            "check",
            "sync",
            "dispatch-interviews",
            "send-interview-notices",
            "sync-original-resumes",
            "queue-existing-interviews",
            "status",
            "self-test",
        ),
    )
    parser.add_argument("--config", default="config.json", help="path to local config JSON")
    args = parser.parse_args()
    if args.command == "self-test":
        self_test()
        return 0
    config_path = Path(args.config).resolve()
    config, root = load_config(config_path)
    log_file = resolve_path(root, str(config["runtime"].get("log_file", "logs/sync.log")))
    configure_logging(log_file)
    if args.command == "authorize-gmail":
        GmailAuth(config, root).authorize()
        return 0
    if args.command == "check":
        validate_remote(config, root)
        print("Gmail and Feishu connection check passed.")
        return 0
    if args.command == "sync":
        return run_sync(config, root, config_path)
    if args.command == "dispatch-interviews":
        return dispatch_interviews(config, root)
    if args.command == "send-interview-notices":
        return dispatch_interview_notices(config, root)
    if args.command == "sync-original-resumes":
        return sync_original_resumes(config, root, config_path)
    if args.command == "queue-existing-interviews":
        return queue_existing_interviews(config, root)
    if args.command == "status":
        database_file = resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
        if not database_file.exists():
            print("No local synchronization history yet.")
            return 0
        state = SyncState(database_file)
        print(json.dumps(dict(state.summary()), ensure_ascii=False))
        return 0
    return 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SyncError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)
