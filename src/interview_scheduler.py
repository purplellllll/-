"""Feishu direct-message interview scheduling for the local recruiting workflow.

Newly synchronized candidates are queued for an invitation.  When enabled, the
coordinator randomly selects one configured interviewer and sends that person a
direct message.  The listener accepts only that selected person's reply,
validates the submitted time against configured windows, and writes it to the
sheet.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import secrets
import sqlite3
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any
from urllib import parse

import recruitment_sync as sync


SHANGHAI = timezone(timedelta(hours=8))
COMMAND = re.compile(
    r"^\s*(?:/)?(?:面试|interview)\s+(INT-[A-Z0-9]+)\s+(\d{4}-\d{2}-\d{2}[ T]\d{1,2}:\s*\d{2})\s*$",
    flags=re.I,
)
RESUME_REVIEWER_BIND_COMMAND = re.compile(r"^\s*/?(?:设置简历审核|绑定简历审核)\s*$")
RESUME_REVIEWER_UNBIND_COMMAND = re.compile(r"^\s*/?(?:取消简历审核|解绑简历审核)\s*$")
OFFER_REVIEW_COMMAND = re.compile(
    r"^\s*/?(批准Offer|同意Offer|拒绝Offer|驳回Offer)\s+(OFFER-[A-Z0-9]+)\s*$",
    flags=re.I,
)
RESUME_SUFFIXES = {".pdf", ".docx", ".jpg", ".jpeg", ".png", ".webp", ".bmp"}
FEISHU_FILE_SIZE_LIMIT = 30 * 1024 * 1024


class InvitationStore:
    def __init__(self, database_file: Path):
        database_file.parent.mkdir(parents=True, exist_ok=True)
        self.connection = sqlite3.connect(database_file)
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS interview_invitation (
                invitation_id TEXT PRIMARY KEY,
                gmail_message_id TEXT UNIQUE NOT NULL,
                sheet_range TEXT NOT NULL,
                candidate_json TEXT NOT NULL,
                status TEXT NOT NULL,
                feishu_message_id TEXT,
                resume_delivery_json TEXT,
                card_nonce TEXT,
                scheduled_at TEXT,
                scheduled_by TEXT,
                created_at INTEGER NOT NULL,
                updated_at INTEGER NOT NULL
            )
            """
        )
        # Existing installations predate the separate resume-delivery state.
        # Add it in place so previously queued invitations remain usable.
        columns = {str(row[1]) for row in self.connection.execute("PRAGMA table_info(interview_invitation)").fetchall()}
        if "resume_delivery_json" not in columns:
            self.connection.execute("ALTER TABLE interview_invitation ADD COLUMN resume_delivery_json TEXT")
        if "card_nonce" not in columns:
            self.connection.execute("ALTER TABLE interview_invitation ADD COLUMN card_nonce TEXT")
        self.connection.execute(
            """
            CREATE TABLE IF NOT EXISTS interview_event_state (
                event_id TEXT PRIMARY KEY,
                processed_at INTEGER NOT NULL
            )
            """
        )
        self.connection.commit()

    def queue(self, gmail_message_id: str, sheet_range: str, candidate: dict[str, str]) -> str:
        invitation_id = "INT-" + re.sub(r"[^A-Z0-9]", "", gmail_message_id.upper())[-12:]
        now = int(time.time())
        self.connection.execute(
            """
            INSERT INTO interview_invitation(
                invitation_id, gmail_message_id, sheet_range, candidate_json, status, created_at, updated_at
            ) VALUES (?, ?, ?, ?, 'pending', ?, ?)
            ON CONFLICT(gmail_message_id) DO UPDATE SET
                sheet_range=excluded.sheet_range,
                candidate_json=excluded.candidate_json,
                updated_at=excluded.updated_at
            """,
            (invitation_id, gmail_message_id, sheet_range, json.dumps(candidate, ensure_ascii=False), now, now),
        )
        self.connection.commit()
        return invitation_id

    def pending(self) -> list[dict[str, Any]]:
        rows = self.connection.execute(
            """
            SELECT invitation_id, gmail_message_id, sheet_range, candidate_json,
                   feishu_message_id, resume_delivery_json, card_nonce, created_at
            FROM interview_invitation
            WHERE status IN ('pending', 'send_failed', 'sent')
            ORDER BY created_at
            """
        ).fetchall()
        return [
            {
                "invitation_id": row[0],
                "gmail_message_id": row[1],
                "sheet_range": row[2],
                "candidate": json.loads(row[3]),
                "deliveries": self._deliveries_from_value(row[4]),
                "resume_deliveries": self._resume_deliveries_from_value(row[5]),
                "card_nonce": str(row[6] or ""),
                "created_at": int(row[7]),
            }
            for row in rows
        ]

    @staticmethod
    def _deliveries_from_value(value: str | None) -> dict[str, str]:
        if not value:
            return {}
        try:
            decoded = json.loads(value)
        except (TypeError, json.JSONDecodeError):
            # The previous group-chat implementation stored one raw message ID.
            # It deliberately does not count as an individual delivery.
            return {}
        if not isinstance(decoded, dict):
            return {}
        return {str(email).lower(): str(message_id) for email, message_id in decoded.items() if message_id}

    def mark_sent(self, invitation_id: str, deliveries: dict[str, str], card_nonce: str | None = None) -> None:
        if card_nonce is None:
            self.connection.execute(
                "UPDATE interview_invitation SET status='sent', feishu_message_id=?, updated_at=? WHERE invitation_id=?",
                (json.dumps(deliveries, ensure_ascii=False, sort_keys=True), int(time.time()), invitation_id),
            )
        else:
            self.connection.execute(
                "UPDATE interview_invitation SET status='sent', feishu_message_id=?, card_nonce=?, updated_at=? WHERE invitation_id=?",
                (json.dumps(deliveries, ensure_ascii=False, sort_keys=True), card_nonce, int(time.time()), invitation_id),
            )
        self.connection.commit()

    @staticmethod
    def _resume_deliveries_from_value(value: str | None) -> dict[str, dict[str, Any]]:
        if not value:
            return {}
        try:
            decoded = json.loads(value)
        except (TypeError, json.JSONDecodeError):
            return {}
        if not isinstance(decoded, dict):
            return {}
        result: dict[str, dict[str, Any]] = {}
        for recipient, item in decoded.items():
            if not isinstance(item, dict):
                continue
            status = str(item.get("status", "")).strip()
            if not status:
                continue
            result[str(recipient)] = {
                "status": status,
                "message_id": str(item.get("message_id", "")).strip(),
                "file_name": str(item.get("file_name", "")).strip(),
            }
        return result

    def mark_resume_delivery(
        self,
        invitation_id: str,
        deliveries: dict[str, dict[str, Any]],
    ) -> None:
        self.connection.execute(
            "UPDATE interview_invitation SET resume_delivery_json=?, updated_at=? WHERE invitation_id=?",
            (json.dumps(deliveries, ensure_ascii=False, sort_keys=True), int(time.time()), invitation_id),
        )
        self.connection.commit()

    def mark_send_failed(self, invitation_id: str) -> None:
        self.connection.execute(
            "UPDATE interview_invitation SET status='send_failed', updated_at=? WHERE invitation_id=?",
            (int(time.time()), invitation_id),
        )
        self.connection.commit()

    def get(self, invitation_id: str) -> dict[str, Any] | None:
        row = self.connection.execute(
            """
            SELECT invitation_id, sheet_range, candidate_json, status, scheduled_at, scheduled_by, feishu_message_id, card_nonce, created_at, updated_at
            FROM interview_invitation WHERE invitation_id=?
            """,
            (invitation_id,),
        ).fetchone()
        if not row:
            return None
        return {
            "invitation_id": row[0],
            "sheet_range": row[1],
            "candidate": json.loads(row[2]),
            "status": row[3],
            "scheduled_at": row[4],
            "scheduled_by": row[5],
            "deliveries": self._deliveries_from_value(row[6]),
            "card_nonce": str(row[7] or ""),
            "created_at": int(row[8]),
            "issued_at": int(row[9]),
        }

    def claim(self, invitation_id: str, scheduled_at: str, scheduled_by: str) -> bool:
        cursor = self.connection.execute(
            """
            UPDATE interview_invitation
            SET status='updating', scheduled_at=?, scheduled_by=?, updated_at=?
            WHERE invitation_id=? AND status='sent'
            """,
            (scheduled_at, scheduled_by, int(time.time()), invitation_id),
        )
        self.connection.commit()
        return cursor.rowcount == 1

    def complete(self, invitation_id: str) -> None:
        self.connection.execute(
            "UPDATE interview_invitation SET status='scheduled', updated_at=? WHERE invitation_id=?",
            (int(time.time()), invitation_id),
        )
        self.connection.commit()

    def restore_sent(self, invitation_id: str) -> None:
        self.connection.execute(
            "UPDATE interview_invitation SET status='sent', updated_at=? WHERE invitation_id=?",
            (int(time.time()), invitation_id),
        )
        self.connection.commit()

    def record_event(self, event_id: str) -> bool:
        if not event_id:
            return True
        cursor = self.connection.execute(
            "INSERT OR IGNORE INTO interview_event_state(event_id, processed_at) VALUES (?, ?)",
            (event_id, int(time.time())),
        )
        self.connection.commit()
        return cursor.rowcount == 1

    def close(self) -> None:
        self.connection.close()


class InterviewCoordinator:
    def __init__(self, config: dict[str, Any], root: Path, config_path: Path | None = None):
        self.config = config
        self.root = root
        self.config_path = config_path or root / "config.json"
        self.options = config.get("interview", {}) if isinstance(config.get("interview", {}), dict) else {}
        database_file = sync.resolve_path(root, str(config["runtime"].get("database_file", "data/sync-state.db")))
        self.store = InvitationStore(database_file)
        self.sync_state = sync.SyncState(database_file)
        self.feishu = sync.FeishuClient(config)

    def queue_candidate(self, gmail_message_id: str, sheet_range: str, candidate: dict[str, str]) -> str:
        return self.store.queue(gmail_message_id, sheet_range, candidate)

    def close(self) -> None:
        self.store.close()
        self.sync_state.connection.close()

    def _windows(self) -> list[tuple[datetime, datetime]]:
        parsed: list[tuple[datetime, datetime]] = []
        for item in self.options.get("time_windows", []):
            if not isinstance(item, dict):
                continue
            try:
                start = self._parse_time(str(item["start"]))
                end = self._parse_time(str(item["end"]))
            except (KeyError, ValueError):
                continue
            if end > start:
                parsed.append((start, end))
        return parsed

    @staticmethod
    def _parse_time(value: str) -> datetime:
        normalized = value.strip().replace("T", " ")
        parsed = datetime.fromisoformat(normalized)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=SHANGHAI)
        return parsed.astimezone(SHANGHAI)

    def _manual_interviewers(self) -> list[dict[str, str]]:
        """Return the uniquely configured direct-message recipients by email."""
        result: list[dict[str, str]] = []
        seen: set[str] = set()
        configured = self.options.get("interviewers", [])
        if not isinstance(configured, list):
            return result
        for item in configured:
            if not isinstance(item, dict):
                continue
            email = str(item.get("email", "")).strip().lower()
            if not email or "@" not in email or email in seen:
                continue
            seen.add(email)
            interviewer = {"id": email, "id_type": "email", "name": str(item.get("name", "")).strip()}
            if self._is_excluded_interviewer(interviewer):
                continue
            result.append(interviewer)
        return result

    def _excluded_interviewers(self) -> set[str]:
        """Return configured interviewer names or IDs that must not be selected."""
        configured = self.options.get("excluded_interviewers", [])
        if not isinstance(configured, list):
            return set()
        return {str(item).strip().casefold() for item in configured if str(item).strip()}

    def _is_excluded_interviewer(self, interviewer: dict[str, str]) -> bool:
        excluded = self._excluded_interviewers()
        if not excluded:
            return False
        return any(
            str(interviewer.get(key, "")).strip().casefold() in excluded
            for key in ("id", "name")
        )

    def _recipient_group_chat_id(self) -> str:
        group = self.options.get("recipient_group", {})
        if not isinstance(group, dict):
            return ""
        return str(group.get("chat_id", "")).strip()

    def _group_interviewers(self, chat_id: str) -> list[dict[str, str]]:
        """Load human members of the configured group as direct-message targets.

        The Feishu API deliberately omits robot members from this endpoint, so
        the selection is exactly "all group members except robots".
        """
        result: list[dict[str, str]] = []
        page_token = ""
        seen: set[str] = set()
        while True:
            query = parse.urlencode(
                {
                    "member_id_type": "open_id",
                    "page_size": 100,
                    "page_token": page_token,
                }
            )
            response = sync.json_request(
                sync.FEISHU_API + f"/im/v1/chats/{parse.quote(chat_id)}/members?{query}",
                headers={"Authorization": f"Bearer {self.feishu.token()}"},
            )
            if response.get("code") != 0:
                raise sync.SyncError(
                    f"Feishu group member lookup failed: {response.get('msg', 'unknown error')} "
                    f"(code {response.get('code')})"
                )
            data = response.get("data", {}) if isinstance(response.get("data", {}), dict) else {}
            for item in data.get("items", []):
                if not isinstance(item, dict):
                    continue
                member_id = str(item.get("member_id", "")).strip()
                if not member_id or member_id in seen:
                    continue
                seen.add(member_id)
                interviewer = {
                    "id": member_id,
                    "id_type": "open_id",
                    "name": str(item.get("name", "")).strip(),
                }
                if self._is_excluded_interviewer(interviewer):
                    continue
                result.append(interviewer)
            if not data.get("has_more"):
                break
            next_page_token = str(data.get("page_token", "")).strip()
            if not next_page_token or next_page_token == page_token:
                raise sync.SyncError("Feishu group member pagination stopped unexpectedly")
            page_token = next_page_token
        return result

    def _interviewers(self) -> list[dict[str, str]]:
        group_chat_id = self._recipient_group_chat_id()
        if group_chat_id:
            return self._group_interviewers(group_chat_id)
        return self._manual_interviewers()

    def _interviewer_display_name(self, interviewer_id: str) -> str:
        """Return a human-readable interviewer name for spreadsheet output."""
        try:
            for interviewer in self._interviewers():
                if interviewer.get("id") == interviewer_id:
                    name = str(interviewer.get("name", "")).strip()
                    if name:
                        return name
                    if interviewer.get("id_type") == "email":
                        return interviewer_id
                    break
        except sync.SyncError as exc:
            logging.warning("Could not resolve interviewer display name: %s", exc)
        return "未识别面试官"

    def _ready(self) -> bool:
        has_recipient_source = bool(self._recipient_group_chat_id() or self._manual_interviewers())
        return bool(self.options.get("enabled") and has_recipient_source and self._windows())

    def _invitation_start(self, invitation: dict[str, Any] | None = None) -> datetime | None:
        if not invitation:
            return None
        timestamp = invitation.get("issued_at") or invitation.get("created_at")
        if not timestamp:
            return None
        try:
            return datetime.fromtimestamp(int(timestamp), tz=SHANGHAI).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
        except (TypeError, ValueError, OSError):
            return None

    def _available_dates(self, invitation: dict[str, Any] | None = None) -> list[datetime]:
        """Return selectable dates, never before the invitation was issued."""
        duration = timedelta(minutes=int(self.options.get("duration_minutes", 30)))
        invitation_day = self._invitation_start(invitation)
        dates: dict[str, datetime] = {}
        for start, end in self._windows():
            last_start = end - duration
            if last_start < start:
                continue
            cursor = start.replace(hour=0, minute=0, second=0, microsecond=0)
            if invitation_day and cursor < invitation_day:
                cursor = invitation_day
            while cursor.date() <= last_start.date():
                dates.setdefault(cursor.strftime("%Y-%m-%d"), cursor)
                cursor += timedelta(days=1)
        return [dates[key] for key in sorted(dates)]

    def _invitation_card(self, invitation: dict[str, Any]) -> dict[str, Any]:
        """Build a fixed-year, month/day/time appointment card (JSON 2.0)."""
        candidate = invitation["candidate"]
        duration = int(self.options.get("duration_minutes", 30))
        dates = self._available_dates(invitation)
        if not dates:
            raise sync.SyncError("No selectable interview date exists in interview.time_windows")
        years = {value.year for value in dates}
        fixed_year = str(next(iter(years))) if len(years) == 1 else "已限定日期"
        weekdays = "一二三四五六日"
        options = [
            {
                "text": {"tag": "plain_text", "content": f"{value.month}月{value.day}日（周{weekdays[value.weekday()]}）"},
                "value": value.strftime("%Y-%m-%d"),
            }
            for value in dates
        ]
        return {
            "schema": "2.0",
            # Feishu requires update_multi to be enabled for this interactive
            # card schema, even though the card is delivered to one interviewer.
            "config": {"enable_forward": False, "update_multi": True, "width_mode": "fill"},
            "header": {
                "title": {"tag": "plain_text", "content": "招新面试时间确认"},
                "subtitle": {"tag": "plain_text", "content": f"固定 {fixed_year} 年"},
                "template": "blue",
            },
            "body": {
                "elements": [
                    {
                        "tag": "markdown",
                        "content": (
                            f"**候选人**：{candidate.get('name', '未识别')}\n"
                            f"**面试时长**：{duration} 分钟\n"
                            "请选择月/日和时间后确认。"
                        ),
                    },
                    {"tag": "hr"},
                    {
                        "tag": "form",
                        "name": "interview_schedule",
                        "elements": [
                            {
                                "tag": "select_static",
                                "name": "interview_date",
                                "required": True,
                                "placeholder": {"tag": "plain_text", "content": "选择月 / 日"},
                                "options": options,
                                "initial_option": options[0]["value"],
                            },
                            {
                                "tag": "picker_time",
                                "name": "interview_time",
                                "required": True,
                                "placeholder": {"tag": "plain_text", "content": "选择时间"},
                                "initial_time": "09:00",
                            },
                            {
                                "tag": "button",
                                "name": "confirm_interview",
                                "text": {"tag": "plain_text", "content": "确认面试时间"},
                                "type": "primary",
                                "form_action_type": "submit",
                                "behaviors": [
                                    {
                                        "type": "callback",
                                        "value": {
                                            "action": "confirm_interview",
                                            "invitation_id": invitation["invitation_id"],
                                            # A new card gets a new nonce.  Once a
                                            # replacement card is issued, a
                                            # callback from the old card can no
                                            # longer change the appointment.
                                            "card_nonce": str(invitation.get("card_nonce", "")),
                                        },
                                    }
                                ],
                            },
                        ],
                    },
                ]
            },
        }

    def _send_message(self, receive_id: str, msg_type: str, content: dict[str, Any], receive_id_type: str = "chat_id") -> str:
        response = sync.json_request(
            sync.FEISHU_API + f"/im/v1/messages?receive_id_type={parse.quote(receive_id_type)}",
            method="POST",
            headers={"Authorization": f"Bearer {self.feishu.token()}"},
            # The IM API expects ``content`` to be a JSON string.  Escaping
            # non-ASCII characters in that inner JSON avoids double-decoding
            # mojibake in some Windows/Feishu client paths; Feishu decodes the
            # JSON escapes before rendering the message or card.
            payload={"receive_id": receive_id, "msg_type": msg_type, "content": json.dumps(content, ensure_ascii=True)},
        )
        if response.get("code") != 0 or not response.get("data", {}).get("message_id"):
            raise sync.SyncError(f"Feishu interview message failed: {response.get('msg', 'unknown error')}")
        return str(response["data"]["message_id"])

    def _send_text(self, receive_id: str, text: str, receive_id_type: str = "chat_id") -> str:
        return self._send_message(receive_id, "text", {"text": text}, receive_id_type)

    def _send_card(self, receive_id: str, card: dict[str, Any], receive_id_type: str = "open_id") -> str:
        return self._send_message(receive_id, "interactive", card, receive_id_type)

    def _resume_is_enabled(self) -> bool:
        """Whether the original resume accompanies a direct invitation."""
        return bool(self.options.get("send_resume_with_invitation", True))

    def _resume_reviewer(self) -> dict[str, str] | None:
        """Return the one user allowed to receive an audit copy of resumes."""
        configured = self.options.get("resume_reviewer", {})
        if not isinstance(configured, dict):
            return None
        open_id = str(configured.get("open_id", "")).strip()
        if not open_id:
            return None
        return {"id": open_id, "id_type": "open_id", "name": str(configured.get("name", "")).strip()}

    def _bind_resume_reviewer(self, open_id: str) -> None:
        """Persist the opt-in audit recipient from that user's private chat."""
        interview = self.config.setdefault("interview", {})
        reviewer = interview.setdefault("resume_reviewer", {})
        if not isinstance(reviewer, dict):
            reviewer = {}
            interview["resume_reviewer"] = reviewer
        reviewer["open_id"] = open_id
        reviewer["name"] = ""
        sync.write_private_json(self.config_path, self.config)
        self.options = interview

    def _unbind_resume_reviewer(self, open_id: str) -> bool:
        reviewer = self._resume_reviewer()
        if reviewer is None or reviewer["id"] != open_id:
            return False
        interview = self.config.setdefault("interview", {})
        interview["resume_reviewer"] = {"open_id": "", "name": ""}
        sync.write_private_json(self.config_path, self.config)
        self.options = interview
        return True

    @staticmethod
    def _resume_rank(filename: str) -> tuple[int, str]:
        """Prefer files explicitly named as a resume when several are attached."""
        normalized = filename.lower()
        is_named_resume = any(word in normalized for word in ("简历", "resume", "curriculum vitae", " cv"))
        return (0 if is_named_resume else 1, normalized)

    def _choose_resume_file(self, files: list[tuple[str, bytes]]) -> tuple[str, bytes] | None:
        candidates = [
            (name, raw)
            for name, raw in files
            if Path(name).suffix.lower() in RESUME_SUFFIXES and raw
        ]
        if not candidates:
            return None
        name, raw = min(candidates, key=lambda item: self._resume_rank(item[0]))
        if len(raw) > FEISHU_FILE_SIZE_LIMIT:
            raise sync.SyncError(f"The resume attachment '{name}' exceeds Feishu's 30 MB file limit")
        return name, raw

    def _backup_resume_file(self, gmail_message_id: str) -> tuple[str, bytes] | None:
        """Load the locally archived original file without contacting Gmail."""
        backup = sync.CandidateBackup(self.config, self.root)
        directory = backup.root / backup._safe_component(gmail_message_id, "gmail-message")
        manifest_file = directory / "candidate.json"
        if not manifest_file.exists():
            return None
        try:
            manifest = json.loads(manifest_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            logging.warning("Could not read the local resume archive for one invitation: %s", exc)
            return None
        attachments = manifest.get("resume_attachments", []) if isinstance(manifest, dict) else []
        files: list[tuple[str, bytes]] = []
        for item in attachments:
            if not isinstance(item, dict):
                continue
            filename = str(item.get("original_name", "")).strip()
            backup_name = str(item.get("backup_name", "")).strip()
            if Path(filename).suffix.lower() not in RESUME_SUFFIXES or not backup_name:
                continue
            path = directory / backup_name
            try:
                path.resolve().relative_to(directory.resolve())
                files.append((filename, path.read_bytes()))
            except (OSError, ValueError):
                logging.warning("A locally archived resume file is unavailable for one invitation.")
        return self._choose_resume_file(files)

    def _gmail_resume_file(self, gmail_message_id: str) -> tuple[str, bytes] | None:
        """Fall back to the original Gmail message before an archive is ready."""
        gmail = sync.GmailClient(sync.GmailAuth(self.config, self.root))
        message = gmail.message(gmail_message_id)
        _body, attachments = sync.message_content(gmail, message)
        return self._choose_resume_file(attachments)

    def _resume_file(self, gmail_message_id: str) -> tuple[str, bytes] | None:
        archived = self._backup_resume_file(gmail_message_id)
        return archived if archived is not None else self._gmail_resume_file(gmail_message_id)

    @staticmethod
    def _feishu_file_type(filename: str) -> str:
        """Map supported resume extensions to Feishu's upload type values."""
        suffix = Path(filename).suffix.lower()
        if suffix == ".pdf":
            return "pdf"
        if suffix == ".docx":
            return "doc"
        return "stream"

    def _send_resume_file(self, receive_id: str, receive_id_type: str, filename: str, content: bytes) -> str:
        upload = sync.multipart_file_request(
            sync.FEISHU_API + "/im/v1/files",
            headers={"Authorization": f"Bearer {self.feishu.token()}"},
            fields={"file_type": self._feishu_file_type(filename), "file_name": filename},
            file_field="file",
            filename=filename,
            content=content,
        )
        file_key = str(upload.get("data", {}).get("file_key", "")).strip()
        if upload.get("code") != 0 or not file_key:
            raise sync.SyncError(f"Feishu resume upload failed: {upload.get('msg', 'unknown error')}")
        return self._send_message(
            receive_id,
            "file",
            {"file_key": file_key, "file_name": filename},
            receive_id_type,
        )

    def _deliver_resume_if_needed(
        self,
        invitation: dict[str, Any],
        delivery_key: str,
        receive_id: str,
        receive_id_type: str,
        result: dict[str, int],
    ) -> None:
        if not self._resume_is_enabled():
            return
        resume_deliveries = dict(invitation.get("resume_deliveries", {}))
        reviewer = self._resume_reviewer()
        reviewer_key = f"reviewer:open_id:{reviewer['id']}" if reviewer and reviewer["id"] != receive_id else ""
        recipients = [(delivery_key, receive_id, receive_id_type, "interviewer")]
        if reviewer and reviewer_key:
            recipients.append((reviewer_key, reviewer["id"], reviewer["id_type"], "reviewer"))
        pending = [
            recipient
            for recipient in recipients
            if resume_deliveries.get(recipient[0], {}).get("status") not in {"sent", "missing"}
        ]
        if not pending:
            return
        try:
            resume = self._resume_file(str(invitation["gmail_message_id"]))
            if resume is None:
                for pending_key, _recipient_id, _recipient_type, role in pending:
                    resume_deliveries[pending_key] = {"status": "missing", "message_id": "", "file_name": ""}
                    result["resume_missing"] += 1
                    if role == "reviewer":
                        result["reviewer_resume_missing"] += 1
                self.store.mark_resume_delivery(invitation["invitation_id"], resume_deliveries)
                logging.warning("No PDF, DOC, or DOCX attachment was found for one interview invitation.")
                return
            filename, content = resume
            for pending_key, recipient_id, recipient_type, role in pending:
                try:
                    message_id = self._send_resume_file(recipient_id, recipient_type, filename, content)
                    resume_deliveries[pending_key] = {
                        "status": "sent",
                        "message_id": message_id,
                        "file_name": filename,
                    }
                    result["resume_sent"] += 1
                    if role == "reviewer":
                        result["reviewer_resume_sent"] += 1
                except sync.SyncError as exc:
                    resume_deliveries[pending_key] = {"status": "failed", "message_id": "", "file_name": ""}
                    result["resume_failed"] += 1
                    result["failed"] += 1
                    if role == "reviewer":
                        result["reviewer_resume_failed"] += 1
                    logging.error("Interview resume delivery failed and will retry: %s", exc)
            self.store.mark_resume_delivery(invitation["invitation_id"], resume_deliveries)
        except sync.SyncError as exc:
            # Keep the selected recipient(s) fixed and retry only the file on
            # the next dispatch instead of selecting another interviewer.
            for pending_key, _recipient_id, _recipient_type, role in pending:
                resume_deliveries[pending_key] = {"status": "failed", "message_id": "", "file_name": ""}
                if role == "reviewer":
                    result["reviewer_resume_failed"] += 1
            self.store.mark_resume_delivery(invitation["invitation_id"], resume_deliveries)
            result["resume_failed"] += 1
            result["failed"] += 1
            logging.error("Interview resume delivery failed and will retry: %s", exc)

    def dispatch_pending(self) -> dict[str, int]:
        result = {
            "sent": 0,
            "failed": 0,
            "waiting_for_configuration": 0,
            "resume_sent": 0,
            "resume_missing": 0,
            "resume_failed": 0,
            "reviewer_resume_sent": 0,
            "reviewer_resume_missing": 0,
            "reviewer_resume_failed": 0,
        }
        pending = self.store.pending()
        if not pending:
            return result
        if not self._ready():
            result["waiting_for_configuration"] = len(pending)
            return result
        try:
            interviewers = self._interviewers()
        except sync.SyncError as exc:
            logging.error("Interview recipient lookup failed: %s", exc)
            result["failed"] = len(pending)
            return result
        if not interviewers:
            logging.warning("Interview recipient source currently has no human members.")
            result["waiting_for_configuration"] = len(pending)
            return result
        for invitation in pending:
            deliveries = dict(invitation["deliveries"])
            # The first successful delivery fixes the one randomly selected
            # interviewer for this candidate.  Later polling runs never pick a
            # second person or resend the same card.  A failed resume upload is
            # retried only to that already selected person.
            if deliveries:
                self.store.mark_sent(invitation["invitation_id"], deliveries)
                delivery_key, _card_message_id = next(iter(deliveries.items()))
                try:
                    recipient_type, recipient_id = delivery_key.split(":", 1)
                except ValueError:
                    logging.error("Stored interview delivery has an invalid recipient key.")
                    result["resume_failed"] += 1
                    continue
                self._deliver_resume_if_needed(
                    invitation,
                    delivery_key,
                    recipient_id,
                    recipient_type,
                    result,
                )
                continue
            interviewer = secrets.choice(interviewers)
            recipient_id = interviewer["id"]
            recipient_type = interviewer["id_type"]
            interviewer_name = str(interviewer.get("name", "")).strip() or "未识别面试官"
            delivery_key = f"{recipient_type}:{recipient_id}"
            sent_this_round = 0
            failed_this_round = False
            try:
                # The earliest selectable day is the day this invitation is
                # actually sent, rather than the time at which the resume was
                # first queued locally.
                card_nonce = secrets.token_hex(16)
                card_invitation = {
                    **invitation,
                    "issued_at": int(time.time()),
                    "card_nonce": card_nonce,
                }
                message_id = self._send_card(
                    recipient_id,
                    self._invitation_card(card_invitation),
                    receive_id_type=recipient_type,
                )
                deliveries[delivery_key] = message_id
                sent_this_round = 1
            except sync.SyncError as exc:
                logging.error("Interview invitation send failed for %s: %s", recipient_id, exc)
                failed_this_round = True
                result["failed"] += 1
            if deliveries:
                self.store.mark_sent(invitation["invitation_id"], deliveries, card_nonce=card_nonce)
                result["sent"] += sent_this_round
                try:
                    self._write_interview_values(
                        invitation["sheet_range"],
                        (("status", "已发送，待确认"), ("interviewer", interviewer_name)),
                    )
                except sync.SyncError as exc:
                    logging.error("Interview invitation status write failed: %s", exc)
                self._deliver_resume_if_needed(
                    invitation,
                    delivery_key,
                    recipient_id,
                    recipient_type,
                    result,
                )
            elif failed_this_round:
                self.store.mark_send_failed(invitation["invitation_id"])
        return result

    def _within_window(self, value: datetime, invitation: dict[str, Any] | None = None) -> bool:
        duration = timedelta(minutes=int(self.options.get("duration_minutes", 30)))
        invitation_day = self._invitation_start(invitation)
        for start, end in self._windows():
            effective_start = max(start, invitation_day) if invitation_day else start
            if effective_start <= value and value + duration <= end:
                return True
        return False

    def _write_interview_values(self, sheet_range: str, updates: tuple[tuple[str, str], ...]) -> None:
        match = re.search(r"![A-Z]+(\d+):", sheet_range)
        if not match:
            raise sync.SyncError("Stored Feishu range has no row number")
        row_number = match.group(1)
        context = self.feishu.sheet_context()
        fields = self.options.get("fields", {})
        if any(key not in fields for key, _value in updates):
            raise sync.SyncError("interview.fields is incomplete in config.json")
        for key, value in updates:
            header = str(fields[key])
            headers = list(context["headers"])
            if header not in headers:
                raise sync.SyncError(f"Interview sheet header is missing: {header}")
            column = sync.column_name(headers.index(header) + 1)
            value_range = f"{context['sheet_id']}!{column}{row_number}:{column}{row_number}"
            self.feishu.call(
                "PUT",
                f"/sheets/v2/spreadsheets/{parse.quote(str(context['spreadsheet_token']))}/values",
                payload={"valueRange": {"range": value_range, "values": [[value]]}},
            )

    def _write_schedule(self, sheet_range: str, scheduled_at: str, interviewer_open_id: str) -> None:
        interviewer_name = self._interviewer_display_name(interviewer_open_id)
        self._write_interview_values(
            sheet_range,
            (
            ("scheduled_at", scheduled_at),
            ("status", "已确定"),
            ("interviewer", interviewer_name),
            ),
        )

    def _reply(self, chat_id: str, text: str) -> None:
        try:
            self._send_text(chat_id, text)
        except sync.SyncError as exc:
            logging.error("Interview reply failed: %s", exc)

    def _workflow_guide(self) -> str:
        """Explain the automatic interview workflow to a newly bound group."""
        configured_windows = self.options.get("time_windows", [])
        windows: list[str] = []
        if isinstance(configured_windows, list):
            for item in configured_windows:
                if not isinstance(item, dict):
                    continue
                end = str(item.get("end", "")).strip()
                if end:
                    windows.append(f"该份邀请发起当天 至 {end}")
        window_text = "；".join(windows) if windows else "以机器人邀请卡显示的范围为准"
        duration = max(1, int(self.options.get("duration_minutes", 30)))
        return (
            "【招新面试流程说明】\n"
            "本群已自动设为面试官名单来源。\n\n"
            "流程：\n"
            "1. 候选人简历从 Gmail 同步到飞书表格。\n"
            "2. 机器人从本群的人类成员中随机选择一位，仅向该面试官私聊发送预约卡片。群内不会群发候选人信息。\n"
            f"3. 被选中的面试官在卡片中选择面试日期和时间（单次 {duration} 分钟；当前可选范围：{window_text}），再点击确认。\n"
            "4. 面试官自行预约腾讯会议，并将会议链接填写到飞书表格对应行；负责招新的同学会先在机器人私聊中审核 Offer，批准后系统才通过招新 Gmail 通知候选人。\n"
            "5. 面试由该面试官与一位负责招新的同学共同参与，并共同确定面试分数；分数填写在表格最后的「分数」列。\n\n"
            "面试官需要做什么：\n"
            "- 留意与机器人的私聊；收到【招新面试时间确认】卡片后选择时间并点击确认。\n"
            "- 不需要在本群回复时间；只有被分配到该候选人的面试官可以确认。\n"
            "- 确认后自行预约腾讯会议，并在飞书表格对应行填写有效的会议链接。\n"
            "- 与负责招新的同学共同完成面试并确定分数，随后填写表格最后的「分数」列。"
        )

    def _bind_recipient_group(self, chat_id: str, group_name: str = "") -> str:
        interview = self.config.setdefault("interview", {})
        group = interview.setdefault("recipient_group", {})
        if not isinstance(group, dict):
            group = {}
            interview["recipient_group"] = group
        group["chat_id"] = chat_id
        group["name"] = group_name.strip()
        sync.write_private_json(self.config_path, self.config)
        self.options = interview
        return "面试官名单群已自动绑定。"

    def handle_bot_added(self, payload: dict[str, Any]) -> None:
        """Bind the first group that adds the bot, then explain the workflow."""
        header = payload.get("header", {}) if isinstance(payload.get("header", {}), dict) else {}
        if not self.store.record_event(str(header.get("event_id", ""))):
            return
        event = payload.get("event", payload) if isinstance(payload, dict) else {}
        if not isinstance(event, dict):
            return
        chat_id = str(event.get("chat_id", "")).strip()
        group_name = str(event.get("name", "")).strip()
        if not chat_id:
            logging.warning("Ignored bot-added event without a chat_id.")
            return
        existing_group = self._recipient_group_chat_id()
        if existing_group and existing_group != chat_id:
            logging.warning("Ignored automatic interview-recipient binding from a second group.")
            self._reply(chat_id, "机器人已绑定一个面试官名单群；不会自动覆盖现有绑定。")
            return
        if not existing_group:
            logging.info("Automatically binding interview-recipient group after bot-added event.")
            self._bind_recipient_group(chat_id, group_name)
        self._reply(chat_id, self._workflow_guide())

    @staticmethod
    def _card_callback_response(toast_type: str, content: str) -> Any:
        from lark_oapi.event.callback.model.p2_card_action_trigger import P2CardActionTriggerResponse

        return P2CardActionTriggerResponse({"toast": {"type": toast_type, "content": content}})

    def _handle_offer_card_action(self, event: Any, value: dict[str, Any], operator_open_id: str) -> Any:
        """Handle the two-button Offer review card for the configured reviewer."""
        review_id = str(value.get("review_id", "")).strip()
        if not review_id or not operator_open_id:
            return self._card_callback_response("error", "无法确认审核人或审核编号，请稍后重试。")
        review = self.sync_state.offer_review_by_id(review_id)
        if review is None:
            return self._card_callback_response("error", "未找到该 Offer 审核，可能已被替换。")
        if review.get("reviewer_open_id") != operator_open_id:
            return self._card_callback_response("error", "你不是该 Offer 的指定审核人。")
        if review.get("status") != "pending":
            status_text = {"approved": "已通过", "rejected": "已拒绝"}.get(review.get("status", ""), "已处理")
            return self._card_callback_response("warning", f"该 Offer {status_text}，无需重复操作。")
        approved = value.get("action") == "approve_offer"
        status = "approved" if approved else "rejected"
        decision_message_id = str(getattr(event, "message_id", "") or f"card:{int(time.time())}")
        self.sync_state.decide_offer_review(review_id, status, decision_message_id)
        if approved:
            return self._card_callback_response("success", "已通过 Offer；下一次同步将发送给候选人。")
        return self._card_callback_response("success", "已拒绝 Offer；系统不会发送给候选人。")

    def handle_card_action(self, callback: Any) -> Any:
        """Handle interview-time forms and the two-button Offer review card."""
        event = getattr(callback, "event", None)
        action = getattr(event, "action", None)
        operator = getattr(event, "operator", None)
        value = getattr(action, "value", None) or {}
        if not isinstance(value, dict):
            return self._card_callback_response("warning", "未识别的卡片操作。")
        operator_open_id = str(getattr(operator, "open_id", "") or "").strip()
        action_name = str(value.get("action", "")).strip()
        if action_name in {"approve_offer", "reject_offer"}:
            return self._handle_offer_card_action(event, value, operator_open_id)
        if action_name != "confirm_interview":
            return self._card_callback_response("warning", "未识别的卡片操作。")
        invitation_id = str(value.get("invitation_id", "")).strip()
        interviewer_open_id = operator_open_id
        if not invitation_id or not interviewer_open_id:
            return self._card_callback_response("error", "无法确认操作人或候选人，请稍后重试。")
        invitation = self.store.get(invitation_id)
        if invitation is None:
            return self._card_callback_response("error", "未找到该面试邀请。")
        expected_nonce = str(invitation.get("card_nonce", "")).strip()
        provided_nonce = str(value.get("card_nonce", "")).strip()
        if expected_nonce and provided_nonce != expected_nonce:
            return self._card_callback_response("error", "这张面试卡已失效，请使用最新卡片。")
        selected_delivery_key = f"open_id:{interviewer_open_id}"
        if self._recipient_group_chat_id() and selected_delivery_key not in invitation["deliveries"]:
            return self._card_callback_response("error", "该候选人已随机分配给另一位面试官。")
        form_value = getattr(action, "form_value", None) or {}
        if not isinstance(form_value, dict):
            return self._card_callback_response("error", "请选择日期和时间后再确认。")
        date_value = str(form_value.get("interview_date", "")).strip()
        raw_time = str(form_value.get("interview_time", "")).strip()
        time_match = re.search(r"\b(\d{1,2}:\d{2})\b", raw_time)
        if date_value not in {value.strftime("%Y-%m-%d") for value in self._available_dates(invitation)} or not time_match:
            return self._card_callback_response("error", "请选择卡片中提供的日期和时间。")
        try:
            chosen = self._parse_time(f"{date_value} {time_match.group(1)}")
        except ValueError:
            return self._card_callback_response("error", "时间格式无效，请重新选择。")
        if not self._within_window(chosen, invitation):
            return self._card_callback_response("error", "该时间不在允许的面试时段内，请重新选择。")
        scheduled_at = chosen.strftime("%Y-%m-%d %H:%M")
        if not self.store.claim(invitation_id, scheduled_at, interviewer_open_id):
            return self._card_callback_response("warning", "该候选人的面试时间已确认。")
        try:
            self._write_schedule(invitation["sheet_range"], scheduled_at, interviewer_open_id)
            self.store.complete(invitation_id)
            return self._card_callback_response("success", f"已确认面试时间：{scheduled_at}")
        except sync.SyncError as exc:
            logging.error("Card schedule write failed: %s", exc)
            self.store.restore_sent(invitation_id)
            return self._card_callback_response("error", "写入面试表失败，请稍后再次确认。")

    def handle_message(self, payload: dict[str, Any]) -> None:
        header = payload.get("header", {}) if isinstance(payload.get("header", {}), dict) else {}
        event = payload.get("event", payload) if isinstance(payload, dict) else {}
        if not self.store.record_event(str(header.get("event_id", ""))):
            return
        message = event.get("message", {}) if isinstance(event.get("message", {}), dict) else {}
        sender = event.get("sender", {}) if isinstance(event.get("sender", {}), dict) else {}
        sender_id = sender.get("sender_id", {}) if isinstance(sender.get("sender_id", {}), dict) else {}
        chat_id = str(message.get("chat_id", ""))
        interviewer_open_id = str(sender_id.get("open_id", ""))
        try:
            content = json.loads(str(message.get("content", "{}")))
        except json.JSONDecodeError:
            return
        text = re.sub(r"<at[^>]*>.*?</at>", "", str(content.get("text", "")), flags=re.S).strip()
        normalized_text = re.sub(r"\s+", " ", text).strip()
        chat_type = str(message.get("chat_type", "")).lower()
        if chat_type not in ("", "p2p") or not chat_id or not interviewer_open_id:
            return
        if RESUME_REVIEWER_BIND_COMMAND.fullmatch(normalized_text):
            self._bind_resume_reviewer(interviewer_open_id)
            self._reply(chat_id, "已设置为简历审核接收人。之后每次向面试官发送原始简历时，机器人都会同步私发一份给你。")
            return
        if RESUME_REVIEWER_UNBIND_COMMAND.fullmatch(normalized_text):
            if self._unbind_resume_reviewer(interviewer_open_id):
                self._reply(chat_id, "已取消简历审核接收。之后不会再向你同步简历文件。")
            else:
                self._reply(chat_id, "你当前不是简历审核接收人。")
            return
        review_command = OFFER_REVIEW_COMMAND.fullmatch(normalized_text)
        if review_command:
            action, review_id = review_command.groups()
            review = self.sync_state.offer_review_by_id(review_id)
            if review is None:
                self._reply(chat_id, "未找到该 Offer 审核编号，请检查后重试。")
                return
            if review.get("reviewer_open_id") != interviewer_open_id:
                self._reply(chat_id, "该 Offer 已分配给其他审核人，不能由你操作。")
                return
            if review.get("status") != "pending":
                status_text = {
                    "approved": "已批准",
                    "rejected": "已拒绝",
                    "superseded": "已失效",
                }.get(review.get("status", ""), "已处理")
                self._reply(chat_id, f"该 Offer {status_text}，无需重复操作。")
                return
            status = "approved" if action.lower() in {"批准offer", "同意offer"} else "rejected"
            self.sync_state.decide_offer_review(
                review_id,
                status,
                f"feishu:{header.get('event_id', '')}",
            )
            candidate_name = review.get("candidate_name") or "该候选人"
            if status == "approved":
                self._reply(chat_id, f"已批准 {candidate_name} 的 Offer；下一次同步将发送至候选人邮箱。")
            else:
                self._reply(chat_id, f"已拒绝 {candidate_name} 的 Offer；系统不会发送。")
            return
        if not self._ready():
            return
        command = COMMAND.fullmatch(normalized_text)
        if not command:
            return
        invitation_id, submitted_time = command.groups()
        submitted_time = re.sub(r":\s+", ":", submitted_time)
        try:
            chosen = self._parse_time(submitted_time)
        except ValueError:
            self._reply(chat_id, "时间格式错误，请使用 YYYY-MM-DD HH:MM。")
            return
        invitation = self.store.get(invitation_id)
        if invitation is None:
            self._reply(chat_id, "未找到该面试编号，请检查后重试。")
            return
        if chosen.strftime("%Y-%m-%d") not in {value.strftime("%Y-%m-%d") for value in self._available_dates(invitation)}:
            self._reply(chat_id, "该日期不在这份邀请的可选范围内，请重新填写。")
            return
        if not self._within_window(chosen, invitation):
            self._reply(chat_id, "该时间不在允许的面试时段内，请重新填写。")
            return
        selected_delivery_key = f"open_id:{interviewer_open_id}"
        if self._recipient_group_chat_id() and selected_delivery_key not in invitation["deliveries"]:
            self._reply(chat_id, "该候选人已随机分配给另一位面试官，不能由你确认时间。")
            return
        if not self.store.claim(invitation_id, chosen.strftime("%Y-%m-%d %H:%M"), interviewer_open_id):
            self._reply(chat_id, "该候选人的面试时间已被其他面试官确认。")
            return
        try:
            scheduled_at = chosen.strftime("%Y-%m-%d %H:%M")
            self._write_schedule(invitation["sheet_range"], scheduled_at, interviewer_open_id)
            self.store.complete(invitation_id)
            self._reply(chat_id, f"已确认 {invitation['candidate'].get('name', '该候选人')} 的面试时间：{scheduled_at}。")
        except sync.SyncError as exc:
            logging.error("Interview schedule write failed: %s", exc)
            self.store.restore_sent(invitation_id)
            self._reply(chat_id, "写入表格失败，请稍后按相同格式重新提交。")


def listener_main(config_file: Path) -> int:
    try:
        import lark_oapi as lark
    except ImportError as exc:
        raise sync.SyncError("Feishu listener requires lark-oapi. Run: python -m pip install -r requirements.txt") from exc
    config, root = sync.load_config(config_file.resolve())
    log_file = sync.resolve_path(root, str(config["runtime"].get("log_file", "logs/interview-listener.log")))
    sync.configure_logging(log_file)
    coordinator = InterviewCoordinator(config, root, config_file.resolve())

    def handle_event(data: Any) -> None:
        try:
            payload = json.loads(lark.JSON.marshal(data))
            event = payload.get("event", payload) if isinstance(payload, dict) else {}
            header = payload.get("header", {}) if isinstance(payload, dict) else {}
            message = event.get("message", {}) if isinstance(event, dict) else {}
            logging.info(
                "Feishu message event received: event_id=%s chat_type=%s",
                str(header.get("event_id", "")),
                str(message.get("chat_type", "")),
            )
            coordinator.handle_message(payload)
        except Exception as exc:  # Keep the long connection alive after a bad message.
            logging.exception("Interview listener could not process a Feishu message: %s", exc)

    def handle_bot_added(data: Any) -> None:
        try:
            payload = json.loads(lark.JSON.marshal(data))
            event = payload.get("event", payload) if isinstance(payload, dict) else {}
            logging.info("Feishu bot-added event received for chat_id=%s", str(event.get("chat_id", "")))
            coordinator.handle_bot_added(payload)
        except Exception as exc:  # Keep the long connection alive after a bad event.
            logging.exception("Interview listener could not process a bot-added event: %s", exc)

    def handle_card_action(data: Any) -> Any:
        try:
            logging.info("Feishu interview card action received.")
            return coordinator.handle_card_action(data)
        except Exception as exc:  # Return a valid callback response within the platform timeout.
            logging.exception("Interview listener could not process a card action: %s", exc)
            return coordinator._card_callback_response("error", "处理预约失败，请稍后重试。")

    handler = (
        lark.EventDispatcherHandler.builder("", "")
        .register_p2_im_message_receive_v1(handle_event)
        .register_p2_im_chat_member_bot_added_v1(handle_bot_added)
        .register_p2_card_action_trigger(handle_card_action)
        .build()
    )
    client = lark.ws.Client(
        str(config["feishu"]["app_id"]),
        str(config["feishu"]["app_secret"]),
        event_handler=handler,
        log_level=lark.LogLevel.INFO,
    )
    logging.info("Feishu interview listener started.")
    client.start()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="Run the local Feishu interview scheduling listener")
    parser.add_argument("command", choices=("listen",))
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()
    return listener_main(Path(args.config))


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except sync.SyncError as exc:
        print(f"ERROR: {exc}")
        raise SystemExit(2)
