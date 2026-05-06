"""Read messages from the running Outlook client via COM (pywin32).

Only usable from the thread that called :meth:`OutlookReader.connect`, since
COM apartment-threaded objects cannot cross thread boundaries.
"""
from __future__ import annotations

import json
import logging
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator

from . import db

log = logging.getLogger(__name__)

# Target names/offices for reply detection
_REPLY_TARGETS = [
    "office of the chairperson",
    "francis saturnino juan",
    "noel salavanera",
]

# Keywords in email body that indicate an attached letter should be checked
_LETTER_KEYWORDS = [
    "attached letter",
    "attached herewith",
    "please find attached",
    "kindly check the attached",
    "please check the attached",
    "enclosed letter",
    "for your review",
    "for your consideration",
    "for your action",
    "for your reference",
    "see attached",
    "pls see attached",
    "please see attached",
    "attached for your",
    "letter attached",
]


def _extract_text_from_attachment(file_path: Path) -> str:
    """Extract text content from PDF or Word document."""
    suffix = file_path.suffix.lower()
    try:
        if suffix == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(str(file_path))
            return "\n".join(page.extract_text() or "" for page in reader.pages[:5])
        elif suffix in (".docx", ".doc"):
            from docx import Document
            doc = Document(str(file_path))
            return "\n".join(p.text for p in doc.paragraphs[:100])
    except Exception:
        log.debug("Could not extract text from %s", file_path.name, exc_info=True)
    return ""


# Single-line patterns that indicate the start of a signature or reply chain.
# These are searched line-by-line, so they must not span newlines.
_SIGNATURE_MARKERS = [
    r"^--\s*$",                    # -- (standard sig separator)
    r"^_{3,}",                     # ___ line
    r"^-{3,}",                     # --- line
    r"^Regards,?\s*$",
    r"^Best regards,?\s*$",
    r"^Kind regards,?\s*$",
    r"^Warm regards,?\s*$",
    r"^Sincerely,?\s*$",
    r"^Thank you,?\s*$",
    r"^Thanks,?\s*$",
    r"^Thanks and regards,?\s*$",
    r"^Respectfully,?\s*$",
    r"^Respectfully yours,?\s*$",
    r"^Very respectfully,?\s*$",
    r"^Sent from my ",
    r"^Sent from Mail for ",
    r"^Get Outlook for ",
    r"^On .* wrote:\s*$",          # Gmail-style reply header
    r"^DISCLAIMER",
    r"^CONFIDENTIALITY",
    r"^This email and any",
    r"^This message is intended",
    r"^NOTE: This e-mail",
]
_SIG_RE = re.compile("|".join(f"(?:{p})" for p in _SIGNATURE_MARKERS), re.IGNORECASE)

# Outlook-style reply header spans multiple lines (From: ... \n Sent: ...),
# so it has to be matched against the whole body before line-splitting.
_OUTLOOK_REPLY_HEADER_RE = re.compile(
    r"^From:.*\r?\n\s*Sent:.*$",
    re.IGNORECASE | re.MULTILINE,
)


def _clean_body(body: str) -> str:
    """Extract only the main message content, stripping signatures, disclaimers, and reply chains."""
    if not body:
        return ""
    body = body.strip()
    # Cut at the first Outlook-style reply header (multi-line "From: ...\nSent: ...").
    m = _OUTLOOK_REPLY_HEADER_RE.search(body)
    if m:
        body = body[: m.start()]
    # Find the first single-line signature/reply marker and cut there.
    clean_lines = []
    for line in body.splitlines():
        if _SIG_RE.search(line):
            break
        clean_lines.append(line)
    result = "\n".join(clean_lines).strip()
    if len(result) > 1000:
        result = result[:1000]
    return result


def _text_mentions_targets(text: str) -> bool:
    """Check if text contains any of the target names/offices."""
    text_lower = text.lower()
    return any(target in text_lower for target in _REPLY_TARGETS)


def _body_instructs_to_check_letter(body: str) -> bool:
    """Check if the email body instructs to check an attached letter."""
    body_lower = body.lower()
    return any(kw in body_lower for kw in _LETTER_KEYWORDS)


def _check_requires_reply(item: Any) -> bool:
    """Determine if this email requires a reply from the target people.

    Logic:
    1. Read the email body
    2. If body instructs to check the letter (attachment), check attachment content
    3. If attachment is addressed to Office of the Chairperson, Francis Saturnino Juan,
       or Noel Salavanera → requires reply
    """
    body = (_safe(item, "Body", "") or "").strip()
    if not body:
        return False

    # Step 1: Check if the body itself instructs to check an attached letter
    if not _body_instructs_to_check_letter(body):
        return False

    # Step 2: Body says to check attachment — now read the attachments
    try:
        attachments = getattr(item, "Attachments", None)
        if not attachments:
            return False
        count = int(getattr(attachments, "Count", 0) or 0)
        if count == 0:
            return False
        for i in range(1, count + 1):  # COM collections are 1-indexed
            att = attachments.Item(i)
            filename = getattr(att, "FileName", "") or ""
            suffix = Path(filename).suffix.lower()
            if suffix not in (".pdf", ".docx", ".doc"):
                continue
            # Save to temp file, extract text, check for target names
            with tempfile.TemporaryDirectory() as tmpdir:
                save_path = Path(tmpdir) / filename
                att.SaveAsFile(str(save_path))
                text = _extract_text_from_attachment(save_path)
                if _text_mentions_targets(text):
                    return True
    except Exception:
        log.debug("Error checking attachments for targets", exc_info=True)
    return False


# Outlook constants (from MSDN)
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT_MAIL = 5
OL_MAIL_ITEM = 43  # MailItem.Class
OL_RECIPIENT_TO = 1
OL_RECIPIENT_CC = 2
OL_RECIPIENT_BCC = 3

# MAPI property tags for SMTP addresses (work for both Exchange and external senders)
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"

_FOLDER_MAP = {"inbox": OL_FOLDER_INBOX, "sentitems": OL_FOLDER_SENT_MAIL}


class OutlookError(RuntimeError):
    pass


def _safe(obj: Any, attr: str, default: Any = None) -> Any:
    try:
        return getattr(obj, attr, default)
    except Exception:  # noqa: BLE001
        return default


def _to_utc(dt: Any) -> datetime:
    """Normalize a pywintypes.datetime (or datetime) into tz-aware UTC."""
    if dt is None:
        return datetime.now(timezone.utc)
    if not isinstance(dt, datetime):
        try:
            dt = datetime.fromtimestamp(float(dt), tz=timezone.utc)
        except Exception:  # noqa: BLE001
            return datetime.now(timezone.utc)
    if dt.tzinfo is None:
        # Outlook returns local time as naive; attach the system local tz then convert
        return dt.astimezone(timezone.utc)
    return dt.astimezone(timezone.utc)


def _sender_smtp(item: Any) -> str:
    try:
        smtp = item.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)
        if smtp:
            return str(smtp)
    except Exception:  # noqa: BLE001
        pass
    return _safe(item, "SenderEmailAddress", "") or ""


def _recipient_smtp(recip: Any) -> str:
    try:
        smtp = recip.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        if smtp:
            return str(smtp)
    except Exception:  # noqa: BLE001
        pass
    return _safe(recip, "Address", "") or ""


def _collect_recipients(item: Any, type_code: int) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    try:
        recipients = item.Recipients
        count = int(recipients.Count)
    except Exception:  # noqa: BLE001
        return out
    for i in range(1, count + 1):
        try:
            r = recipients.Item(i)
            if int(_safe(r, "Type", 0)) != type_code:
                continue
            out.append({
                "name": _safe(r, "Name", "") or "",
                "address": _recipient_smtp(r),
            })
        except Exception:  # noqa: BLE001
            continue
    return out


class OutlookReader:
    def __init__(self, shared_mailbox_smtp: str) -> None:
        self._smtp = shared_mailbox_smtp
        self._ns: Any = None
        self._folders: dict[str, Any] = {}
        self._com_initialized = False

    def connect(self) -> None:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        self._com_initialized = True

        try:
            app = win32com.client.Dispatch("Outlook.Application")
            self._ns = app.GetNamespace("MAPI")
        except Exception as exc:  # noqa: BLE001
            raise OutlookError(
                "Could not attach to Outlook. Is Outlook installed on this machine?"
            ) from exc

        # Strategy 1: Try GetSharedDefaultFolder (works for other users' shared mailboxes)
        # Strategy 2: Fall back to GetDefaultFolder (works for the signed-in user's own mailbox)
        # Strategy 3: Navigate via Inbox.Store.GetDefaultFolder for Sent Items
        #             (GetSharedDefaultFolder doesn't support olFolderSentMail)

        inbox = None
        recipient = self._ns.CreateRecipient(self._smtp)
        resolved = False
        try:
            resolved = recipient.Resolve()
        except Exception:  # noqa: BLE001
            pass

        if resolved:
            # Try shared-folder access for Inbox
            try:
                inbox = self._ns.GetSharedDefaultFolder(recipient, OL_FOLDER_INBOX)
                self._folders["inbox"] = inbox
            except Exception:  # noqa: BLE001
                pass

        # Fallback: use default folders (the signed-in user's own mailbox)
        if "inbox" not in self._folders:
            try:
                inbox = self._ns.GetDefaultFolder(OL_FOLDER_INBOX)
                self._folders["inbox"] = inbox
            except Exception as exc:  # noqa: BLE001
                raise OutlookError(
                    f"Could not open Inbox for {self._smtp!r}. "
                    f"Make sure Outlook is running and the mailbox is accessible."
                ) from exc

        # Sent Items: GetSharedDefaultFolder doesn't support olFolderSentMail,
        # so navigate via the Inbox's Store object instead.
        try:
            store = inbox.Store
            self._folders["sentitems"] = store.GetDefaultFolder(OL_FOLDER_SENT_MAIL)
        except Exception:  # noqa: BLE001
            # Last resort: walk up from Inbox to the store root and find Sent Items by name
            try:
                store_root = inbox.Parent
                for name in ("Sent Items", "Sent", "Sent Mail"):
                    try:
                        self._folders["sentitems"] = store_root.Folders[name]
                        break
                    except Exception:  # noqa: BLE001
                        continue
            except Exception:  # noqa: BLE001
                pass

        if "sentitems" not in self._folders:
            log.warning(
                "Could not open Sent Items for %s — only inbound mail will be tracked.",
                self._smtp,
            )

        log.info(
            "Connected to Outlook (folders: %s)",
            ", ".join(self._folders.keys()),
        )

    def disconnect(self) -> None:
        self._folders.clear()
        self._ns = None
        if self._com_initialized:
            try:
                import pythoncom

                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
            self._com_initialized = False

    def iter_since(
        self,
        folder_key: str,
        since: datetime | None,
        direction: str,
    ) -> Iterator[dict[str, Any]]:
        """Yield message rows from a folder with ReceivedTime strictly after ``since``.

        Items are enumerated newest-first; iteration stops as soon as we see a
        message at or before ``since``.
        """
        folder = self._folders.get(folder_key)
        if folder is None:
            return  # folder wasn't available at connect time; skip silently

        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)  # True = descending
        except Exception as exc:  # noqa: BLE001
            raise OutlookError(f"Failed to enumerate {folder_key} items") from exc

        # COM collection is 1-indexed. Using GetFirst/GetNext avoids a Python
        # iteration protocol quirk with some Outlook item types.
        item = items.GetFirst()
        while item is not None:
            try:
                if int(_safe(item, "Class", 0)) != OL_MAIL_ITEM:
                    item = items.GetNext()
                    continue
                received = _to_utc(_safe(item, "ReceivedTime"))
                if since is not None and received <= since:
                    return
                row = self._map_item(item, folder_key, direction, received)
                if row is not None:
                    yield row
            except Exception:  # noqa: BLE001
                log.exception("Skipping malformed item in %s", folder_key)
            item = items.GetNext()

    def _map_item(
        self,
        item: Any,
        folder_key: str,
        direction: str,
        received_utc: datetime,
    ) -> dict[str, Any] | None:
        entry_id = _safe(item, "EntryID")
        if not entry_id:
            return None
        subject = _safe(item, "Subject", "") or ""
        subject_lower = subject.lower()
        return {
            "id": str(entry_id),
            "conversation_id": _safe(item, "ConversationID") or _safe(item, "ConversationTopic", ""),
            "direction": direction,
            "folder": folder_key,
            "sender_name": _safe(item, "SenderName", "") or "",
            "sender_address": _sender_smtp(item),
            "to_recipients": json.dumps(
                _collect_recipients(item, OL_RECIPIENT_TO), ensure_ascii=False
            ),
            "cc_recipients": json.dumps(
                _collect_recipients(item, OL_RECIPIENT_CC), ensure_ascii=False
            ),
            "bcc_recipients": json.dumps(
                _collect_recipients(item, OL_RECIPIENT_BCC), ensure_ascii=False
            ),
            "subject": subject,
            "received_at": received_utc.isoformat(),
            "is_reply": 1 if subject_lower.startswith("re:") else 0,
            "is_forward": 1 if subject_lower.startswith(("fw:", "fwd:")) else 0,
            "has_attachments": 1 if int(getattr(_safe(item, "Attachments"), "Count", 0) or 0) > 0 else 0,
            "requires_reply": 1 if _check_requires_reply(item) else 0,
            "body_preview": _clean_body((_safe(item, "Body", "") or "")),
            "ingested_at": db.now_utc_iso(),
        }
