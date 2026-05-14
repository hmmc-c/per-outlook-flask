"""Thin wrapper around Outlook's COM API (MAPI namespace).

All functions assume the caller has called ``pythoncom.CoInitialize()`` for the
current thread. ``with_com`` is provided as a convenience for Flask routes.
"""

from __future__ import annotations

import functools
import re
from contextlib import contextmanager
from typing import Any, Callable, Iterable, Iterator

try:  # pragma: no cover - import only works on Windows with Outlook installed
    import pythoncom
    import win32com.client
    _HAS_OUTLOOK = True
except Exception:  # noqa: BLE001
    pythoncom = None  # type: ignore[assignment]
    win32com = None  # type: ignore[assignment]
    _HAS_OUTLOOK = False


# Property tag for full internet (transport) headers; lets us inspect List-Id etc.
PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"

# Outlook item class constant; only mail items are searchable here.
MAIL_ITEM_CLASS = 43


class OutlookUnavailable(RuntimeError):
    """Raised when pywin32/Outlook cannot be reached."""


@contextmanager
def with_com() -> Iterator[None]:
    """Initialize COM for the current thread (Flask uses a worker pool)."""
    if not _HAS_OUTLOOK:
        raise OutlookUnavailable(
            "pywin32 is not installed or Outlook is not available on this host."
        )
    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def com(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator: run a function inside a COM-initialized block."""

    @functools.wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        with with_com():
            return func(*args, **kwargs)

    return wrapper


def _get_namespace():
    app = win32com.client.Dispatch("Outlook.Application")
    return app.GetNamespace("MAPI")


def _folder_key(folder) -> str:
    """Stable identifier for a folder across calls."""
    return f"{folder.EntryID}|{folder.StoreID}"


def _resolve_folder(ns, key: str):
    entry_id, store_id = key.split("|", 1)
    return ns.GetFolderFromID(entry_id, store_id)


def _safe_count(folder) -> int:
    try:
        return int(folder.Items.Count)
    except Exception:  # noqa: BLE001
        return 0


def list_folders() -> list[dict[str, Any]]:
    """Return every mail folder across every store as a flat list.

    Each entry has ``id``, ``name``, ``path`` (Outlook ``FolderPath``), ``store``,
    and ``depth`` for tree rendering in the UI.
    """
    ns = _get_namespace()
    result: list[dict[str, Any]] = []

    def walk(folder, depth: int, store_name: str) -> None:
        result.append(
            {
                "id": _folder_key(folder),
                "name": folder.Name,
                "path": folder.FolderPath,
                "store": store_name,
                "depth": depth,
                "count": _safe_count(folder),
            }
        )
        for sub in folder.Folders:
            walk(sub, depth + 1, store_name)

    for store in ns.Folders:
        walk(store, 0, store.Name)

    return result


def _headers_for(item) -> str:
    try:
        return item.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS) or ""
    except Exception:  # noqa: BLE001
        return ""


_LIST_HEADER_RE = re.compile(
    r"^(List-Unsubscribe|List-Id|List-Post|Precedence|X-Mailer-Recipient):",
    re.IGNORECASE | re.MULTILINE,
)
_MAILING_LIST_SENDER_RE = re.compile(
    r"(no[-_.]?reply|do[-_.]?not[-_.]?reply|newsletter|mailer|notification|updates|info|news)@",
    re.IGNORECASE,
)


def _sender_email(item) -> str:
    """Best-effort SMTP address for the sender (handles Exchange senders)."""
    try:
        addr = item.SenderEmailAddress or ""
    except Exception:  # noqa: BLE001
        addr = ""
    if addr and "@" in addr:
        return addr
    try:
        sender = item.Sender
        if sender is not None:
            exch_user = sender.GetExchangeUser()
            if exch_user is not None and exch_user.PrimarySmtpAddress:
                return exch_user.PrimarySmtpAddress
    except Exception:  # noqa: BLE001
        pass
    return addr


def _recipients(item) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    try:
        for r in item.Recipients:
            try:
                addr = r.Address or ""
                if "@" not in addr:
                    try:
                        exch_user = r.AddressEntry.GetExchangeUser()
                        if exch_user and exch_user.PrimarySmtpAddress:
                            addr = exch_user.PrimarySmtpAddress
                    except Exception:  # noqa: BLE001
                        pass
                out.append({"name": r.Name or "", "email": addr})
            except Exception:  # noqa: BLE001
                continue
    except Exception:  # noqa: BLE001
        pass
    return out


def _message_snapshot(item, folder_id: str, folder_path: str) -> dict[str, Any]:
    """Materialise the subset of fields the search/UI needs."""
    headers = _headers_for(item)
    has_list_header = bool(_LIST_HEADER_RE.search(headers))
    sender_email = _sender_email(item)
    is_mailing_list = has_list_header or bool(
        sender_email and _MAILING_LIST_SENDER_RE.search(sender_email)
    )

    try:
        received = item.ReceivedTime
        received_iso = received.isoformat() if received else None
    except Exception:  # noqa: BLE001
        received_iso = None

    try:
        body = item.Body or ""
    except Exception:  # noqa: BLE001
        body = ""

    try:
        subject = item.Subject or ""
    except Exception:  # noqa: BLE001
        subject = ""

    try:
        sender_name = item.SenderName or ""
    except Exception:  # noqa: BLE001
        sender_name = ""

    try:
        unread = bool(item.UnRead)
    except Exception:  # noqa: BLE001
        unread = False

    try:
        has_attachments = item.Attachments.Count > 0
    except Exception:  # noqa: BLE001
        has_attachments = False

    try:
        to_field = item.To or ""
    except Exception:  # noqa: BLE001
        to_field = ""

    return {
        "id": f"{item.EntryID}|{item.Parent.StoreID}",
        "entry_id": item.EntryID,
        "store_id": item.Parent.StoreID,
        "folder_id": folder_id,
        "folder_path": folder_path,
        "subject": subject,
        "sender_name": sender_name,
        "sender_email": sender_email,
        "to_line": to_field,
        "recipients": _recipients(item),
        "body": body,
        "received": received_iso,
        "unread": unread,
        "has_attachments": has_attachments,
        "has_list_header": has_list_header,
        "is_mailing_list": is_mailing_list,
    }


def _iter_folder_items(folder, since: str | None, limit: int) -> Iterator[Any]:
    """Yield mail items in the folder, newest first, optionally date-filtered.

    ``Restrict`` is used for the date filter so Outlook itself does the work.
    """
    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    if since:
        # Outlook expects "MM/DD/YYYY HH:MM AM/PM" for the Restrict clause.
        items = items.Restrict(f"[ReceivedTime] >= '{since}'")
        items.Sort("[ReceivedTime]", True)
    count = 0
    for item in items:
        if count >= limit:
            return
        try:
            cls = item.Class
        except Exception:  # noqa: BLE001
            continue
        if cls != MAIL_ITEM_CLASS:
            continue
        yield item
        count += 1


def fetch_messages(
    folder_ids: Iterable[str],
    *,
    since: str | None = None,
    per_folder_limit: int = 500,
) -> list[dict[str, Any]]:
    """Pull message snapshots from each selected folder."""
    ns = _get_namespace()
    snapshots: list[dict[str, Any]] = []
    for fid in folder_ids:
        try:
            folder = _resolve_folder(ns, fid)
        except Exception:  # noqa: BLE001
            continue
        folder_path = folder.FolderPath
        for item in _iter_folder_items(folder, since, per_folder_limit):
            try:
                snapshots.append(_message_snapshot(item, fid, folder_path))
            except Exception:  # noqa: BLE001
                # Skip messages we can't read (corrupted, encrypted, etc.).
                continue
    snapshots.sort(key=lambda m: m["received"] or "", reverse=True)
    return snapshots


def move_messages(message_ids: Iterable[str], target_folder_id: str) -> dict[str, Any]:
    """Move each ``EntryID|StoreID`` to ``target_folder_id``."""
    ns = _get_namespace()
    target = _resolve_folder(ns, target_folder_id)
    moved: list[str] = []
    errors: list[dict[str, str]] = []
    for mid in message_ids:
        try:
            entry_id, store_id = mid.split("|", 1)
            item = ns.GetItemFromID(entry_id, store_id)
            item.Move(target)
            moved.append(mid)
        except Exception as exc:  # noqa: BLE001
            errors.append({"id": mid, "error": str(exc)})
    return {"moved": moved, "errors": errors}
