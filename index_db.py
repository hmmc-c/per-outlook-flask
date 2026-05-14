"""SQLite-backed local index of Outlook messages.

Each operation opens its own connection (cheap, and keeps Flask's worker pool
+ background indexer threads happy). WAL mode lets searches run while the
indexer is writing.
"""

from __future__ import annotations

import json
import re
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator


SCHEMA = """
CREATE TABLE IF NOT EXISTS folders (
  id TEXT PRIMARY KEY,
  store TEXT NOT NULL,
  path TEXT NOT NULL,
  last_indexed_at TEXT,
  message_count INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS messages (
  id TEXT PRIMARY KEY,
  folder_id TEXT NOT NULL,
  entry_id TEXT NOT NULL,
  store_id TEXT NOT NULL,
  folder_path TEXT NOT NULL,
  subject TEXT,
  sender_name TEXT,
  sender_email TEXT,
  to_line TEXT,
  recipients_json TEXT,
  recipient_emails TEXT,
  received TEXT,
  unread INTEGER DEFAULT 0,
  has_attachments INTEGER DEFAULT 0,
  has_list_header INTEGER DEFAULT 0,
  is_mailing_list INTEGER DEFAULT 0,
  body TEXT
);

CREATE INDEX IF NOT EXISTS idx_messages_folder ON messages(folder_id);
CREATE INDEX IF NOT EXISTS idx_messages_received ON messages(received);
CREATE INDEX IF NOT EXISTS idx_messages_sender ON messages(sender_email);
"""


def _regexp(pattern: str | None, value: str | None) -> bool:
    if not pattern or value is None:
        return False
    try:
        return re.search(pattern, value, re.IGNORECASE) is not None
    except re.error:
        return False


class IndexDB:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._init()

    @contextmanager
    def _conn(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.create_function("REGEXP", 2, _regexp)
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init(self) -> None:
        with self._conn() as c:
            c.execute("PRAGMA journal_mode=WAL")
            c.executescript(SCHEMA)

    # ----- folders -----

    def upsert_folder(self, folder_id: str, store: str, path: str) -> None:
        with self._conn() as c:
            c.execute(
                """
                INSERT INTO folders(id, store, path) VALUES(?,?,?)
                ON CONFLICT(id) DO UPDATE SET store=excluded.store, path=excluded.path
                """,
                (folder_id, store, path),
            )

    def mark_indexed(self, folder_id: str, message_count: int, when_iso: str) -> None:
        with self._conn() as c:
            c.execute(
                "UPDATE folders SET last_indexed_at=?, message_count=? WHERE id=?",
                (when_iso, message_count, folder_id),
            )

    def folders_status(self) -> list[dict[str, Any]]:
        with self._conn() as c:
            return [dict(r) for r in c.execute("SELECT * FROM folders ORDER BY path")]

    def folder_status(self, folder_id: str) -> dict[str, Any] | None:
        with self._conn() as c:
            row = c.execute(
                "SELECT * FROM folders WHERE id=?", (folder_id,)
            ).fetchone()
        return dict(row) if row else None

    # ----- messages -----

    def delete_folder_messages(self, folder_id: str) -> None:
        with self._conn() as c:
            c.execute("DELETE FROM messages WHERE folder_id=?", (folder_id,))

    def upsert_messages(self, snapshots: Iterable[dict[str, Any]]) -> None:
        rows = [
            (
                s["id"], s["folder_id"], s["entry_id"], s["store_id"], s["folder_path"],
                s.get("subject"),
                s.get("sender_name"),
                s.get("sender_email"),
                s.get("to_line"),
                json.dumps(s.get("recipients") or []),
                " ".join(r.get("email", "") for r in (s.get("recipients") or [])),
                s.get("received"),
                int(bool(s.get("unread"))),
                int(bool(s.get("has_attachments"))),
                int(bool(s.get("has_list_header"))),
                int(bool(s.get("is_mailing_list"))),
                s.get("body"),
            )
            for s in snapshots
        ]
        if not rows:
            return
        with self._conn() as c:
            c.executemany(
                """
                INSERT INTO messages(id, folder_id, entry_id, store_id, folder_path,
                  subject, sender_name, sender_email, to_line, recipients_json,
                  recipient_emails, received, unread, has_attachments,
                  has_list_header, is_mailing_list, body)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(id) DO UPDATE SET
                  folder_id=excluded.folder_id,
                  folder_path=excluded.folder_path,
                  subject=excluded.subject,
                  sender_name=excluded.sender_name,
                  sender_email=excluded.sender_email,
                  to_line=excluded.to_line,
                  recipients_json=excluded.recipients_json,
                  recipient_emails=excluded.recipient_emails,
                  received=excluded.received,
                  unread=excluded.unread,
                  has_attachments=excluded.has_attachments,
                  has_list_header=excluded.has_list_header,
                  is_mailing_list=excluded.is_mailing_list,
                  body=excluded.body
                """,
                rows,
            )

    def delete_messages(self, message_ids: Iterable[str]) -> None:
        ids = [(mid,) for mid in message_ids]
        if not ids:
            return
        with self._conn() as c:
            c.executemany("DELETE FROM messages WHERE id=?", ids)

    def count_messages(self, folder_id: str | None = None) -> int:
        with self._conn() as c:
            if folder_id:
                row = c.execute(
                    "SELECT COUNT(*) AS c FROM messages WHERE folder_id=?",
                    (folder_id,),
                ).fetchone()
            else:
                row = c.execute("SELECT COUNT(*) AS c FROM messages").fetchone()
        return int(row["c"])

    def count_in_scope(
        self,
        folder_ids: list[str] | None,
        since: str | None,
    ) -> int:
        clauses: list[str] = []
        params: list[Any] = []
        if folder_ids:
            placeholders = ",".join("?" * len(folder_ids))
            clauses.append(f"folder_id IN ({placeholders})")
            params.extend(folder_ids)
        if since:
            clauses.append("received >= ?")
            params.append(since)
        where = " AND ".join(clauses) if clauses else "1=1"
        with self._conn() as c:
            row = c.execute(
                f"SELECT COUNT(*) AS c FROM messages WHERE {where}", params
            ).fetchone()
        return int(row["c"])

    def search(
        self,
        where_sql: str,
        params: list[Any],
        *,
        limit: int = 1000,
    ) -> list[dict[str, Any]]:
        sql = (
            f"SELECT * FROM messages WHERE {where_sql} "
            f"ORDER BY received DESC LIMIT ?"
        )
        with self._conn() as c:
            return [dict(r) for r in c.execute(sql, list(params) + [limit])]
