"""Background indexer: pulls messages from Outlook into the SQLite IndexDB.

Each folder is reindexed in its own thread. ``Indexer.status()`` returns a
snapshot of all in-flight + recent jobs so the UI can poll for progress.
"""

from __future__ import annotations

import threading
from datetime import datetime, timezone
from typing import Any, Iterable

import outlook_client
from index_db import IndexDB


BATCH_SIZE = 200


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class Indexer:
    def __init__(self, db: IndexDB) -> None:
        self.db = db
        self._jobs: dict[str, dict[str, Any]] = {}
        self._lock = threading.Lock()

    def status(self) -> dict[str, dict[str, Any]]:
        with self._lock:
            return {fid: dict(info) for fid, info in self._jobs.items()}

    def _set(self, folder_id: str, **fields: Any) -> None:
        with self._lock:
            self._jobs.setdefault(folder_id, {})
            self._jobs[folder_id].update(fields)

    def _is_running(self, folder_id: str) -> bool:
        with self._lock:
            cur = self._jobs.get(folder_id)
            return bool(cur and cur.get("state") in ("queued", "running"))

    def start(self, folder_ids: Iterable[str]) -> list[str]:
        started: list[str] = []
        for fid in folder_ids:
            if self._is_running(fid):
                continue
            self._set(
                fid,
                state="queued",
                progress=0,
                total=0,
                error=None,
                started_at=None,
                finished_at=None,
            )
            t = threading.Thread(target=self._run, args=(fid,), daemon=True)
            t.start()
            started.append(fid)
        return started

    def _run(self, folder_id: str) -> None:
        try:
            self._set(folder_id, state="running", started_at=_now_iso())
            with outlook_client.with_com():
                ns = outlook_client.get_namespace()
                folder = outlook_client.resolve_folder(ns, folder_id)
                store_name = ""
                try:
                    store_name = folder.Store.DisplayName
                except Exception:  # noqa: BLE001
                    pass
                self.db.upsert_folder(folder_id, store_name, folder.FolderPath)

                items = folder.Items
                items.Sort("[ReceivedTime]", True)
                try:
                    total = int(items.Count)
                except Exception:  # noqa: BLE001
                    total = 0
                self._set(folder_id, total=total)

                # Full reindex: clear then repopulate. Simpler than diffing and
                # safe because EntryIDs are stable inside a folder.
                self.db.delete_folder_messages(folder_id)

                batch: list[dict[str, Any]] = []
                processed = 0
                for item in items:
                    processed += 1
                    try:
                        if item.Class != outlook_client.MAIL_ITEM_CLASS:
                            continue
                        snap = outlook_client.message_snapshot(
                            item, folder_id, folder.FolderPath
                        )
                        batch.append(snap)
                    except Exception:  # noqa: BLE001
                        # Skip messages we can't materialise (corrupt, encrypted, …).
                        continue
                    if len(batch) >= BATCH_SIZE:
                        self.db.upsert_messages(batch)
                        batch = []
                        self._set(folder_id, progress=processed)
                if batch:
                    self.db.upsert_messages(batch)

                indexed = self.db.count_messages(folder_id)
                self.db.mark_indexed(folder_id, indexed, _now_iso())
                self._set(
                    folder_id,
                    state="done",
                    progress=processed,
                    indexed=indexed,
                    finished_at=_now_iso(),
                )
        except outlook_client.OutlookUnavailable as exc:
            self._set(folder_id, state="error", error=str(exc), finished_at=_now_iso())
        except Exception as exc:  # noqa: BLE001
            self._set(folder_id, state="error", error=str(exc), finished_at=_now_iso())
