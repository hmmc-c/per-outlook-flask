"""JSON-file persistence for saved searches.

Each saved search bundles a query tree, the folders to search across, and an
optional action (currently only ``move`` to a target folder). Stored as a flat
list keyed by uuid.
"""

from __future__ import annotations

import json
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


_LOCK = threading.Lock()


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


class SavedSearchStore:
    def __init__(self, path: Path) -> None:
        self.path = path

    def _read(self) -> list[dict[str, Any]]:
        if not self.path.exists():
            return []
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return []

    def _write(self, items: list[dict[str, Any]]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        tmp.write_text(json.dumps(items, indent=2), encoding="utf-8")
        tmp.replace(self.path)

    def list(self) -> list[dict[str, Any]]:
        with _LOCK:
            return self._read()

    def get(self, search_id: str) -> dict[str, Any] | None:
        with _LOCK:
            for item in self._read():
                if item.get("id") == search_id:
                    return item
            return None

    def create(self, payload: dict[str, Any]) -> dict[str, Any]:
        with _LOCK:
            items = self._read()
            entry = {
                "id": str(uuid.uuid4()),
                "name": payload.get("name") or "Untitled search",
                "folders": payload.get("folders") or [],
                "query": payload.get("query") or {"type": "group", "operator": "AND", "children": []},
                "action": payload.get("action"),
                "since": payload.get("since"),
                "per_folder_limit": payload.get("per_folder_limit") or 500,
                "created_at": _now_iso(),
                "updated_at": _now_iso(),
            }
            items.append(entry)
            self._write(items)
            return entry

    def update(self, search_id: str, payload: dict[str, Any]) -> dict[str, Any] | None:
        with _LOCK:
            items = self._read()
            for idx, item in enumerate(items):
                if item.get("id") != search_id:
                    continue
                merged = {**item, **{k: v for k, v in payload.items() if k != "id"}}
                merged["updated_at"] = _now_iso()
                items[idx] = merged
                self._write(items)
                return merged
        return None

    def delete(self, search_id: str) -> bool:
        with _LOCK:
            items = self._read()
            filtered = [i for i in items if i.get("id") != search_id]
            if len(filtered) == len(items):
                return False
            self._write(filtered)
            return True
