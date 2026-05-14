"""Flask front-end for searching and moving local Outlook mail.

Searches run against a local SQLite index (built per-folder via
``/api/index``) for speed. Move operations still talk to Outlook over COM and
keep the index in sync by deleting moved messages from it.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request

import outlook_client
import search as search_module
import search_sql
from index_db import IndexDB
from indexer import Indexer
from storage import SavedSearchStore


BASE_DIR = Path(__file__).resolve().parent
SAVED_SEARCHES_PATH = Path(
    os.environ.get("OUTLOOK_SAVED_SEARCHES", BASE_DIR / "saved_searches.json")
)
INDEX_DB_PATH = Path(
    os.environ.get("OUTLOOK_INDEX_DB", BASE_DIR / "outlook_index.sqlite3")
)


def create_app() -> Flask:
    app = Flask(__name__)
    store = SavedSearchStore(SAVED_SEARCHES_PATH)
    db = IndexDB(INDEX_DB_PATH)
    indexer = Indexer(db)

    @app.get("/")
    def index() -> str:
        return render_template(
            "index.html",
            field_kinds=search_module.FIELD_KINDS,
        )

    @app.get("/api/status")
    def status() -> Any:
        return jsonify(
            {
                "outlook_available": outlook_client._HAS_OUTLOOK,
                "index_path": str(INDEX_DB_PATH),
            }
        )

    @app.get("/api/folders")
    def folders() -> Any:
        try:
            with outlook_client.with_com():
                data = outlook_client.list_folders()
            index_status = {f["id"]: f for f in db.folders_status()}
            for f in data:
                meta = index_status.get(f["id"])
                if meta:
                    f["indexed_count"] = meta.get("message_count")
                    f["last_indexed_at"] = meta.get("last_indexed_at")
                else:
                    f["indexed_count"] = None
                    f["last_indexed_at"] = None
            return jsonify({"folders": data})
        except outlook_client.OutlookUnavailable as exc:
            return jsonify({"error": str(exc)}), 503
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Failed to list folders: {exc}"}), 500

    # ---------- index ----------

    @app.post("/api/index")
    def start_index() -> Any:
        payload = request.get_json(force=True, silent=True) or {}
        folder_ids = payload.get("folders") or []
        if not folder_ids:
            return jsonify({"error": "Pick at least one folder to index."}), 400
        if not outlook_client._HAS_OUTLOOK:
            return jsonify({"error": "Outlook is not available."}), 503
        started = indexer.start(folder_ids)
        return jsonify({"started": started, "jobs": indexer.status()}), 202

    @app.get("/api/index/status")
    def index_status() -> Any:
        return jsonify(
            {
                "jobs": indexer.status(),
                "folders": db.folders_status(),
                "total_messages": db.count_messages(),
            }
        )

    # ---------- search ----------

    @app.post("/api/search")
    def run_search() -> Any:
        payload = request.get_json(force=True, silent=True) or {}
        folder_ids = payload.get("folders") or []
        query = payload.get("query") or {
            "type": "group",
            "operator": "AND",
            "children": [],
        }
        since = payload.get("since") or None
        result_limit = int(payload.get("result_limit") or 1000)

        try:
            where, params = search_sql.to_sql(
                query, folder_ids=folder_ids or None, since=since
            )
            scope_count = db.count_in_scope(folder_ids or None, since)
            rows = db.search(where, params, limit=result_limit)
        except ValueError as exc:
            return jsonify({"error": f"Invalid query: {exc}"}), 400
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Search failed: {exc}"}), 500

        return jsonify(
            {
                "scanned": scope_count,
                "matched": len(rows),
                "truncated": len(rows) >= result_limit,
                "results": [
                    {
                        "id": r["id"],
                        "folder_path": r["folder_path"],
                        "subject": r["subject"],
                        "sender_name": r["sender_name"],
                        "sender_email": r["sender_email"],
                        "to_line": r["to_line"],
                        "received": r["received"],
                        "unread": bool(r["unread"]),
                        "has_attachments": bool(r["has_attachments"]),
                        "is_mailing_list": bool(r["is_mailing_list"]),
                    }
                    for r in rows
                ],
            }
        )

    @app.post("/api/move")
    def move() -> Any:
        payload = request.get_json(force=True, silent=True) or {}
        message_ids = payload.get("message_ids") or []
        target = payload.get("target_folder_id")
        if not message_ids:
            return jsonify({"error": "No messages selected."}), 400
        if not target:
            return jsonify({"error": "No target folder selected."}), 400
        try:
            with outlook_client.with_com():
                result = outlook_client.move_messages(message_ids, target)
            # Drop moved IDs from the index — their EntryIDs change after the
            # move and the destination folder may not be indexed yet.
            if result.get("moved"):
                db.delete_messages(result["moved"])
            return jsonify(result)
        except outlook_client.OutlookUnavailable as exc:
            return jsonify({"error": str(exc)}), 503
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Move failed: {exc}"}), 500

    # ---------- saved searches ----------

    @app.get("/api/saved-searches")
    def list_saved() -> Any:
        return jsonify({"items": store.list()})

    @app.post("/api/saved-searches")
    def create_saved() -> Any:
        payload = request.get_json(force=True, silent=True) or {}
        return jsonify(store.create(payload)), 201

    @app.put("/api/saved-searches/<search_id>")
    def update_saved(search_id: str) -> Any:
        payload = request.get_json(force=True, silent=True) or {}
        updated = store.update(search_id, payload)
        if updated is None:
            return jsonify({"error": "Not found"}), 404
        return jsonify(updated)

    @app.delete("/api/saved-searches/<search_id>")
    def delete_saved(search_id: str) -> Any:
        if not store.delete(search_id):
            return jsonify({"error": "Not found"}), 404
        return jsonify({"ok": True})

    return app


app = create_app()


if __name__ == "__main__":
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("DEBUG", "0") == "1"
    app.run(host=host, port=port, debug=debug)
