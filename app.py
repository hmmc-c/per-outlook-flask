"""Flask front-end for searching and moving local Outlook mail."""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, render_template, request

import outlook_client
import search as search_module
from storage import SavedSearchStore


BASE_DIR = Path(__file__).resolve().parent
SAVED_SEARCHES_PATH = Path(
    os.environ.get("OUTLOOK_SAVED_SEARCHES", BASE_DIR / "saved_searches.json")
)


def create_app() -> Flask:
    app = Flask(__name__)
    store = SavedSearchStore(SAVED_SEARCHES_PATH)

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
            }
        )

    @app.get("/api/folders")
    def folders() -> Any:
        try:
            with outlook_client.with_com():
                data = outlook_client.list_folders()
            return jsonify({"folders": data})
        except outlook_client.OutlookUnavailable as exc:
            return jsonify({"error": str(exc)}), 503
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Failed to list folders: {exc}"}), 500

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
        per_folder_limit = int(payload.get("per_folder_limit") or 500)
        result_limit = int(payload.get("result_limit") or 1000)

        if not folder_ids:
            return jsonify({"error": "Pick at least one folder to search."}), 400

        try:
            with outlook_client.with_com():
                messages = outlook_client.fetch_messages(
                    folder_ids,
                    since=since,
                    per_folder_limit=per_folder_limit,
                )
            matched = search_module.filter_messages(query, messages)
        except outlook_client.OutlookUnavailable as exc:
            return jsonify({"error": str(exc)}), 503
        except ValueError as exc:
            return jsonify({"error": f"Invalid query: {exc}"}), 400
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Search failed: {exc}"}), 500

        truncated = len(matched) > result_limit
        return jsonify(
            {
                "scanned": len(messages),
                "matched": len(matched),
                "truncated": truncated,
                "results": [
                    {
                        "id": m["id"],
                        "folder_path": m["folder_path"],
                        "subject": m["subject"],
                        "sender_name": m["sender_name"],
                        "sender_email": m["sender_email"],
                        "to_line": m["to_line"],
                        "received": m["received"],
                        "unread": m["unread"],
                        "has_attachments": m["has_attachments"],
                        "is_mailing_list": m["is_mailing_list"],
                    }
                    for m in matched[:result_limit]
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
            return jsonify(result)
        except outlook_client.OutlookUnavailable as exc:
            return jsonify({"error": str(exc)}), 503
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Move failed: {exc}"}), 500

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
