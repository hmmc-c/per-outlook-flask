"""Translate the saved-search query tree into a SQL WHERE clause for IndexDB.

The dialect mirrors ``search.py`` (the in-memory evaluator) so the same
front-end builder can target either backend.
"""

from __future__ import annotations

from typing import Any, Mapping


TEXT_COLUMNS = {
    "subject": "subject",
    "body": "body",
    "any_text": "(IFNULL(subject,'') || char(10) || IFNULL(body,''))",
    "from_name": "sender_name",
    "from_email": "sender_email",
    "from_any": "(IFNULL(sender_name,'') || ' <' || IFNULL(sender_email,'') || '>')",
    "to_line": "to_line",
    "recipient_email": "recipient_emails",
    "recipient_any": "recipient_emails",
    "folder": "folder_path",
}

BOOL_COLUMNS = {
    "is_mailing_list": "is_mailing_list",
    "has_list_header": "has_list_header",
    "unread": "unread",
    "has_attachments": "has_attachments",
}

DATE_COLUMNS = {
    "received": "received",
}


def _text_clause(col: str, op: str, value: str) -> tuple[str, list[Any]]:
    val = (value or "").lower()
    if op == "contains":
        return f"LOWER(IFNULL({col},'')) LIKE ?", [f"%{val}%"]
    if op == "not_contains":
        return f"LOWER(IFNULL({col},'')) NOT LIKE ?", [f"%{val}%"]
    if op == "equals":
        return f"LOWER(IFNULL({col},'')) = ?", [val]
    if op == "starts_with":
        return f"LOWER(IFNULL({col},'')) LIKE ?", [f"{val}%"]
    if op == "ends_with":
        return f"LOWER(IFNULL({col},'')) LIKE ?", [f"%{val}"]
    if op == "regex":
        return f"({col} IS NOT NULL AND {col} REGEXP ?)", [value or ""]
    raise ValueError(f"Unknown text operator: {op}")


def _bool_clause(col: str, op: str) -> tuple[str, list[Any]]:
    if op == "is_true":
        return f"{col} = 1", []
    if op == "is_false":
        return f"({col} = 0 OR {col} IS NULL)", []
    raise ValueError(f"Unknown bool operator: {op}")


def _date_clause(col: str, op: str, value: str) -> tuple[str, list[Any]]:
    if not value:
        return "1=0", []
    if op == "before":
        return f"{col} < ?", [value]
    if op == "after":
        return f"{col} > ?", [value]
    if op == "on":
        return f"date({col}) = date(?)", [value]
    raise ValueError(f"Unknown date operator: {op}")


def _build(node: Mapping[str, Any]) -> tuple[str, list[Any]]:
    node_type = node.get("type")
    if node_type == "rule":
        field = node.get("field", "")
        op = node.get("operator", "")
        value = node.get("value", "") or ""
        if field in TEXT_COLUMNS:
            sql, params = _text_clause(TEXT_COLUMNS[field], op, str(value))
        elif field in BOOL_COLUMNS:
            sql, params = _bool_clause(BOOL_COLUMNS[field], op)
        elif field in DATE_COLUMNS:
            sql, params = _date_clause(DATE_COLUMNS[field], op, str(value))
        else:
            raise ValueError(f"Unknown field: {field}")
        if node.get("negate"):
            sql = f"NOT ({sql})"
        return sql, params
    if node_type == "group":
        children = node.get("children") or []
        if not children:
            return "1=1", []
        op = (node.get("operator") or "AND").upper()
        joiner = " AND " if op == "AND" else " OR "
        parts: list[str] = []
        params: list[Any] = []
        for c in children:
            s, p = _build(c)
            parts.append(f"({s})")
            params.extend(p)
        sql = joiner.join(parts)
        if node.get("negate"):
            sql = f"NOT ({sql})"
        return sql, params
    raise ValueError(f"Unknown node type: {node_type!r}")


def to_sql(
    query: Mapping[str, Any],
    *,
    folder_ids: list[str] | None = None,
    since: str | None = None,
) -> tuple[str, list[Any]]:
    sql, params = _build(query)
    parts = [f"({sql})"]
    if folder_ids:
        placeholders = ",".join("?" * len(folder_ids))
        parts.append(f"folder_id IN ({placeholders})")
        params.extend(folder_ids)
    if since:
        parts.append("received >= ?")
        params.append(since)
    return " AND ".join(parts), params
