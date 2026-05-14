"""Evaluate a saved-search query against materialised message snapshots.

A query is a tree of nodes:

    {"type": "group", "operator": "AND" | "OR", "negate": bool, "children": [...]}
    {"type": "rule", "field": "...", "operator": "...", "value": "...", "negate": bool}

The dialect intentionally stays small — the UI builds these and the backend
runs them in Python after the COM layer has produced message dicts.
"""

from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Any, Callable, Mapping


TEXT_FIELDS = {
    "subject": lambda m: m.get("subject", ""),
    "body": lambda m: m.get("body", ""),
    "any_text": lambda m: f"{m.get('subject', '')}\n{m.get('body', '')}",
    "from_name": lambda m: m.get("sender_name", ""),
    "from_email": lambda m: m.get("sender_email", ""),
    "from_any": lambda m: f"{m.get('sender_name', '')} <{m.get('sender_email', '')}>",
    "to_line": lambda m: m.get("to_line", ""),
    "recipient_email": lambda m: " ".join(
        r.get("email", "") for r in m.get("recipients", [])
    ),
    "recipient_any": lambda m: " ".join(
        f"{r.get('name', '')} <{r.get('email', '')}>"
        for r in m.get("recipients", [])
    ),
    "folder": lambda m: m.get("folder_path", ""),
}

BOOL_FIELDS = {
    "is_mailing_list": lambda m: bool(m.get("is_mailing_list")),
    "has_list_header": lambda m: bool(m.get("has_list_header")),
    "unread": lambda m: bool(m.get("unread")),
    "has_attachments": lambda m: bool(m.get("has_attachments")),
}

DATE_FIELDS = {
    "received": lambda m: m.get("received"),
}


FIELD_KINDS: dict[str, str] = (
    {k: "text" for k in TEXT_FIELDS}
    | {k: "bool" for k in BOOL_FIELDS}
    | {k: "date" for k in DATE_FIELDS}
)


def _text(value: str | None) -> str:
    return (value or "").lower()


def _eval_text(field_value: str, op: str, target: str) -> bool:
    fv = _text(field_value)
    tv = (target or "").lower()
    if op == "contains":
        return tv in fv
    if op == "not_contains":
        return tv not in fv
    if op == "equals":
        return fv == tv
    if op == "starts_with":
        return fv.startswith(tv)
    if op == "ends_with":
        return fv.endswith(tv)
    if op == "regex":
        try:
            return re.search(target or "", field_value or "", re.IGNORECASE) is not None
        except re.error:
            return False
    raise ValueError(f"Unknown text operator: {op}")


def _eval_bool(field_value: bool, op: str) -> bool:
    if op == "is_true":
        return bool(field_value)
    if op == "is_false":
        return not bool(field_value)
    raise ValueError(f"Unknown bool operator: {op}")


def _parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def _eval_date(field_value: str | None, op: str, target: str) -> bool:
    fv = _parse_date(field_value)
    tv = _parse_date(target)
    if fv is None or tv is None:
        return False
    if op == "before":
        return fv < tv
    if op == "after":
        return fv > tv
    if op == "on":
        return fv.date() == tv.date()
    raise ValueError(f"Unknown date operator: {op}")


def _eval_rule(rule: Mapping[str, Any], message: Mapping[str, Any]) -> bool:
    field = rule.get("field", "")
    op = rule.get("operator", "")
    value = rule.get("value", "")

    if field in TEXT_FIELDS:
        result = _eval_text(TEXT_FIELDS[field](message), op, str(value))
    elif field in BOOL_FIELDS:
        result = _eval_bool(BOOL_FIELDS[field](message), op)
    elif field in DATE_FIELDS:
        result = _eval_date(DATE_FIELDS[field](message), op, str(value))
    else:
        raise ValueError(f"Unknown field: {field}")

    if rule.get("negate"):
        result = not result
    return result


def evaluate(node: Mapping[str, Any], message: Mapping[str, Any]) -> bool:
    node_type = node.get("type")
    if node_type == "rule":
        return _eval_rule(node, message)
    if node_type == "group":
        children = node.get("children") or []
        if not children:
            # Empty group matches everything by convention.
            return True
        operator = (node.get("operator") or "AND").upper()
        if operator == "AND":
            result = all(evaluate(c, message) for c in children)
        elif operator == "OR":
            result = any(evaluate(c, message) for c in children)
        else:
            raise ValueError(f"Unknown group operator: {operator}")
        if node.get("negate"):
            result = not result
        return result
    raise ValueError(f"Unknown node type: {node_type!r}")


def filter_messages(
    query: Mapping[str, Any],
    messages: list[Mapping[str, Any]],
) -> list[Mapping[str, Any]]:
    return [m for m in messages if evaluate(query, m)]


__all__ = [
    "FIELD_KINDS",
    "TEXT_FIELDS",
    "BOOL_FIELDS",
    "DATE_FIELDS",
    "evaluate",
    "filter_messages",
]
