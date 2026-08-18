"""Data-access helpers and normalization utilities."""

from __future__ import annotations

import re
import sqlite3
from datetime import date, datetime
from typing import Optional


def normalize_name(value: str) -> str:
    """Case-insensitive, whitespace-insensitive normalization for matching."""
    if value is None:
        return ""
    value = value.strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def parse_dob(value) -> Optional[date]:
    """Accept an ISO string, date, or datetime and return a ``date``."""
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            continue
    return None


def client_to_dict(row: sqlite3.Row) -> dict:
    d = dict(row)
    d["date_of_birth"] = parse_dob(d.get("date_of_birth"))
    return d


def intake_to_dict(row: sqlite3.Row) -> dict:
    return dict(row)


def visit_to_dict(row: sqlite3.Row) -> dict:
    return dict(row)
