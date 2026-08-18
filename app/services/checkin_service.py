"""Visit / check-in records. Every physical visit produces a Visit row."""

from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import List, Optional

from ..database import now_iso
from ..mapping import DEFAULT_SERVICES


def create_visit(conn: sqlite3.Connection, client_id: int,
                 monthly_intake_id: Optional[int] = None,
                 visitor_type: Optional[str] = None,
                 services: Optional[str] = None,
                 when: Optional[datetime] = None,
                 dedupe_seconds: int = 90) -> tuple[int, bool]:
    """Create a visit for *today*.

    A short dedupe window guards against a double-clicked ``CHECK IN`` button:
    if the same client already has a visit within ``dedupe_seconds``, the
    existing visit is returned instead of creating a duplicate.
    Returns (visit_id, created).
    """
    when = when or datetime.now()
    visit_date = when.strftime("%Y-%m-%d")
    visit_time = when.strftime("%H:%M")

    recent = conn.execute(
        """
        SELECT id, created_at FROM visits
        WHERE client_id = ? AND visit_date = ?
        ORDER BY id DESC LIMIT 1
        """,
        (client_id, visit_date),
    ).fetchone()
    if recent is not None:
        try:
            prev = datetime.fromisoformat(recent["created_at"])
            if (when - prev).total_seconds() < dedupe_seconds:
                return recent["id"], False
        except (ValueError, TypeError):
            pass

    cur = conn.execute(
        """
        INSERT INTO visits
            (client_id, monthly_intake_id, visit_date, visit_time,
             reporting_year, reporting_month, visitor_type, services, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (client_id, monthly_intake_id, visit_date, visit_time,
         when.year, when.month, visitor_type,
         services or DEFAULT_SERVICES, now_iso()),
    )
    return cur.lastrowid, True


def list_visits(conn: sqlite3.Connection, client_id: int) -> List[sqlite3.Row]:
    cur = conn.execute(
        """
        SELECT * FROM visits WHERE client_id = ?
        ORDER BY visit_date DESC, visit_time DESC
        """,
        (client_id,),
    )
    return cur.fetchall()
