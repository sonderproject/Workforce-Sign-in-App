"""Monthly intake records: one per client per calendar month."""

from __future__ import annotations

import sqlite3
from typing import List, Optional

from ..database import now_iso

INTAKE_FIELDS = (
    "veteran", "hispanic", "race", "house_size",
    "household_type", "population_category", "female_head", "disabled",
)


def get_intake_for_month(conn: sqlite3.Connection, client_id: int, year: int,
                         month: int) -> Optional[sqlite3.Row]:
    cur = conn.execute(
        """
        SELECT * FROM monthly_intakes
        WHERE client_id = ? AND reporting_year = ? AND reporting_month = ?
        """,
        (client_id, year, month),
    )
    return cur.fetchone()


def has_completed_intake(conn: sqlite3.Connection, client_id: int, year: int,
                         month: int) -> bool:
    return get_intake_for_month(conn, client_id, year, month) is not None


def get_latest_intake(conn: sqlite3.Connection, client_id: int) -> Optional[sqlite3.Row]:
    """Most recent intake, used to prefill a new month's form."""
    cur = conn.execute(
        """
        SELECT * FROM monthly_intakes
        WHERE client_id = ?
        ORDER BY reporting_year DESC, reporting_month DESC, id DESC
        LIMIT 1
        """,
        (client_id,),
    )
    return cur.fetchone()


def list_intakes(conn: sqlite3.Connection, client_id: int) -> List[sqlite3.Row]:
    cur = conn.execute(
        """
        SELECT * FROM monthly_intakes WHERE client_id = ?
        ORDER BY reporting_year DESC, reporting_month DESC
        """,
        (client_id,),
    )
    return cur.fetchall()


def create_or_get_intake(conn: sqlite3.Connection, client_id: int, year: int,
                         month: int, data: dict) -> tuple[int, bool]:
    """Create the month's intake if absent.

    Uniqueness (client_id + year + month) is enforced by the DB. This is
    idempotent: a duplicate submission returns the existing record instead of
    creating a second row. Returns (intake_id, created).
    """
    existing = get_intake_for_month(conn, client_id, year, month)
    if existing is not None:
        return existing["id"], False

    ts = now_iso()
    values = {f: data.get(f) for f in INTAKE_FIELDS}
    try:
        cur = conn.execute(
            """
            INSERT INTO monthly_intakes
                (client_id, reporting_year, reporting_month,
                 veteran, hispanic, race, house_size, household_type,
                 population_category, female_head, disabled, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (client_id, year, month,
             int(bool(values["veteran"])), int(bool(values["hispanic"])),
             values["race"], values["house_size"], values["household_type"],
             values["population_category"], int(bool(values["female_head"])),
             int(bool(values["disabled"])), ts, ts),
        )
        return cur.lastrowid, True
    except sqlite3.IntegrityError:
        # Lost a race against a concurrent identical submission.
        existing = get_intake_for_month(conn, client_id, year, month)
        return existing["id"], False


def update_intake(conn: sqlite3.Connection, intake_id: int, data: dict) -> None:
    """Correct an existing intake (admin edit). Does NOT touch other months."""
    conn.execute(
        """
        UPDATE monthly_intakes
           SET veteran = ?, hispanic = ?, race = ?, house_size = ?,
               household_type = ?, population_category = ?, female_head = ?,
               disabled = ?, updated_at = ?
         WHERE id = ?
        """,
        (int(bool(data.get("veteran"))), int(bool(data.get("hispanic"))),
         data.get("race"), data.get("house_size"), data.get("household_type"),
         data.get("population_category"), int(bool(data.get("female_head"))),
         int(bool(data.get("disabled"))), now_iso(), intake_id),
    )
