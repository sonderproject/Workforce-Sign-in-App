"""Client identity: matching, creation, updates, history."""

from __future__ import annotations

import sqlite3
from typing import List, Optional

from ..database import now_iso
from ..models import normalize_name


def find_matches(conn: sqlite3.Connection, first_name: str, last_name: str,
                 dob_iso: str) -> List[sqlite3.Row]:
    """Match on normalized first + last name + DOB (case/whitespace-insensitive)."""
    cur = conn.execute(
        """
        SELECT * FROM clients
        WHERE first_norm = ? AND last_norm = ? AND date_of_birth = ?
        ORDER BY id
        """,
        (normalize_name(first_name), normalize_name(last_name), dob_iso),
    )
    return cur.fetchall()


def search_clients(conn: sqlite3.Connection, term: str) -> List[sqlite3.Row]:
    """Loose search used by the admin dashboard."""
    like = f"%{normalize_name(term)}%"
    cur = conn.execute(
        """
        SELECT * FROM clients
        WHERE first_norm LIKE ? OR last_norm LIKE ?
           OR (last_norm || ' ' || first_norm) LIKE ?
           OR (first_norm || ' ' || last_norm) LIKE ?
        ORDER BY last_norm, first_norm
        LIMIT 200
        """,
        (like, like, like, like),
    )
    return cur.fetchall()


def get_client(conn: sqlite3.Connection, client_id: int) -> Optional[sqlite3.Row]:
    cur = conn.execute("SELECT * FROM clients WHERE id = ?", (client_id,))
    return cur.fetchone()


def create_client(conn: sqlite3.Connection, first_name: str, last_name: str,
                  dob_iso: str, uid: Optional[str] = None) -> int:
    ts = now_iso()
    cur = conn.execute(
        """
        INSERT INTO clients
            (first_name, last_name, date_of_birth, uid,
             first_norm, last_norm, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (first_name.strip(), last_name.strip(), dob_iso, (uid or "").strip() or None,
         normalize_name(first_name), normalize_name(last_name), ts, ts),
    )
    return cur.lastrowid


def update_client(conn: sqlite3.Connection, client_id: int, first_name: str,
                  last_name: str, dob_iso: str, uid: Optional[str] = None) -> None:
    conn.execute(
        """
        UPDATE clients
           SET first_name = ?, last_name = ?, date_of_birth = ?, uid = ?,
               first_norm = ?, last_norm = ?, updated_at = ?
         WHERE id = ?
        """,
        (first_name.strip(), last_name.strip(), dob_iso, (uid or "").strip() or None,
         normalize_name(first_name), normalize_name(last_name), now_iso(), client_id),
    )
