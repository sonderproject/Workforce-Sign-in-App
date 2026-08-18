"""SQLite connection and schema management."""

from __future__ import annotations

import os
import sqlite3
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
DEFAULT_DB_PATH = os.path.join(DATA_DIR, "app.db")


SCHEMA = """
CREATE TABLE IF NOT EXISTS clients (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    first_name    TEXT NOT NULL,
    last_name     TEXT NOT NULL,
    date_of_birth TEXT NOT NULL,            -- ISO YYYY-MM-DD
    uid           TEXT,
    first_norm    TEXT NOT NULL,            -- normalized for matching
    last_norm     TEXT NOT NULL,
    created_at    TEXT NOT NULL,
    updated_at    TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_clients_match
    ON clients (last_norm, first_norm, date_of_birth);

CREATE TABLE IF NOT EXISTS monthly_intakes (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id             INTEGER NOT NULL REFERENCES clients(id),
    reporting_year        INTEGER NOT NULL,
    reporting_month       INTEGER NOT NULL,
    veteran               INTEGER NOT NULL DEFAULT 0,
    hispanic              INTEGER NOT NULL DEFAULT 0,
    race                  TEXT,               -- white/black/asian/american_indian/other_multi
    house_size            INTEGER,
    household_type        TEXT,               -- single_non_elderly/elderly/single_parent/two_parent/other
    population_category   TEXT,               -- adult/tay/family_with_minor/senior
    female_head           INTEGER NOT NULL DEFAULT 0,
    disabled              INTEGER NOT NULL DEFAULT 0,
    created_at            TEXT NOT NULL,
    updated_at            TEXT NOT NULL,
    UNIQUE (client_id, reporting_year, reporting_month)
);

CREATE TABLE IF NOT EXISTS visits (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id         INTEGER NOT NULL REFERENCES clients(id),
    monthly_intake_id INTEGER REFERENCES monthly_intakes(id),
    visit_date        TEXT NOT NULL,          -- ISO YYYY-MM-DD
    visit_time        TEXT NOT NULL,          -- HH:MM
    reporting_year    INTEGER NOT NULL,
    reporting_month   INTEGER NOT NULL,
    visitor_type      TEXT,
    services          TEXT,
    created_at        TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_visits_client ON visits (client_id);
CREATE INDEX IF NOT EXISTS idx_visits_month ON visits (reporting_year, reporting_month);
"""


def get_connection(db_path: str = DEFAULT_DB_PATH) -> sqlite3.Connection:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(db_path: str = DEFAULT_DB_PATH) -> None:
    conn = get_connection(db_path)
    try:
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")
