"""Akses database SQLite untuk Universal Invoice Maker.

Skema hanya berisi proyek, item RAB, riwayat, dan pengaturan. Tidak ada tabel
kendaraan, termin, atau alokasi.
"""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Sequence

from .models import Item, ParsedItem, Project
from .utils import date_to_iso

SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    client_name TEXT NOT NULL DEFAULT '',
    client_address TEXT NOT NULL DEFAULT '',
    po_prefix TEXT NOT NULL DEFAULT '',
    invoice_date TEXT NOT NULL,
    date_format TEXT NOT NULL DEFAULT 'full',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    row_no INTEGER NOT NULL DEFAULT 0,
    name TEXT NOT NULL,
    unit TEXT NOT NULL DEFAULT '',
    unit_price INTEGER NOT NULL DEFAULT 0,
    total_qty INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER REFERENCES projects(id) ON DELETE CASCADE,
    ts TEXT NOT NULL,
    action TEXT NOT NULL,
    detail TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_items_project ON items(project_id);
"""


class Database:
    def __init__(self, path: str | Path):
        self.path = str(path)
        if self.path != ":memory:":
            Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    # ------------------------------------------------------------------ utils
    def close(self) -> None:
        self.conn.close()

    @contextmanager
    def transaction(self):
        try:
            self.conn.execute("BEGIN")
            yield self.conn
            self.conn.commit()
        except Exception:
            self.conn.rollback()
            raise

    def _now(self) -> str:
        return datetime.now().isoformat(timespec="seconds")

    # --------------------------------------------------------------- projects
    def list_projects(self) -> list[Project]:
        rows = self.conn.execute(
            "SELECT * FROM projects ORDER BY id DESC"
        ).fetchall()
        return [Project.from_row(r) for r in rows]

    def get_project(self, project_id: int) -> Project | None:
        row = self.conn.execute(
            "SELECT * FROM projects WHERE id = ?", (project_id,)
        ).fetchone()
        return Project.from_row(row) if row else None

    def create_project(
        self,
        name: str,
        client_name: str = "",
        client_address: str = "",
        po_prefix: str = "",
        invoice_date: date | None = None,
        date_format: str = "full",
    ) -> int:
        invoice_date = invoice_date or date.today()
        with self.transaction() as conn:
            cur = conn.execute(
                """INSERT INTO projects
                   (name, client_name, client_address, po_prefix,
                    invoice_date, date_format, created_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    name.strip() or "Proyek Baru",
                    client_name,
                    client_address,
                    po_prefix,
                    date_to_iso(invoice_date),
                    date_format or "full",
                    self._now(),
                ),
            )
            project_id = int(cur.lastrowid)
            conn.execute(
                "INSERT INTO audit_log (project_id, ts, action, detail) VALUES (?, ?, ?, ?)",
                (project_id, self._now(), "buat_proyek", f"Proyek '{name}' dibuat"),
            )
        return project_id

    def update_project(
        self,
        project_id: int,
        *,
        name: str | None = None,
        client_name: str | None = None,
        client_address: str | None = None,
        po_prefix: str | None = None,
        invoice_date: date | None = None,
        date_format: str | None = None,
    ) -> None:
        fields: list[str] = []
        values: list[object] = []
        for key, value in (
            ("name", name),
            ("client_name", client_name),
            ("client_address", client_address),
            ("po_prefix", po_prefix),
            ("date_format", date_format),
        ):
            if value is not None:
                fields.append(f"{key} = ?")
                values.append(value)
        if invoice_date is not None:
            fields.append("invoice_date = ?")
            values.append(date_to_iso(invoice_date))
        if not fields:
            return
        values.append(project_id)
        with self.transaction() as conn:
            conn.execute(
                f"UPDATE projects SET {', '.join(fields)} WHERE id = ?", values
            )

    def delete_project(self, project_id: int) -> None:
        with self.transaction() as conn:
            conn.execute("DELETE FROM projects WHERE id = ?", (project_id,))

    # ------------------------------------------------------------------ items
    def list_items(self, project_id: int) -> list[Item]:
        rows = self.conn.execute(
            "SELECT * FROM items WHERE project_id = ? ORDER BY row_no, id",
            (project_id,),
        ).fetchall()
        return [Item.from_row(r) for r in rows]

    def replace_rab(self, project_id: int, items: Sequence[ParsedItem]) -> int:
        """Ganti seluruh item RAB proyek. Return jumlah item yang disimpan."""
        with self.transaction() as conn:
            conn.execute("DELETE FROM items WHERE project_id = ?", (project_id,))
            for index, item in enumerate(items, start=1):
                conn.execute(
                    """INSERT INTO items
                       (project_id, row_no, name, unit, unit_price, total_qty)
                       VALUES (?, ?, ?, ?, ?, ?)""",
                    (
                        project_id,
                        item.source_row or index,
                        item.name,
                        item.unit,
                        int(item.unit_price),
                        int(item.total_qty),
                    ),
                )
            conn.execute(
                "INSERT INTO audit_log (project_id, ts, action, detail) VALUES (?, ?, ?, ?)",
                (
                    project_id,
                    self._now(),
                    "import_rab",
                    f"{len(items)} baris RAB diimpor",
                ),
            )
        return len(items)

    # --------------------------------------------------------------- audit log
    def log(self, project_id: int | None, action: str, detail: str = "") -> None:
        with self.transaction() as conn:
            conn.execute(
                "INSERT INTO audit_log (project_id, ts, action, detail) VALUES (?, ?, ?, ?)",
                (project_id, self._now(), action, detail),
            )

    def list_log(self, project_id: int, limit: int = 300) -> list[sqlite3.Row]:
        return self.conn.execute(
            """SELECT * FROM audit_log WHERE project_id = ?
               ORDER BY id DESC LIMIT ?""",
            (project_id, limit),
        ).fetchall()

    # ---------------------------------------------------------------- settings
    def get_setting(self, key: str, default: str = "") -> str:
        row = self.conn.execute(
            "SELECT value FROM settings WHERE key = ?", (key,)
        ).fetchone()
        return row["value"] if row else default

    def set_setting(self, key: str, value: str) -> None:
        with self.transaction() as conn:
            conn.execute(
                """INSERT INTO settings (key, value) VALUES (?, ?)
                   ON CONFLICT(key) DO UPDATE SET value = excluded.value""",
                (key, value),
            )
