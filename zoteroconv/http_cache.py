"""Tiny SQLite cache for CrossRef / PubMed lookups (spec: never re-hit public
APIs for the same query; Zotero API responses are NOT cached -- live state)."""

from __future__ import annotations

import json
import sqlite3
import time
from pathlib import Path
from typing import Any

SCHEMA = """
CREATE TABLE IF NOT EXISTS http_cache (
    kind       TEXT NOT NULL,
    key        TEXT NOT NULL,
    payload    TEXT NOT NULL,
    fetched_at REAL NOT NULL,
    PRIMARY KEY (kind, key)
)
"""


class HttpCache:
    """
    ``db_path=None`` disables caching entirely (every get is a miss, every set
    is a no-op) so tests and --no-cache runs share one code path.
    """

    def __init__(self, db_path: str | Path | None):
        self.db_path: Path | None = Path(db_path) if db_path is not None else None
        self._conn: sqlite3.Connection | None = None
        if self.db_path is not None:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            self._conn = sqlite3.connect(str(self.db_path))
            self._conn.execute(SCHEMA)
            self._conn.commit()

    @property
    def enabled(self) -> bool:
        return self._conn is not None

    def get(self, kind: str, key: str) -> Any | None:
        if self._conn is None:
            return None
        row = self._conn.execute(
            "SELECT payload FROM http_cache WHERE kind=? AND key=?", (kind, key)
        ).fetchone()
        if row is None:
            return None
        return json.loads(row[0])

    def set(self, kind: str, key: str, payload: Any) -> None:
        if self._conn is None:
            return
        self._conn.execute(
            "INSERT OR REPLACE INTO http_cache (kind, key, payload, fetched_at) "
            "VALUES (?, ?, ?, ?)",
            (kind, key, json.dumps(payload), time.time()),
        )
        self._conn.commit()

    def close(self) -> None:
        if self._conn is not None:
            self._conn.close()
            self._conn = None

    def __enter__(self) -> "HttpCache":
        return self

    def __exit__(self, *exc) -> None:
        self.close()
