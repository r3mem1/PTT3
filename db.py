"""SQLite-хранилище настроек пользователей и истории генераций.

Лёгкая замена глобальным .env-настройкам: каждый пользователь может задать
своё число слайдов/тон/язык/тему оформления через /settings, а история
последних тем позволяет повторить генерацию одной кнопкой (см. bot.py).
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import aiosqlite

from config import settings

DB_PATH: Path = settings.output_dir / "bot.sqlite3"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    slides_count INTEGER,
    tone TEXT,
    lang TEXT,
    theme_preset TEXT,
    logo_path TEXT
);

CREATE TABLE IF NOT EXISTS generations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    topic TEXT NOT NULL,
    created_at REAL NOT NULL
);
"""

DEFAULT_TONE = "business"
DEFAULT_LANG = "ru"
DEFAULT_THEME = "warm"
HISTORY_LIMIT = 10


async def init_db() -> None:
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.executescript(_SCHEMA)
        await db.commit()


async def get_user_settings(user_id: int) -> dict[str, Any]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT slides_count, tone, lang, theme_preset, logo_path FROM users WHERE user_id = ?",
            (user_id,),
        ) as cursor:
            row = await cursor.fetchone()

    if row is None:
        return {
            "slides_count": settings.slides_count,
            "tone": DEFAULT_TONE,
            "lang": DEFAULT_LANG,
            "theme_preset": DEFAULT_THEME,
            "logo_path": None,
        }
    return {
        "slides_count": row["slides_count"] or settings.slides_count,
        "tone": row["tone"] or DEFAULT_TONE,
        "lang": row["lang"] or DEFAULT_LANG,
        "theme_preset": row["theme_preset"] or DEFAULT_THEME,
        "logo_path": row["logo_path"],
    }


async def set_user_setting(user_id: int, **fields: Any) -> None:
    current = await get_user_settings(user_id)
    current.update({k: v for k, v in fields.items() if v is not None or k == "logo_path"})
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            """
            INSERT INTO users (user_id, slides_count, tone, lang, theme_preset, logo_path)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(user_id) DO UPDATE SET
                slides_count = excluded.slides_count,
                tone = excluded.tone,
                lang = excluded.lang,
                theme_preset = excluded.theme_preset,
                logo_path = excluded.logo_path
            """,
            (
                user_id,
                current["slides_count"],
                current["tone"],
                current["lang"],
                current["theme_preset"],
                current["logo_path"],
            ),
        )
        await db.commit()


async def add_generation(user_id: int, topic: str) -> None:
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT INTO generations (user_id, topic, created_at) VALUES (?, ?, ?)",
            (user_id, topic, time.time()),
        )
        # Держим только последние HISTORY_LIMIT записей на пользователя, чтобы таблица не росла бесконечно.
        await db.execute(
            """
            DELETE FROM generations
            WHERE user_id = ? AND id NOT IN (
                SELECT id FROM generations WHERE user_id = ? ORDER BY created_at DESC LIMIT ?
            )
            """,
            (user_id, user_id, HISTORY_LIMIT),
        )
        await db.commit()


async def get_history(user_id: int, limit: int = HISTORY_LIMIT) -> list[dict[str, Any]]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT topic, created_at FROM generations WHERE user_id = ? ORDER BY created_at DESC LIMIT ?",
            (user_id, limit),
        ) as cursor:
            rows = await cursor.fetchall()
    return [{"topic": r["topic"], "created_at": r["created_at"]} for r in rows]
