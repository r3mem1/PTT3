"""Единая точка конфигурации: переменные окружения и пути.

Все модули читают настройки отсюда, а не вызывают load_dotenv() по отдельности.
Так проще тестировать и невозможно «забыть» проверить токен.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")


def _env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def _env_int(name: str, default: int) -> int:
    raw = _env(name)
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError as exc:
        raise RuntimeError(f"{name} должен быть целым числом, получено: {raw!r}") from exc
    if value < 1:
        raise RuntimeError(f"{name} должен быть >= 1")
    return value


@dataclass(frozen=True)
class Settings:
    telegram_bot_token: str
    openrouter_api_key: str
    pexels_api_key: str
    openrouter_url: str
    model: str
    slides_count: int
    template_path: Path
    output_dir: Path
    image_dir: Path
    logo_dir: Path
    log_level: str
    topic_min_len: int
    topic_max_len: int


def load_settings() -> Settings:
    token = _env("TELEGRAM_BOT_TOKEN")
    if not token:
        raise RuntimeError("Не задан TELEGRAM_BOT_TOKEN. Скопируйте .env.example в .env.")

    openrouter_key = _env("OPENROUTER_API_KEY")
    if not openrouter_key:
        raise RuntimeError("Не задан OPENROUTER_API_KEY. Скопируйте .env.example в .env.")

    template_path = Path(_env("TEMPLATE_PATH") or str(BASE_DIR / "templates" / "template.pptx"))
    if not template_path.is_file():
        raise RuntimeError(f"Не найден шаблон презентации: {template_path}")

    output_dir = Path(_env("OUTPUT_DIR") or str(BASE_DIR / "output"))
    image_dir = output_dir / "images"
    logo_dir = output_dir / "logos"

    return Settings(
        telegram_bot_token=token,
        openrouter_api_key=openrouter_key,
        pexels_api_key=_env("PEXELS_API_KEY"),
        openrouter_url=_env("OPENROUTER_URL") or "https://openrouter.ai/api/v1/chat/completions",
        model=_env("OPENROUTER_MODEL") or "inclusionai/ling-3.0-flash-vl:free",
        slides_count=_env_int("SLIDES_COUNT", 6),
        template_path=template_path,
        output_dir=output_dir,
        image_dir=image_dir,
        logo_dir=logo_dir,
        log_level=_env("LOG_LEVEL") or "INFO",
        topic_min_len=_env_int("TOPIC_MIN_LEN", 3),
        topic_max_len=_env_int("TOPIC_MAX_LEN", 300),
    )


settings = load_settings()
