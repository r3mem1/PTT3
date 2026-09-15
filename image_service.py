"""
Поиск и скачивание изображений под тему слайда через Pexels API
(бесплатный ключ: https://www.pexels.com/api/).

Улучшения:
- берём несколько кандидатов и выбираем не всегда первый, а тот, чьи
  пропорции ближе к нужным (landscape ~16:10) — меньше случайных
  нерелевантных/неудобных для обрезки картинок;
- кешируем байты картинки по нормализованному запросу в памяти процесса,
  чтобы одинаковый image_query на разных слайдах/генерациях не долбил
  сеть повторно (LRU на ограниченное число запросов).
"""

from __future__ import annotations

import logging
import uuid
from collections import OrderedDict
from pathlib import Path

import aiohttp
from PIL import Image, ImageStat

from config import settings

log = logging.getLogger(__name__)

PEXELS_SEARCH_URL = "https://api.pexels.com/v1/search"
REQUEST_TIMEOUT = aiohttp.ClientTimeout(total=20)
MAX_IMAGE_BYTES = 8 * 1024 * 1024
CANDIDATES_PER_QUERY = 3
TARGET_RATIO = 16 / 10  # ориентир: почти все картинки в колоде landscape

_CACHE_MAX_ITEMS = 200
_image_bytes_cache: "OrderedDict[str, bytes]" = OrderedDict()


def _normalize_query(query: str) -> str:
    return " ".join(query.strip().lower().split())


def _cache_get(key: str) -> bytes | None:
    if key not in _image_bytes_cache:
        return None
    _image_bytes_cache.move_to_end(key)
    return _image_bytes_cache[key]


def _cache_put(key: str, content: bytes) -> None:
    _image_bytes_cache[key] = content
    _image_bytes_cache.move_to_end(key)
    while len(_image_bytes_cache) > _CACHE_MAX_ITEMS:
        _image_bytes_cache.popitem(last=False)


def _pick_best_photo(photos: list[dict]) -> dict | None:
    if not photos:
        return None

    def score(photo: dict) -> float:
        w = photo.get("width") or 0
        h = photo.get("height") or 0
        if not w or not h:
            return float("inf")
        ratio = w / h
        return abs(ratio - TARGET_RATIO)

    return min(photos, key=score)


async def _write_temp_file(content: bytes) -> Path:
    settings.image_dir.mkdir(parents=True, exist_ok=True)
    file_path = settings.image_dir / f"{uuid.uuid4().hex}.jpg"
    file_path.write_bytes(content)
    return file_path


async def find_and_download_image(query: str) -> Path | None:
    """Ищет картинку по смыслу запроса и скачивает её на диск.

    Возвращает путь к файлу или None, если картинку не удалось получить.
    Слайд в этом случае собирается без изображения.
    """
    query = query.strip()
    if not query:
        return None

    cache_key = _normalize_query(query)
    cached = _cache_get(cache_key)
    if cached is not None:
        log.debug("Картинка для %r взята из кеша, сеть не дёргаем", query)
        return await _write_temp_file(cached)

    if not settings.pexels_api_key:
        log.warning("PEXELS_API_KEY не задан в .env — картинки не ищутся.")
        return None

    try:
        async with aiohttp.ClientSession(timeout=REQUEST_TIMEOUT) as session:
            params = {"query": query, "per_page": CANDIDATES_PER_QUERY, "orientation": "landscape"}
            headers = {"Authorization": settings.pexels_api_key}

            async with session.get(PEXELS_SEARCH_URL, params=params, headers=headers) as resp:
                if resp.status != 200:
                    body = await resp.text()
                    log.warning(
                        "Pexels статус %s для запроса %r: %s",
                        resp.status,
                        query,
                        body[:200],
                    )
                    return None
                data = await resp.json()

            photos = data.get("photos") or []
            best = _pick_best_photo(photos)
            if best is None:
                log.warning("Pexels не нашёл картинок по запросу %r", query)
                return None

            image_url = best["src"]["large"]

            async with session.get(image_url) as img_resp:
                if img_resp.status != 200:
                    log.warning("Не удалось скачать картинку %s (статус %s)", image_url, img_resp.status)
                    return None
                content = await img_resp.read()
                if len(content) > MAX_IMAGE_BYTES:
                    log.warning("Картинка слишком большая (%s байт), пропускаю", len(content))
                    return None

    except aiohttp.ClientError as e:
        log.warning("Сетевая ошибка при обращении к Pexels: %s", e)
        return None
    except TimeoutError:
        log.warning("Pexels не ответил вовремя для запроса %r", query)
        return None

    _cache_put(cache_key, content)
    file_path = await _write_temp_file(content)
    log.info("Картинка для %r сохранена: %s", query, file_path)
    return file_path


def analyze_brightness(path: Path) -> float:
    """Средняя яркость картинки (0-255) — используется для подбора светлого
    или тёмного текста поверх кастомного фона презентации."""
    with Image.open(str(path)) as im:
        return ImageStat.Stat(im.convert("L")).mean[0]
