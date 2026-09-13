"""Типы данных, которые ожидаем от ИИ и передаём в сборщик .pptx."""

from __future__ import annotations

from typing import Any, Literal, TypedDict

SlideType = Literal["title", "content", "section", "closing"]
SlideLayout = Literal[
    "text",
    "image_left",
    "image_right",
    "cards",
    "stat",
    "quote",
]


class CardData(TypedDict, total=False):
    title: str
    body: str


class SlideData(TypedDict, total=False):
    type: SlideType
    layout: SlideLayout
    title: str
    subtitle: str
    kicker: str
    bullets: list[str]
    image_query: str
    metric: str
    metric_label: str
    cards: list[CardData]


class PresentationData(TypedDict):
    title: str
    slides: list[SlideData]


ALLOWED_TYPES = {"title", "content", "section", "closing"}
ALLOWED_LAYOUTS = {"text", "image_left", "image_right", "cards", "stat", "quote"}

BULLET_CHAR_LIMIT = 220


def _trim_bullet(text: str, limit: int = BULLET_CHAR_LIMIT) -> str:
    """Мягко обрезает длинный буллет по границе слова, чтобы он не вылезал
    за пределы карточки/колонки на слайде (авто-подгонка шрифта в
    pptx_builder помогает, но не спасает от откровенно длинных абзацев)."""
    text = text.strip()
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(",.;:—-")
    return (cut or text[:limit]) + "…"


def _as_cards(item: dict) -> list[CardData]:
    raw = item.get("cards")
    cards: list[CardData] = []
    if isinstance(raw, list):
        for card in raw:
            if not isinstance(card, dict):
                continue
            title = str(card.get("title") or "").strip()
            body = str(card.get("body") or card.get("text") or "").strip()
            if title or body:
                cards.append({"title": title or body, "body": body if title else ""})
    if cards:
        return cards[:3]

    bullets = item.get("bullets") or []
    if not isinstance(bullets, list):
        return []
    for bullet in bullets[:3]:
        text = str(bullet).strip()
        if " — " in text:
            title, body = text.split(" — ", 1)
        elif ": " in text:
            title, body = text.split(": ", 1)
        else:
            title, body = text, ""
        cards.append({"title": title.strip(), "body": body.strip()})
    return cards


def normalize_presentation(data: Any) -> PresentationData:
    """Приводит сырой JSON модели к ожидаемой форме."""
    if not isinstance(data, dict):
        raise ValueError("Ответ ИИ должен быть JSON-объектом.")

    title = str(data.get("title") or "presentation").strip() or "presentation"
    raw_slides = data.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        raise ValueError("В ответе ИИ нет ни одного слайда.")

    slides: list[SlideData] = []
    for item in raw_slides:
        if not isinstance(item, dict):
            continue

        slide_type = item.get("type", "content")
        if slide_type not in ALLOWED_TYPES:
            slide_type = "content"

        layout = item.get("layout", "text")
        if slide_type in {"section", "closing"}:
            layout = "quote" if slide_type == "section" else "text"
        elif layout not in ALLOWED_LAYOUTS:
            layout = "text"

        bullets = item.get("bullets") or []
        if not isinstance(bullets, list):
            bullets = [str(bullets)]
        bullets = [_trim_bullet(str(b)) for b in bullets if str(b).strip()][:4]

        slide: SlideData = {
            "type": slide_type,  # type: ignore[typeddict-item]
            "layout": layout,  # type: ignore[typeddict-item]
            "title": str(item.get("title") or "Без названия").strip()[:90],
            "subtitle": str(item.get("subtitle") or "").strip()[:180],
            "kicker": str(item.get("kicker") or "").strip()[:48],
            "bullets": bullets,
            "metric": str(item.get("metric") or "").strip()[:24],
            "metric_label": str(item.get("metric_label") or "").strip()[:80],
            "cards": _as_cards(item),
        }
        query = str(item.get("image_query") or "").strip()
        if query:
            slide["image_query"] = query
        slides.append(slide)

    if not slides:
        raise ValueError("После проверки не осталось валидных слайдов.")

    return {"title": title, "slides": slides}
