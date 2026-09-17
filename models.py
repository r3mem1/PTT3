"""Типы данных, которые ожидаем от ИИ и передаём в сборщик .pptx."""

from __future__ import annotations

from typing import Any, Literal, TypedDict

from text_cleanup import clean_text

SlideType = Literal["title", "content", "section", "closing"]
SlideLayout = Literal[
    "text",
    "image_left",
    "image_right",
    "cards",
    "stat",
    "quote",
    "divider",
    "chart",
    "timeline",
    "comparison",
    "team",
    "big_photo",
]


class CardData(TypedDict, total=False):
    title: str
    body: str


class ChartSeries(TypedDict):
    name: str
    values: list[float]


class ChartData(TypedDict):
    type: Literal["bar", "pie", "line"]
    categories: list[str]
    series: list[ChartSeries]


class TimelineItem(TypedDict, total=False):
    label: str
    title: str
    body: str


class TeamMember(TypedDict, total=False):
    name: str
    role: str
    note: str


class ComparisonSide(TypedDict, total=False):
    title: str
    bullets: list[str]


class ComparisonData(TypedDict, total=False):
    left: ComparisonSide
    right: ComparisonSide


class SlideData(TypedDict, total=False):
    type: SlideType
    layout: SlideLayout
    title: str
    subtitle: str
    bullets: list[str]
    body: str
    image_query: str
    metric: str
    metric_label: str
    cards: list[CardData]
    chart: ChartData
    timeline: list[TimelineItem]
    team: list[TeamMember]
    comparison: ComparisonData


class PresentationData(TypedDict):
    title: str
    slides: list[SlideData]


ALLOWED_TYPES = {"title", "content", "section", "closing"}
ALLOWED_LAYOUTS = {
    "text",
    "image_left",
    "image_right",
    "cards",
    "stat",
    "quote",
    "divider",
    "chart",
    "timeline",
    "comparison",
    "team",
    "big_photo",
}

CHART_TYPES = {"bar", "pie", "line"}

BULLET_CHAR_LIMIT = 220

# Детерминированный предохранитель для визуальных форматов (см. ai_service.py):
# сколько буллетов и какой их длины оставлять, независимо от того, насколько
# точно модель следует текстовым инструкциям формата в промпте.
_FORMAT_BULLET_LIMITS = {
    "image_only": (1, 90),
    "image_heavy": (2, 160),
    "balanced": (4, BULLET_CHAR_LIMIT),
}
# То же самое, но для связного абзаца в поле "body" (альтернатива буллетам).
_FORMAT_BODY_LIMITS = {
    "image_only": 110,
    "image_heavy": 220,
    "balanced": 420,
}
DEFAULT_VISUAL_FORMAT = "balanced"


def _trim_bullet(text: str, limit: int = BULLET_CHAR_LIMIT) -> str:
    """Чистит воду/markdown и мягко обрезает буллет по границе слова, чтобы
    он не вылезал за пределы карточки/колонки на слайде (авто-подгонка
    шрифта в pptx_builder помогает, но не спасает от откровенно длинных
    абзацев)."""
    text = clean_text(text.strip())
    if len(text) <= limit:
        return text
    cut = text[:limit].rsplit(" ", 1)[0].rstrip(",.;:—-")
    return (cut or text[:limit]) + "…"


def _clean_field(value: Any, limit: int) -> str:
    return clean_text(str(value or "").strip())[:limit]


def _trim_paragraph(text: str, limit: int) -> str:
    """Как _trim_bullet, но не режет по первому предложению — "body" это
    осознанно связный абзац из нескольких предложений, а не один буллет."""
    text = clean_text(text.strip(), single_idea=False)
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
            title = _clean_field(card.get("title"), 60)
            body = _clean_field(card.get("body") or card.get("text"), 220)
            if title or body:
                cards.append({"title": title or body, "body": body if title else ""})
    if cards:
        return cards[:3]

    bullets = item.get("bullets") or []
    if not isinstance(bullets, list):
        return []
    for bullet in bullets[:3]:
        text = clean_text(str(bullet).strip())
        if " — " in text:
            title, body = text.split(" — ", 1)
        elif ": " in text:
            title, body = text.split(": ", 1)
        else:
            title, body = text, ""
        cards.append({"title": title.strip()[:60], "body": body.strip()[:220]})
    return cards


def _as_chart(item: dict) -> ChartData | None:
    """Валидирует поле chart; возвращает None, если данных недостаточно
    для осмысленной диаграммы (тогда слайд деградирует до layout "text")."""
    raw = item.get("chart")
    if not isinstance(raw, dict):
        return None

    chart_type = raw.get("type")
    if chart_type not in CHART_TYPES:
        chart_type = "bar"

    categories_raw = raw.get("categories")
    if not isinstance(categories_raw, list):
        return None
    categories = [_clean_field(c, 24) for c in categories_raw if str(c).strip()][:6]
    if len(categories) < 2:
        return None

    series_raw = raw.get("series")
    if not isinstance(series_raw, list):
        return None

    series: list[ChartSeries] = []
    for s in series_raw[:3]:
        if not isinstance(s, dict):
            continue
        values_raw = s.get("values")
        if not isinstance(values_raw, list) or len(values_raw) != len(categories):
            continue
        try:
            values = [float(v) for v in values_raw]
        except (TypeError, ValueError):
            continue
        name = _clean_field(s.get("name"), 30) or "Значение"
        series.append({"name": name, "values": values})

    if not series:
        return None

    return {"type": chart_type, "categories": categories, "series": series}


def _as_timeline(item: dict) -> list[TimelineItem] | None:
    raw = item.get("timeline")
    if not isinstance(raw, list):
        return None

    steps: list[TimelineItem] = []
    for entry in raw[:5]:
        if not isinstance(entry, dict):
            continue
        label = _clean_field(entry.get("label"), 24)
        title = _clean_field(entry.get("title"), 60)
        if not (label and title):
            continue
        steps.append({"label": label, "title": title, "body": _clean_field(entry.get("body"), 140)})

    return steps if len(steps) >= 3 else None


def _as_team(item: dict) -> list[TeamMember] | None:
    raw = item.get("team")
    if not isinstance(raw, list):
        return None

    members: list[TeamMember] = []
    for entry in raw[:6]:
        if not isinstance(entry, dict):
            continue
        name = _clean_field(entry.get("name"), 40)
        if not name:
            continue
        members.append(
            {
                "name": name,
                "role": _clean_field(entry.get("role"), 40),
                "note": _clean_field(entry.get("note"), 100),
            }
        )

    return members if len(members) >= 2 else None


def _as_comparison_side(raw: Any) -> ComparisonSide | None:
    if not isinstance(raw, dict):
        return None
    title = _clean_field(raw.get("title"), 40)
    if not title:
        return None
    bullets_raw = raw.get("bullets") or []
    bullets = [_trim_bullet(str(b), 140) for b in bullets_raw if str(b).strip()][:4] if isinstance(bullets_raw, list) else []
    return {"title": title, "bullets": bullets}


def _as_comparison(item: dict) -> ComparisonData | None:
    raw = item.get("comparison")
    if not isinstance(raw, dict):
        return None
    left = _as_comparison_side(raw.get("left"))
    right = _as_comparison_side(raw.get("right"))
    if left is None or right is None:
        return None
    return {"left": left, "right": right}


def normalize_presentation(data: Any, *, visual_format: str = DEFAULT_VISUAL_FORMAT) -> PresentationData:
    """Приводит сырой JSON модели к ожидаемой форме."""
    if not isinstance(data, dict):
        raise ValueError("Ответ ИИ должен быть JSON-объектом.")

    title = clean_text(str(data.get("title") or "presentation").strip()) or "presentation"
    raw_slides = data.get("slides")
    if not isinstance(raw_slides, list) or not raw_slides:
        raise ValueError("В ответе ИИ нет ни одного слайда.")

    max_bullets, bullet_limit = _FORMAT_BULLET_LIMITS.get(
        visual_format, _FORMAT_BULLET_LIMITS[DEFAULT_VISUAL_FORMAT]
    )
    body_limit = _FORMAT_BODY_LIMITS.get(visual_format, _FORMAT_BODY_LIMITS[DEFAULT_VISUAL_FORMAT])

    slides: list[SlideData] = []
    for item in raw_slides:
        if not isinstance(item, dict):
            continue

        slide_type = item.get("type", "content")
        if slide_type not in ALLOWED_TYPES:
            slide_type = "content"

        layout = item.get("layout", "text")
        if slide_type == "section":
            layout = "quote" if layout == "quote" else "divider"
        elif slide_type == "closing":
            layout = "text"
        elif layout not in ALLOWED_LAYOUTS:
            layout = "text"

        bullets = item.get("bullets") or []
        if not isinstance(bullets, list):
            bullets = [str(bullets)]
        bullets = [_trim_bullet(str(b), bullet_limit) for b in bullets if str(b).strip()][:max_bullets]

        slide: SlideData = {
            "type": slide_type,  # type: ignore[typeddict-item]
            "layout": layout,  # type: ignore[typeddict-item]
            "title": _clean_field(item.get("title") or "Без названия", 90),
            "subtitle": _clean_field(item.get("subtitle"), 180),
            "bullets": bullets,
            "metric": _clean_field(item.get("metric"), 24),
            "metric_label": _clean_field(item.get("metric_label"), 80),
            "cards": _as_cards(item),
        }
        query = str(item.get("image_query") or "").strip()
        if query:
            slide["image_query"] = query

        body_raw = item.get("body")
        if body_raw and str(body_raw).strip():
            slide["body"] = _trim_paragraph(str(body_raw), body_limit)

        if layout == "chart":
            chart = _as_chart(item)
            if chart is None:
                slide["layout"] = "text"  # type: ignore[typeddict-item]
            else:
                slide["chart"] = chart
        elif layout == "timeline":
            timeline = _as_timeline(item)
            if timeline is None:
                slide["layout"] = "text"  # type: ignore[typeddict-item]
            else:
                slide["timeline"] = timeline
        elif layout == "team":
            team = _as_team(item)
            if team is None:
                slide["layout"] = "text"  # type: ignore[typeddict-item]
            else:
                slide["team"] = team
        elif layout == "comparison":
            comparison = _as_comparison(item)
            if comparison is None:
                slide["layout"] = "cards"  # type: ignore[typeddict-item]
            else:
                slide["comparison"] = comparison
        elif layout == "big_photo" and not query:
            slide["layout"] = "text"  # type: ignore[typeddict-item]

        slides.append(slide)

    if not slides:
        raise ValueError("После проверки не осталось валидных слайдов.")

    return {"title": title, "slides": slides}
