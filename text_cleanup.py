"""Пост-обработка текста слайдов: убирает воду и протекший markdown до того,
как models.py обрежет строки по длине.

Работает по правилам (без второго вызова ИИ) — быстро, детерминированно,
достаточно для соблюдения принципа "один слайд - одна мысль".
"""

from __future__ import annotations

import re

# Вводные клише, которые ничего не добавляют к мысли — вырезаем целиком,
# независимо от регистра, в начале предложения.
_FILLER_PATTERNS = [
    r"стоит отметить,?\s*что\s*",
    r"важно понимать,?\s*что\s*",
    r"важно отметить,?\s*что\s*",
    r"как известно,?\s*",
    r"как мы знаем,?\s*",
    r"в данном контексте,?\s*",
    r"необходимо подчеркнуть,?\s*что\s*",
    r"следует (отметить|подчеркнуть|сказать),?\s*что\s*",
    r"нельзя не отметить,?\s*что\s*",
    r"it is important to note that\s*",
    r"it should be noted that\s*",
    r"in today'?s world,?\s*",
    r"as we all know,?\s*",
    r"needless to say,?\s*",
]
_FILLER_RE = re.compile("|".join(_FILLER_PATTERNS), re.IGNORECASE)

# Протёкший markdown/маркеры списков, которые модель иногда вставляет
# несмотря на просьбу вернуть чистый JSON без разметки.
_MD_BOLD_RE = re.compile(r"\*\*(.*?)\*\*|__(.*?)__")
_MD_LEADING_RE = re.compile(r"^\s*(?:[-*•]|\d{1,2}[.)])\s+")
_WS_RE = re.compile(r"[ \t]{2,}")
_PUNCT_RE = re.compile(r"([!?.,;:])\1{1,}")

# Грубое разбиение на предложения — достаточно для "оставить только первое".
_SENTENCE_SPLIT_RE = re.compile(r"(?<=[.!?…])\s+")

ONE_IDEA_LENGTH_THRESHOLD = 140


def clean_text(text: str, *, single_idea: bool = True) -> str:
    """Убирает воду/протёкший markdown.

    Если single_idea=True (буллеты, заголовки), длинный текст из 2+
    предложений обрезается до первого — "один слайд/пункт - одна мысль".
    Если single_idea=False (связный абзац в поле "body"), несколько
    предложений — это ожидаемая форма, обрезка по предложениям не нужна.
    """
    if not text:
        return text

    cleaned = text.strip()
    cleaned = _MD_LEADING_RE.sub("", cleaned)
    cleaned = _MD_BOLD_RE.sub(lambda m: m.group(1) or m.group(2) or "", cleaned)
    cleaned = _FILLER_RE.sub("", cleaned)
    cleaned = _PUNCT_RE.sub(r"\1", cleaned)
    cleaned = _WS_RE.sub(" ", cleaned).strip()
    if not cleaned:
        return cleaned

    # Делаем первую букву заглавной, если вырезка клише её "съела".
    cleaned = cleaned[0].upper() + cleaned[1:] if cleaned[0].isalpha() else cleaned

    if single_idea and len(cleaned) > ONE_IDEA_LENGTH_THRESHOLD:
        sentences = _SENTENCE_SPLIT_RE.split(cleaned)
        if len(sentences) > 1 and sentences[0].strip():
            cleaned = sentences[0].strip()

    return cleaned
