"""Визуальный язык презентации.

Раньше был один жёстко заданный набор цветов (THEME). Теперь это набор
именованных пресетов (THEME_PRESETS) — пользователь выбирает пресет в
/settings, а pptx_builder.py получает нужный словарь через get_theme().

THEME оставлен как алиас на пресет "warm" для обратной совместимости
(если что-то импортирует THEME напрямую).
"""

from pptx.dml.color import RGBColor
from pptx.util import Emu, Inches, Pt


def _preset(
    *,
    bg,
    bg_dark,
    panel,
    card,
    accent,
    accent_soft,
    title_color,
    body_color,
    muted_color,
    on_dark,
    on_dark_muted,
    heading_font: str = "Calibri",
    body_font: str = "Calibri",
) -> dict:
    return {
        "bg": bg,
        "bg_dark": bg_dark,
        "panel": panel,
        "card": card,
        "accent": accent,
        "accent_soft": accent_soft,
        "title_color": title_color,
        "body_color": body_color,
        "muted_color": muted_color,
        "on_dark": on_dark,
        "on_dark_muted": on_dark_muted,
        "heading_font": heading_font,
        "body_font": body_font,
    }


# Тёплый «бумажный» фон + графитовый текст + терракотовый акцент (исходный вариант).
THEME_WARM = _preset(
    bg=RGBColor(0xF6, 0xF3, 0xEE),
    bg_dark=RGBColor(0x12, 0x18, 0x22),
    panel=RGBColor(0x17, 0x1F, 0x2B),
    card=RGBColor(0xFF, 0xFC, 0xF8),
    accent=RGBColor(0xC4, 0x5C, 0x26),
    accent_soft=RGBColor(0xE8, 0xD5, 0xC4),
    title_color=RGBColor(0x14, 0x18, 0x20),
    body_color=RGBColor(0x3A, 0x40, 0x4C),
    muted_color=RGBColor(0x7A, 0x82, 0x90),
    on_dark=RGBColor(0xF6, 0xF3, 0xEE),
    on_dark_muted=RGBColor(0xB7, 0xBE, 0xC8),
)

# Холодный корпоративный синий — для официальных/финансовых тем.
THEME_CORPORATE = _preset(
    bg=RGBColor(0xF3, 0xF6, 0xFA),
    bg_dark=RGBColor(0x0B, 0x16, 0x2A),
    panel=RGBColor(0x10, 0x1D, 0x36),
    card=RGBColor(0xFF, 0xFF, 0xFF),
    accent=RGBColor(0x1F, 0x6F, 0xEB),
    accent_soft=RGBColor(0xC9, 0xDC, 0xF7),
    title_color=RGBColor(0x0E, 0x16, 0x28),
    body_color=RGBColor(0x33, 0x3F, 0x52),
    muted_color=RGBColor(0x74, 0x82, 0x96),
    on_dark=RGBColor(0xF3, 0xF6, 0xFA),
    on_dark_muted=RGBColor(0xA9, 0xB8, 0xD1),
)

# Тёмный премиум — золотой акцент на графитовом фоне.
THEME_DARK = _preset(
    bg=RGBColor(0x18, 0x18, 0x1C),
    bg_dark=RGBColor(0x0A, 0x0A, 0x0C),
    panel=RGBColor(0x20, 0x20, 0x25),
    card=RGBColor(0x26, 0x26, 0x2C),
    accent=RGBColor(0xD4, 0xAF, 0x37),
    accent_soft=RGBColor(0x4A, 0x42, 0x2A),
    title_color=RGBColor(0xF2, 0xF0, 0xEB),
    body_color=RGBColor(0xC7, 0xC3, 0xBA),
    muted_color=RGBColor(0x86, 0x83, 0x7C),
    on_dark=RGBColor(0xF2, 0xF0, 0xEB),
    on_dark_muted=RGBColor(0x9A, 0x97, 0x90),
)

THEME_PRESETS = {
    "warm": THEME_WARM,
    "corporate": THEME_CORPORATE,
    "dark": THEME_DARK,
}

THEME_PRESET_LABELS = {
    "warm": "🟧 Тёплый терракотовый",
    "corporate": "🟦 Холодный корпоративный",
    "dark": "⬛ Тёмный премиум",
}

DEFAULT_THEME_PRESET = "warm"

# Алиас для обратной совместимости с кодом, который импортирует THEME напрямую.
THEME = THEME_PRESETS[DEFAULT_THEME_PRESET]


def get_theme(preset: str | None) -> dict:
    return THEME_PRESETS.get(preset or DEFAULT_THEME_PRESET, THEME_PRESETS[DEFAULT_THEME_PRESET])


# Раскладочные константы — не зависят от цветовой темы.
MARGIN = Inches(0.55)
GUTTER = Inches(0.28)
RAIL = Emu(90000)
FOOTER_H = Inches(0.32)
TITLE_SIZE = Pt(32)
BODY_SIZE = Pt(16)
METRIC_SIZE = Pt(66)
LOGO_SIZE = Inches(0.42)
