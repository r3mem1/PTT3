"""
Собирает .pptx: раскладки рисуются шейпами поверх Blank-макета,
чтобы колода выглядела как единый дизайн, а не набор стандартных
плейсхолдеров PowerPoint.
"""

from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Length, Pt

from config import settings
from image_service import find_and_download_image
from models import PresentationData, SlideData
from theme import (
    BODY_SIZE,
    FOOTER_H,
    GUTTER,
    KICKER_SIZE,
    LOGO_SIZE,
    MARGIN,
    METRIC_SIZE,
    RAIL,
    TITLE_SIZE,
    get_theme,
)

log = logging.getLogger(__name__)

LAYOUT_BLANK = 6


@dataclass
class BuildContext:
    prs: Presentation
    deck_title: str
    index: int
    total: int
    used_images: list[Path]
    theme: dict = field(default_factory=lambda: get_theme(None))
    logo_path: Path | None = None


def _fill_solid(shape, color: RGBColor) -> None:
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False


def _slide_bg(slide, color: RGBColor) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _strip_placeholders(slide) -> None:
    for ph in list(slide.placeholders):
        el = ph._element
        el.getparent().remove(el)


def _set_run(paragraph, text: str, *, font: str, size, color: RGBColor, bold=False, italic=False) -> None:
    paragraph.text = text
    for run in paragraph.runs:
        run.font.name = font
        run.font.size = size
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color


def _textbox(
    slide,
    left,
    top,
    width,
    height,
    text: str,
    *,
    font: str,
    size,
    color: RGBColor,
    bold: bool = False,
    italic: bool = False,
    align=None,
    anchor: str = "t",
    shrink_to_fit: bool = False,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    if shrink_to_fit:
        # Длинный буллет/заголовок ужимается по шрифту вместо того, чтобы
        # вылезти за пределы карточки/колонки.
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    _set_run(p, text, font=font, size=size, color=color, bold=bold, italic=italic)
    tf._txBody.bodyPr.set("anchor", anchor)
    return box


def _kicker_text(slide_data: SlideData, fallback: str) -> str:
    return (slide_data.get("kicker") or fallback).upper()


def _add_logo(slide, ctx: BuildContext) -> None:
    """Небольшой логотип в правом нижнем углу, если пользователь его загрузил."""
    if ctx.logo_path is None or not Path(ctx.logo_path).is_file():
        return
    left = ctx.prs.slide_width - MARGIN - LOGO_SIZE
    top = ctx.prs.slide_height - FOOTER_H - LOGO_SIZE - Inches(0.08)
    try:
        with Image.open(str(ctx.logo_path)) as im:
            w, h = im.size
        ratio = w / h
        if ratio >= 1:
            box_w, box_h = LOGO_SIZE, LOGO_SIZE / ratio
        else:
            box_w, box_h = LOGO_SIZE * ratio, LOGO_SIZE
        pic = slide.shapes.add_picture(str(ctx.logo_path), left, top, width=box_w, height=box_h)
        pic.line.fill.background()
    except Exception:
        log.debug("Не удалось наложить логотип %s", ctx.logo_path)


def _add_footer(slide, ctx: BuildContext, *, on_dark: bool = False) -> None:
    theme = ctx.theme
    color = theme["on_dark_muted"] if on_dark else theme["muted_color"]
    top = ctx.prs.slide_height - FOOTER_H
    width = ctx.prs.slide_width - MARGIN * 2
    label = ctx.deck_title[:42]
    _textbox(
        slide, MARGIN, top, width * 0.75, FOOTER_H, label,
        font=theme["body_font"], size=Pt(10), color=color,
    )
    _textbox(
        slide, MARGIN, top, width, FOOTER_H, f"{ctx.index:02d}  /  {ctx.total:02d}",
        font=theme["body_font"], size=Pt(10), color=color, align=PP_ALIGN.RIGHT,
    )
    _add_logo(slide, ctx)


def _add_rail(slide, prs, color: RGBColor) -> None:
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, RAIL, prs.slide_height)
    _fill_solid(bar, color)


def _picture(slide, image_path: Path, left, top, width, height):
    path = str(image_path)
    with Image.open(path) as im:
        img_w, img_h = im.size
    img_ratio = img_w / img_h
    box_ratio = width / height
    pic = slide.shapes.add_picture(path, left, top, width=width, height=height)
    if img_ratio > box_ratio:
        crop = (1 - box_ratio / img_ratio) / 2
        pic.crop_left = crop
        pic.crop_right = crop
    else:
        crop = (1 - img_ratio / box_ratio) / 2
        pic.crop_top = crop
        pic.crop_bottom = crop
    pic.line.fill.background()
    return pic


def _card_shape(slide, left, top, width, height, theme: dict):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    _fill_solid(shape, theme["card"])
    try:
        shape.adjustments[0] = 0.06
    except (AttributeError, IndexError):
        pass
    shape.line.color.rgb = theme["accent_soft"]
    shape.line.width = Pt(1)
    return shape


def _decorative_panel(slide, prs, left, top, width, height, theme: dict) -> None:
    """Заполняет правую панель титульного слайда геометрией, если картинки нет.

    Пустая заливка выглядит недоделанной — несколько полупрозрачно-имитирующих
    кругов акцентным/мягким цветом дают ощущение осознанного дизайна.
    """
    base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    _fill_solid(base, theme["bg_dark"])

    big = int(min(width, height) * 0.9)
    circle1 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left + width - big // 2, top - big // 4, big, big
    )
    _fill_solid(circle1, theme["panel"])

    mid = int(min(width, height) * 0.55)
    circle2 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, left + int(width * 0.1), top + height - mid // 2, mid, mid
    )
    _fill_solid(circle2, theme["accent"])

    small = int(min(width, height) * 0.22)
    circle3 = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        left + int(width * 0.55),
        top + int(height * 0.35),
        small,
        small,
    )
    _fill_solid(circle3, theme["accent_soft"])


async def _maybe_image(slide_data: SlideData, ctx: BuildContext) -> Path | None:
    query = slide_data.get("image_query")
    if not query:
        return None
    path = await find_and_download_image(query)
    if path is not None:
        ctx.used_images.append(Path(path))
    return path


def _new_slide(ctx: BuildContext, *, dark: bool = False):
    slide = ctx.prs.slides.add_slide(ctx.prs.slide_layouts[LAYOUT_BLANK])
    _strip_placeholders(slide)
    _slide_bg(slide, ctx.theme["bg_dark"] if dark else ctx.theme["bg"])
    return slide


def _add_title_block(
    slide, prs, slide_data: SlideData, *, kicker: str, theme: dict, on_dark: bool = False
) -> Length:
    title_color = theme["on_dark"] if on_dark else theme["title_color"]
    kicker_color = theme["accent"] if not on_dark else theme["accent_soft"]
    left = MARGIN + Inches(0.12)
    width = prs.slide_width - MARGIN * 2
    _textbox(
        slide, left, Inches(0.38), width, Inches(0.32), kicker,
        font=theme["body_font"], size=KICKER_SIZE, color=kicker_color, bold=True,
    )
    _textbox(
        slide, left, Inches(0.68), width, Inches(1.15), slide_data.get("title") or "",
        font=theme["heading_font"], size=TITLE_SIZE, color=title_color, bold=True,
        shrink_to_fit=True,
    )
    return Inches(1.95)


async def _add_title_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx, dark=True)
    prs = ctx.prs
    panel_w = int(prs.slide_width * 0.56)
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, panel_w, prs.slide_height)
    _fill_solid(panel, theme["panel"])
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, panel_w, 0, RAIL, prs.slide_height)
    _fill_solid(accent, theme["accent"])

    image_path = await _maybe_image(slide_data, ctx)
    right_w = prs.slide_width - panel_w
    if image_path is not None:
        _picture(slide, image_path, panel_w, 0, right_w, prs.slide_height)
    else:
        # Раньше здесь оставалось пустое цветное поле — теперь декоративная геометрия.
        _decorative_panel(slide, prs, panel_w, 0, right_w, prs.slide_height, theme)

    pad = Inches(0.7)
    _textbox(
        slide, pad, Inches(1.55), panel_w - pad * 1.4, Inches(0.35),
        _kicker_text(slide_data, "Briefing"),
        font=theme["body_font"], size=KICKER_SIZE, color=theme["accent_soft"], bold=True,
    )
    _textbox(
        slide, pad, Inches(2.0), panel_w - pad * 1.4, Inches(2.4),
        slide_data.get("title") or ctx.deck_title,
        font=theme["heading_font"], size=Pt(40), color=theme["on_dark"], bold=True,
        shrink_to_fit=True,
    )
    subtitle = slide_data.get("subtitle") or ""
    if subtitle:
        _textbox(
            slide, pad, Inches(4.6), panel_w - pad * 1.4, Inches(1.4), subtitle,
            font=theme["body_font"], size=Pt(16), color=theme["on_dark_muted"],
        )
    _add_footer(slide, ctx, on_dark=True)


def _add_points(slide, left, top, width, bullets: list[str], theme: dict, *, light: bool = True) -> None:
    row_h = Inches(0.92)
    num_color = theme["accent"]
    text_color = theme["body_color"] if light else theme["on_dark"]
    for i, bullet in enumerate(bullets[:4]):
        y = top + row_h * i
        _textbox(
            slide, left, y, Inches(0.55), Inches(0.4), f"{i + 1:02d}",
            font=theme["heading_font"], size=Pt(16), color=num_color, bold=True,
        )
        _textbox(
            slide, left + Inches(0.62), y, width - Inches(0.62), Inches(0.82), bullet,
            font=theme["body_font"], size=BODY_SIZE, color=text_color, shrink_to_fit=True,
        )


def _add_text_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx)
    _add_rail(slide, ctx.prs, theme["accent"])
    content_top = _add_title_block(
        slide, ctx.prs, slide_data, kicker=_kicker_text(slide_data, "Key points"), theme=theme
    )
    bullets = slide_data.get("bullets") or []
    if bullets:
        _add_points(slide, MARGIN + Inches(0.12), content_top, ctx.prs.slide_width - MARGIN * 2, bullets, theme)
    _add_footer(slide, ctx)


async def _add_image_text_slide(ctx: BuildContext, slide_data: SlideData, image_on_left: bool) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx)
    _add_rail(slide, ctx.prs, theme["accent"])
    content_top = _add_title_block(
        slide, ctx.prs, slide_data, kicker=_kicker_text(slide_data, "Insight"), theme=theme
    )

    gap = GUTTER
    body_top = content_top
    body_h = ctx.prs.slide_height - body_top - FOOTER_H - Inches(0.15)
    col_w = (ctx.prs.slide_width - MARGIN * 2 - gap) * 0.46
    text_w = ctx.prs.slide_width - MARGIN * 2 - gap - col_w
    img_left = MARGIN if image_on_left else MARGIN + text_w + gap
    text_left = MARGIN + col_w + gap if image_on_left else MARGIN

    image_path = await _maybe_image(slide_data, ctx)
    if image_path is not None:
        pic_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, img_left, body_top, col_w, body_h)
        _fill_solid(pic_box, theme["accent_soft"])
        _picture(slide, image_path, img_left, body_top, col_w, body_h)
    else:
        text_left = MARGIN
        text_w = ctx.prs.slide_width - MARGIN * 2

    bullets = slide_data.get("bullets") or []
    if bullets:
        _add_points(slide, text_left, body_top, text_w, bullets, theme)
    _add_footer(slide, ctx)


def _add_cards_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx)
    _add_rail(slide, ctx.prs, theme["accent"])
    content_top = _add_title_block(
        slide, ctx.prs, slide_data, kicker=_kicker_text(slide_data, "Framework"), theme=theme
    )
    cards = slide_data.get("cards") or []
    if not cards:
        bullets = slide_data.get("bullets") or []
        cards = [{"title": b, "body": ""} for b in bullets[:3]]

    n = max(1, min(3, len(cards)))
    available = ctx.prs.slide_width - MARGIN * 2
    card_w = (available - GUTTER * (n - 1)) / n
    height = ctx.prs.slide_height - content_top - FOOTER_H - Inches(0.2)
    for i, card in enumerate(cards[:n]):
        left = MARGIN + i * (card_w + GUTTER)
        _card_shape(slide, left, content_top, card_w, height, theme)
        inner = left + Inches(0.28)
        inner_w = card_w - Inches(0.56)
        _textbox(
            slide, inner, content_top + Inches(0.28), inner_w, Inches(0.36), f"{i + 1:02d}",
            font=theme["heading_font"], size=Pt(14), color=theme["accent"], bold=True,
        )
        _textbox(
            slide, inner, content_top + Inches(0.7), inner_w, Inches(1.3), card.get("title") or "",
            font=theme["heading_font"], size=Pt(20), color=theme["title_color"], bold=True,
            shrink_to_fit=True,
        )
        _textbox(
            slide, inner, content_top + Inches(2.1), inner_w, height - Inches(2.4), card.get("body") or "",
            font=theme["body_font"], size=Pt(14), color=theme["body_color"], shrink_to_fit=True,
        )
    _add_footer(slide, ctx)


def _add_stat_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx)
    _add_rail(slide, ctx.prs, theme["accent"])
    _add_title_block(slide, ctx.prs, slide_data, kicker=_kicker_text(slide_data, "Signal"), theme=theme)
    metric = slide_data.get("metric") or (slide_data.get("bullets") or ["—"])[0]
    label = slide_data.get("metric_label") or slide_data.get("subtitle") or ""
    _textbox(
        slide, MARGIN, Inches(2.15), ctx.prs.slide_width - MARGIN * 2, Inches(1.6), metric,
        font=theme["heading_font"], size=METRIC_SIZE, color=theme["accent"], bold=True,
        shrink_to_fit=True,
    )
    if label:
        _textbox(
            slide, MARGIN, Inches(3.75), ctx.prs.slide_width - MARGIN * 2, Inches(0.7), label,
            font=theme["heading_font"], size=Pt(22), color=theme["title_color"], bold=True,
            shrink_to_fit=True,
        )
    bullets = slide_data.get("bullets") or []
    if bullets:
        _textbox(
            slide, MARGIN, Inches(4.55), ctx.prs.slide_width - MARGIN * 2, Inches(1.6), bullets[0],
            font=theme["body_font"], size=Pt(16), color=theme["body_color"], shrink_to_fit=True,
        )
    _add_footer(slide, ctx)


def _add_quote_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx, dark=True)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, ctx.prs.slide_width, RAIL)
    _fill_solid(bar, theme["accent"])
    quote = slide_data.get("title") or ""
    _textbox(
        slide, Inches(1.1), Inches(2.1), ctx.prs.slide_width - Inches(2.2), Inches(2.8),
        f"“{quote}”" if quote and not quote.startswith("“") else quote,
        font=theme["heading_font"], size=Pt(32), color=theme["on_dark"], bold=True,
        shrink_to_fit=True,
    )
    attribution = slide_data.get("subtitle") or slide_data.get("kicker") or ""
    if attribution:
        _textbox(
            slide, Inches(1.1), Inches(5.1), ctx.prs.slide_width - Inches(2.2), Inches(0.6),
            attribution,
            font=theme["body_font"], size=Pt(16), color=theme["accent_soft"], italic=True,
        )
    _add_footer(slide, ctx, on_dark=True)


def _add_closing_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    theme = ctx.theme
    slide = _new_slide(ctx, dark=True)
    _add_rail(slide, ctx.prs, theme["accent"])
    _textbox(
        slide, MARGIN, Inches(0.55), ctx.prs.slide_width - MARGIN * 2, Inches(0.35),
        _kicker_text(slide_data, "Next move"),
        font=theme["body_font"], size=KICKER_SIZE, color=theme["accent_soft"], bold=True,
    )
    _textbox(
        slide, MARGIN, Inches(1.0), ctx.prs.slide_width - MARGIN * 2, Inches(1.6),
        slide_data.get("title") or "Что делать дальше",
        font=theme["heading_font"], size=Pt(36), color=theme["on_dark"], bold=True,
        shrink_to_fit=True,
    )
    bullets = slide_data.get("bullets") or []
    if bullets:
        _add_points(slide, MARGIN, Inches(2.8), ctx.prs.slide_width - MARGIN * 2, bullets, theme, light=False)
    _add_footer(slide, ctx, on_dark=True)


async def _dispatch_slide(ctx: BuildContext, slide_data: SlideData) -> None:
    slide_type = slide_data.get("type", "content")
    layout = slide_data.get("layout", "text")

    if slide_type == "title":
        await _add_title_slide(ctx, slide_data)
        return
    if slide_type == "closing":
        _add_closing_slide(ctx, slide_data)
        return
    if slide_type == "section" or layout == "quote":
        _add_quote_slide(ctx, slide_data)
        return
    if layout == "cards":
        _add_cards_slide(ctx, slide_data)
        return
    if layout == "stat":
        _add_stat_slide(ctx, slide_data)
        return
    if layout == "image_left":
        await _add_image_text_slide(ctx, slide_data, image_on_left=True)
        return
    if layout == "image_right":
        await _add_image_text_slide(ctx, slide_data, image_on_left=False)
        return
    _add_text_slide(ctx, slide_data)


def _safe_filename(title: str) -> str:
    cleaned = "".join(c for c in title if c.isalnum() or c in " _-").strip()
    return cleaned[:40] or "presentation"


async def build_presentation(
    data: PresentationData,
    user_id: int,
    *,
    theme_preset: str | None = None,
    logo_path: Path | None = None,
) -> str:
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    prs = Presentation(str(settings.template_path))
    used_images: list[Path] = []
    slides = data.get("slides") or []
    if not slides:
        raise ValueError("В ответе ИИ нет ни одного слайда.")

    ctx = BuildContext(
        prs=prs,
        deck_title=data.get("title") or "presentation",
        index=1,
        total=len(slides),
        used_images=used_images,
        theme=get_theme(theme_preset),
        logo_path=logo_path,
    )

    try:
        for i, slide_data in enumerate(slides, start=1):
            ctx.index = i
            await _dispatch_slide(ctx, slide_data)

        file_name = f"{user_id}_{_safe_filename(ctx.deck_title)}_{uuid.uuid4().hex[:6]}.pptx"
        file_path = settings.output_dir / file_name
        prs.save(str(file_path))
        log.info("Сохранён файл %s (%s слайдов)", file_path, len(slides))
        return str(file_path)
    finally:
        for image_path in used_images:
            try:
                image_path.unlink(missing_ok=True)
            except OSError:
                log.debug("Не удалось удалить временную картинку %s", image_path)
