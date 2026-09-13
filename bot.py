"""
Telegram-бот: пользователь присылает тему (или .docx документ) — бот
собирает структуру презентации через ИИ и отправляет .pptx файл в ответ.

Запуск: python bot.py
"""

from __future__ import annotations

import asyncio
import logging
import uuid
from datetime import datetime
from pathlib import Path

from aiogram import Bot, Dispatcher, F, types
from aiogram.exceptions import TelegramBadRequest
from aiogram.filters import Command, CommandStart, StateFilter
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardMarkup,
)

import db
from ai_service import (
    AIServiceError,
    extract_json,
    generate_presentation_text,
    generate_presentation_text_from_document,
)
from config import settings
from docx_service import DocxServiceError, extract_structured_text
from models import PresentationData, normalize_presentation
from pptx_builder import build_presentation
from theme import THEME_PRESET_LABELS

logging.basicConfig(
    level=getattr(logging, settings.log_level.upper(), logging.INFO),
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
log = logging.getLogger("presentation_bot")

bot = Bot(token=settings.telegram_bot_token)
dp = Dispatcher(storage=MemoryStorage())

# Один пользователь — одна генерация за раз (иначе легко сжечь квоту API).
# Общий для генерации по теме, по документу и повтора из истории.
_busy_users: set[int] = set()
_busy_lock = asyncio.Lock()

MAX_DOC_BYTES = 15 * 1024 * 1024
SLIDES_PRESETS = (5, 8, 12)
TONE_LABELS = {
    "business": "Деловой",
    "startup": "Стартап-питч",
    "academic": "Учебный",
    "minimal": "Минимал",
}


class SettingsStates(StatesGroup):
    waiting_custom_slides = State()


class DocStates(StatesGroup):
    waiting_edit_feedback = State()


class LogoStates(StatesGroup):
    waiting_logo = State()


# ---------------------------------------------------------------------------
# Вспомогательные функции: занятость пользователя, безопасное редактирование
# ---------------------------------------------------------------------------


async def _try_acquire(user_id: int) -> bool:
    async with _busy_lock:
        if user_id in _busy_users:
            return False
        _busy_users.add(user_id)
        return True


def _release(user_id: int) -> None:
    _busy_users.discard(user_id)


async def _safe_edit(message: types.Message, text: str, **kwargs) -> None:
    """edit_text, который не падает, если текст/клавиатура не изменились."""
    try:
        await message.edit_text(text, **kwargs)
    except TelegramBadRequest as e:
        if "message is not modified" not in str(e):
            raise


# ---------------------------------------------------------------------------
# Клавиатуры
# ---------------------------------------------------------------------------


def main_reply_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="☰ Меню")]], resize_keyboard=True)


def main_menu_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="🆕 Новая презентация", callback_data="menu:new")],
            [InlineKeyboardButton(text="📄 Презентация из Word-документа", callback_data="menu:from_doc")],
            [InlineKeyboardButton(text="⚙️ Настройки", callback_data="menu:settings")],
            [InlineKeyboardButton(text="🕘 История", callback_data="menu:history")],
            [InlineKeyboardButton(text="🖼 Логотип", callback_data="menu:logo")],
            [InlineKeyboardButton(text="❓ Помощь", callback_data="menu:help")],
        ]
    )


def settings_menu_kb(user_settings: dict) -> InlineKeyboardMarkup:
    def mark(cond: bool) -> str:
        return "✅ " if cond else ""

    slide_row = [
        InlineKeyboardButton(
            text=f"{mark(user_settings['slides_count'] == n)}{n}", callback_data=f"settings:slides:{n}"
        )
        for n in SLIDES_PRESETS
    ]
    slide_row.append(InlineKeyboardButton(text="✏️ своё", callback_data="settings:slides:custom"))

    tone_buttons = [
        InlineKeyboardButton(text=f"{mark(user_settings['tone'] == key)}{label}", callback_data=f"settings:tone:{key}")
        for key, label in TONE_LABELS.items()
    ]

    lang_row = [
        InlineKeyboardButton(text=f"{mark(user_settings['lang'] == 'ru')}Русский", callback_data="settings:lang:ru"),
        InlineKeyboardButton(text=f"{mark(user_settings['lang'] == 'en')}English", callback_data="settings:lang:en"),
    ]

    theme_row = [
        InlineKeyboardButton(
            text=f"{mark(user_settings['theme_preset'] == key)}{label}", callback_data=f"settings:theme:{key}"
        )
        for key, label in THEME_PRESET_LABELS.items()
    ]

    return InlineKeyboardMarkup(
        inline_keyboard=[
            slide_row,
            tone_buttons[:2],
            tone_buttons[2:],
            lang_row,
            theme_row,
            [InlineKeyboardButton(text="⬅️ В меню", callback_data="menu:root")],
        ]
    )


def draft_review_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Собрать как есть", callback_data="doc:confirm")],
            [InlineKeyboardButton(text="✏️ Доработать", callback_data="doc:edit")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data="doc:cancel")],
        ]
    )


def _draft_preview_text(draft: PresentationData) -> str:
    lines = [f"📋 Черновик презентации «{draft['title']}»:", ""]
    for i, slide in enumerate(draft["slides"], start=1):
        title = slide.get("title") or "(без названия)"
        lines.append(f"{i}. {title}")
    lines.append("")
    lines.append("Собрать как есть или доработать?")
    return "\n".join(lines)


def _help_text() -> str:
    return (
        "Отправьте тему текстом — например «Как ИИ меняет маркетинг» — или пришлите .docx документ, "
        "и я сделаю по нему краткую выжимку в виде слайдов.\n\n"
        "/menu — главное меню\n"
        "/settings — число слайдов, стиль, язык, тема оформления\n"
        "/logo_remove — убрать загруженный логотип\n"
        "/cancel — отменить текущее действие"
    )


# ---------------------------------------------------------------------------
# Генерация презентации по теме — общий пайплайн (используется и для
# обычной генерации, и для повтора из истории).
# ---------------------------------------------------------------------------


async def _generate_and_send(message: types.Message, user_id: int, topic: str, status: types.Message) -> None:
    file_path: str | None = None
    try:
        user_settings = await db.get_user_settings(user_id)

        await status.edit_text("⏳ Анализирую тему и собираю структуру...")
        raw_text = await generate_presentation_text(
            topic, user_settings["slides_count"], tone=user_settings["tone"], lang=user_settings["lang"]
        )
        data = normalize_presentation(extract_json(raw_text))

        await status.edit_text("🖼 Подбираю изображения и собираю слайды...")
        logo_path = Path(user_settings["logo_path"]) if user_settings.get("logo_path") else None
        file_path = await build_presentation(
            data, user_id, theme_preset=user_settings["theme_preset"], logo_path=logo_path
        )

        await message.answer_document(types.FSInputFile(file_path), caption=data["title"][:200])
        await db.add_generation(user_id, topic)
        log.info("Презентация готова для user_id=%s: %s", user_id, file_path)

    except AIServiceError as e:
        log.warning("AIServiceError для пользователя %s: %s", user_id, e)
        await message.answer(
            "⚠️ Не получилось сгенерировать презентацию. "
            "Попробуйте другую формулировку темы или повторите попытку."
        )
    except ValueError as e:
        log.warning("Невалидные данные презентации для %s: %s", user_id, e)
        await message.answer("⚠️ ИИ вернул странную структуру. Попробуйте ещё раз.")
    except Exception:
        log.exception("Неожиданная ошибка при генерации презентации user_id=%s", user_id)
        await message.answer("⚠️ Что-то пошло не так. Попробуйте ещё раз чуть позже.")
    finally:
        try:
            await status.delete()
        except Exception:
            log.debug("Не удалось удалить статусное сообщение")
        if file_path:
            Path(file_path).unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Базовые команды и меню
# ---------------------------------------------------------------------------


@dp.message(CommandStart())
async def start(message: types.Message) -> None:
    await message.answer(
        "Привет! Я собираю презентации через ИИ.\n\n"
        "• Пришлите тему текстом — например: «Как ИИ меняет маркетинг»\n"
        "• Или пришлите .docx документ — сделаю выжимку в виде слайдов\n"
        "• В «⚙️ Настройках» можно поменять число слайдов, стиль, язык и тему оформления",
        reply_markup=main_reply_kb(),
    )
    await message.answer("Главное меню:", reply_markup=main_menu_kb())


@dp.message(Command("help"))
async def help_cmd(message: types.Message) -> None:
    await message.answer(_help_text())


@dp.message(Command("menu"))
async def menu_cmd(message: types.Message) -> None:
    await message.answer("Главное меню:", reply_markup=main_menu_kb())


@dp.message(F.text == "☰ Меню")
async def menu_button(message: types.Message) -> None:
    await message.answer("Главное меню:", reply_markup=main_menu_kb())


@dp.message(Command("settings"))
async def settings_cmd(message: types.Message) -> None:
    user_settings = await db.get_user_settings(message.from_user.id)
    await message.answer("⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))


@dp.message(Command("logo_remove"))
async def cmd_logo_remove(message: types.Message) -> None:
    await db.set_user_setting(message.from_user.id, logo_path=None)
    await message.answer("Логотип убран.")


@dp.message(Command("cancel"))
async def cmd_cancel(message: types.Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("Отменено.")


# ---------------------------------------------------------------------------
# Callback-и главного меню
# ---------------------------------------------------------------------------


@dp.callback_query(F.data == "menu:root")
async def cb_menu_root(callback: types.CallbackQuery) -> None:
    await _safe_edit(callback.message, "Главное меню:", reply_markup=main_menu_kb())
    await callback.answer()


@dp.callback_query(F.data == "menu:new")
async def cb_menu_new(callback: types.CallbackQuery) -> None:
    await _safe_edit(callback.message, "Отправьте тему презентации одним сообщением.")
    await callback.answer()


@dp.callback_query(F.data == "menu:help")
async def cb_menu_help(callback: types.CallbackQuery) -> None:
    await _safe_edit(callback.message, _help_text())
    await callback.answer()


@dp.callback_query(F.data == "menu:from_doc")
async def cb_menu_from_doc(callback: types.CallbackQuery) -> None:
    await _safe_edit(
        callback.message,
        "Пришлите .docx файл (Word) — соберу из него краткую выжимку в виде презентации. "
        "Максимальный размер файла — 15 МБ.",
    )
    await callback.answer()


@dp.callback_query(F.data == "menu:settings")
async def cb_menu_settings(callback: types.CallbackQuery) -> None:
    user_settings = await db.get_user_settings(callback.from_user.id)
    await _safe_edit(callback.message, "⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))
    await callback.answer()


@dp.callback_query(F.data == "menu:logo")
async def cb_menu_logo(callback: types.CallbackQuery, state: FSMContext) -> None:
    await state.set_state(LogoStates.waiting_logo)
    await _safe_edit(
        callback.message,
        "Пришлите изображение логотипа (фото или файл-картинка) — буду накладывать его в угол каждого слайда.\n"
        "Чтобы убрать логотип позже — команда /logo_remove.",
    )
    await callback.answer()


@dp.callback_query(F.data == "menu:history")
async def cb_menu_history(callback: types.CallbackQuery) -> None:
    history = await db.get_history(callback.from_user.id)
    if not history:
        await _safe_edit(callback.message, "История пуста — вы ещё не создавали презентаций.", reply_markup=main_menu_kb())
        await callback.answer()
        return

    rows = []
    for i, item in enumerate(history):
        dt = datetime.fromtimestamp(item["created_at"]).strftime("%d.%m %H:%M")
        label = f"{dt} — {item['topic'][:40]}"
        rows.append([InlineKeyboardButton(text=label, callback_data=f"history:repeat:{i}")])
    rows.append([InlineKeyboardButton(text="⬅️ В меню", callback_data="menu:root")])

    await _safe_edit(
        callback.message, "Последние темы — нажмите, чтобы повторить генерацию:", reply_markup=InlineKeyboardMarkup(inline_keyboard=rows)
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("history:repeat:"))
async def cb_history_repeat(callback: types.CallbackQuery) -> None:
    idx = int(callback.data.split(":")[-1])
    history = await db.get_history(callback.from_user.id)
    if idx >= len(history):
        await callback.answer("Эта запись больше недоступна.", show_alert=True)
        return

    topic = history[idx]["topic"]
    user_id = callback.from_user.id
    if not await _try_acquire(user_id):
        await callback.answer("Уже собираю презентацию. Дождитесь файла.", show_alert=True)
        return

    await callback.answer("Запускаю генерацию…")
    status = await callback.message.answer(f"⏳ Повторяю тему «{topic[:60]}»...")
    try:
        await _generate_and_send(callback.message, user_id, topic, status)
    finally:
        _release(user_id)


# ---------------------------------------------------------------------------
# Настройки: число слайдов / тон / язык / тема оформления
# ---------------------------------------------------------------------------


@dp.callback_query(F.data.startswith("settings:slides:"))
async def cb_settings_slides(callback: types.CallbackQuery, state: FSMContext) -> None:
    value = callback.data.split(":")[-1]
    if value == "custom":
        await state.set_state(SettingsStates.waiting_custom_slides)
        await _safe_edit(callback.message, "Пришлите число слайдов текстом (от 3 до 20).")
        await callback.answer()
        return

    await db.set_user_setting(callback.from_user.id, slides_count=int(value))
    user_settings = await db.get_user_settings(callback.from_user.id)
    await _safe_edit(callback.message, "⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))
    await callback.answer(f"Слайдов: {value}")


@dp.message(SettingsStates.waiting_custom_slides, F.text)
async def on_custom_slides(message: types.Message, state: FSMContext) -> None:
    raw = message.text.strip()
    if not raw.isdigit() or not (3 <= int(raw) <= 20):
        await message.answer("Нужно целое число от 3 до 20. Попробуйте снова, или /cancel для отмены.")
        return
    await db.set_user_setting(message.from_user.id, slides_count=int(raw))
    await state.clear()
    user_settings = await db.get_user_settings(message.from_user.id)
    await message.answer("Готово ✅")
    await message.answer("⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))


@dp.callback_query(F.data.startswith("settings:tone:"))
async def cb_settings_tone(callback: types.CallbackQuery) -> None:
    tone = callback.data.split(":")[-1]
    await db.set_user_setting(callback.from_user.id, tone=tone)
    user_settings = await db.get_user_settings(callback.from_user.id)
    await _safe_edit(callback.message, "⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))
    await callback.answer("Стиль обновлён")


@dp.callback_query(F.data.startswith("settings:lang:"))
async def cb_settings_lang(callback: types.CallbackQuery) -> None:
    lang = callback.data.split(":")[-1]
    await db.set_user_setting(callback.from_user.id, lang=lang)
    user_settings = await db.get_user_settings(callback.from_user.id)
    await _safe_edit(callback.message, "⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))
    await callback.answer("Язык обновлён")


@dp.callback_query(F.data.startswith("settings:theme:"))
async def cb_settings_theme(callback: types.CallbackQuery) -> None:
    theme_preset = callback.data.split(":")[-1]
    await db.set_user_setting(callback.from_user.id, theme_preset=theme_preset)
    user_settings = await db.get_user_settings(callback.from_user.id)
    await _safe_edit(callback.message, "⚙️ Настройки генерации:", reply_markup=settings_menu_kb(user_settings))
    await callback.answer("Тема оформления обновлена")


# ---------------------------------------------------------------------------
# Логотип
# ---------------------------------------------------------------------------


@dp.message(LogoStates.waiting_logo, F.photo | F.document)
async def on_logo_upload(message: types.Message, state: FSMContext) -> None:
    if message.photo:
        file_id = message.photo[-1].file_id
    elif message.document and (message.document.mime_type or "").startswith("image/"):
        file_id = message.document.file_id
    else:
        await message.answer("Нужно изображение (фото или файл-картинка PNG/JPG), или /cancel для отмены.")
        return

    settings.logo_dir.mkdir(parents=True, exist_ok=True)
    dest = settings.logo_dir / f"{message.from_user.id}.png"
    tg_file = await bot.get_file(file_id)
    await bot.download_file(tg_file.file_path, destination=str(dest))

    await db.set_user_setting(message.from_user.id, logo_path=str(dest))
    await state.clear()
    await message.answer("Логотип сохранён ✅ Буду накладывать его на каждый слайд. Убрать — /logo_remove.")


@dp.message(LogoStates.waiting_logo)
async def on_logo_upload_invalid(message: types.Message) -> None:
    await message.answer("Пришлите изображение (фото или файл-картинку), или /cancel для отмены.")


# ---------------------------------------------------------------------------
# Импорт Word-документа → черновик → (доработка) → сборка
# ---------------------------------------------------------------------------


@dp.message(F.document, StateFilter(None))
async def handle_document(message: types.Message, state: FSMContext) -> None:
    if not message.from_user or not message.document:
        return

    document = message.document
    filename = document.file_name or ""
    if not filename.lower().endswith(".docx"):
        await message.answer(
            "Пока поддерживаются только файлы .docx (Word). Пришлите тему текстом или .docx документ."
        )
        return

    if document.file_size and document.file_size > MAX_DOC_BYTES:
        await message.answer("Файл слишком большой (максимум 15 МБ).")
        return

    user_id = message.from_user.id
    if not await _try_acquire(user_id):
        await message.answer("Уже собираю презентацию. Дождитесь файла, затем пришлите новый документ.")
        return

    status = await message.answer("📄 Читаю документ...")
    tmp_path: Path | None = None
    try:
        settings.output_dir.mkdir(parents=True, exist_ok=True)
        tmp_path = settings.output_dir / f"doc_{user_id}_{uuid.uuid4().hex[:8]}.docx"
        tg_file = await bot.get_file(document.file_id)
        await bot.download_file(tg_file.file_path, destination=str(tmp_path))

        doc_text, truncated = extract_structured_text(tmp_path)
        user_settings = await db.get_user_settings(user_id)

        await status.edit_text("🧠 Делаю выжимку и продумываю структуру слайдов...")
        raw_text = await generate_presentation_text_from_document(
            doc_text, user_settings["slides_count"], tone=user_settings["tone"], lang=user_settings["lang"]
        )
        draft = normalize_presentation(extract_json(raw_text))

        await state.update_data(
            doc_text=doc_text,
            slides_count=user_settings["slides_count"],
            tone=user_settings["tone"],
            lang=user_settings["lang"],
            theme_preset=user_settings["theme_preset"],
            logo_path=user_settings.get("logo_path"),
            draft=draft,
        )

        await status.delete()
        note = "\n\n⚠️ Документ длинный, обработана только первая часть." if truncated else ""
        await message.answer(_draft_preview_text(draft) + note, reply_markup=draft_review_kb())

    except DocxServiceError as e:
        await status.edit_text(f"⚠️ {e}")
    except AIServiceError:
        await status.edit_text("⚠️ Не получилось сделать выжимку. Попробуйте другой документ или повторите попытку.")
    except ValueError:
        await status.edit_text("⚠️ ИИ вернул странную структуру. Попробуйте ещё раз.")
    except Exception:
        log.exception("Ошибка обработки документа user_id=%s", user_id)
        await status.edit_text("⚠️ Что-то пошло не так. Попробуйте ещё раз чуть позже.")
    finally:
        _release(user_id)
        if tmp_path is not None:
            tmp_path.unlink(missing_ok=True)


@dp.callback_query(F.data == "doc:confirm")
async def cb_doc_confirm(callback: types.CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    draft = data.get("draft")
    if not draft:
        await callback.answer("Черновик устарел, пришлите документ заново.", show_alert=True)
        return

    user_id = callback.from_user.id
    if not await _try_acquire(user_id):
        await callback.answer("Уже собираю презентацию.", show_alert=True)
        return

    await callback.answer()
    try:
        await callback.message.edit_reply_markup(reply_markup=None)
    except TelegramBadRequest:
        pass

    status = await callback.message.answer("🖼 Подбираю изображения и собираю слайды...")
    file_path: str | None = None
    try:
        logo_path = Path(data["logo_path"]) if data.get("logo_path") else None
        file_path = await build_presentation(
            draft, user_id, theme_preset=data.get("theme_preset"), logo_path=logo_path
        )
        await callback.message.answer_document(types.FSInputFile(file_path), caption=draft["title"][:200])
        await db.add_generation(user_id, draft["title"])
    except Exception:
        log.exception("Ошибка сборки презентации из документа user_id=%s", user_id)
        await callback.message.answer("⚠️ Не получилось собрать файл. Попробуйте ещё раз.")
    finally:
        _release(user_id)
        try:
            await status.delete()
        except Exception:
            pass
        if file_path:
            Path(file_path).unlink(missing_ok=True)
        await state.clear()


@dp.callback_query(F.data == "doc:cancel")
async def cb_doc_cancel(callback: types.CallbackQuery, state: FSMContext) -> None:
    await state.clear()
    await _safe_edit(callback.message, "Отменено. Пришлите тему текстом или новый .docx, когда будете готовы.")
    await callback.answer()


@dp.callback_query(F.data == "doc:edit")
async def cb_doc_edit(callback: types.CallbackQuery, state: FSMContext) -> None:
    data = await state.get_data()
    if not data.get("draft"):
        await callback.answer("Черновик устарел.", show_alert=True)
        return
    await state.set_state(DocStates.waiting_edit_feedback)
    try:
        await callback.message.edit_reply_markup(reply_markup=None)
    except TelegramBadRequest:
        pass
    await callback.message.answer(
        "Напишите, что поправить — например: «убери слайд про историю компании, добавь риски и сроки», "
        "и я пересоберу структуру."
    )
    await callback.answer()


@dp.message(DocStates.waiting_edit_feedback, F.text)
async def on_doc_edit_feedback(message: types.Message, state: FSMContext) -> None:
    data = await state.get_data()
    doc_text = data.get("doc_text")
    if not doc_text:
        await message.answer("Черновик устарел, пришлите документ заново.")
        await state.clear()
        return

    status = await message.answer("🧠 Пересобираю структуру с учётом правок...")
    try:
        raw_text = await generate_presentation_text_from_document(
            doc_text,
            data["slides_count"],
            tone=data["tone"],
            lang=data["lang"],
            feedback=message.text.strip(),
        )
        draft = normalize_presentation(extract_json(raw_text))
        await state.update_data(draft=draft)
        await state.set_state(None)  # черновик снова готов к подтверждению, обычный текст не должен считаться правкой
        await status.delete()
        await message.answer(_draft_preview_text(draft), reply_markup=draft_review_kb())
    except AIServiceError:
        await status.edit_text("⚠️ Не получилось доработать структуру. Попробуйте ещё раз описать правки.")
    except ValueError:
        await status.edit_text("⚠️ ИИ вернул странную структуру. Попробуйте ещё раз.")


# ---------------------------------------------------------------------------
# Генерация презентации по теме (обычный текст)
# ---------------------------------------------------------------------------


@dp.message(F.text, StateFilter(None))
async def handle_topic(message: types.Message) -> None:
    if not message.from_user or not message.text:
        return

    topic = message.text.strip()
    user_id = message.from_user.id

    if topic.startswith("/"):
        await message.answer("Неизвестная команда. Напишите /menu или пришлите тему текстом.")
        return

    if len(topic) < settings.topic_min_len:
        await message.answer("Тема слишком короткая, опишите её подробнее.")
        return

    if len(topic) > settings.topic_max_len:
        await message.answer("Тема слишком длинная, сократите её до нескольких предложений.")
        return

    if not await _try_acquire(user_id):
        await message.answer("Уже собираю презентацию. Дождитесь файла, затем пришлите новую тему.")
        return

    status = await message.answer("⏳ Генерирую презентацию, это займёт 10-30 секунд...")
    try:
        await _generate_and_send(message, user_id, topic, status)
    finally:
        _release(user_id)


@dp.message()
async def handle_other(message: types.Message) -> None:
    await message.answer("Пришлите тему презентации текстом, .docx документ, или откройте /menu.")


async def main() -> None:
    await db.init_db()
    log.info("Бот запускается, модель=%s, слайдов по умолчанию=%s", settings.model, settings.slides_count)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
