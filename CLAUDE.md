# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Telegram bot (aiogram 3) that generates a `.pptx` presentation from either a text topic or an
uploaded `.docx` document. Slide structure/copy comes from an LLM via OpenRouter; stock photos come
from Pexels; the deck is assembled with `python-pptx` on top of `templates/template.pptx`.

There is no test suite, linter, or build step in this repo — it's a single flat package of Python
modules run directly with `python bot.py`.

## Setup and running

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Copy `.env.example` to `.env` and fill in `TELEGRAM_BOT_TOKEN`, `OPENROUTER_API_KEY`, and optionally
`PEXELS_API_KEY` (if empty, slides build without photos, with a decorative fallback on the title slide).

Run the bot:

```powershell
python bot.py
```

`output/bot.sqlite3` is created automatically on first run (user settings + topic history).

There are no automated tests; verify changes by running the bot and exercising the relevant Telegram
flow (see Architecture below for which module owns which behavior).

## Architecture

**Data flow:** Telegram input → `ai_service.py` builds a prompt and calls OpenRouter →
`ai_service.extract_json` pulls a JSON object out of the raw model text → `models.normalize_presentation`
validates/clamps it into a `PresentationData`/`SlideData` (`models.py`) → `pptx_builder.build_presentation`
renders it slide-by-slide into a `.pptx`, fetching images per slide via `image_service.py` and applying a
color theme from `theme.py`.

- **`config.py`** — the only place that reads `.env` / calls `load_dotenv`. Every other module imports
  the singleton `settings` from here rather than touching `os.environ` directly.
- **`bot.py`** — all Telegram handlers, FSM states (`SettingsStates`, `DocStates`, `LogoStates`), and
  keyboards live in this one file. Key flows:
  - Plain text message with no active state → treated as a presentation topic (`handle_topic`).
  - `.docx` upload with no active state → `handle_document`: extracts text via `docx_service`, builds a
    draft slide outline via `ai_service.build_prompt_from_document`, and shows an inline
    Собрать/Доработать/Отмена review step *before* rendering the `.pptx`. "Доработать" loops back into
    the same AI call with free-text `feedback` appended, reusing the identical JSON schema as topic-based
    generation — this is why `pptx_builder.py` didn't need any document-specific code path.
  - `/settings` — per-user slide count, tone, output language, and theme preset, persisted via `db.py`.
  - A simple per-user in-memory lock (`_try_acquire`/`_release` in `bot.py`) prevents a user from
    triggering two concurrent generations.
- **`ai_service.py`** — owns both prompt templates (`build_prompt` for topics, `build_prompt_from_document`
  for docx summarization) and OpenRouter calling. Both prompts emit the *same* slide JSON schema
  (`type`/`layout`/`title`/`bullets`/`cards`/`metric`/...) so downstream code is shared. `image_query`
  fields are always requested in English regardless of output language, since Pexels search works better
  that way. All network calls are `aiohttp`-async by design — a sync `requests` call would block the bot's
  event loop for the duration of the LLM response.
- **`models.py`** — `normalize_presentation` is the trust boundary between "whatever JSON the LLM returned"
  and the rest of the app: it clamps string lengths, whitelists `type`/`layout` values, and derives `cards`
  from bullets if the model didn't return structured cards. Never assume AI output is well-formed past
  this function.
- **`pptx_builder.py`** — one `_add_*_slide` function per layout (`text`, `image_left`/`image_right`,
  `cards`, `stat`, `quote`, plus `title`/`closing`), dispatched by `_dispatch_slide` based on
  `slide["type"]`/`slide["layout"]`. Shared visual primitives (textboxes, logo, footer, decorative rail)
  are factored into helpers above the per-layout functions. Layout constants (margins, font sizes) live in
  `theme.py` alongside the color presets, not in `pptx_builder.py` itself.
- **`theme.py`** — `THEME_PRESETS` (`warm`/`corporate`/`dark`) selected per-user via `/settings`;
  `get_theme(preset)` is the accessor `pptx_builder.py` uses. `THEME` is a back-compat alias for `warm`.
- **`image_service.py`** — Pexels search + download, with an in-process LRU byte-cache keyed by normalized
  query text (so repeated `image_query` values across slides/generations don't hit the network twice) and
  a best-photo heuristic that picks the closest to a 16:10 landscape ratio rather than always the first
  result.
- **`docx_service.py`** — `.docx` → plain text with `## `-prefixed heading markers and tables flattened to
  `|`-joined rows, truncated to `MAX_CHARS` to bound prompt/token size. This structure-preserving text is
  what `ai_service.build_prompt_from_document` summarizes from.
- **`db.py`** — SQLite (`aiosqlite`) for two things only: per-user settings (`users` table) and a capped
  topic history (`generations` table, last `HISTORY_LIMIT` per user). Not a general-purpose store.

## Known, intentional gaps

- **No access control / rate limiting** ("Модуль 1" from the original spec is deliberately unimplemented):
  the bot responds to any Telegram user with no limits. Don't assume auth/throttling exists anywhere in
  `bot.py`.
- Raw API errors (OpenRouter/Pexels balance, key issues, etc.) are logged, never shown to the end user —
  preserve this when touching error handling in `ai_service.py` / `image_service.py`.
- The prompt in `ai_service.build_prompt` explicitly instructs the model to treat the user-supplied topic
  as inert content, not instructions — a basic prompt-injection guard. Keep this framing if you edit the
  prompt.
