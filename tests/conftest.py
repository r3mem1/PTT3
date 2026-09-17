"""Даёт тестам работать сразу после git clone, без реального .env.

config.py требует TELEGRAM_BOT_TOKEN и OPENROUTER_API_KEY уже на этапе
импорта модуля (см. load_settings()). Юнит-тесты не обращаются ни к
Telegram, ни к OpenRouter, поэтому подставляем фиктивные значения — но
только если пользователь не задал реальные (например, через свой .env).
"""

import os
import sys
from pathlib import Path

# Модули бота (models.py, pptx_builder.py, ...) лежат плоско в корне репо,
# а не пакетом - добавляем корень в sys.path, чтобы `import models` работал
# независимо от того, как именно запущен pytest (из корня или из tests/).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

os.environ.setdefault("TELEGRAM_BOT_TOKEN", "test-token")
os.environ.setdefault("OPENROUTER_API_KEY", "test-key")
