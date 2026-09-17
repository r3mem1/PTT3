FROM python:3.12-slim

# Пересобирается только когда меняется requirements.txt — слой с зависимостями кешируется отдельно от кода.
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Бот пишет во время работы только сюда (настройки/история/временные файлы) — не в системные директории.
RUN useradd --create-home --uid 1000 bot \
    && mkdir -p /app/output \
    && chown -R bot:bot /app
USER bot

# Это long-polling бот, а не HTTP-сервер — портов не открываем.
CMD ["python", "bot.py"]
