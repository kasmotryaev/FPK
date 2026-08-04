FROM mirror.gcr.io/library/python:3.11-slim

WORKDIR /app

# ffmpeg нужен для faster-whisper (голосовые сообщения в чате)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ffmpeg \
    && rm -rf /var/lib/apt/lists/*

# Зависимости Python
# gunicorn   — production-сервер вместо Flask dev-сервера
# pyxlsb     — чтение файлов .xlsb (отчёты трудозатрат)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn pyxlsb

# Код приложения (данные не копируются — они живут в томе)
COPY app/ app/
COPY run.py .

# Точка входа: инициализирует БД при первом старте, затем запускает gunicorn
COPY docker/entrypoint.sh /entrypoint.sh
RUN chmod +x /entrypoint.sh

# Директория данных создаётся заранее; при запуске будет перекрыта томом
RUN mkdir -p app/data/uploads

EXPOSE 5001

# Переменные окружения — переопределите в docker-compose.yml или при запуске
ENV OLLAMA_HOST="http://host.docker.internal:11434"
ENV OLLAMA_MODEL="qwen3:8b"
ENV OLLAMA_TIMEOUT="60"

ENTRYPOINT ["/entrypoint.sh"]
