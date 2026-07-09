#!/bin/bash
# Перезапуск веб-портала ФП-Контроль (прямой запуск, без launchd).

PORTAL_DIR="$(cd "$(dirname "$0")" && pwd)"
PID_FILE="$PORTAL_DIR/logs/web.pid"
LOG_OUT="$PORTAL_DIR/logs/web.out.log"
LOG_ERR="$PORTAL_DIR/logs/web.err.log"

echo "Перезапуск портала..."

# Убиваем старый процесс по PID-файлу
if [ -f "$PID_FILE" ]; then
    OLD_PID=$(cat "$PID_FILE")
    if kill -0 "$OLD_PID" 2>/dev/null; then
        echo "  Останавливаю старый процесс (PID $OLD_PID)..."
        kill "$OLD_PID" 2>/dev/null
        sleep 1
    fi
    rm -f "$PID_FILE"
fi

# Страхуемся: убиваем любые остатки
EXTRA=$(pgrep -f "python3.*run\.py" 2>/dev/null)
if [ -n "$EXTRA" ]; then
    kill $EXTRA 2>/dev/null
    sleep 1
fi

# Запускаем Flask в фоне, перенаправляем вывод в логи
cd "$PORTAL_DIR"
nohup /usr/bin/python3 run.py >> "$LOG_OUT" 2>> "$LOG_ERR" &
NEW_PID=$!
echo "$NEW_PID" > "$PID_FILE"

echo "  Flask запущен (PID $NEW_PID)"
sleep 4

# Проверяем — процесс ещё жив?
if kill -0 "$NEW_PID" 2>/dev/null; then
    echo ""
    echo "✓ Портал работает."
    echo "  Откройте в браузере: http://127.0.0.1:5001"
    echo "  Трудозатраты:       http://127.0.0.1:5001/timesheets"
else
    echo ""
    echo "✗ Процесс упал. Последние строки лога ошибок:"
    echo "---"
    tail -n 20 "$LOG_ERR"
    echo "---"
fi

read -p "Нажмите Enter, чтобы закрыть окно..."
