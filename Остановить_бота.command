#!/bin/bash
# Остановка Telegram-бота ФП-портала (через launchd).
# Бот не перезапустится сам, пока вы не нажмёте "Запустить_бота.command".

LABEL="com.fpportal.bot"
DOMAIN="gui/$(id -u)"

echo "Останавливаю бота..."

if launchctl print "$DOMAIN/$LABEL" >/dev/null 2>&1; then
    launchctl bootout "$DOMAIN/$LABEL"
    echo "Бот остановлен."
else
    echo "Бот уже не запущен."
fi

read -p "Нажмите Enter, чтобы закрыть окно..."
