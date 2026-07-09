#!/bin/bash
# Запуск Telegram-бота ФП-портала (через launchd).
# Если бот уже запущен — перезапускает его (подхватит свежий код).

LABEL="com.fpportal.bot"
DOMAIN="gui/$(id -u)"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

echo "Запуск бота..."

if launchctl print "$DOMAIN/$LABEL" >/dev/null 2>&1; then
    echo "Бот уже загружен в launchd — перезапускаю, чтобы подхватить свежий код."
    launchctl kickstart -k "$DOMAIN/$LABEL"
else
    if [ ! -f "$PLIST" ]; then
        echo "Не найден конфиг: $PLIST"
        echo "Сначала нужно один раз установить автозапуск (см. README, раздел про launchd)."
        read -p "Нажмите Enter, чтобы закрыть окно..."
        exit 1
    fi
    launchctl bootstrap "$DOMAIN" "$PLIST"
fi

sleep 1
echo ""
echo "Статус бота:"
launchctl print "$DOMAIN/$LABEL" 2>/dev/null | grep -E "state =|pid ="

echo ""
echo "Готово. Бот запущен."
read -p "Нажмите Enter, чтобы закрыть окно..."
