#!/bin/bash
# Сохраняет все изменения в GitHub. Запускайте после каждой сессии с Claude.

cd "$(dirname "$0")"
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║   Сохранение изменений на GitHub         ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# ── Проверяем что репозиторий настроен ────────────────────────
if [ ! -d ".git" ]; then
  echo "❌  Сначала запустите: Настроить_GitHub.command"
  echo "    Затем:             Опубликовать_на_GitHub.command"
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

if ! git remote get-url origin &>/dev/null; then
  echo "❌  Запустите: Опубликовать_на_GitHub.command"
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

# ── Проверяем наличие изменений ───────────────────────────────
git add .

CHANGED=$(git diff --cached --name-only)
if [ -z "$CHANGED" ]; then
  echo "ℹ️   Нет изменений для сохранения."
  LAST=$(git log -1 --format="%ci — %s" 2>/dev/null)
  echo "    Последний коммит: $LAST"
  echo ""
  read -p "Нажмите Enter для выхода..." _; exit 0
fi

# ── Защита от личных данных ───────────────────────────────────
if echo "$CHANGED" | grep -qE "token\.txt|fp_portal\.db|/data/uploads/"; then
  echo "🚨  СТОП! В коммит попали личные данные:"
  echo "$CHANGED" | grep -E "token\.txt|fp_portal\.db|/data/uploads/"
  git reset HEAD
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

# ── Показываем что изменилось ──────────────────────────────────
echo "📝  Изменённые файлы:"
echo "$CHANGED" | sed 's/^/    /'
echo ""

# ── Коммит и пуш ──────────────────────────────────────────────
COUNT=$(echo "$CHANGED" | wc -l | tr -d ' ')
git commit -m "Обновление $(date '+%d.%m.%Y %H:%M') · $COUNT файл(ов)"

echo "⏳  Отправляем на GitHub..."
git push

if [ $? -eq 0 ]; then
  echo ""
  echo "╔══════════════════════════════════════════╗"
  echo "║  ✅  Сохранено на GitHub!               ║"
  echo "╚══════════════════════════════════════════╝"
  GH_USER=$(cat ~/.config/fp-portal/github_user 2>/dev/null)
  [ -n "$GH_USER" ] && echo "    https://github.com/$GH_USER/FPK"
else
  echo "❌  Ошибка. Запустите Настроить_GitHub.command если проблема повторяется."
fi

echo ""
read -p "Нажмите Enter для выхода..." _
