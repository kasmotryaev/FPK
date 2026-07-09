#!/bin/bash
# Первая публикация проекта на GitHub. Запускать один раз после Настроить_GitHub.command.

cd "$(dirname "$0")"
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║   Публикация ФП-Контроль на GitHub       ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# ── Читаем сохранённые данные ──────────────────────────────────
GH_USER=$(cat ~/.config/fp-portal/github_user 2>/dev/null)
GH_TOKEN=$(cat ~/.config/fp-portal/github_token 2>/dev/null)

if [ -z "$GH_USER" ] || [ -z "$GH_TOKEN" ]; then
  echo "❌  Настройка не выполнена."
  echo "    Сначала запустите: Настроить_GitHub.command"
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

echo "✅  Аккаунт: @$GH_USER"

# ── Инициализируем git ────────────────────────────────────────
if [ ! -d ".git" ]; then
  git init -b main
  echo "✅  git init выполнен"
fi

# ── Добавляем файлы ───────────────────────────────────────────
git add .

# ── Защита от попадания личных данных ─────────────────────────
STAGED=$(git diff --cached --name-only)
if echo "$STAGED" | grep -qE "token\.txt|fp_portal\.db|/data/uploads/"; then
  echo "🚨  В коммит попали личные данные — прерываем."
  git reset HEAD
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

git commit -m "Initial commit — ФП-Контроль" 2>/dev/null || echo "ℹ️  Файлы уже закоммичены"

# ── Создаём приватный репозиторий через API ───────────────────
echo ""
echo "⏳  Создаём приватный репозиторий FPK на GitHub..."

CREATE_RESP=$(curl -sf -X POST \
  -H "Authorization: token $GH_TOKEN" \
  -H "Content-Type: application/json" \
  https://api.github.com/user/repos \
  -d '{"name":"FPK","private":true,"description":"ФП-Контроль — мониторинг финансового плана"}' \
  2>/dev/null)

REPO_URL=$(echo "$CREATE_RESP" | python3 -c "import sys,json; print(json.load(sys.stdin).get('clone_url',''))" 2>/dev/null)

if [ -z "$REPO_URL" ]; then
  # Репозиторий уже существует?
  REPO_URL="https://github.com/$GH_USER/FPK.git"
  echo "ℹ️  Репозиторий уже существует, используем: $REPO_URL"
else
  echo "✅  Репозиторий создан: https://github.com/$GH_USER/FPK"
fi

# ── Настраиваем remote и пушим ───────────────────────────────
git remote remove origin 2>/dev/null || true
git remote add origin "$REPO_URL"

echo "⏳  Отправляем код на GitHub..."
git push -u origin main

if [ $? -eq 0 ]; then
  echo ""
  echo "╔══════════════════════════════════════════╗"
  echo "║  ✅  Опубликовано!                       ║"
  echo "╚══════════════════════════════════════════╝"
  echo ""
  echo "    https://github.com/$GH_USER/FPK"
  echo ""
  open "https://github.com/$GH_USER/FPK"
else
  echo ""
  echo "❌  Ошибка при отправке."
  echo "    Попробуйте повторно или запустите Настроить_GitHub.command"
fi

echo ""
read -p "Нажмите Enter для выхода..." _
