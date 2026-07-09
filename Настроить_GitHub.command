#!/bin/bash
# Одноразовая настройка — запустите один раз, затем используйте Сохранить_на_GitHub.command

cd "$(dirname "$0")"
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║   Настройка GitHub для ФП-Контроль       ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# ── Проверяем git ──────────────────────────────────────────────
if ! command -v git &>/dev/null; then
  echo "⏳  Устанавливаем git (Xcode Command Line Tools)..."
  xcode-select --install
  echo "    После установки запустите этот скрипт снова."
  read -p "Нажмите Enter для выхода..." _; exit 1
fi
echo "✅  git: $(git --version)"

# ── Логин GitHub ───────────────────────────────────────────────
echo ""
echo "Введите ваш логин на GitHub (то, что в адресе github.com/ВАШ_ЛОГИН):"
read -r GH_USER
if [ -z "$GH_USER" ]; then
  echo "❌  Логин не введён."; read -p "Enter для выхода..." _; exit 1
fi

# ── Создаём токен ──────────────────────────────────────────────
echo ""
echo "Сейчас откроется браузер на странице создания токена GitHub."
echo ""
echo "На странице:"
echo "  1. Убедитесь, что стоит галочка «repo» (остальное не нужно)"
echo "  2. Нажмите «Generate token» внизу"
echo "  3. Скопируйте токен (начинается с ghp_)"
echo ""
read -p "Нажмите Enter — откроем браузер..." _

open "https://github.com/settings/tokens/new?scopes=repo&description=fp-portal-$(date +%Y%m)"

echo ""
echo "Вставьте токен (символы не будут видны — это нормально):"
read -rs GH_TOKEN
echo ""

if [ -z "$GH_TOKEN" ]; then
  echo "❌  Токен не введён."; read -p "Enter для выхода..." _; exit 1
fi

# ── Сохраняем в связку ключей macOS ───────────────────────────
git config --global credential.helper osxkeychain
printf "protocol=https\nhost=github.com\nusername=%s\npassword=%s\n\n" \
  "$GH_USER" "$GH_TOKEN" | git credential-osxkeychain store

# ── Проверяем токен через API ──────────────────────────────────
echo "⏳  Проверяем токен..."
API_RESP=$(curl -sf -H "Authorization: token $GH_TOKEN" https://api.github.com/user 2>/dev/null)
if [ $? -ne 0 ] || echo "$API_RESP" | grep -q '"message"'; then
  echo "❌  Токен не работает. Убедитесь, что скопировали его полностью."
  read -p "Нажмите Enter для выхода..." _; exit 1
fi

GH_NAME=$(echo "$API_RESP" | python3 -c "import sys,json; u=json.load(sys.stdin); print(u.get('name') or u.get('login',''))" 2>/dev/null)
echo "✅  GitHub: авторизован как «$GH_NAME» (@$GH_USER)"

# ── Сохраняем логин для других скриптов ───────────────────────
mkdir -p ~/.config/fp-portal
echo "$GH_USER" > ~/.config/fp-portal/github_user
echo "$GH_TOKEN" > ~/.config/fp-portal/github_token
chmod 600 ~/.config/fp-portal/github_token

# ── Настраиваем автора коммитов ───────────────────────────────
GH_EMAIL=$(curl -sf -H "Authorization: token $GH_TOKEN" https://api.github.com/user/emails \
  | python3 -c "import sys,json; emails=json.load(sys.stdin); print(next((e['email'] for e in emails if e.get('primary')), ''))" 2>/dev/null)

[ -n "$GH_NAME" ]  && git config --global user.name  "$GH_NAME"
[ -n "$GH_EMAIL" ] && git config --global user.email "$GH_EMAIL"
echo "✅  Автор коммитов: $(git config --global user.name) <$(git config --global user.email)>"

echo ""
echo "╔══════════════════════════════════════════════════════════╗"
echo "║  ✅  Настройка завершена!                               ║"
echo "║                                                          ║"
echo "║  Следующий шаг:                                         ║"
echo "║     Опубликовать_на_GitHub.command                       ║"
echo "╚══════════════════════════════════════════════════════════╝"
echo ""
read -p "Нажмите Enter для выхода..." _
