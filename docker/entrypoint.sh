#!/bin/sh
set -e

echo "========================================"
echo "  ФП-Контроль"
echo "========================================"

# Инициализируем БД (создаёт таблицы, если их ещё нет)
echo "Инициализация базы данных..."
python -c "
import sys
sys.path.insert(0, '/app')
from app.db import init_db
init_db()
print('База данных готова.')
"

echo "Запуск сервера на http://0.0.0.0:5001 ..."
exec gunicorn \
  --bind 0.0.0.0:5001 \
  --workers 2 \
  --timeout 120 \
  --access-logfile - \
  --error-logfile - \
  "app.main:app"
