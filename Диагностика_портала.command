#!/bin/bash
# Диагностика веб-портала ФП-Контроль
# Запустите двойным кликом. Выведет причину проблемы.

PORTAL_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PORTAL_DIR"

echo "=== Диагностика портала ФП-Контроль ==="
echo ""

echo "1) Python:"
/usr/bin/python3 --version 2>&1

echo ""
echo "2) Проверка импортов (Flask, openpyxl):"
/usr/bin/python3 -c "
import sys, importlib
mods = ['flask', 'openpyxl', 'werkzeug']
for m in mods:
    try:
        mod = importlib.import_module(m)
        print(f'  OK  {m} ({getattr(mod, \"__version__\", \"?\")})')
    except ImportError as e:
        print(f'  ERR {m}: {e}')
"

echo ""
echo "3) Загрузка приложения:"
/usr/bin/python3 -c "
import sys
sys.path.insert(0, '.')
try:
    from app.ts_parser import parse_ts_file, parse_employees_file
    print('  OK  ts_parser')
except Exception as e:
    print(f'  ERR ts_parser: {e}')

try:
    from app.db import init_db, DB_PATH
    print(f'  OK  db (БД: {DB_PATH})')
except Exception as e:
    print(f'  ERR db: {e}')

try:
    from app.main import app
    routes = [str(r) for r in app.url_map.iter_rules() if 'timesheets' in str(r)]
    print(f'  OK  main (маршруты /timesheets: {routes})')
except Exception as e:
    import traceback
    print(f'  ERR main:')
    traceback.print_exc()
" 2>&1

echo ""
echo "4) Проверка init_db (новые таблицы):"
/usr/bin/python3 -c "
import sys, sqlite3
sys.path.insert(0, '.')
from app.db import DB_PATH
conn = sqlite3.connect(str(DB_PATH))
tables = [r[0] for r in conn.execute(\"SELECT name FROM sqlite_master WHERE type='table' ORDER BY name\").fetchall()]
conn.close()
ts_tables = [t for t in tables if t.startswith('ts_') or t == 'my_employees']
if ts_tables:
    print(f'  OK  Таблицы трудозатрат найдены: {ts_tables}')
else:
    print(f'  WARN Таблиц трудозатрат нет (будут созданы при запуске)')
print(f'  Все таблицы: {tables}')
" 2>&1

echo ""
echo "5) Статус службы launchd:"
launchctl print "gui/$(id -u)/com.fpportal.web" 2>/dev/null | grep -E "state =|pid =|last exit" || echo "  Служба не зарегистрирована"

echo ""
echo "6) Процессы Flask:"
pgrep -a -f "python3.*run\.py" 2>/dev/null || echo "  Нет запущенных процессов Flask"

echo ""
echo "7) Порт 5000:"
lsof -iTCP:5000 -sTCP:LISTEN 2>/dev/null || echo "  Порт 5000 свободен (Flask не слушает)"

echo ""
echo "=== Конец диагностики ==="
echo ""
read -p "Нажмите Enter, чтобы закрыть окно..."
