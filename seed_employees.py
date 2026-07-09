#!/usr/bin/env python3
"""
Загружает список сотрудников из Excel-файла прямо в БД портала.
Запускать из папки fp-portal:
    python3 seed_employees.py "путь/к/файлу.xlsx"
или без аргументов — ищет файл рядом со скриптом.
"""
import sys, os, sqlite3, glob

BASE = os.path.dirname(os.path.abspath(__file__))
DB   = os.path.join(BASE, "app", "data", "fp_portal.db")

# ─── Найти xlsx ───────────────────────────────────────────────────────────────
if len(sys.argv) > 1:
    xl_path = sys.argv[1]
else:
    # Ищем файл «Список моих сотрудников» рядом
    patterns = [
        os.path.join(BASE, "*сотрудник*.xlsx"),
        os.path.join(BASE, "*сотрудник*.xls"),
        os.path.join(os.path.expanduser("~/Downloads"), "*сотрудник*.xlsx"),
    ]
    xl_path = None
    for p in patterns:
        found = glob.glob(p)
        if found:
            xl_path = found[0]
            break
    if not xl_path:
        print("Файл со списком сотрудников не найден.")
        print("Использование: python3 seed_employees.py «путь к файлу.xlsx»")
        sys.exit(1)

print(f"Файл: {xl_path}")
print(f"БД:   {DB}")

# ─── Читаем имена ─────────────────────────────────────────────────────────────
import openpyxl
wb = openpyxl.load_workbook(xl_path, data_only=True)
ws = wb.active
names = []
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        continue  # заголовок
    v = row[0] if row else None
    if v and str(v).strip():
        names.append(str(v).strip())
wb.close()
print(f"Сотрудников в файле: {len(names)}")

# ─── Вставляем в БД ───────────────────────────────────────────────────────────
conn = sqlite3.connect(DB, timeout=10)
added, skipped = 0, 0
for name in names:
    try:
        conn.execute("INSERT INTO my_employees (full_name) VALUES (?)", (name,))
        added += 1
    except sqlite3.IntegrityError:
        skipped += 1
conn.commit()
total = conn.execute("SELECT COUNT(*) FROM my_employees").fetchone()[0]
conn.close()

print(f"Добавлено: {added}, уже было: {skipped}")
print(f"Итого в справочнике: {total} сотрудников")
