# -*- coding: utf-8 -*-
r"""
Обновление финплана в ФП-Контроле из Ситуационного центра — одной командой, по требованию.

Тянет данные СЦ по REST (см. tools/sc_export.py), кладёт .xlsx в папку выгрузок и сразу
загружает его тем же импортёром, что и загрузка файла через интерфейс. Расписания намеренно
нет: запускается руками, когда нужен свежий срез.

Токен Confluence — в окружении (FP_CONFLUENCE_PAT или FP_CONFLUENCE_PAT_FILE), см. sc_export.py.

Запуск:
    py tools\update_from_sc.py --pm 12345
    py tools\update_from_sc.py --all-pm --quarter 2026-Q3
    py tools\update_from_sc.py --pm 12345 --keep-dir D:\Выгрузки
"""
import argparse
import datetime
import os
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))

import sc_export
from app.db import get_conn, get_setting
from app.importer import import_excel

DEFAULT_KEEP_DIR = os.environ.get(
    "FP_SC_EXPORT_DIR",
    str(pathlib.Path(__file__).resolve().parent.parent / "app" / "data" / "sc_exports"),
)


def owner_user_id():
    conn = get_conn()
    row = conn.execute("SELECT id FROM users WHERE role = 'owner' ORDER BY id LIMIT 1").fetchone()
    conn.close()
    return row["id"] if row else 1


def main():
    parser = sc_export.add_common_arguments(
        argparse.ArgumentParser(description="Обновить ФП в портале из Ситуационного центра")
    )
    parser.add_argument("--quarter", help="метка квартала загрузки, например 2026-Q3")
    parser.add_argument("--keep-dir", default=DEFAULT_KEEP_DIR, help="куда класть выгруженный файл")
    args = parser.parse_args()

    pat = sc_export.read_pat()

    if args.list_pm:
        for pm_id, name in sc_export.list_product_managers(pat, args.department_id):
            print(f"{pm_id}\t{name}")
        return

    pm_id = sc_export.resolve_pm_id(args)

    keep_dir = pathlib.Path(args.keep_dir)
    keep_dir.mkdir(parents=True, exist_ok=True)
    out_path = keep_dir / f"СЦ ФП {datetime.datetime.now():%Y-%m-%d %H-%M}.xlsx"

    wb, totals = sc_export.build_workbook(pat, args, pm_id)
    wb.save(out_path)
    sc_export.print_totals(totals)
    print("файл:", out_path)

    conn = get_conn()
    quarter = args.quarter or get_setting(conn, "current_quarter_label")
    conn.close()

    stats = import_excel(str(out_path), out_path.name, owner_user_id(), quarter_label=quarter)
    print("квартал загрузки:", quarter or "определён импортёром")
    for key in ("rows_total", "rows_new", "rows_updated", "rows_unchanged", "rows_deactivated"):
        print(f"  {key}: {stats.get(key)}")
    if stats.get("header_mismatches"):
        print("  ВНИМАНИЕ, не совпали заголовки:", stats["header_mismatches"])


if __name__ == "__main__":
    main()
