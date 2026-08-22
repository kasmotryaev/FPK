# -*- coding: utf-8 -*-
r"""
Выгрузка финплана из Ситуационного центра ПЦ напрямую по REST — вместо ручного экспорта
со страницы Confluence.

Результат — .xlsx в том формате, который читает импортёр ФП-Контроля (app/importer.py),
причём с колонками «Стратегическое решение» и «РП», которых в ручной выгрузке нет.

Как работает:
  1. Страница СЦ — это макрос плагина; данные он тянет через REST, авторизуя запросы
     параметром pageHash. Сам pageHash лежит в свойстве страницы
     /rest/api/content/<pageId>/property/ds_page_hash и читается обычным токеном Confluence.
  2. По каждому портфелю (Факт / 0-100 / Возможности) запрашивается детализация — либо
     по конкретному руководителю продукта, либо по всему ПЦ.

Токен Confluence (personal access token) берётся из окружения — в коде и в репозитории
секретов нет:
    set FP_CONFLUENCE_PAT=<токен>
либо путь к JSON-файлу, в котором где-то лежит ключ CONFLUENCE_PAT (например конфиг MCP):
    set FP_CONFLUENCE_PAT_FILE=C:\path\to\config.json

Запуск:
    py tools\sc_export.py --list-pm                  # узнать id руководителей продукта
    py tools\sc_export.py --pm 12345                 # выгрузка по одному РП
    py tools\sc_export.py --all-pm --out plan.xlsx   # весь ПЦ
"""
import argparse
import datetime
import json
import os
import pathlib
import urllib.request

import openpyxl

HOST = os.environ.get("FP_CONFLUENCE_HOST", "https://conf.diasoft.ru")
API = "/rest/dscore/1.0/api/extsource/api/dsconfluence/v1"

# Значения по умолчанию — ПЦ «Финансовые рынки». Переопределяются ключами командной строки.
DEFAULT_SC_PAGE_ID = 97452636   # страница «Ситуационный Центр ПЦ»
DEFAULT_SOLUTION_ID = 460       # решение (ПЦ) в терминах СЦ
DEFAULT_DEPARTMENT_ID = 1551    # подразделение — нужно для справочника РП
DEFAULT_PC_NAME = "ПЦ Финансовые рынки"

# портфель → (сегмент URL, поле суммы в ответе, значение колонки «Портфель»)
PORTFOLIOS = (
    ("fact", "fact_no_tax", "Факт"),
    ("forecast-0-100", "amount", "0-100"),
    ("capability", "amount", "Возможности"),
)

HEADERS = [
    "Месяц", "Стратегическое решение", "ПЦ", "РП", "Раздел ФП", "Наименование клиента",
    "Номер проекта", "Наименование проекта", "Менеджер проекта", "Номер договора", "Номер ЭЗ",
    "СДЗ", "ДПА", "фДЗ", "Способ учета", "0-100 от МП", "Комментарий МП к СДЗ",
    "Сумма из СRM, руб.", "Сумма по 0-100, руб.", "Примечание", "Портфель", "Колодец", "Риск",
]

# Confluence — внутренний хост, прокси для него не нужен.
_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def read_pat():
    """Токен Confluence: из FP_CONFLUENCE_PAT или из JSON-файла FP_CONFLUENCE_PAT_FILE."""
    pat = os.environ.get("FP_CONFLUENCE_PAT") or os.environ.get("CONFLUENCE_PAT")
    if pat:
        return pat

    path = os.environ.get("FP_CONFLUENCE_PAT_FILE")
    if path:
        found = []

        def walk(node):
            if isinstance(node, dict):
                for key, value in node.items():
                    if key == "CONFLUENCE_PAT" and isinstance(value, str) and value:
                        found.append(value)
                    else:
                        walk(value)
            elif isinstance(node, list):
                for item in node:
                    walk(item)

        walk(json.loads(pathlib.Path(path).read_text(encoding="utf-8")))
        if found:
            return found[0]

    raise SystemExit(
        "Не найден токен Confluence. Задайте FP_CONFLUENCE_PAT=<токен> "
        "или FP_CONFLUENCE_PAT_FILE=<путь к json с ключом CONFLUENCE_PAT>."
    )


def get_json(path, pat):
    req = urllib.request.Request(
        HOST + path,
        headers={"Authorization": f"Bearer {pat}", "Accept": "application/json",
                 "User-Agent": "Mozilla/5.0"},
    )
    with _opener.open(req, timeout=180) as resp:
        return json.loads(resp.read().decode("utf-8", "replace"))


def page_hash(pat, page_id):
    """Ключ, которым плагин СЦ авторизует запросы к данным. Хранится в свойстве страницы."""
    return get_json(f"/rest/api/content/{page_id}/property/ds_page_hash", pat)["value"]


def list_product_managers(pat, department_id):
    """Справочник руководителей продукта подразделения: (id, ФИО)."""
    data = get_json(
        f"/rest/dscore/1.0/api/extsource/api/ems/employees"
        f"?filters=status_id:eq:2%7ctop_department_id:eq:{department_id}"
        f"%7cis_team_product_manager:eq:true",
        pat,
    )
    rows = data.get("data") if isinstance(data, dict) else data
    return [(r.get("id"), r.get("fullName") or r.get("brief")) for r in (rows or [])]


def fetch_items(pat, phash, portfolio, solution_id, pm_id):
    """Детализация портфеля: строки по конкретному РП либо по всему ПЦ."""
    base = f"{API}/fm/analyze/solutions/{solution_id}/{portfolio}"
    path = f"{base}/product-managers/{pm_id}/items" if pm_id else f"{base}/items"
    data = get_json(f"{path}?pageHash={phash}", pat)
    items = []
    for dim in data.get("dimensions", []):
        items.extend(dim.get("items", []))
    return items


def as_date(value):
    if not value:
        return None
    try:
        return datetime.datetime.fromisoformat(str(value)).date()
    except ValueError:
        return value


def plain(value):
    """Часть полей СЦ приходит объектами вида {id, name, color} — в ячейку кладём name."""
    if isinstance(value, dict):
        return value.get("name")
    return value


def to_row(item, amount_field, portfolio_name, pc_name):
    amount = item.get(amount_field)
    if amount is None:
        amount = item.get("amount") or item.get("fact_no_tax") or 0
    crm = item.get("forecast_no_tax")
    if crm is None:
        crm = item.get("fact_no_tax") or amount
    risk = item.get("risk")
    if risk is None:
        risk = "Да" if item.get("has_risk") else "Нет"
    return [plain(v) for v in (
        item.get("month_name"),
        item.get("strategic_solution_name"),
        pc_name,
        item.get("product_manager_name"),
        item.get("section_name"),
        item.get("client_name"),
        item.get("project_id"),
        item.get("project_name"),
        item.get("project_manager_name"),
        item.get("findoc_num"),
        item.get("close_stage_num"),
        as_date(item.get("close_stage_cdz")),
        as_date(item.get("dpa")),
        as_date(item.get("fdz")),
        item.get("organization_name"),
        item.get("manager_probability"),
        item.get("close_stage_comment"),
        float(crm or 0),
        float(amount or 0),
        item.get("comment"),
        portfolio_name,
        item.get("full_detailing"),
        risk,
    )]


def build_workbook(pat, args, pm_id):
    """Собирает книгу и возвращает (workbook, {портфель: (строк, сумма)})."""
    phash = page_hash(pat, args.page_id)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные"
    ws.append(HEADERS)

    totals = {}
    for portfolio, amount_field, portfolio_name in PORTFOLIOS:
        items = fetch_items(pat, phash, portfolio, args.solution_id, pm_id)
        for item in items:
            ws.append(to_row(item, amount_field, portfolio_name, args.pc_name))
        totals[portfolio_name] = (
            len(items),
            sum(float(i.get(amount_field) or 0) for i in items),
        )
    return wb, totals


def add_common_arguments(parser):
    """Ключи, общие для этого скрипта и tools/update_from_sc.py."""
    parser.add_argument("--pm", type=int, default=os.environ.get("FP_SC_PM_ID"),
                        help="EMS-id руководителя продукта (или переменная FP_SC_PM_ID)")
    parser.add_argument("--all-pm", action="store_true", help="весь ПЦ, без фильтра по РП")
    parser.add_argument("--list-pm", action="store_true", help="показать id руководителей продукта и выйти")
    parser.add_argument("--page-id", type=int, default=DEFAULT_SC_PAGE_ID, help="pageId страницы СЦ")
    parser.add_argument("--solution-id", type=int, default=DEFAULT_SOLUTION_ID, help="id решения (ПЦ) в СЦ")
    parser.add_argument("--department-id", type=int, default=DEFAULT_DEPARTMENT_ID,
                        help="id подразделения для справочника РП")
    parser.add_argument("--pc-name", default=DEFAULT_PC_NAME, help="значение колонки «ПЦ» в файле")
    return parser


def resolve_pm_id(args):
    """Проверяет, что понятно, чью детализацию тянуть."""
    if args.all_pm:
        return None
    if not args.pm:
        raise SystemExit(
            "Укажите --pm <id> (или FP_SC_PM_ID), либо --all-pm для всего ПЦ. "
            "Список id: --list-pm"
        )
    return int(args.pm)


def print_totals(totals):
    for name, (count, amount) in totals.items():
        print(f"{name}: строк {count}, сумма {amount:,.2f}".replace(",", " "))


def main():
    parser = add_common_arguments(
        argparse.ArgumentParser(description="Выгрузка ФП из Ситуационного центра")
    )
    parser.add_argument("--out", help="путь к .xlsx (по умолчанию — рядом, с датой в имени)")
    args = parser.parse_args()

    pat = read_pat()

    if args.list_pm:
        for pm_id, name in list_product_managers(pat, args.department_id):
            print(f"{pm_id}\t{name}")
        return

    pm_id = resolve_pm_id(args)
    wb, totals = build_workbook(pat, args, pm_id)

    out = args.out or f"СЦ ФП {datetime.date.today():%Y-%m-%d}.xlsx"
    wb.save(out)

    print_totals(totals)
    print("файл:", pathlib.Path(out).resolve())


if __name__ == "__main__":
    main()
