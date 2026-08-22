"""Регрессионные тесты колонки «Стратегическое решение» в загрузке ФП."""
import tempfile
import unittest
from pathlib import Path

import openpyxl

import app.db as db
from app.importer import EXPECTED_HEADERS, import_excel


def _write_file(path, strategic_column=True, strategic_value="Решение А"):
    """Создаёт файл ФП с одной строкой — с колонкой стратрешения или без неё (старый формат)."""
    headers = list(EXPECTED_HEADERS)
    values = {
        "Месяц": "Август",
        "Стратегическое решение": strategic_value,
        "ПЦ": "ПЦ Финансовые рынки",
        "Раздел ФП": "Проекты",
        "Наименование клиента": "ТЕСТ-БАНК",
        "Номер проекта": "3855884",
        "Наименование проекта": "Проект 1",
        "Номер договора": "1 к 0404/25",
        "Номер ЭЗ": "3",
        "Сумма по 0-100, руб": 1000.0,
        "Портфель": "0-100",
    }
    if not strategic_column:
        headers.remove("Стратегическое решение")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные"
    ws.append(headers)
    ws.append([values.get(h) for h in headers])
    wb.save(path)


class StrategicSolutionImportTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "strategic-test.db"
        db.init_db()

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def _import(self, **kwargs):
        path = Path(self.temp_dir.name) / "fp.xlsx"
        _write_file(path, **kwargs)
        import_excel(str(path), "fp.xlsx", user_id=1, quarter_label="2026-Q3")
        conn = db.get_conn()
        row = conn.execute("SELECT * FROM fp_rows WHERE is_active = 1").fetchone()
        conn.close()
        return row

    def test_strategic_solution_is_stored(self):
        row = self._import()
        self.assertEqual(row["strategic_solution"], "Решение А")

    def test_old_format_without_column_imports_without_error(self):
        row = self._import(strategic_column=False)
        self.assertIsNone(row["strategic_solution"])
        self.assertEqual(row["client_name"], "ТЕСТ-БАНК")

    def test_change_of_solution_is_tracked_as_event(self):
        self._import()
        path = Path(self.temp_dir.name) / "fp.xlsx"
        _write_file(path, strategic_value="Решение Б")
        import_excel(str(path), "fp.xlsx", user_id=1, quarter_label="2026-Q3")

        conn = db.get_conn()
        event = conn.execute(
            "SELECT * FROM row_events WHERE event_type = 'field_changed' ORDER BY id DESC"
        ).fetchone()
        row = conn.execute("SELECT * FROM fp_rows WHERE is_active = 1").fetchone()
        conn.close()

        self.assertEqual(row["strategic_solution"], "Решение Б")
        self.assertIsNotNone(event)
        self.assertEqual(event["field_label"], "Стратегическое решение")
        self.assertEqual(event["old_value"], "Решение А")
        self.assertEqual(event["new_value"], "Решение Б")


class StrategicSolutionViewTest(unittest.TestCase):
    """Разрез по блокам стратегии на дашборде и фильтр в детализации."""

    def setUp(self):
        import app.main as main

        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "strategic-view.db"
        db.init_db()

        import datetime

        conn = db.get_conn()
        quarter = db.compute_quarter_label(datetime.date.today())
        conn.execute(
            """INSERT INTO fp_rows (row_key, quarter_label, month, section, client_name,
                                    project_name, amount_0_100, portfolio, strategic_solution)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            ("view-row", quarter, _current_month_name(), "Проекты", "ТЕСТ-БАНК",
             "Проект 1", 1000.0, "0-100", "Решение А"),
        )
        conn.commit()
        conn.close()

        main.app.config.update(TESTING=True, SECRET_KEY="strategic-view-secret")
        self.client = main.app.test_client()
        self.client.post("/login", data={"username": "owner", "password": "owner123"})

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        # На Windows файл базы после запросов через test_client может остаться занятым —
        # для теста это неважно, временную папку подчистит система.
        try:
            self.temp_dir.cleanup()
        except (PermissionError, NotADirectoryError):
            pass

    def test_dashboard_shows_strategic_block(self):
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("Разбивка по блокам стратегии", page)
        self.assertIn("Решение А", page)

    def test_rows_list_filters_by_strategic(self):
        # Имя клиента есть и в списке значений фильтра, поэтому проверяем не его,
        # а признак пустой выдачи: по «своему» решению строка находится, по чужому — нет.
        empty_marker = "Нет строк по выбранным фильтрам"

        matched = self.client.get("/rows?strategic=Решение А").get_data(as_text=True)
        self.assertNotIn(empty_marker, matched)

        other = self.client.get("/rows?strategic=Решение Б").get_data(as_text=True)
        self.assertIn(empty_marker, other)


def _current_month_name():
    import datetime

    months = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
              "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    return months[datetime.date.today().month - 1]


if __name__ == "__main__":
    unittest.main()
