"""Ручная корректировка факта: ввод суммы с разрядами и снятие корректировки.

Поле перевели с input type="number" на текстовое с форматированием по разрядам
(class="js-amount"), поэтому на сервер приходит строка вида «1 200 000» — в том числе
с неразрывным пробелом, если значение вставили из Excel или Confluence.
"""
import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main


class FactCorrectionTest(unittest.TestCase):
    QUARTER = "2026-Q3"

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "fact-correction.db"
        db.init_db()

        conn = db.get_conn()
        conn.execute(
            """INSERT INTO fp_rows (row_key, quarter_label, month, section, client_name,
                                    project_name, amount_0_100, portfolio)
               VALUES (?,?,?,?,?,?,?,?)""",
            ("corr-row", self.QUARTER, "Август", "Проекты", "ТЕСТ-БАНК", "Проект 1", 1000.0, "Факт"),
        )
        conn.commit()
        conn.close()

        main.app.config.update(TESTING=True, SECRET_KEY="fact-correction-secret")
        self.client = main.app.test_client()
        self.client.post("/login", data={"username": "owner", "password": "owner123"})
        self.client.post("/set-quarter-view", data={"mode": "quarter", "quarter": self.QUARTER})

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        try:
            self.temp_dir.cleanup()
        except (PermissionError, NotADirectoryError):
            pass

    def _saved(self):
        conn = db.get_conn()
        value = db.get_setting(conn, f"fact_correction:{self.QUARTER}")
        conn.close()
        return value

    def _post(self, amount, note=""):
        return self.client.post("/save-fact-correction", data={
            "quarter_label": self.QUARTER,
            "correction_amount": amount,
            "correction_note": note,
        })

    def test_amount_with_separators(self):
        for value in ("1 200 000", "1 200 000", "1 200 000"):
            with self.subTest(value=value):
                self._post(value)
                self.assertEqual(float(self._saved()), 1200000.0)

    def test_negative_amount(self):
        self._post("-450 000")
        self.assertEqual(float(self._saved()), -450000.0)

    def test_empty_value_clears_correction(self):
        self._post("1 200 000")
        self._post("")
        self.assertEqual(float(self._saved()), 0.0)

    def test_field_is_rendered_with_separators(self):
        self._post("1 200 000", note="ожидаемое поступление")
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn('value="1 200 000"', page)
        self.assertNotIn('type="number" name="correction_amount"', page)

    def test_wrong_amount_does_not_crash(self):
        response = self._post("не число")
        self.assertEqual(response.status_code, 302)
        self.assertIsNone(self._saved())


if __name__ == "__main__":
    unittest.main()
