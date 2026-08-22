"""Регрессионный тест сохранения целевой суммы квартала (POST /target).

Ошибка, которую тест закрывает: соединение с базой закрывалось до INSERT, и любая попытка
задать цель на дашборде отваливалась 500-й («Cannot operate on a closed database»).
"""
import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main


class SetTargetTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "target-test.db"
        db.init_db()

        main.app.config.update(TESTING=True, SECRET_KEY="target-secret")
        self.client = main.app.test_client()
        self.client.post("/login", data={"username": "owner", "password": "owner123"})

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        try:
            self.temp_dir.cleanup()
        except (PermissionError, NotADirectoryError):
            pass

    def _saved_target(self):
        conn = db.get_conn()
        row = conn.execute("SELECT target_amount FROM quarter_targets ORDER BY id DESC LIMIT 1").fetchone()
        conn.close()
        return row["target_amount"] if row else None

    def test_target_is_saved(self):
        response = self.client.post("/target", data={"target_amount": "605 000 000"})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self._saved_target(), 605000000.0)

    def test_target_is_updated_on_second_submit(self):
        self.client.post("/target", data={"target_amount": "100000"})
        self.client.post("/target", data={"target_amount": "200000"})
        self.assertEqual(self._saved_target(), 200000.0)

    def test_wrong_amount_does_not_crash(self):
        response = self.client.post("/target", data={"target_amount": "не число"})
        self.assertEqual(response.status_code, 302)
        self.assertIsNone(self._saved_target())

    def test_amount_with_thousand_separators(self):
        """Поле форматируется по разрядам — на сервер приходит строка с пробелами.

        Браузер может подставить обычный, неразрывный или узкий пробел — принимаем любой.
        """
        for value in ("151 000 000", "151 000 000", "151 000 000"):
            with self.subTest(value=value):
                self.client.post("/target", data={"target_amount": value})
                self.assertEqual(self._saved_target(), 151000000.0)

    def test_target_field_is_rendered_with_separators(self):
        self.client.post("/target", data={"target_amount": "151 000 000"})
        page = self.client.get("/").get_data(as_text=True)
        self.assertIn('value="151 000 000"', page)
        self.assertNotIn('value="151000000"', page)


if __name__ == "__main__":
    unittest.main()
