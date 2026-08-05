"""Регрессионные тесты загрузки списка сотрудников из Excel."""
import io
import tempfile
import unittest
from pathlib import Path

import openpyxl

import app.db as db
import app.main as main


class EmployeeImportTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        self.original_upload_dir = main.UPLOAD_DIR
        db.DB_PATH = Path(self.temp_dir.name) / "employee-import-test.db"
        main.UPLOAD_DIR = Path(self.temp_dir.name) / "uploads"
        main.UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        db.init_db()

        main.app.config.update(TESTING=True, SECRET_KEY="employee-import-secret")
        self.client = main.app.test_client()
        response = self.client.post(
            "/login",
            data={"username": "owner", "password": "owner123"},
            follow_redirects=False,
        )
        self.assertEqual(response.status_code, 302)

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        main.UPLOAD_DIR = self.original_upload_dir
        self.temp_dir.cleanup()

    @staticmethod
    def _employees_xlsx():
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Сотрудник"])
        sheet.append(["ИВАНОВ ИВАН ИВАНОВИЧ"])
        sheet.append(["ПЕТРОВА АННА СЕРГЕЕВНА"])
        payload = io.BytesIO()
        workbook.save(payload)
        workbook.close()
        payload.seek(0)
        return payload

    def test_cyrillic_xlsx_filename_is_imported(self):
        response = self.client.post(
            "/timesheets/employees",
            data={"emp_file": (self._employees_xlsx(), "Сотрудники.xlsx")},
            content_type="multipart/form-data",
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        conn = db.get_conn()
        names = {
            row["full_name"]
            for row in conn.execute("SELECT full_name FROM my_employees").fetchall()
        }
        conn.close()
        self.assertEqual(names, {
            "ИВАНОВ ИВАН ИВАНОВИЧ",
            "ПЕТРОВА АННА СЕРГЕЕВНА",
        })
        self.assertEqual(list(main.UPLOAD_DIR.iterdir()), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
