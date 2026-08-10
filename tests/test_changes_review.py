"""Регрессионные тесты отметки изменений ФП как просмотренных."""
import tempfile
import unittest
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

import app.db as db
import app.main as main


class ChangesReviewTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "changes-review-test.db"
        db.init_db()

        main.app.config.update(TESTING=True, SECRET_KEY="changes-review-secret")
        self.client = main.app.test_client()
        response = self.client.post(
            "/login",
            data={"username": "owner", "password": "owner123"},
            follow_redirects=False,
        )
        self.assertEqual(response.status_code, 302)

        conn = db.get_conn()
        row = conn.execute(
            """INSERT INTO fp_rows
               (row_key, section, client_name, project_name, amount_0_100, portfolio)
               VALUES (?, ?, ?, ?, ?, ?)""",
            ("review-row", "Проекты", "Тестовый клиент", "Тестовый проект", 1000, "0-100"),
        )
        event = conn.execute(
            """INSERT INTO row_events
               (fp_row_id, event_type, section, client_name, project_name, amount_after, portfolio)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (row.lastrowid, "new", "Проекты", "Тестовый клиент", "Тестовый проект", 1000, "0-100"),
        )
        self.row_id = row.lastrowid
        self.event_id = event.lastrowid
        conn.commit()
        conn.close()

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def test_ajax_review_returns_timestamp_and_reviewer(self):
        response = self.client.post(
            f"/changes/{self.event_id}/review",
            data={"tab": "new", "section": ["Проекты", "Лицензии"], "sort": "amount"},
            headers={"X-Requested-With": "XMLHttpRequest", "Accept": "application/json"},
        )

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload["ok"])
        self.assertTrue(payload["reviewed_at"])
        self.assertEqual(payload["reviewed_by_name"], "Руководитель продукта")

        conn = db.get_conn()
        saved = conn.execute(
            "SELECT reviewed_at, reviewed_by FROM row_events WHERE id = ?",
            (self.event_id,),
        ).fetchone()
        conn.close()
        self.assertTrue(saved["reviewed_at"])
        self.assertIsNotNone(saved["reviewed_by"])

    def test_fallback_redirect_preserves_filters(self):
        response = self.client.post(
            f"/changes/{self.event_id}/review",
            data={
                "tab": "other",
                "show": "all",
                "import_log_id": "17",
                "section": ["Проекты", "Лицензии"],
                "sort": "amount",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 302)
        query = parse_qs(urlsplit(response.headers["Location"]).query)
        self.assertEqual(query["tab"], ["other"])
        self.assertEqual(query["show"], ["all"])
        self.assertEqual(query["import_log_id"], ["17"])
        self.assertEqual(query["section"], ["Проекты", "Лицензии"])
        self.assertEqual(query["sort"], ["amount"])

    def test_field_change_displays_current_amount(self):
        conn = db.get_conn()
        conn.execute(
            """UPDATE fp_rows SET amount_0_100 = ? WHERE id = ?""",
            (4756, self.row_id),
        )
        conn.execute(
            """INSERT INTO row_events
               (fp_row_id, event_type, field_label, old_value, new_value,
                section, client_name, project_name, portfolio)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (self.row_id, "field_changed", "СДЗ", "2026-08-01", "2026-08-05",
             "Проекты", "Тестовый клиент", "Тестовый проект", "0-100"),
        )
        conn.commit()
        conn.close()

        response = self.client.get("/changes?tab=other")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("<em>сумма 0‑100: <strong>4 756 ₽</strong></em>", html)


if __name__ == "__main__":
    unittest.main(verbosity=2)
