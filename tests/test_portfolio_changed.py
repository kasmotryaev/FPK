"""Regression tests for forecast-to-fact change events."""
import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main


LEGACY_ROW_EVENTS_SQL = """
CREATE TABLE row_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fp_row_id INTEGER NOT NULL REFERENCES fp_rows(id) ON DELETE CASCADE,
    import_log_id INTEGER REFERENCES import_log(id),
    event_type TEXT NOT NULL CHECK(event_type IN
        ('new','zeroed','deactivated','reactivated','amount_changed','field_changed')),
    field_label TEXT,
    old_value TEXT,
    new_value TEXT,
    amount_before REAL,
    amount_after REAL,
    month TEXT,
    client_name TEXT,
    project_name TEXT,
    section TEXT,
    portfolio TEXT,
    contract_num TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    reviewed_at TEXT,
    reviewed_by INTEGER REFERENCES users(id)
)
"""


class PortfolioChangedTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "portfolio-changed-test.db"
        db.init_db()

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def test_legacy_event_table_is_migrated_without_data_loss(self):
        conn = db.get_conn()
        conn.execute("DROP TABLE row_events")
        conn.execute(LEGACY_ROW_EVENTS_SQL)
        row_id = conn.execute(
            "INSERT INTO fp_rows (row_key, section, amount_0_100) VALUES (?, ?, ?)",
            ("legacy-row", "Проекты", 1000),
        ).lastrowid
        conn.execute(
            "INSERT INTO row_events (fp_row_id, event_type, amount_after) VALUES (?, ?, ?)",
            (row_id, "new", 1000),
        )
        conn.commit()
        conn.close()

        db.init_db()

        conn = db.get_conn()
        preserved = conn.execute(
            "SELECT COUNT(*) AS count FROM row_events WHERE event_type = 'new'"
        ).fetchone()["count"]
        conn.execute(
            "INSERT INTO row_events (fp_row_id, event_type, old_value, new_value) VALUES (?, ?, ?, ?)",
            (row_id, "portfolio_changed", "0-100", "Факт"),
        )
        conn.execute(
            "INSERT INTO row_events (fp_row_id, event_type, old_value, new_value) VALUES (?, ?, ?, ?)",
            (row_id, "month_changed", "2026-08", "2026-09"),
        )
        conn.commit()
        fact_count = conn.execute(
            "SELECT COUNT(*) AS count FROM row_events WHERE event_type = 'portfolio_changed'"
        ).fetchone()["count"]
        month_count = conn.execute(
            "SELECT COUNT(*) AS count FROM row_events WHERE event_type = 'month_changed'"
        ).fetchone()["count"]
        conn.close()

        self.assertEqual(preserved, 1)
        self.assertEqual(fact_count, 1)
        self.assertEqual(month_count, 1)

    def test_fact_tab_is_visible(self):
        main.app.config.update(TESTING=True, SECRET_KEY="portfolio-changed-secret")
        client = main.app.test_client()
        login = client.post(
            "/login",
            data={"username": "owner", "password": "owner123"},
            follow_redirects=False,
        )
        self.assertEqual(login.status_code, 302)

        response = client.get("/changes?tab=fact")

        self.assertEqual(response.status_code, 200)
        self.assertIn("Перешло в Факт", response.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
