import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main
from app.importer import _save_import_money_snapshot


class MoneyChartTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "money-chart.db"
        db.init_db()
        main.app.config.update(TESTING=True, SECRET_KEY="money-chart-secret")

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def test_import_snapshot_stores_only_portfolio_totals(self):
        conn = db.get_conn()
        conn.executemany(
            """
            INSERT INTO fp_rows (row_key, quarter_label, portfolio, amount_0_100, is_active)
            VALUES (?, '2026-Q3', ?, ?, ?)
            """,
            [
                ("fact", "Факт", 120.0, 1),
                ("plan", "0-100", 80.0, 1),
                ("opportunities", "Возможности", 30.0, 1),
                ("inactive", "Факт", 999.0, 0),
            ],
        )
        cursor = conn.execute(
            "INSERT INTO import_log (filename, imported_at) VALUES (?, ?)",
            ("test.xlsx", "2026-08-14 10:00:00"),
        )
        _save_import_money_snapshot(conn.cursor(), cursor.lastrowid, "2026-Q3")
        conn.commit()

        snapshot = conn.execute(
            "SELECT * FROM import_money_snapshots WHERE import_log_id = ?",
            (cursor.lastrowid,),
        ).fetchone()
        conn.close()

        self.assertEqual(snapshot["quarter_label"], "2026-Q3")
        self.assertEqual(snapshot["fact_amount"], 120.0)
        self.assertEqual(snapshot["plan_amount"], 80.0)
        self.assertEqual(snapshot["opportunities_amount"], 30.0)

    def test_history_aggregates_latest_state_for_fiscal_year(self):
        conn = db.get_conn()
        imports = []
        for filename, imported_at in (
            ("q2-a.xlsx", "2026-05-01 10:00:00"),
            ("q2-b.xlsx", "2026-05-10 10:00:00"),
            ("q3-a.xlsx", "2026-08-01 10:00:00"),
        ):
            cursor = conn.execute(
                "INSERT INTO import_log (filename, imported_at) VALUES (?, ?)",
                (filename, imported_at),
            )
            imports.append(cursor.lastrowid)
        conn.executemany(
            """
            INSERT INTO import_money_snapshots
                (import_log_id, quarter_label, fact_amount, plan_amount, opportunities_amount)
            VALUES (?, ?, ?, ?, ?)
            """,
            [
                (imports[0], "2026-Q2", 40.0, 60.0, 10.0),
                (imports[1], "2026-Q2", 50.0, 55.0, 12.0),
                (imports[2], "2026-Q3", 20.0, 30.0, 5.0),
            ],
        )
        conn.commit()

        history = main._get_money_history(conn, ["2026-Q2", "2026-Q3"])
        chart = main._prepare_money_chart(history)
        conn.close()

        self.assertEqual(len(history), 3)
        self.assertEqual(history[-1]["fact"], 70.0)
        self.assertEqual(history[-1]["plan"], 85.0)
        self.assertEqual(history[-1]["fact_plan"], 155.0)
        self.assertEqual(history[-1]["opportunities"], 17.0)
        self.assertEqual(chart["max_label"], "155 ₽")
        self.assertEqual(len(chart["series"]), 4)

    def test_dashboard_contains_collapsible_chart_section(self):
        conn = db.get_conn()
        user = conn.execute("SELECT id, full_name, role FROM users ORDER BY id LIMIT 1").fetchone()
        conn.close()
        client = main.app.test_client()
        with client.session_transaction() as session:
            session["user_id"] = user["id"]
            session["full_name"] = user["full_name"]
            session["role"] = user["role"]

        response = client.get("/")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('id="money-chart-details"', html)
        self.assertIn("Развернуть график", Path("app/static/css/style.css").read_text(encoding="utf-8"))
        self.assertIn("История сумм начнёт накапливаться", html)


    def test_dashboard_renders_saved_history(self):
        conn = db.get_conn()
        user = conn.execute("SELECT id, full_name, role FROM users ORDER BY id LIMIT 1").fetchone()
        import_row = conn.execute(
            "INSERT INTO import_log (filename, imported_at) VALUES (?, ?)",
            ("quarter.xlsx", "2026-08-14 10:00:00"),
        )
        conn.execute(
            """
            INSERT INTO import_money_snapshots
                (import_log_id, quarter_label, fact_amount, plan_amount, opportunities_amount)
            VALUES (?, ?, ?, ?, ?)
            """,
            (import_row.lastrowid, "2026-Q3", 70.0, 85.0, 17.0),
        )
        conn.commit()
        conn.close()

        client = main.app.test_client()
        with client.session_transaction() as session:
            session["user_id"] = user["id"]
            session["full_name"] = user["full_name"]
            session["role"] = user["role"]
            session["view_mode"] = "quarter"
            session["view_quarter"] = "2026-Q3"

        response = client.get("/")

        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn('class="money-chart-series"', html)
        self.assertIn("155 ₽", html)

if __name__ == "__main__":
    unittest.main(verbosity=2)
