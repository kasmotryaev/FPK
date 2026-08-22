"""Цель периода привязана к метке квартала/финансового года, а не к перечню месяцев.

Старое поведение: ключом цели была строка вида «Июль-Август-Сентябрь». Из-за этого цель
«терялась» при переключении вида квартал ↔ финансовый год и при появлении в файле строк
нового месяца (ключ становился «Июль-Август-Сентябрь-Октябрь»).
"""
import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main


class TargetPeriodKeyTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "target-key.db"
        db.init_db()

        conn = db.get_conn()
        for idx, month in enumerate(("Июль", "Август", "Сентябрь")):
            conn.execute(
                """INSERT INTO fp_rows (row_key, quarter_label, month, section, client_name,
                                        project_name, amount_0_100, portfolio)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (f"row-{idx}", "2026-Q3", month, "Проекты", "ТЕСТ-БАНК", "Проект 1", 1000.0, "0-100"),
            )
        conn.commit()
        conn.close()

        main.app.config.update(TESTING=True, SECRET_KEY="target-key-secret")
        self.client = main.app.test_client()
        self.client.post("/login", data={"username": "owner", "password": "owner123"})
        # Смотрим конкретный квартал, а не «по умолчанию»
        self.client.post("/set-quarter-view", data={"mode": "quarter", "quarter": "2026-Q3"})

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        try:
            self.temp_dir.cleanup()
        except (PermissionError, NotADirectoryError):
            pass

    def _targets(self):
        conn = db.get_conn()
        rows = conn.execute("SELECT period_label, target_amount FROM quarter_targets").fetchall()
        conn.close()
        return {r["period_label"]: r["target_amount"] for r in rows}

    def test_quarter_target_is_stored_under_quarter_label(self):
        self.client.post("/target", data={"target_amount": "605000000"})
        self.assertEqual(self._targets(), {"2026-Q3": 605000000.0})

    def test_target_survives_new_month_in_data(self):
        self.client.post("/target", data={"target_amount": "605000000"})

        conn = db.get_conn()
        conn.execute(
            """INSERT INTO fp_rows (row_key, quarter_label, month, section, client_name,
                                    project_name, amount_0_100, portfolio)
               VALUES (?,?,?,?,?,?,?,?)""",
            ("row-oct", "2026-Q3", "Октябрь", "Проекты", "ТЕСТ-БАНК", "Проект 1", 500.0, "0-100"),
        )
        conn.commit()
        conn.close()

        page = self.client.get("/").get_data(as_text=True)
        self.assertIn("605 000 000 ₽", page)

    def test_fy_target_is_separate_and_falls_back_to_sum_of_quarters(self):
        self.client.post("/target", data={"target_amount": "605000000"})  # цель квартала
        self.client.post("/set-quarter-view", data={"mode": "fy", "fy": "FY2026"})

        conn = db.get_conn()
        vqls = db.fiscal_year_quarters("FY2026")
        amount, key, derived = main.read_target(conn, vqls, "fy")
        conn.close()
        self.assertEqual(key, "FY2026")
        self.assertTrue(derived)
        self.assertEqual(amount, 605000000.0)  # пока своей цели у ФГ нет — сумма кварталов

        self.client.post("/target", data={"target_amount": "2000000000"})
        targets = self._targets()
        self.assertEqual(targets["FY2026"], 2000000000.0)
        self.assertEqual(targets["2026-Q3"], 605000000.0)  # цель квартала не затёрта

    def test_legacy_month_key_is_migrated(self):
        conn = db.get_conn()
        conn.execute(
            "INSERT INTO quarter_targets (period_label, target_amount) VALUES (?,?)",
            ("Июль-Август-Сентябрь", 123456.0),
        )
        conn.commit()

        moved = main.migrate_target_keys(conn)
        rows = conn.execute("SELECT period_label, target_amount FROM quarter_targets").fetchall()
        conn.close()

        self.assertEqual(moved, 1)
        self.assertEqual({r["period_label"]: r["target_amount"] for r in rows}, {"2026-Q3": 123456.0})

    def test_migration_keeps_unmatched_rows(self):
        conn = db.get_conn()
        conn.execute(
            "INSERT INTO quarter_targets (period_label, target_amount) VALUES (?,?)",
            ("Январь-Февраль-Март", 777.0),
        )
        conn.commit()

        main.migrate_target_keys(conn)
        rows = conn.execute("SELECT period_label FROM quarter_targets").fetchall()
        conn.close()

        self.assertIn("Январь-Февраль-Март", [r["period_label"] for r in rows])


if __name__ == "__main__":
    unittest.main()
