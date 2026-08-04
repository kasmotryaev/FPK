"""Регрессионные тесты изменения пароля через страницу пользователей."""
import tempfile
import unittest
from pathlib import Path

import app.db as db
import app.main as main


class PasswordChangeTest(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_path = db.DB_PATH
        db.DB_PATH = Path(self.temp_dir.name) / "password-test.db"
        db.init_db()

        main.app.config.update(TESTING=True, SECRET_KEY="password-test-secret")
        self.client = main.app.test_client()

    def tearDown(self):
        db.DB_PATH = self.original_db_path
        self.temp_dir.cleanup()

    def login(self, password):
        return self.client.post(
            "/login",
            data={"username": "owner", "password": password},
            follow_redirects=False,
        )

    def owner_id(self):
        with self.client.session_transaction() as current_session:
            return current_session["user_id"]

    def change_password(self, password, confirmation):
        return self.client.post(
            f"/admin/users/{self.owner_id()}/edit",
            data={
                "username": "owner",
                "full_name": "Руководитель продукта",
                "password": password,
                "password_confirm": confirmation,
                "role": "owner",
                "team_lead_name": "",
            },
            follow_redirects=False,
        )

    def test_new_password_replaces_default_password(self):
        self.assertEqual(self.login("owner123").status_code, 302)

        response = self.change_password("NewPassword-2026!", "NewPassword-2026!")
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers["Location"].endswith("/login"))

        self.assertEqual(self.login("NewPassword-2026!").status_code, 302)
        self.client.get("/logout")
        self.assertEqual(self.login("owner123").status_code, 200)

    def test_mismatched_confirmation_keeps_old_password(self):
        self.assertEqual(self.login("owner123").status_code, 302)

        response = self.change_password("NewPassword-2026!", "different-password")
        self.assertEqual(response.status_code, 302)
        self.assertTrue(response.headers["Location"].endswith("/admin/users"))

        self.client.get("/logout")
        self.assertEqual(self.login("owner123").status_code, 302)
        self.client.get("/logout")
        self.assertEqual(self.login("NewPassword-2026!").status_code, 200)


if __name__ == "__main__":
    unittest.main(verbosity=2)
