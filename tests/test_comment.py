"""
Тесты сохранности комментариев к строкам ФП.
  T1 — шаблон rows_list.html использует r.internal_comment, а не r.comment (баг)
  T2 — повторный импорт (тот же row_key) не трогает internal_comment
  T3 — смена row_key внутри квартала (другой месяц, тот же проект+раздел)
       → комментарий переносится через same_q_user_data
  T4 — первый импорт нового квартала → комментарий переносится через carry_over_map
"""
import sys, os, re, types, unittest, tempfile, shutil, sqlite3

# ── заглушка werkzeug (нет в sandbox-python, но нужна db.py при импорте) ─────
def _stub_werkzeug():
    try:
        import werkzeug.security
        return
    except ImportError:
        pass

    _ws     = types.ModuleType("werkzeug")
    _ws_sec = types.ModuleType("werkzeug.security")
    _ws_sec.generate_password_hash = lambda pw, **kw: "hashed:" + pw
    _ws_sec.check_password_hash    = lambda h, pw: h == "hashed:" + pw
    _ws.security = _ws_sec
    sys.modules.setdefault("werkzeug",          _ws)
    sys.modules.setdefault("werkzeug.security", _ws_sec)

_stub_werkzeug()

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import openpyxl
import app.importer as _imp_mod
import app.db       as _db_mod
from app.importer import import_excel

# ── минимальная схема ─────────────────────────────────────────────────────────

_SCHEMA = """
PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON;
CREATE TABLE IF NOT EXISTS app_settings  (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS users         (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, full_name TEXT, password_hash TEXT, role TEXT DEFAULT 'viewer');
CREATE TABLE IF NOT EXISTS fp_rows       (id INTEGER PRIMARY KEY AUTOINCREMENT, row_key TEXT NOT NULL UNIQUE, quarter_label TEXT, month TEXT, pc TEXT, section TEXT, client_name TEXT, project_num TEXT, project_name TEXT, project_manager TEXT, contract_num TEXT, ez_num TEXT, sdz_date TEXT, dpa_date TEXT, fdz_date TEXT, accounting_entity TEXT, mp_0_100 INTEGER, mp_comment TEXT, crm_amount REAL DEFAULT 0, amount_0_100 REAL DEFAULT 0, note TEXT, portfolio TEXT, kolodec TEXT, is_active INTEGER DEFAULT 1, internal_comment TEXT, risk_level INTEGER DEFAULT 0, responsible_user_id INTEGER, last_seen_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS obligations   (id INTEGER PRIMARY KEY AUTOINCREMENT, fp_row_id INTEGER NOT NULL REFERENCES fp_rows(id), title TEXT NOT NULL, description TEXT DEFAULT '', responsible_type TEXT NOT NULL DEFAULT 'team_lead', responsible_name TEXT, due_date TEXT, status TEXT NOT NULL DEFAULT 'not_started', created_by INTEGER, completed_at TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS import_log    (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, rows_total INTEGER DEFAULT 0, rows_new INTEGER DEFAULT 0, rows_updated INTEGER DEFAULT 0, rows_deactivated INTEGER DEFAULT 0, imported_by INTEGER, imported_at TEXT DEFAULT CURRENT_TIMESTAMP, diff_json TEXT);
CREATE TABLE IF NOT EXISTS row_events    (id INTEGER PRIMARY KEY AUTOINCREMENT, fp_row_id INTEGER, import_log_id INTEGER, event_type TEXT, field_label TEXT, old_value TEXT, new_value TEXT, amount_before REAL, amount_after REAL, month TEXT, client_name TEXT, project_name TEXT, section TEXT, portfolio TEXT, contract_num TEXT, reviewed_at TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS import_money_snapshots (id INTEGER PRIMARY KEY AUTOINCREMENT, import_log_id INTEGER UNIQUE NOT NULL REFERENCES import_log(id), quarter_label TEXT NOT NULL, fact_amount REAL NOT NULL DEFAULT 0, plan_amount REAL NOT NULL DEFAULT 0, opportunities_amount REAL NOT NULL DEFAULT 0, created_at TEXT DEFAULT CURRENT_TIMESTAMP);
"""


def _make_db():
    tmp  = tempfile.mkdtemp()
    path = os.path.join(tmp, "t.db")
    c = sqlite3.connect(path)
    c.row_factory = sqlite3.Row
    c.executescript(_SCHEMA)
    c.commit(); c.close()

    def gc():
        c2 = sqlite3.connect(path)
        c2.row_factory = sqlite3.Row
        c2.execute("PRAGMA foreign_keys=ON")
        return c2

    return tmp, gc


def _xlsx(tmp, project_num="P001", section="Тест", month="01"):
    wb = openpyxl.Workbook(); ws = wb.active
    ws.append(["Месяц","Стратегическое решение","ПЦ","РП","Раздел ФП",
               "Наименование клиента","Номер проекта","Наименование проекта",
               "Менеджер проекта","Номер договора","Номер ЭЗ","СДЗ","ДПА",
               "ФДЗ","Способ учета","0-100 от МП","Комментарий МП к СДЗ",
               "Сумма CRM, руб","Сумма по 0-100, руб","Примечание","Портфель",
               "Колодец","Внешний ID клиента"])
    ws.append([month,"","ПЦ1","",section,"Клиент А",project_num,"Проект А",
               "","","",None,None,None,"",None,"",0,100.0,"","0-100","",""])
    p = os.path.join(tmp, "fp.xlsx"); wb.save(p); return p


def _do_import(gc, tmp, quarter_label,
               project_num="P001", section="Тест", month="01"):
    orig_imp, orig_db = _imp_mod.get_conn, _db_mod.get_conn
    _imp_mod.get_conn = _db_mod.get_conn = gc
    try:
        return import_excel(
            _xlsx(tmp, project_num, section, month),
            "fp.xlsx", 1, quarter_label=quarter_label,
        )
    finally:
        _imp_mod.get_conn, _db_mod.get_conn = orig_imp, orig_db


# ── тесты ─────────────────────────────────────────────────────────────────────

class T1_TemplateFieldName(unittest.TestCase):
    """
    БАРЬЕРНЫЙ ТЕСТ.
    rows_list.html должен читать r.internal_comment, а не r.comment.
    r.comment не существует в fp_rows → textarea всегда пустая →
    пользователь жмёт «Сохранить» → comment='' → internal_comment затирается.
    """
    _TMPL = os.path.join(ROOT, "app", "templates", "rows_list.html")

    def _html(self):
        with open(self._TMPL, encoding="utf-8") as f:
            return f.read()

    def test_no_r_dot_comment(self):
        bad = re.findall(r'\br\.comment\b', self._html())
        self.assertEqual(bad, [],
            "rows_list.html содержит r.comment — несуществующая колонка! "
            "Используйте r.internal_comment.")

    def test_r_internal_comment_present(self):
        good = re.findall(r'\br\.internal_comment\b', self._html())
        self.assertGreater(len(good), 0,
            "rows_list.html не использует r.internal_comment.")


class T1b_CommentPopoverVisibility(unittest.TestCase):
    _CSS = os.path.join(ROOT, "app", "static", "css", "style.css")

    def test_open_client_card_does_not_clip_editor(self):
        with open(self._CSS, encoding="utf-8") as f:
            css = f.read()
        self.assertRegex(
            css,
            r"\.fp-cli-block\[open\]\s*\{[^}]*overflow:\s*visible",
            "Open client card must not clip the inline comment editor.",
        )


class T2_CommentSurvivesReimport(unittest.TestCase):
    """Повторный импорт (тот же row_key) не трогает internal_comment."""

    def setUp(self):
        self.tmp, self.gc = _make_db()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_comment_preserved(self):
        _do_import(self.gc, self.tmp, "2025Q1")

        conn = self.gc()
        row_id = conn.execute(
            "SELECT id FROM fp_rows WHERE is_active=1 LIMIT 1"
        ).fetchone()["id"]
        conn.execute("UPDATE fp_rows SET internal_comment='Мой комментарий' WHERE id=?",
                     (row_id,))
        conn.commit(); conn.close()

        _do_import(self.gc, self.tmp, "2025Q1")  # повторный импорт

        conn = self.gc()
        val = conn.execute("SELECT internal_comment FROM fp_rows WHERE id=?",
                           (row_id,)).fetchone()["internal_comment"]
        conn.close()
        self.assertEqual(val, "Мой комментарий",
                         "Комментарий должен сохраниться при повторном импорте")


class T3_CommentCarriedOnKeyChange(unittest.TestCase):
    """
    Если row_key изменился внутри квартала (изменился месяц, но project_num
    и section те же) — комментарий переносится через same_q_user_data.
    """

    def setUp(self):
        self.tmp, self.gc = _make_db()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_comment_carried(self):
        _do_import(self.gc, self.tmp, "2025Q1", month="01")

        conn = self.gc()
        old = conn.execute(
            "SELECT id FROM fp_rows WHERE is_active=1 LIMIT 1"
        ).fetchone()
        conn.execute("UPDATE fp_rows SET internal_comment='Важный' WHERE id=?",
                     (old["id"],))
        conn.commit(); conn.close()

        # Другой месяц → другой row_key, но (project_num, section) те же
        _do_import(self.gc, self.tmp, "2025Q1", month="02")

        conn = self.gc()
        new = conn.execute(
            "SELECT internal_comment FROM fp_rows WHERE is_active=1 LIMIT 1"
        ).fetchone()
        conn.close()
        self.assertEqual(new["internal_comment"], "Важный",
                         "Комментарий должен перенестись при смене row_key внутри квартала")


class T4_CommentCarriedToNewQuarter(unittest.TestCase):
    """При первом импорте нового квартала комментарий переносится через carry_over_map."""

    def setUp(self):
        self.tmp, self.gc = _make_db()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_comment_carried(self):
        _do_import(self.gc, self.tmp, "2025Q1")

        conn = self.gc()
        q1 = conn.execute(
            "SELECT id FROM fp_rows WHERE quarter_label='2025Q1' AND is_active=1 LIMIT 1"
        ).fetchone()
        conn.execute("UPDATE fp_rows SET internal_comment='Q1 комментарий' WHERE id=?",
                     (q1["id"],))
        conn.commit(); conn.close()

        _do_import(self.gc, self.tmp, "2025Q2")

        conn = self.gc()
        q2 = conn.execute(
            "SELECT internal_comment FROM fp_rows "
            "WHERE quarter_label='2025Q2' AND is_active=1 LIMIT 1"
        ).fetchone()
        conn.close()
        self.assertEqual(q2["internal_comment"], "Q1 комментарий",
                         "Комментарий должен перенестись из Q1 в Q2")


if __name__ == "__main__":
    unittest.main(verbosity=2)
