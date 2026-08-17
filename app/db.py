import sqlite3
import hashlib
import time
import datetime
from pathlib import Path
from werkzeug.security import generate_password_hash

DB_PATH = Path(__file__).parent / "data" / "fp_portal.db"

# Список руководителей команд задаётся через страницу «Пользователи» в интерфейсе.
# Оставьте пустым — пользователей создаёт owner после первого входа.
TEAM_LEADS = []
OWNER_LABEL = "Руководитель продукта"
ALL_RESPONSIBLE = [OWNER_LABEL] + TEAM_LEADS

# Разделы ФП, ответственность по умолчанию на Owner (руководитель продукта)
OWNER_SECTIONS = {"Докупки", "Лицензии"}
# Разделы, требующие назначения одного из руководителей команд
TEAM_LEAD_SECTIONS = {"Проекты", "Заказные доработки"}
# Разделы только для отслеживания — ответственный не требуется, в риски не попадают
TRACKING_ONLY_SECTIONS = {"Сопровождение", "TaaS"}
# Разделы основного фокуса работы (0‑100 / Возможности) на дашборде квартала.
# Сопровождение и TaaS включены сюда же, чтобы их прогресс по плану/возможностям тоже было
# удобно отслеживать в этом блоке (отдельно от них же ниже есть детальная разбивка по клиентам).
FOCUS_SECTIONS = {"Проекты", "Заказные доработки", "Лицензии", "Сопровождение", "TaaS"}
# Категории для агрегированной таблицы «Детализация» по клиентам: раздел → категория
AGGREGATE_CATEGORY_MAP = {
    "Проекты": "Проекты",
    "Заказные доработки": "Заказные доработки",
    "Докупки": "Докупки",
    "Лицензии": "Докупки",
}
AGGREGATE_CATEGORY_ORDER = ["Проекты", "Заказные доработки", "Докупки"]
# Обратное отображение для ссылок «в детализацию»: категория → разделы, которые в неё входят
AGGREGATE_CATEGORY_SECTIONS = {
    "Проекты": ["Проекты"],
    "Заказные доработки": ["Заказные доработки"],
    "Докупки": ["Докупки", "Лицензии"],
}
# Уровни риска по статье ФП (risk_level): 0 — отметки нет, 1-3 — градация с примерной
# вероятностью фактического поступления денег. Используется и в детализации (выбор уровня
# на строке), и в агрегированной таблице (бейдж максимального уровня риска в группе).
RISK_LEVELS = {
    1: {"label": "Низкий риск", "probability": 90, "css": "low"},
    2: {"label": "Средний риск", "probability": 70, "css": "medium"},
    3: {"label": "Высокий риск", "probability": 50, "css": "high"},
}


# Папка проекта лежит в «Документах», синхронизируемых через iCloud Drive. Из-за этого
# macOS иногда на доли секунды отказывает в доступе к файлу базы — sqlite3 отдаёт
# "authorization denied", пока идёт фоновая синхронизация/выгрузка файла. Это не повреждение
# данных, проходит само. Несколько попыток с короткой паузой избавляют пользователя от
# случайной 500-й ошибки на таких кратковременных сбоях.
_CONNECT_RETRIES = 5
_CONNECT_RETRY_DELAY = 0.4  # секунды


def get_conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(_CONNECT_RETRIES):
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA foreign_keys = ON")
            conn.execute("PRAGMA journal_mode = WAL")   # лучше при параллельных чтениях
            conn.execute("PRAGMA cache_size = -8000")   # 8 МБ кэш страниц
            conn.execute("PRAGMA synchronous = NORMAL") # безопасно + быстрее FULL
            # SQLite's built-in LOWER()/LIKE case-folding is ASCII-only and doesn't handle Cyrillic.
            # Override LOWER with Python's Unicode-aware str.lower() so LOWER(x) LIKE LOWER(?) works
            # correctly for Cyrillic text (client names, project names, etc).
            conn.create_function("LOWER", 1, lambda s: s.lower() if isinstance(s, str) else s)
            return conn
        except sqlite3.DatabaseError as e:
            if "authorization denied" in str(e).lower() and attempt < _CONNECT_RETRIES - 1:
                time.sleep(_CONNECT_RETRY_DELAY)
                continue
            raise


def row_key(month, client, project_num, section, contract_num, ez_num, portfolio="", accounting_entity="", occurrence=0, quarter_label=""):
    """Вычисляет уникальный ключ строки ФП. Параметр quarter_label включён в хэш начиная с v2
    миграции, что позволяет одной и той же строке существовать в разных кварталах независимо."""
    raw = "|".join(str(x or "") for x in [month, client, project_num, section, contract_num, ez_num, portfolio, accounting_entity, occurrence, quarter_label])
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:24]


def compute_quarter_label(d):
    """«2026-Q3» и т.п. -- метка календарного квартала по дате d (datetime.date).
    Используется при первом запуске (инициализация current_quarter_label) и при миграции базы.
    Переход между кварталами в портале управляется ВРУЧНУЮ через «Настройки → Зафиксировать
    квартал», поэтому эта функция больше не определяет момент перехода при загрузке Excel."""
    q = (d.month - 1) // 3 + 1
    return f"{d.year}-Q{q}"


def next_quarter_label(label):
    """'2026-Q2' → '2026-Q3', '2026-Q4' → '2027-Q1'"""
    year, q = label.split("-Q")
    year, q = int(year), int(q)
    if q == 4:
        return f"{year + 1}-Q1"
    return f"{year}-Q{q + 1}"


def prev_quarter_label(label):
    """'2026-Q2' → '2026-Q1', '2026-Q1' → '2025-Q4'"""
    year, q = label.split("-Q")
    year, q = int(year), int(q)
    if q == 1:
        return f"{year - 1}-Q4"
    return f"{year}-Q{q - 1}"


def quarter_options(current_label, back=3, forward=3):
    """Список меток кварталов вокруг current_label для выпадающего списка в Настройках."""
    options = [current_label]
    q = current_label
    for _ in range(back):
        q = prev_quarter_label(q)
        options.insert(0, q)
    q = current_label
    for _ in range(forward):
        q = next_quarter_label(q)
        options.append(q)
    return options


def get_setting(conn, key, default=None):
    row = conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else default


def set_setting(conn, key, value):
    conn.execute(
        "INSERT INTO app_settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )


def calendar_to_fiscal(quarter_label):
    """Конвертирует метку календарного квартала в параметры финансового года.
    Финансовый год начинается 1 апреля.

    Примеры:
      '2026-Q2' (апр–июн 2026) → fy_key='FY2026', fy_label='ФГ 2026–2027', fq_num=1
      '2026-Q3' (июл–сен 2026) → fy_key='FY2026', fy_label='ФГ 2026–2027', fq_num=2
      '2026-Q4' (окт–дек 2026) → fy_key='FY2026', fy_label='ФГ 2026–2027', fq_num=3
      '2027-Q1' (янв–мар 2027) → fy_key='FY2026', fy_label='ФГ 2026–2027', fq_num=4

    Возвращает: (fy_key: str, fy_label: str, fq_num: int)
    """
    year, q = quarter_label.split("-Q")
    year, q = int(year), int(q)
    if q == 1:          # Январь–Март: 4-й квартал финансового года (предыдущий April start)
        fy_start = year - 1
        fq = 4
    elif q == 2:        # Апрель–Июнь: ФКВ 1
        fy_start = year
        fq = 1
    elif q == 3:        # Июль–Сентябрь: ФКВ 2
        fy_start = year
        fq = 2
    else:               # q == 4, Октябрь–Декабрь: ФКВ 3
        fy_start = year
        fq = 3
    fy_key = f"FY{fy_start}"
    fy_label = f"ФГ {fy_start}–{fy_start + 1}"  # en-dash для диапазона годов
    return fy_key, fy_label, fq


def fiscal_year_quarters(fy_key):
    """Возвращает список calendar quarter_label для данного финансового года.
    'FY2026' → ['2026-Q2', '2026-Q3', '2026-Q4', '2027-Q1']
    """
    start_year = int(fy_key[2:])
    return [
        f"{start_year}-Q2",
        f"{start_year}-Q3",
        f"{start_year}-Q4",
        f"{start_year + 1}-Q1",
    ]


def get_available_quarters(conn):
    """Возвращает список кварталов для переключателя вида.
    Включает кварталы с активными строками + current_quarter_label из настроек
    (даже если для него ещё нет строк — чтобы пустой квартал тоже был виден).
    Каждый элемент: {'label', 'count', 'is_finalized', 'fy_key', 'fy_label', 'fq_num'}
    """
    rows = conn.execute("""
        SELECT quarter_label, COUNT(*) as cnt
        FROM fp_rows WHERE is_active = 1 AND quarter_label IS NOT NULL AND quarter_label != ''
        GROUP BY quarter_label ORDER BY quarter_label
    """).fetchall()
    result = {r["quarter_label"]: r["cnt"] for r in rows}

    fin_str = get_setting(conn, "finalized_quarters") or ""
    finalized = {q.strip() for q in fin_str.split(",") if q.strip()}

    current = get_setting(conn, "current_quarter_label")
    if current and current not in result:
        result[current] = 0  # показываем, даже если данных ещё нет

    # Финализированные кварталы показываем всегда (даже пустые), иначе кнопка снятия
    # фиксации не появится, когда пользователь случайно зафиксировал пустой квартал.
    for q in finalized:
        if q not in result:
            result[q] = 0

    items = []
    for q in sorted(result.keys()):
        fy_key, fy_label, fq_num = calendar_to_fiscal(q)
        items.append({
            "label": q,
            "count": result[q],
            "is_finalized": q in finalized,
            "fy_key": fy_key,
            "fy_label": fy_label,
            "fq_num": fq_num,
        })
    return items


def get_available_fiscal_years(conn):
    """Список финансовых годов, производных из get_available_quarters.
    Каждый элемент: {'key': 'FY2026', 'label': 'ФГ 2026–2027', 'count': int}
    """
    quarters = get_available_quarters(conn)
    fy_map = {}
    for q in quarters:
        fk, fl = q["fy_key"], q["fy_label"]
        if fk not in fy_map:
            fy_map[fk] = {"key": fk, "label": fl, "count": 0}
        fy_map[fk]["count"] += q["count"]
    return [fy_map[k] for k in sorted(fy_map.keys())]


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        full_name TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK(role IN ('owner','team_lead','viewer')),
        team_lead_name TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS fp_rows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_key TEXT UNIQUE NOT NULL,
        month TEXT,
        pc TEXT,
        section TEXT,
        client_name TEXT,
        project_num TEXT,
        project_name TEXT,
        project_manager TEXT,
        contract_num TEXT,
        ez_num TEXT,
        sdz_date TEXT,
        dpa_date TEXT,
        fdz_date TEXT,
        accounting_entity TEXT,
        mp_0_100 INTEGER,
        mp_comment TEXT,
        crm_amount REAL,
        amount_0_100 REAL,
        note TEXT,
        portfolio TEXT,
        kolodec TEXT,
        is_active INTEGER DEFAULT 1,
        first_seen_at TEXT DEFAULT CURRENT_TIMESTAMP,
        last_seen_at TEXT DEFAULT CURRENT_TIMESTAMP,
        last_import_id INTEGER,
        responsible_user_id INTEGER REFERENCES users(id),
        is_risk INTEGER NOT NULL DEFAULT 0,
        risk_level INTEGER NOT NULL DEFAULT 0,
        quarter_label TEXT
    )
    """)

    # Ключ-значение для общих настроек портала, не привязанных к конкретной строке/пользователю.
    # Сейчас единственное применение -- current_quarter_label (см. compute_quarter_label выше и
    # переход на новый квартал в app/importer.py).
    cur.execute("""
    CREATE TABLE IF NOT EXISTS app_settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS obligations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fp_row_id INTEGER NOT NULL REFERENCES fp_rows(id) ON DELETE CASCADE,
        title TEXT NOT NULL,
        description TEXT,
        responsible_type TEXT NOT NULL CHECK(responsible_type IN ('owner','team_lead')),
        responsible_name TEXT NOT NULL,
        due_date TEXT,
        status TEXT NOT NULL DEFAULT 'not_started' CHECK(status IN ('not_started','in_progress','done','blocked')),
        created_by INTEGER REFERENCES users(id),
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
        completed_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS obligation_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        obligation_id INTEGER NOT NULL REFERENCES obligations(id) ON DELETE CASCADE,
        user_id INTEGER REFERENCES users(id),
        action TEXT NOT NULL,
        details TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS import_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        rows_total INTEGER,
        rows_new INTEGER,
        rows_updated INTEGER,
        rows_deactivated INTEGER,
        imported_by INTEGER REFERENCES users(id),
        imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
        diff_json TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS import_money_snapshots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_log_id INTEGER UNIQUE NOT NULL REFERENCES import_log(id) ON DELETE CASCADE,
        quarter_label TEXT NOT NULL,
        fact_amount REAL NOT NULL DEFAULT 0,
        plan_amount REAL NOT NULL DEFAULT 0,
        opportunities_amount REAL NOT NULL DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_import_money_snapshots_quarter ON import_money_snapshots(quarter_label, import_log_id)")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS quarter_targets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        period_label TEXT UNIQUE NOT NULL,
        target_amount REAL NOT NULL,
        set_by INTEGER REFERENCES users(id),
        updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Журнал событий по строкам ФП — постоянная (не привязанная к одной загрузке) лента
    # изменений: новые статьи, обнулённые/пропавшие суммы, реактивации, прочие изменения полей.
    # Каждое событие можно отметить просмотренным (reviewed_at/reviewed_by) — это и есть
    # очередь на просмотр, которая копится между загрузками.
    cur.execute("""
    CREATE TABLE IF NOT EXISTS row_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fp_row_id INTEGER NOT NULL REFERENCES fp_rows(id) ON DELETE CASCADE,
        import_log_id INTEGER REFERENCES import_log(id),
        event_type TEXT NOT NULL CHECK(event_type IN
            ('new','zeroed','deactivated','reactivated','amount_changed','field_changed','portfolio_changed','month_changed')),
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
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_row_events_type_reviewed ON row_events(event_type, reviewed_at)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_row_events_import ON row_events(import_log_id)")

    # Личные сохранённые наборы фильтров на странице «Статьи ФП» — каждый пользователь
    # видит и применяет только свои.
    cur.execute("""
    CREATE TABLE IF NOT EXISTS saved_filters (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        name TEXT NOT NULL,
        query_string TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_saved_filters_user ON saved_filters(user_id)")

    # Индексы на fp_rows/obligations -- ускоряют фильтрацию и сортировку на страницах
    # «Дашборд» и «Детализация» (WHERE is_active=1 AND month IN (...), фильтры по клиенту/
    # разделу/портфелю/менеджеру) и подгрузку обязательств по строкам (WHERE fp_row_id IN (...)).
    # Сейчас, при текущем объёме данных, это не критично, но не даёт деградировать по скорости
    # с ростом базы -- и ничего не меняет в содержимом/поведении страниц.
    # Составной индекс покрывает основной WHERE на всех страницах:
    # is_active=1 AND quarter_label IN (...) AND month IN (...)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fp_rows_active_ql_month ON fp_rows(is_active, quarter_label, month)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fp_rows_client ON fp_rows(client_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fp_rows_section ON fp_rows(section)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fp_rows_portfolio ON fp_rows(portfolio)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_fp_rows_manager ON fp_rows(project_manager)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_obligations_fp_row ON obligations(fp_row_id)")

    conn.commit()

    # SQLite cannot extend a CHECK constraint with ALTER TABLE. Rebuild legacy
    # row_events tables so existing installations can store portfolio changes.
    row_events_schema = cur.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'row_events'"
    ).fetchone()
    row_events_sql = (row_events_schema["sql"] or "") if row_events_schema else ""
    if row_events_schema and (
        "portfolio_changed" not in row_events_sql or "month_changed" not in row_events_sql
    ):
        conn.execute("PRAGMA foreign_keys = OFF")
        try:
            conn.execute("BEGIN")
            cur.execute("ALTER TABLE row_events RENAME TO row_events_legacy")
            cur.execute("""
            CREATE TABLE row_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fp_row_id INTEGER NOT NULL REFERENCES fp_rows(id) ON DELETE CASCADE,
                import_log_id INTEGER REFERENCES import_log(id),
                event_type TEXT NOT NULL CHECK(event_type IN
                    ('new','zeroed','deactivated','reactivated','amount_changed','field_changed','portfolio_changed','month_changed')),
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
            """)
            cur.execute("""
                INSERT INTO row_events (
                    id, fp_row_id, import_log_id, event_type, field_label,
                    old_value, new_value, amount_before, amount_after, month,
                    client_name, project_name, section, portfolio, contract_num,
                    created_at, reviewed_at, reviewed_by
                )
                SELECT
                    id, fp_row_id, import_log_id, event_type, field_label,
                    old_value, new_value, amount_before, amount_after, month,
                    client_name, project_name, section, portfolio, contract_num,
                    created_at, reviewed_at, reviewed_by
                FROM row_events_legacy
            """)
            cur.execute("DROP TABLE row_events_legacy")
            cur.execute("CREATE INDEX idx_row_events_type_reviewed ON row_events(event_type, reviewed_at)")
            cur.execute("CREATE INDEX idx_row_events_import ON row_events(import_log_id)")
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.execute("PRAGMA foreign_keys = ON")

    # Миграция: добавляем diff_json в существующие старые базы, если его не было
    cols = [c["name"] for c in cur.execute("PRAGMA table_info(import_log)").fetchall()]
    if "diff_json" not in cols:
        cur.execute("ALTER TABLE import_log ADD COLUMN diff_json TEXT")
        conn.commit()

    cols_fp = [c["name"] for c in cur.execute("PRAGMA table_info(fp_rows)").fetchall()]
    if "responsible_user_id" not in cols_fp:
        cur.execute("ALTER TABLE fp_rows ADD COLUMN responsible_user_id INTEGER REFERENCES users(id)")
        conn.commit()
    if "internal_comment" not in cols_fp:
        cur.execute("ALTER TABLE fp_rows ADD COLUMN internal_comment TEXT")
        conn.commit()
    if "is_risk" not in cols_fp:
        cur.execute("ALTER TABLE fp_rows ADD COLUMN is_risk INTEGER NOT NULL DEFAULT 0")
        conn.commit()
    if "risk_level" not in cols_fp:
        cur.execute("ALTER TABLE fp_rows ADD COLUMN risk_level INTEGER NOT NULL DEFAULT 0")
        conn.commit()
        # Перенос со старого булевого флага «под риском»: считаем как средний уровень,
        # чтобы не терять уже сделанные отметки — точный уровень владелец сможет уточнить вручную.
        cur.execute("UPDATE fp_rows SET risk_level = 2 WHERE is_risk = 1")
        conn.commit()
    if "quarter_label" not in cols_fp:
        cur.execute("ALTER TABLE fp_rows ADD COLUMN quarter_label TEXT")
        conn.commit()
        # До этой миграции квартал в данных никак не отмечался -- год по названию месяца
        # восстановить нельзя (в файле только «Апрель», без года). Поэтому весь существующий
        # массив (и активные, и уже отключившиеся строки) одной меткой относим к текущему
        # календарному кварталу на момент миграции -- это и есть тот квартал, который реально
        # сейчас ведётся в портале. Все ПОСЛЕДУЮЩИЕ загрузки и переходы между кварталами уже
        # будут отмечаться корректно (см. app/importer.py).
        current_label = compute_quarter_label(datetime.date.today())
        cur.execute("UPDATE fp_rows SET quarter_label = ? WHERE quarter_label IS NULL", (current_label,))
        conn.commit()
        if get_setting(conn, "current_quarter_label") is None:
            set_setting(conn, "current_quarter_label", current_label)
            conn.commit()

    # Миграция row_key v2: включаем quarter_label в хэш, чтобы одни и те же строки в разных
    # кварталах не коллидировали по UNIQUE(row_key). Пересчитываем ключи всех существующих строк.
    # Запускается однократно (флаг row_key_v2 в app_settings).
    if not get_setting(conn, "row_key_v2"):
        all_rows = cur.execute(
            """SELECT id, month, client_name, project_num, section, contract_num, ez_num,
               portfolio, accounting_entity, quarter_label FROM fp_rows ORDER BY id"""
        ).fetchall()
        occ_counter = {}
        updates = []
        for r in all_rows:
            base = (
                r["month"], r["client_name"], r["project_num"], r["section"],
                r["contract_num"], r["ez_num"], r["portfolio"] or "", r["accounting_entity"] or "",
                r["quarter_label"] or "",
            )
            occ = occ_counter.get(base, 0)
            occ_counter[base] = occ + 1
            new_key = row_key(
                r["month"], r["client_name"], r["project_num"], r["section"],
                r["contract_num"], r["ez_num"], r["portfolio"] or "", r["accounting_entity"] or "",
                occ, r["quarter_label"] or "",
            )
            updates.append((new_key, r["id"]))
        for new_key, row_id in updates:
            cur.execute("UPDATE fp_rows SET row_key = ? WHERE id = ?", (new_key, row_id))
        set_setting(conn, "row_key_v2", "1")
        conn.commit()

    cols_users = [c["name"] for c in cur.execute("PRAGMA table_info(users)").fetchall()]
    if "telegram_chat_id" not in cols_users:
        cur.execute("ALTER TABLE users ADD COLUMN telegram_chat_id TEXT")
        conn.commit()
    if "telegram_link_code" not in cols_users:
        cur.execute("ALTER TABLE users ADD COLUMN telegram_link_code TEXT")
        conn.commit()

    # ─── Трудозатраты ─────────────────────────────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS ts_imports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT NOT NULL,
        period_label TEXT,
        imported_by INTEGER REFERENCES users(id),
        imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
        rows_total INTEGER DEFAULT 0,
        rp_filter TEXT,
        pc_filter TEXT,
        file_type TEXT DEFAULT 'office'
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS ts_rows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_id INTEGER NOT NULL REFERENCES ts_imports(id) ON DELETE CASCADE,
        rp TEXT,
        rp_product TEXT,
        dept TEXT,
        division TEXT,
        employee TEXT,
        project TEXT,
        task TEXT,
        client TEXT,
        project_type TEXT,
        work_type TEXT,
        hours REAL NOT NULL DEFAULT 0
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_import ON ts_rows(import_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_dept ON ts_rows(import_id, dept)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_emp ON ts_rows(import_id, employee)")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS my_employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        full_name TEXT NOT NULL,
        imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(full_name)
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_my_employees_name ON my_employees(full_name)")
    conn.commit()

    # Миграция: если ts_rows содержит старую колонку row_level — пересоздаём таблицы
    ts_cols = [c["name"] for c in cur.execute("PRAGMA table_info(ts_rows)").fetchall()]
    if "row_level" in ts_cols:
        cur.execute("DROP TABLE IF EXISTS ts_rows")
        cur.execute("DROP TABLE IF EXISTS ts_imports")
        cur.execute("""
        CREATE TABLE ts_imports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            period_label TEXT,
            imported_by INTEGER REFERENCES users(id),
            imported_at TEXT DEFAULT CURRENT_TIMESTAMP,
            rows_total INTEGER DEFAULT 0,
            rp_filter TEXT,
            pc_filter TEXT
        )
        """)
        cur.execute("""
        CREATE TABLE ts_rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            import_id INTEGER NOT NULL REFERENCES ts_imports(id) ON DELETE CASCADE,
            rp TEXT,
            rp_product TEXT,
            dept TEXT,
            division TEXT,
            employee TEXT,
            project TEXT,
            task TEXT,
            client TEXT,
            project_type TEXT,
            work_type TEXT,
            hours REAL NOT NULL DEFAULT 0
        )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_import ON ts_rows(import_id)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_dept ON ts_rows(import_id, dept)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_ts_rows_emp ON ts_rows(import_id, employee)")
        conn.commit()

    # Миграция: добавляем pc_filter в ts_imports если его нет
    ts_imp_cols = [c["name"] for c in cur.execute("PRAGMA table_info(ts_imports)").fetchall()]
    if "pc_filter" not in ts_imp_cols:
        cur.execute("ALTER TABLE ts_imports ADD COLUMN pc_filter TEXT")
        conn.commit()
    # Миграция: добавляем file_type в ts_imports если его нет
    ts_imp_cols = [c["name"] for c in cur.execute("PRAGMA table_info(ts_imports)").fetchall()]
    if "file_type" not in ts_imp_cols:
        cur.execute("ALTER TABLE ts_imports ADD COLUMN file_type TEXT DEFAULT 'office'")
        conn.commit()
    # Миграция: добавляем rp_product в ts_rows если его нет
    ts_row_cols = [c["name"] for c in cur.execute("PRAGMA table_info(ts_rows)").fetchall()]
    if "rp_product" not in ts_row_cols:
        cur.execute("ALTER TABLE ts_rows ADD COLUMN rp_product TEXT")
        conn.commit()
    # ──────────────────────────────────────────────────────────────────────────

    # Создаём дефолтных пользователей при первом запуске
    cur.execute("SELECT COUNT(*) c FROM users")
    if cur.fetchone()["c"] == 0:
        default_users = [
            ("owner", "Руководитель продукта", "owner123", "owner", None),
        ]
        for name in TEAM_LEADS:
            username = name.lower().split()[0]
            default_users.append((username, name, username + "123", "team_lead", name))
        default_users.append(("viewer", "Наблюдатель", "viewer123", "viewer", None))

        for username, full_name, pwd, role, tl_name in default_users:
            cur.execute(
                "INSERT INTO users (username, full_name, password_hash, role, team_lead_name) VALUES (?,?,?,?,?)",
                (username, full_name, generate_password_hash(pwd, method="pbkdf2:sha256"), role, tl_name),
            )
        conn.commit()

    conn.close()


def generate_link_code():
    import secrets
    return secrets.token_hex(4).upper()
