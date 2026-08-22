# -*- coding: utf-8 -*-
"""Программа максимум: сохранение загрузок и сверка с финпланом из СЦ.

Зачем это в портале. ПМ и финплан СЦ — два разных среза одних и тех же денег:
ПМ собирается по ПЦ в разрезе решений (и живёт на стороне менеджеров УП), финплан СЦ —
по руководителю продукта. Расходятся они регулярно, и ровно в местах расхождения прячутся
вопросы к менеджерам: возможность, засчитанная в прогноз; деньги без стратегического решения;
позиции, которых в моём ФП нет вовсе.

Сопоставление идёт по клиенту и кварталу:
  ПМ «Прогноз» за квартал  ↔  fp_rows портфель «0-100» за тот же quarter_label.
«Возможности» и «Факт» показываем рядом, но в дельту не берём — это разные стадии денег.
"""
import json

from app.db import get_conn

# Порог, ниже которого расхождение считаем шумом округления, а не вопросом к менеджеру.
DEFAULT_THRESHOLD = 50_000.0

FORECAST_KIND = "Прогноз"
FACT_KIND = "Факт"

NO_STRATEGIC_LABEL = "— без стратрешения —"

# Чем «наши» деньги под чужим решением отличаются от просто чужих денег того же банка.
# Список тематический, а не про конкретный продукт: правится в настройках блока на странице.
DEFAULT_RELATED_KEYWORDS = (
    "депозитар", "цфа", "dfa", "custody", "перевод цб", "переводов ценных",
    "m2m", "м2м", "реестр", "хранени",
)

PM_ROW_FIELDS = (
    "client_name", "contract_num", "findoc", "presale", "pc", "section",
    "solution", "strategic_solution", "direction", "share", "manager", "kt",
    "directorate", "total_amount", "forecast_amount", "quarter_label", "kind", "amount",
)


def normalize_client(name):
    """Клиент в ПМ и в выгрузке СЦ пишется одинаково, но регистр и пробелы гуляют."""
    return " ".join(str(name or "").split()).upper()


def save_import(parsed, filename, user_id=None, report_label=None, conn=None):
    """Сохраняет разобранный файл ПМ. Возвращает id загрузки."""
    own_conn = conn is None
    conn = conn or get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO pm_imports
               (filename, report_label, strategic_filter, imported_by, rows_total, source_rows, quarters)
               VALUES (?,?,?,?,?,?,?)""",
            (
                filename,
                report_label or "",
                parsed["meta"].get("strategic_filter") or "",
                user_id,
                len(parsed["rows"]),
                parsed["meta"].get("source_rows", 0),
                json.dumps(
                    [{"fq_label": q["fq_label"], "quarter_label": q["quarter_label"]}
                     for q in parsed["quarters"]],
                    ensure_ascii=False,
                ),
            ),
        )
        import_id = cur.lastrowid
        placeholders = ",".join("?" * (len(PM_ROW_FIELDS) + 1))
        cur.executemany(
            f"INSERT INTO pm_rows (import_id, {', '.join(PM_ROW_FIELDS)}) VALUES ({placeholders})",
            [tuple([import_id] + [row.get(f) for f in PM_ROW_FIELDS]) for row in parsed["rows"]],
        )
        conn.commit()
        return import_id
    finally:
        if own_conn:
            conn.close()


def list_imports(conn):
    return conn.execute(
        "SELECT * FROM pm_imports ORDER BY imported_at DESC, id DESC"
    ).fetchall()


def get_import(conn, import_id=None):
    """Загрузка по id; без id — самая свежая. None, если загрузок ещё нет."""
    if import_id:
        return conn.execute("SELECT * FROM pm_imports WHERE id = ?", (import_id,)).fetchone()
    return conn.execute(
        "SELECT * FROM pm_imports ORDER BY imported_at DESC, id DESC LIMIT 1"
    ).fetchone()


def import_quarters(imp):
    """Список кварталов загрузки: [{'fq_label', 'quarter_label'}, ...]."""
    if not imp or not imp["quarters"]:
        return []
    try:
        return json.loads(imp["quarters"])
    except (ValueError, TypeError):
        return []


def strategic_options(conn, import_id):
    """Стратегические решения загрузки с суммами — для выпадающего списка."""
    rows = conn.execute(
        """SELECT COALESCE(NULLIF(TRIM(strategic_solution), ''), '— без стратрешения —') AS name,
                  SUM(amount) AS total
           FROM pm_rows WHERE import_id = ?
           GROUP BY name ORDER BY total DESC""",
        (import_id,),
    ).fetchall()
    return [{"name": r["name"], "total": r["total"] or 0.0} for r in rows]


def _pm_where(strategic):
    """Фильтр по стратрешению ищет и в «Страт. решение», и в «Решение» — так находятся
    строки, где стратегическое поле пустое, а продукт в названии решения указан."""
    if not strategic:
        return "", []
    like = f"%{strategic}%"
    return " AND (strategic_solution LIKE ? OR solution LIKE ?) ", [like, like]


def summary(conn, import_id, strategic=None, quarter=None):
    """Итоги ПМ: факт, прогноз по кварталам, разрез по разделам ФП."""
    cond, params = _pm_where(strategic)

    by_quarter = conn.execute(
        f"""SELECT kind, quarter_label, SUM(amount) AS total, COUNT(*) AS cnt
            FROM pm_rows WHERE import_id = ? {cond}
            GROUP BY kind, quarter_label ORDER BY quarter_label""",
        [import_id] + params,
    ).fetchall()

    q_cond, q_params = ("", [])
    if quarter:
        q_cond, q_params = " AND quarter_label = ? AND kind = ? ", [quarter, FORECAST_KIND]

    by_section = conn.execute(
        f"""SELECT section, SUM(amount) AS total, COUNT(*) AS cnt
            FROM pm_rows WHERE import_id = ? {cond} {q_cond}
            GROUP BY section ORDER BY total DESC""",
        [import_id] + params + q_params,
    ).fetchall()

    by_manager = conn.execute(
        f"""SELECT COALESCE(NULLIF(TRIM(manager), ''), '— не указан —') AS manager,
                   SUM(amount) AS total
            FROM pm_rows WHERE import_id = ? {cond} {q_cond}
            GROUP BY manager ORDER BY total DESC LIMIT 15""",
        [import_id] + params + q_params,
    ).fetchall()

    fact_total = sum(r["total"] or 0 for r in by_quarter if r["kind"] == FACT_KIND)
    quarters = [
        {"quarter_label": r["quarter_label"], "total": r["total"] or 0.0, "cnt": r["cnt"]}
        for r in by_quarter if r["kind"] == FORECAST_KIND
    ]
    return {
        "fact_total": fact_total,
        "forecast_total": sum(q["total"] for q in quarters),
        "grand_total": fact_total + sum(q["total"] for q in quarters),
        "quarters": quarters,
        "sections": [{"section": r["section"] or "—", "total": r["total"] or 0.0, "cnt": r["cnt"]}
                     for r in by_section],
        "managers": [{"manager": r["manager"], "total": r["total"] or 0.0} for r in by_manager],
    }


def _fp_by_client(conn, quarter):
    """{клиент → {портфель → сумма}} из финплана СЦ за квартал."""
    rows = conn.execute(
        """SELECT client_name, portfolio, SUM(amount_0_100) AS total
           FROM fp_rows
           WHERE is_active = 1 AND quarter_label = ?
           GROUP BY client_name, portfolio""",
        (quarter,),
    ).fetchall()
    result = {}
    for r in rows:
        key = normalize_client(r["client_name"])
        result.setdefault(key, {"name": r["client_name"], "Факт": 0.0, "0-100": 0.0, "Возможности": 0.0})
        if r["portfolio"] in result[key]:
            result[key][r["portfolio"]] += r["total"] or 0.0
    return result


def client_comparison(conn, import_id, quarter, strategic=None, threshold=DEFAULT_THRESHOLD):
    """Сверка ПМ ↔ финплан по клиентам за квартал.

    Статусы строк:
      only_pm  — есть в ПМ, в финплане клиента за квартал нет вовсе;
      only_fp  — наоборот, деньги в финплане есть, а в ПМ по этому решению их нет;
      diff     — есть в обоих, но расхождение больше порога;
      ok       — сходится.
    """
    cond, params = _pm_where(strategic)
    pm_rows = conn.execute(
        f"""SELECT client_name, SUM(amount) AS total, COUNT(*) AS cnt
            FROM pm_rows
            WHERE import_id = ? AND kind = ? AND quarter_label = ? {cond}
            GROUP BY client_name""",
        [import_id, FORECAST_KIND, quarter] + params,
    ).fetchall()

    pm_by_client = {}
    for r in pm_rows:
        pm_by_client[normalize_client(r["client_name"])] = {
            "name": r["client_name"], "total": r["total"] or 0.0, "cnt": r["cnt"],
        }

    fp_by_client = _fp_by_client(conn, quarter)

    result = []
    for key in set(pm_by_client) | set(fp_by_client):
        pm = pm_by_client.get(key)
        fp = fp_by_client.get(key)
        pm_total = pm["total"] if pm else 0.0
        fp_0_100 = fp["0-100"] if fp else 0.0
        delta = round(pm_total - fp_0_100, 2)

        if pm and not fp:
            status = "only_pm"
        elif fp and not pm:
            # клиент без денег в финплане за квартал — не повод для строки в сверке
            if abs(fp_0_100) < threshold and abs(fp.get("Возможности", 0.0)) < threshold:
                continue
            status = "only_fp"
        elif abs(delta) > threshold:
            status = "diff"
        else:
            status = "ok"

        result.append({
            "client": (pm or fp)["name"],
            "pm_total": pm_total,
            "pm_rows": pm["cnt"] if pm else 0,
            "fp_fact": fp["Факт"] if fp else 0.0,
            "fp_0_100": fp_0_100,
            "fp_capability": fp["Возможности"] if fp else 0.0,
            "delta": delta,
            "status": status,
        })

    order = {"diff": 0, "only_pm": 1, "only_fp": 2, "ok": 3}
    result.sort(key=lambda r: (order[r["status"]], -abs(r["delta"]), -r["pm_total"]))
    return result


def other_solutions(conn, import_id, quarter, strategic, keywords=None, limit=40):
    """Деньги наших клиентов, лежащие под ДРУГИМ (или пустым) стратегическим решением.

    Именно сюда попадают классические потери разреза: строки без стратрешения и проекты,
    отнесённые на соседний продукт. Клиенты берутся те, что уже есть в нашем периметре.

    Без фильтра по ключевым словам блок утонул бы в чужих деньгах тех же банков (ядро,
    платежи, отчётность), поэтому оставляем только:
      * строки вообще без стратегического решения — их надо проставить в любом случае;
      * строки, где в названии решения есть слово из keywords (тематически наши).
    """
    if not strategic:
        return []

    needles = [k.strip().casefold() for k in (keywords or DEFAULT_RELATED_KEYWORDS) if k.strip()]

    cond, params = _pm_where(strategic)
    our_clients = [
        normalize_client(r["client_name"])
        for r in conn.execute(
            f"SELECT DISTINCT client_name FROM pm_rows WHERE import_id = ? {cond}",
            [import_id] + params,
        ).fetchall()
    ]
    if not our_clients:
        return []

    like = f"%{strategic}%"
    rows = conn.execute(
        """SELECT client_name,
                  COALESCE(NULLIF(TRIM(strategic_solution), ''), ?) AS strategic,
                  COALESCE(NULLIF(TRIM(solution), ''), '—') AS solution,
                  presale, SUM(amount) AS total
           FROM pm_rows
           WHERE import_id = ? AND kind = ? AND quarter_label = ?
             AND strategic_solution NOT LIKE ? AND solution NOT LIKE ?
           GROUP BY client_name, strategic, solution, presale
           ORDER BY total DESC""",
        (NO_STRATEGIC_LABEL, import_id, FORECAST_KIND, quarter, like, like),
    ).fetchall()

    our = set(our_clients)
    out = []
    for r in rows:
        if normalize_client(r["client_name"]) not in our:
            continue
        no_strategic = r["strategic"] == NO_STRATEGIC_LABEL
        text = f"{r['solution']} {r['strategic']}".casefold()
        if not no_strategic and needles and not any(n in text for n in needles):
            continue
        out.append({
            "client": r["client_name"],
            "strategic": r["strategic"],
            "solution": r["solution"],
            "presale": r["presale"] or "",
            "total": r["total"] or 0.0,
        })
        if len(out) >= limit:
            break
    return out


def client_details(conn, import_id, quarter, client, strategic=None):
    """Построчная расшифровка клиента за квартал — чтобы понять, из чего сложилась дельта."""
    cond, params = _pm_where(strategic)
    rows = conn.execute(
        f"""SELECT * FROM pm_rows
            WHERE import_id = ? AND kind = ? AND quarter_label = ?
              AND UPPER(TRIM(client_name)) = ? {cond}
            ORDER BY amount DESC""",
        [import_id, FORECAST_KIND, quarter, normalize_client(client)] + params,
    ).fetchall()
    return rows
