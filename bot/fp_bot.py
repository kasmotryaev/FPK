"""
Telegram-бот для ФП-Контроль.

Запускается отдельным процессом (python3 bot/run_bot.py), рядом с веб-порталом.
Использует ту же базу SQLite app/data/fp_portal.db.

Команды:
  /start             — приветствие и инструкция
  /link КОД          — привязать Telegram-аккаунт к пользователю портала (код выдаёт Owner)
  /my                — список своих обязательств
  /client             — выбрать клиента из списка кнопками (или /client название для поиска)
  /clientobl название — посмотреть и изменить обязательства по клиенту
  /add                — добавить обязательство (диалог по шагам, выбор кнопками)
  /help               — список команд

Любое сообщение без "/" в начале — это вопрос чату портала (тот же ИИ-чат "Спросить
портал", что и на дашборде): answer_portal_query собирает контекст из БД и отвечает
через Ollama, либо локальным поиском, если Ollama недоступна.

Также есть кнопки под каждым обязательством для быстрой смены статуса.
Фоновый поток раз в N минут проверяет дедлайны и шлёт напоминания.
"""
import sys
import time
import html
import datetime
import tempfile
import threading
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from app.db import get_conn, TEAM_LEADS, OWNER_LABEL, ALL_RESPONSIBLE, TEAM_LEAD_SECTIONS
from app.main import answer_portal_query
from app.stt import transcribe, SttError
from bot.telegram_client import TelegramClient, TelegramError, inline_keyboard

REMINDER_CHECK_INTERVAL_SECONDS = 15 * 60  # каждые 15 минут проверяем дедлайны
TELEGRAM_MAX_LEN = 3500  # запас от лимита Telegram (4096) на разбивку длинных ответов


def split_for_telegram(text, limit=TELEGRAM_MAX_LEN):
    """Telegram режет сообщения длиннее ~4096 символов. Бьём по границам строк, чтобы
    не разрывать слова и таблицы посередине строки."""
    text = (text or "").strip()
    if not text:
        return []
    if len(text) <= limit:
        return [text]
    chunks, current = [], ""
    for line in text.split("\n"):
        candidate = f"{current}\n{line}" if current else line
        if len(candidate) > limit and current:
            chunks.append(current)
            current = line
        else:
            current = candidate
    if current:
        chunks.append(current)
    return chunks
STATUS_LABELS = {
    "not_started": "Не начато", "in_progress": "В работе",
    "done": "Выполнено", "blocked": "Блокировано",
}
STATUS_EMOJI = {"not_started": "⬜", "in_progress": "🔵", "done": "✅", "blocked": "⛔"}

# Временное состояние диалога /add для каждого chat_id (в памяти процесса бота)
_add_flow_state = {}

# Предложенное, но ещё не подтверждённое действие через чат (риск/обязательство/статус/
# комментарий) для каждого chat_id -- см. answer_portal_query/_try_handle_chat_action в
# app/main.py. Как и _add_flow_state, живёт только в памяти процесса (теряется при
# перезапуске бота, это нормально -- предложение просто нужно будет повторить).
_pending_action_state = {}


def fmt_money(amount):
    return "{:,.0f}".format(amount or 0).replace(",", " ") + " ₽"


def truncate(text, max_len=24):
    text = text or ""
    if len(text) <= max_len:
        return text
    return text[: max_len - 1] + "…"


def fmt_date_short(date_iso):
    if not date_iso:
        return "—"
    # 2026-06-23 -> 06-23
    parts = date_iso.split("-")
    return f"{parts[1]}-{parts[2]}" if len(parts) == 3 else date_iso


def build_rows_table(rows, obls_by_row, project_width=22, num_width=10):
    """Строит моноширинную таблицу проект/№проекта/сумма/ДПА/обязательства для блока <pre>."""
    header = f"{'Проект':<{project_width}} {'№ проекта':<{num_width}} {'Сумма':>12} {'ДПА':>6} {'Об.':>4}"
    sep = "─" * len(header)
    lines = [header, sep]
    for r in rows:
        obls = obls_by_row.get(r["id"], [])
        obl_str = str(len(obls)) if obls else "–"
        proj = truncate(r["project_name"], project_width)
        proj_num = truncate(r["project_num"] or "—", num_width)
        amount = "{:,.0f}".format(r["amount_0_100"] or 0).replace(",", " ")
        date_short = fmt_date_short(r["dpa_date"])
        lines.append(f"{proj:<{project_width}} {proj_num:<{num_width}} {amount:>12} {date_short:>6} {obl_str:>4}")
    return "\n".join(lines)


def get_user_by_chat_id(chat_id):
    conn = get_conn()
    user = conn.execute("SELECT * FROM users WHERE telegram_chat_id = ?", (str(chat_id),)).fetchone()
    conn.close()
    return user


def link_account(chat_id, code):
    conn = get_conn()
    user = conn.execute("SELECT * FROM users WHERE telegram_link_code = ?", (code.strip().upper(),)).fetchone()
    if not user:
        conn.close()
        return None
    conn.execute(
        "UPDATE users SET telegram_chat_id = ?, telegram_link_code = NULL WHERE id = ?",
        (str(chat_id), user["id"]),
    )
    conn.commit()
    conn.close()
    return user


def get_my_obligations(user):
    conn = get_conn()
    if user["role"] == "team_lead":
        obls = conn.execute("""
            SELECT o.*, f.client_name, f.project_name, f.project_num, f.section, f.amount_0_100, f.contract_num
            FROM obligations o JOIN fp_rows f ON f.id = o.fp_row_id
            WHERE o.responsible_name = ? ORDER BY o.due_date IS NULL, o.due_date ASC
        """, (user["team_lead_name"],)).fetchall()
    elif user["role"] == "owner":
        obls = conn.execute("""
            SELECT o.*, f.client_name, f.project_name, f.project_num, f.section, f.amount_0_100, f.contract_num
            FROM obligations o JOIN fp_rows f ON f.id = o.fp_row_id
            WHERE o.responsible_name = ? ORDER BY o.due_date IS NULL, o.due_date ASC
        """, (OWNER_LABEL,)).fetchall()
    else:
        obls = []
    conn.close()
    return obls


def get_client_info(client_name):
    conn = get_conn()
    rows_all = conn.execute("""
        SELECT * FROM fp_rows WHERE is_active = 1 ORDER BY dpa_date IS NULL, dpa_date ASC
    """).fetchall()
    query_lower = client_name.strip().lower()
    rows = [r for r in rows_all if query_lower in (r["client_name"] or "").lower()]
    row_ids = [r["id"] for r in rows]
    obls_by_row = {}
    if row_ids:
        ph = ",".join("?" * len(row_ids))
        obls = conn.execute(
            f"SELECT * FROM obligations WHERE fp_row_id IN ({ph}) ORDER BY due_date IS NULL, due_date ASC", row_ids
        ).fetchall()
        for o in obls:
            obls_by_row.setdefault(o["fp_row_id"], []).append(o)
    conn.close()
    return rows, obls_by_row


def get_all_clients():
    conn = get_conn()
    rows = conn.execute("SELECT DISTINCT client_name FROM fp_rows WHERE is_active=1 ORDER BY client_name").fetchall()
    conn.close()
    return [r["client_name"] for r in rows]


def today_iso():
    return datetime.date.today().isoformat()


class FPBot:
    def __init__(self, token):
        self.tg = TelegramClient(token)
        self.offset = None
        self._stop = False

    # ---------- Отправка обязательства с кнопками ----------

    def _obligation_text_and_kb(self, o, show_client=True):
        is_overdue = o["due_date"] and o["due_date"] < today_iso() and o["status"] != "done"
        status_str = f"{STATUS_EMOJI.get(o['status'], '')} {STATUS_LABELS.get(o['status'], o['status'])}"
        lines = [f"<b>{o['title']}</b>"]
        if show_client:
            proj_num = o["project_num"] if "project_num" in (o.keys() if hasattr(o, "keys") else []) else None
            proj_label = f"{o['project_name']} (№ {proj_num})" if proj_num else o["project_name"]
            lines.append(f"{o['client_name']} · {proj_label}")
        keys = o.keys() if hasattr(o, "keys") else []
        if "amount_0_100" in keys:
            lines.append(f"Сумма статьи: {fmt_money(o['amount_0_100'])}")
        lines.append(f"Срок: {o['due_date'] or 'не задан'}{' ⚠ ПРОСРОЧЕНО' if is_overdue else ''}")
        lines.append(f"Статус: {status_str}")
        if o["description"]:
            lines.append(f"Комментарий: {o['description']}")
        text = "\n".join(lines)

        buttons = []
        row = []
        for status_key, label in STATUS_LABELS.items():
            if status_key != o["status"]:
                row.append({"text": f"{STATUS_EMOJI[status_key]} {label}", "callback_data": f"status:{o['id']}:{status_key}"})
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)
        return text, inline_keyboard(buttons)

    # ---------- Команды ----------

    def handle_start(self, chat_id):
        user = get_user_by_chat_id(chat_id)
        if user:
            self.tg.send_message(
                chat_id,
                f"Привет, {user['full_name']}! Аккаунт уже привязан.\n\n"
                "/help — список команд. Можно и просто написать вопрос без команды — "
                "отвечает чат портала.",
            )
        else:
            self.tg.send_message(
                chat_id,
                "Добро пожаловать в бот ФП-Контроль.\n\n"
                "Чтобы начать, получите код привязки у руководителя продукта (Owner) "
                "и отправьте команду:\n<code>/link КОД</code>",
            )

    def handle_help(self, chat_id):
        self.tg.send_message(
            chat_id,
            "Доступные команды:\n"
            "/my — мои обязательства\n"
            "/client — выбрать клиента из списка (или /client название для поиска)\n"
            "/clientobl название — посмотреть и изменить обязательства по клиенту\n"
            "/add — добавить обязательство\n"
            "/link КОД — привязать аккаунт\n"
            "/help — эта справка\n\n"
            "Можно просто написать вопрос без команды (например: «какие риски по Сберу?» "
            "или «найди 2606/23-2В-О») — отвечает тот же ИИ-чат, что и «Спросить портал» "
            "на дашборде.\n\n"
            "Текстом же можно и выполнять действия (если аккаунт привязан), например:\n"
            "«поставь высокий риск по Сберу», «добавь обязательство по проекту ... "
            "ответственный Зинин, срок 2026-07-01», «обязательство по Сберу выполнено».\n\n"
            "Перед любым действием бот сначала покажет, что собирается сделать, и ничего не "
            "изменит, пока вы не ответите «да» (отменить — «нет»).\n\n"
            "Можно не печатать, а наговорить вопрос или действие голосовым сообщением — бот "
            "распознает речь, покажет, что услышал, и дальше ответит как на обычный текст.",
        )

    def handle_link(self, chat_id, code):
        if not code:
            self.tg.send_message(chat_id, "Укажите код после команды, например: /link A1B2C3D4")
            return
        user = link_account(chat_id, code)
        if user:
            self.tg.send_message(chat_id, f"Готово! Привязан аккаунт: {user['full_name']}.\n\n/help — список команд")
        else:
            self.tg.send_message(chat_id, "Код не найден или уже использован. Запросите новый код у руководителя продукта.")

    def handle_my(self, chat_id):
        user = get_user_by_chat_id(chat_id)
        if not user:
            self.tg.send_message(chat_id, "Аккаунт не привязан. Используйте /link КОД")
            return
        obls = get_my_obligations(user)
        if not obls:
            self.tg.send_message(chat_id, "У вас нет обязательств.")
            return
        self.tg.send_message(chat_id, f"Ваши обязательства ({len(obls)}):")
        for o in obls:
            text, kb = self._obligation_text_and_kb(o)
            self.tg.send_message(chat_id, text, reply_markup=kb)

    CLIENTS_PER_PAGE = 10

    def _send_client_picker(self, chat_id, page=0, message_id=None):
        clients = get_all_clients()
        start = page * self.CLIENTS_PER_PAGE
        page_clients = clients[start:start + self.CLIENTS_PER_PAGE]
        total_pages = (len(clients) - 1) // self.CLIENTS_PER_PAGE + 1 if clients else 1

        buttons = [[{"text": c, "callback_data": f"pickclient:{c}"}] for c in page_clients]
        nav = []
        if page > 0:
            nav.append({"text": "« Назад", "callback_data": f"clientpage:{page-1}"})
        if start + self.CLIENTS_PER_PAGE < len(clients):
            nav.append({"text": "Вперёд »", "callback_data": f"clientpage:{page+1}"})
        if nav:
            buttons.append(nav)

        text = f"Выберите клиента (стр. {page+1}/{total_pages}):"
        kb = inline_keyboard(buttons)
        if message_id:
            self.tg.edit_message_text(chat_id, message_id, text, reply_markup=kb)
        else:
            self.tg.send_message(chat_id, text, reply_markup=kb)

    def handle_client(self, chat_id, client_query):
        if not client_query:
            self._send_client_picker(chat_id)
            return
        self._send_client_report(chat_id, client_query)

    def _send_client_report(self, chat_id, client_query):
        rows, obls_by_row = get_client_info(client_query)
        if not rows:
            self.tg.send_message(chat_id, f"Клиент по запросу «{client_query}» не найден.")
            return

        clients_found = sorted(set(r["client_name"] for r in rows))
        totals = {"Факт": 0.0, "0-100": 0.0, "Возможности": 0.0}
        for r in rows:
            p = r["portfolio"] or "Прочее"
            totals[p] = totals.get(p, 0.0) + (r["amount_0_100"] or 0)

        lines = [f"<b>{', '.join(clients_found)}</b>", ""]
        lines.append(f"Факт: {fmt_money(totals.get('Факт', 0))}")
        lines.append(f"План 0‑100: {fmt_money(totals.get('0-100', 0))}")
        lines.append(f"Возможности: {fmt_money(totals.get('Возможности', 0))}")
        lines.append("")
        lines.append(f"Строк ФП: {len(rows)}")
        self.tg.send_message(chat_id, "\n".join(lines))

        # Детализация: без Факта, две компактные таблицы (0-100, Возможности) в одном сообщении
        rows_0100 = [r for r in rows if r["portfolio"] == "0-100"]
        rows_opportunities = [r for r in rows if r["portfolio"] == "Возможности"]

        detail_blocks = []
        if rows_0100:
            detail_blocks.append("<b>План 0‑100</b>\n<pre>" + build_rows_table(rows_0100, obls_by_row) + "</pre>")
        if rows_opportunities:
            detail_blocks.append("<b>Возможности</b>\n<pre>" + build_rows_table(rows_opportunities, obls_by_row) + "</pre>")

        if detail_blocks:
            self.tg.send_message(chat_id, "\n\n".join(detail_blocks))
        else:
            self.tg.send_message(chat_id, "По плану 0‑100 и возможностям активных строк нет (только факт).")

        any_obligations = any(obls_by_row.get(r["id"]) for r in rows_0100 + rows_opportunities)
        if any_obligations:
            self.tg.send_message(
                chat_id,
                "Колонка «Об.» — количество обязательств по строке. "
                "Отправьте /clientobl " + clients_found[0] + " чтобы посмотреть и изменить обязательства по каждой строке.",
            )

    def handle_client_obligations(self, chat_id, client_query):
        if not client_query:
            self.tg.send_message(chat_id, "Укажите название клиента, например: /clientobl Ренкап")
            return
        rows, obls_by_row = get_client_info(client_query)
        rows_with_obls = [r for r in rows if obls_by_row.get(r["id"])]
        if not rows_with_obls:
            self.tg.send_message(chat_id, f"По клиенту «{client_query}» нет строк с обязательствами.")
            return
        for r in rows_with_obls:
            portfolio_tag = {"Факт": "✅ Факт", "0-100": "🔵 0-100", "Возможности": "⚪ Возможности"}.get(r["portfolio"], r["portfolio"])
            self.tg.send_message(chat_id, f"<b>{r['project_name']}</b> (№ {r['project_num'] or '—'})\n{portfolio_tag} · {fmt_money(r['amount_0_100'])} · ДПА {r['dpa_date'] or '—'}")
            for o in obls_by_row.get(r["id"], []):
                otext, okb = self._obligation_text_and_kb(o, show_client=False)
                self.tg.send_message(chat_id, otext, reply_markup=okb)

    def handle_add_start(self, chat_id):
        user = get_user_by_chat_id(chat_id)
        if not user:
            self.tg.send_message(chat_id, "Аккаунт не привязан. Используйте /link КОД")
            return
        _add_flow_state[chat_id] = {"step": "client", "data": {}}
        self.tg.send_message(chat_id, "Добавление обязательства.\nВведите название клиента (или часть названия) для поиска статьи:")

    def handle_add_flow(self, chat_id, text, user):
        state = _add_flow_state.get(chat_id)
        if not state:
            return False

        step = state["step"]

        if step == "client":
            conn = get_conn()
            all_rows = conn.execute("SELECT * FROM fp_rows WHERE is_active=1").fetchall()
            conn.close()
            query_lower = text.strip().lower()
            rows = [r for r in all_rows if query_lower in (r["client_name"] or "").lower()][:8]
            if not rows:
                self.tg.send_message(chat_id, "Ничего не найдено, попробуйте другое название клиента или /add заново.")
                _add_flow_state.pop(chat_id, None)
                return True
            buttons = [[{"text": f"{r['client_name']} · {r['project_name'][:24]} №{r['project_num'] or '—'} · {fmt_money(r['amount_0_100'])}",
                         "callback_data": f"addrow:{r['id']}"}] for r in rows]
            self.tg.send_message(chat_id, "Выберите статью:", reply_markup=inline_keyboard(buttons))
            state["step"] = "waiting_row_pick"
            return True

        if step == "title":
            state["data"]["title"] = text
            state["step"] = "waiting_due_date_pick"
            row_id = state["data"]["row_id"]
            conn = get_conn()
            row = conn.execute("SELECT dpa_date FROM fp_rows WHERE id=?", (row_id,)).fetchone()
            conn.close()
            today = datetime.date.today()
            buttons = [
                [{"text": "Сегодня", "callback_data": "adddue:" + today.isoformat()}],
                [{"text": "+3 дня", "callback_data": "adddue:" + (today + datetime.timedelta(days=3)).isoformat()}],
                [{"text": "+7 дней", "callback_data": "adddue:" + (today + datetime.timedelta(days=7)).isoformat()}],
            ]
            if row and row["dpa_date"]:
                buttons.append([{"text": f"По ДПА статьи ({row['dpa_date']})", "callback_data": "adddue:" + row["dpa_date"]}])
            buttons.append([{"text": "Без срока", "callback_data": "adddue:none"}])
            buttons.append([{"text": "Ввести дату вручную", "callback_data": "adddue:manual"}])
            self.tg.send_message(chat_id, "Выберите срок (дедлайн):", reply_markup=inline_keyboard(buttons))
            return True

        if step == "due_date_manual":
            due_date = None if text.strip().lower() in ("нет", "-", "") else text.strip()
            state["data"]["due_date"] = due_date
            self._ask_responsible(chat_id, state)
            return True

        return True

    def _ask_responsible(self, chat_id, state):
        row_id = state["data"]["row_id"]
        conn = get_conn()
        row = conn.execute("SELECT section FROM fp_rows WHERE id=?", (row_id,)).fetchone()
        conn.close()
        if row and row["section"] in TEAM_LEAD_SECTIONS:
            options = TEAM_LEADS
        else:
            options = ALL_RESPONSIBLE
        buttons = [[{"text": name, "callback_data": f"addresp:{name}"}] for name in options]
        self.tg.send_message(chat_id, "Выберите ответственного за обязательство:", reply_markup=inline_keyboard(buttons))
        state["step"] = "waiting_responsible_pick"

    def finalize_add_obligation(self, chat_id, user):
        state = _add_flow_state.get(chat_id)
        if not state:
            return
        data = state["data"]
        responsible_name = data["responsible_name"]
        responsible_type = "owner" if responsible_name == OWNER_LABEL else "team_lead"
        conn = get_conn()
        conn.execute("""
            INSERT INTO obligations (fp_row_id, title, description, responsible_type, responsible_name, due_date, created_by)
            VALUES (?,?,?,?,?,?,?)
        """, (data["row_id"], data["title"], "", responsible_type, responsible_name, data.get("due_date"), user["id"]))
        conn.commit()
        conn.close()
        self.tg.send_message(chat_id, f"Обязательство «{data['title']}» создано, ответственный: {responsible_name}.")
        _add_flow_state.pop(chat_id, None)

    # ---------- Свободный текст -> чат портала ----------

    def handle_chat_question(self, chat_id, text):
        """Сообщение без "/" в начале — вопрос тому же ИИ-чату, что и "Спросить портал"
        на дашборде. Ollama может отвечать не мгновенно, поэтому считаем ответ в фоновом
        потоке, чтобы не блокировать long polling для остальных чатов, и сразу показываем
        индикатор "печатает"."""
        try:
            self.tg.send_chat_action(chat_id, "typing")
        except TelegramError:
            pass
        threading.Thread(target=self._answer_chat_question, args=(chat_id, text), daemon=True).start()

    def _actor_for_chat(self, chat_id):
        """Данные пользователя для проверки прав при действиях через чат (риск, обязательства
        и т.п.) -- те же, что в веб-сессии. None, если телеграм-чат ещё не привязан к
        аккаунту портала (команда /link) -- тогда чат отвечает на вопросы, но не выполняет
        действия (см. answer_portal_query/_try_handle_chat_action в app/main.py)."""
        user = get_user_by_chat_id(chat_id)
        if not user:
            return None
        return {
            "role": user["role"],
            "team_lead_name": user["team_lead_name"],
            "user_id": user["id"],
            "full_name": user["full_name"],
        }

    def _answer_chat_question(self, chat_id, text):
        try:
            answer, pending = answer_portal_query(
                text, actor=self._actor_for_chat(chat_id), pending_action=_pending_action_state.get(chat_id),
            )
        except Exception as e:
            answer, pending = f"Не получилось получить ответ от чата портала: {e}", None
        if pending:
            _pending_action_state[chat_id] = pending
        else:
            _pending_action_state.pop(chat_id, None)
        chunks = split_for_telegram(answer) or ["Не нашёл, что ответить."]
        for chunk in chunks:
            # Моноширенный текст (<pre>) -- так удобнее читать таблицы и цифры в ответе.
            # Содержимое обязательно экранируем (html.escape), иначе символы <, >, & в ответе
            # (имена, комментарии и т.п.) могут сломать разбор HTML-разметки на стороне Telegram.
            wrapped = "<pre>" + html.escape(chunk, quote=False) + "</pre>"
            if not self._send_with_retry(chat_id, wrapped, parse_mode="HTML"):
                break

    def handle_voice_message(self, chat_id, msg):
        """Голосовое сообщение (voice) или аудио-файл (audio) -- скачиваем, распознаём
        локально через Whisper (app/stt.py) и дальше обрабатываем ровно как обычный
        текстовый вопрос/действие (_answer_chat_question), включая подтверждение «да»/«нет»
        перед записью в БД. Делаем всё в фоновом потоке: скачивание + распознавание +
        возможный ответ Ollama вместе могут занять заметное время."""
        try:
            self.tg.send_chat_action(chat_id, "typing")
        except TelegramError:
            pass
        voice = msg.get("voice") or msg.get("audio")
        file_id = voice["file_id"]
        threading.Thread(target=self._answer_voice_message, args=(chat_id, file_id), daemon=True).start()

    def _answer_voice_message(self, chat_id, file_id):
        tmp_path = None
        try:
            file_info = self.tg.get_file(file_id)
            data = self.tg.download_file(file_info["file_path"])
            with tempfile.NamedTemporaryFile(suffix=".oga", delete=False) as tmp:
                tmp.write(data)
                tmp_path = tmp.name
            text = transcribe(tmp_path)
        except SttError as e:
            self._send_with_retry(chat_id, str(e))
            return
        except Exception as e:
            self._send_with_retry(chat_id, f"Не получилось обработать голосовое сообщение: {e}")
            return
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.remove(tmp_path)

        # Показываем, что именно распознали -- голос мог разобраться неправильно, и
        # пользователю важно это видеть до того, как чат предложит действие или ответ.
        self._send_with_retry(chat_id, f"🎙 Распознал: «{html.escape(text, quote=False)}»", parse_mode="HTML")
        self._answer_chat_question(chat_id, text)

    def _send_with_retry(self, chat_id, text, parse_mode=None, attempts=3, delay_seconds=3):
        """Отправляет сообщение с повторными попытками при сетевых сбоях. Пользователь к
        этому моменту уже мог ждать ответ почти минуту (Ollama думает не мгновенно) -- терять
        готовый ответ из-за короткого сетевого сбоя (таймаут до api.telegram.org и т.п.)
        особенно неприятно: со стороны бота показалось "печатает", а реального ответа
        пользователь не увидит вообще, без единой попытки повтора. Возвращает True, если
        сообщение в итоге отправлено."""
        last_error = None
        for attempt in range(1, attempts + 1):
            try:
                self.tg.send_message(chat_id, text, parse_mode=parse_mode)
                return True
            except TelegramError as e:
                last_error = e
                if attempt < attempts:
                    time.sleep(delay_seconds)
        print(f"[bot] Не удалось отправить ответ чата после {attempts} попыток: {last_error}")
        return False

    # ---------- Обработка callback-кнопок ----------

    def handle_callback(self, callback_query):
        chat_id = callback_query["message"]["chat"]["id"]
        message_id = callback_query["message"]["message_id"]
        data = callback_query["data"]
        user = get_user_by_chat_id(chat_id)

        if not user:
            self.tg.answer_callback_query(callback_query["id"], "Аккаунт не привязан")
            return

        if data.startswith("clientpage:"):
            page = int(data.split(":")[1])
            self.tg.answer_callback_query(callback_query["id"])
            self._send_client_picker(chat_id, page=page, message_id=message_id)
            return

        if data.startswith("pickclient:"):
            client_name = data.split(":", 1)[1]
            self.tg.answer_callback_query(callback_query["id"])
            self._send_client_report(chat_id, client_name)
            return

        if data.startswith("status:"):
            _, obl_id, new_status = data.split(":")
            conn = get_conn()
            obl = conn.execute("SELECT * FROM obligations WHERE id=?", (obl_id,)).fetchone()
            can_edit = user["role"] == "owner" or (user["role"] == "team_lead" and obl and obl["responsible_name"] == user["team_lead_name"])
            if not obl or not can_edit:
                conn.close()
                self.tg.answer_callback_query(callback_query["id"], "Нет прав на изменение")
                return
            completed_at = "CURRENT_TIMESTAMP" if new_status == "done" else "NULL"
            conn.execute(f"UPDATE obligations SET status=?, updated_at=CURRENT_TIMESTAMP, completed_at={completed_at} WHERE id=?", (new_status, obl_id))
            conn.commit()
            obl2 = conn.execute("""
                SELECT o.*, f.client_name, f.project_name, f.project_num, f.amount_0_100 FROM obligations o
                JOIN fp_rows f ON f.id = o.fp_row_id WHERE o.id=?
            """, (obl_id,)).fetchone()
            conn.close()
            text, kb = self._obligation_text_and_kb(obl2)
            self.tg.edit_message_text(chat_id, message_id, text, reply_markup=kb)
            self.tg.answer_callback_query(callback_query["id"], f"Статус: {STATUS_LABELS.get(new_status)}")
            return

        if data.startswith("adddue:"):
            value = data.split(":", 1)[1]
            state = _add_flow_state.get(chat_id)
            if not state:
                self.tg.answer_callback_query(callback_query["id"])
                return
            self.tg.answer_callback_query(callback_query["id"])
            if value == "manual":
                state["step"] = "due_date_manual"
                self.tg.send_message(chat_id, "Введите дату в формате ГГГГ-ММ-ДД, или «нет» без срока:")
                return
            state["data"]["due_date"] = None if value == "none" else value
            self._ask_responsible(chat_id, state)
            return

        if data.startswith("addrow:"):
            row_id = int(data.split(":")[1])
            state = _add_flow_state.get(chat_id)
            if not state:
                self.tg.answer_callback_query(callback_query["id"])
                return
            state["data"]["row_id"] = row_id
            state["step"] = "title"
            self.tg.answer_callback_query(callback_query["id"])
            self.tg.send_message(chat_id, "Введите текст обязательства / этапа:")
            return

        if data.startswith("addresp:"):
            responsible_name = data.split(":", 1)[1]
            state = _add_flow_state.get(chat_id)
            if not state:
                self.tg.answer_callback_query(callback_query["id"])
                return
            state["data"]["responsible_name"] = responsible_name
            self.tg.answer_callback_query(callback_query["id"])
            self.finalize_add_obligation(chat_id, user)
            return

    # ---------- Главный цикл polling ----------

    def process_update(self, update):
        if "message" in update:
            msg = update["message"]
            chat_id = msg["chat"]["id"]
            text = (msg.get("text") or "").strip()

            if text.startswith("/start"):
                self.handle_start(chat_id)
            elif text.startswith("/help"):
                self.handle_help(chat_id)
            elif text.startswith("/link"):
                parts = text.split(maxsplit=1)
                code = parts[1] if len(parts) > 1 else ""
                self.handle_link(chat_id, code)
            elif text.startswith("/my"):
                self.handle_my(chat_id)
            elif text.startswith("/clientobl"):
                parts = text.split(maxsplit=1)
                query = parts[1] if len(parts) > 1 else ""
                self.handle_client_obligations(chat_id, query)
            elif text.startswith("/client"):
                parts = text.split(maxsplit=1)
                query = parts[1] if len(parts) > 1 else ""
                self.handle_client(chat_id, query)
            elif text.startswith("/add"):
                self.handle_add_start(chat_id)
            elif chat_id in _add_flow_state:
                user = get_user_by_chat_id(chat_id)
                self.handle_add_flow(chat_id, text, user)
            elif text.startswith("/"):
                self.tg.send_message(chat_id, "Не понял команду. /help — список команд")
            elif text:
                self.handle_chat_question(chat_id, text)
            elif "voice" in msg or "audio" in msg:
                self.handle_voice_message(chat_id, msg)

        elif "callback_query" in update:
            try:
                self.handle_callback(update["callback_query"])
            except Exception as e:
                print(f"[bot] Ошибка обработки callback: {e}")

    def run_polling(self):
        print("[bot] Запуск long polling...")
        me = self.tg.get_me()
        print(f"[bot] Авторизован как @{me.get('username')}")
        while not self._stop:
            try:
                updates = self.tg.get_updates(offset=self.offset, timeout=30)
                for update in updates:
                    self.offset = update["update_id"] + 1
                    try:
                        self.process_update(update)
                    except Exception as e:
                        print(f"[bot] Ошибка обработки update: {e}")
            except TelegramError as e:
                print(f"[bot] Ошибка Telegram API: {e}")
                time.sleep(5)
            except Exception as e:
                print(f"[bot] Неожиданная ошибка polling: {e}")
                time.sleep(5)

    def stop(self):
        self._stop = True

    # ---------- Напоминания (фоновый поток) ----------

    def check_reminders_loop(self):
        sent_today = set()  # (obligation_id, date, kind) чтобы не дублировать в течение дня
        while not self._stop:
            try:
                self._check_reminders_once(sent_today)
            except Exception as e:
                print(f"[bot] Ошибка проверки напоминаний: {e}")
            time.sleep(REMINDER_CHECK_INTERVAL_SECONDS)

    def _check_reminders_once(self, sent_today):
        conn = get_conn()
        today = today_iso()
        obls = conn.execute("""
            SELECT o.*, f.client_name, f.project_name, f.project_num FROM obligations o
            JOIN fp_rows f ON f.id = o.fp_row_id
            WHERE o.status != 'done' AND o.due_date IS NOT NULL
        """).fetchall()

        for o in obls:
            due = datetime.date.fromisoformat(o["due_date"])
            days_left = (due - datetime.date.today()).days
            kind = None
            if days_left < 0:
                kind = "overdue"
            elif days_left <= 3:
                kind = "due_soon"
            if not kind:
                continue

            key = (o["id"], today, kind)
            if key in sent_today:
                continue

            user = conn.execute(
                "SELECT * FROM users WHERE (role='team_lead' AND team_lead_name=?) OR (role='owner' AND ?=?) ",
                (o["responsible_name"], o["responsible_name"], OWNER_LABEL),
            ).fetchone()
            if not user or not user["telegram_chat_id"]:
                continue

            proj_label = f"{o['project_name']} (№ {o['project_num'] or '—'})"
            if kind == "overdue":
                text = f"⚠ Просрочено: «{o['title']}»\n{o['client_name']} · {proj_label}\nСрок был: {o['due_date']}"
            else:
                text = f"🔔 Горящий срок: «{o['title']}»\n{o['client_name']} · {proj_label}\nСрок: {o['due_date']} (осталось {days_left} дн.)"

            try:
                self.tg.send_message(user["telegram_chat_id"], text)
                sent_today.add(key)
            except TelegramError as e:
                print(f"[bot] Не удалось отправить напоминание: {e}")

        conn.close()


def main():
    import os
    token = os.environ.get("FP_BOT_TOKEN")
    if not token:
        token_file = Path(__file__).parent / "token.txt"
        if token_file.exists():
            token = token_file.read_text().strip()
    if not token:
        print("ОШИБКА: не найден токен бота.")
        print("Укажите его в переменной окружения FP_BOT_TOKEN или в файле bot/token.txt")
        sys.exit(1)

    bot = FPBot(token)
    reminder_thread = threading.Thread(target=bot.check_reminders_loop, daemon=True)
    reminder_thread.start()
    try:
        bot.run_polling()
    except KeyboardInterrupt:
        bot.stop()
        print("\n[bot] Остановлен")


if __name__ == "__main__":
    main()
