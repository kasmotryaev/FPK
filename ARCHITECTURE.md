# fp-portal — Архитектура и зависимости продукта

> **OpenSpec-контекст для агентов.**  
> Документ описывает сценарии использования, архитектурные компоненты, карту зависимостей
> и схему данных продукта **fp-portal** — внутреннего веб-портала мониторинга финансового
> плана и трудозатрат продуктового подразделения. Предназначен для чтения агентами (Claude,
> другими LLM) перед изменениями кода: он содержит архитектурные ограничения, которые
> нельзя нарушать.

---

## 1. Краткое описание продукта

**fp-portal** — Flask-приложение, работающее локально на macOS-хосте (или в Docker).
Позволяет руководителю продукта и лидам команд:

- отслеживать финансовый план (ФП) по кварталам и финансовым годам;
- управлять обязательствами по статьям ФП;
- анализировать трудозатраты (ТЗ) в разрезе периода / сотрудника / проекта;
- сравнивать фактические трудозатраты с плановым ФП;
- задавать вопросы данным через чат (локальный LLM + fallback-поиск без сети).

---

## 2. Сценарии использования

### UC-01 · Загрузка финансового плана

**Актор:** `owner`  
**Триггер:** получена новая выгрузка из СЦ (Excel `.xlsx`).  
**Путь:**  
1. Открывает `/import`, выбирает квартал и файл.  
2. `import_excel()` (`app/importer.py`) читает первый лист через `openpyxl`.  
3. Для каждой строки вычисляется `row_key` (SHA-256 по 10 полям включая `quarter_label`).  
4. Строки сравниваются с БД: новые — `INSERT`, изменённые — `UPDATE`, пропавшие —
   `is_active=0`.  
5. Дельта записывается в `import_log` (JSON) и `row_events` (постоянный журнал).  
6. При первой загрузке нового квартала выполняется rollover: обязательства и примечания
   переносятся из предыдущего квартала по совпадению `row_key`.

**Затронутые таблицы:** `fp_rows`, `import_log`, `row_events`, `obligations`,
`app_settings`.  
**Библиотеки:** `openpyxl`.

---

### UC-02 · Дашборд квартала

**Актор:** все роли.  
**Путь:**  
1. `/` → `dashboard()`: читает `fp_rows` для активного периода (месяцы текущего квартала).  
2. Агрегирует суммы по портфелям (Факт / 0-100 / Возможности) и разделам ФП.  
3. Блок **«К получению за 4 дня»**: строки, где `dpa_date` попадает в `[today, today+4]`.  
4. Блок **«Под риском»**: строки с `risk_level ∈ {1,2,3}`.  
5. Блок **«Просрочки / горящие сроки»**: обязательства с `due_date < today` или
   `due_date ≤ today+3`.  
6. Чат «Спросить портал»: AJAX POST `/chat` → `answer_portal_query()` → Ollama или
   fallback.

**Затронутые таблицы:** `fp_rows`, `obligations`, `quarter_targets`.  
**Внешние системы:** Ollama (опционально).

---

### UC-03 · Детализация статей ФП

**Актор:** все роли.  
**Путь:**  
1. `/rows` → `rows_list()`: читает `fp_rows` по активному периоду.  
2. Аккордеон по клиентам; для каждой строки — бейдж числа обязательств (COUNT из БД).  
3. Первый клик на бейдж → AJAX GET `/rows/<id>/obl-panel` → фрагмент с формами
   обязательств (lazy load).  
4. Inline-действия: изменить комментарий (`POST /rows/<id>/comment`), уровень риска
   (`POST /rows/<id>/risk`), добавить обязательство (`POST /rows/<id>/obligations/add`).

**Затронутые таблицы:** `fp_rows`, `obligations`, `obligation_history`, `saved_filters`.

---

### UC-04 · Управление обязательствами

**Актор:** `owner`, `team_lead`.  
**Путь:**  
1. Добавить: `POST /rows/<id>/obligations/add` → `INSERT INTO obligations`.  
2. Сменить статус: `POST /obligations/<id>/status` → `UPDATE obligations.status`,
   запись в `obligation_history`.  
3. Переназначить: `POST /obligations/<id>/reassign`.  
4. Удалить: `POST /obligations/<id>/delete`.  
5. Мои обязательства: `/my` → список своих открытых обязательств.

**Затронутые таблицы:** `obligations`, `obligation_history`.

---

### UC-05 · Журнал изменений ФП

**Актор:** `owner`.  
**Путь:**  
1. `/changes` → список событий из `row_events`, сгруппированных по типу (новые /
   закрытые / прочие).  
2. Отметить просмотренным: `POST /changes/<id>/review` → `reviewed_at`, `reviewed_by`.  
3. Отметить все: `POST /changes/review_all`.

**Затронутые таблицы:** `row_events`, `import_log`.

---

### UC-06 · Загрузка трудозатрат

**Актор:** `owner`.  
**Путь:**  
1. `/import` → форма загрузки файла ТЗ (.xlsx / .xlsb).  
2. `parse_ts_file()` (`app/ts_parser.py`) авто-определяет заголовки и формат
   (`file_type='office'` или `'detail'`).  
3. `.xlsb` → конвертация через `libreoffice` в subprocess → повторный разбор.  
4. `INSERT INTO ts_imports` + `INSERT INTO ts_rows` (batch).  
5. Фильтры РП и ПЦ — из `app_settings` (ключи `ts_rp_filter`, `ts_pc_filter`).

**Затронутые таблицы:** `ts_imports`, `ts_rows`, `app_settings`.  
**Библиотеки:** `openpyxl`, `pyxlsb` (опционально), `subprocess(libreoffice)`.

---

### UC-07 · Анализ трудозатрат (Трудозатраты)

**Актор:** все роли.  
**Путь:**  
1. `/timesheets` → выбор периода (импорта), иерархия `dept → division → employee → tasks`.  
2. Карточка сотрудника `/timesheets/employee?import_id=…&name=…`.  
3. Аналитика по периодам `/timesheets/analytics`.  
4. Diff между двумя импортами `/timesheets/diff`.  
5. Выгрузка в CSV `/timesheets/export`.

**Затронутые таблицы:** `ts_imports`, `ts_rows`, `my_employees`.

---

### UC-08 · Расходы vs Доходы (сравнение ТЗ с ФП)

**Актор:** `owner`, `team_lead`.  
**Путь:**  
1. `/timesheets/compare` → выбор импорта ТЗ и квартала ФП.  
2. `_ts_build_compare()` → для каждого проекта ТЗ ищет соответствие в `fp_rows`:
   — по числовому номеру в начале имени проекта (`_RE_PROJ_NUM`),
   — по точному имени,
   — по нечёткому совпадению клиента (Jaccard по токенам, порог 0.10).  
3. Иерархия: `dept → project (суммы ФП) → employees (задачи)`.  
4. Три блока: все / без «моих» сотрудников / «мои» меньше чужих.  
5. Ставка `₽/день` — из `app_settings` (`ts_daily_rate`).

**Затронутые таблицы:** `ts_rows`, `ts_imports`, `fp_rows`, `my_employees`, `app_settings`.

---

### UC-09 · ИИ-чат и голосовые запросы

**Актор:** все роли.  
**Путь:**  
1. Текстовый вопрос → AJAX POST `/chat` → `answer_portal_query()`:
   - `_build_chat_context()` формирует текстовый срез данных из БД;
   - `_call_ollama()` отправляет контекст + вопрос в Ollama REST API;
   - Fallback: `_answer_portal_query_local()` — поиск по ключевым словам без сети.
2. Действия через чат (добавить обязательство, изменить риск): двухшаговый диалог
   (показать план → `pending_chat_action` в Flask session → подтвердить «да»).  
3. Голосовой вопрос → POST `/chat/voice` → `transcribe()` (faster-whisper, CPU) →
   тот же `answer_portal_query()`.

**Внешние системы:** Ollama (localhost:11434, модель `qwen3:8b`).  
**Библиотеки:** `faster-whisper`, `ffmpeg`.

---

### UC-10 · Telegram-бот

**Актор:** привязанные пользователи портала.  
**Путь:**  
1. Отдельный процесс (`python3 run_bot.py`), читает из той же SQLite.  
2. Polling через `bot/telegram_client.py` (голый urllib, без сторонних SDK).  
3. Команды: `/link КОД` (привязка аккаунта), `/my` (мои обязательства),
   `/client` (поиск клиента кнопками), `/add` (добавить обязательство диалогом).  
4. Любой текст без `/` → `answer_portal_query()` (тот же ИИ-чат).  
5. Голосовые сообщения → `transcribe()` → `answer_portal_query()`.  
6. Фоновый поток (каждые 15 мин): проверяет дедлайны обязательств, шлёт напоминания.

**Внешние системы:** Telegram Bot API (polling), Ollama.  
**Библиотеки:** `faster-whisper`, `ffmpeg`.  
**Ограничение:** `bot/token.txt` — секрет, не должен попадать в git.

---

### UC-11 · Управление кварталами

**Актор:** `owner`.  
**Путь:**  
1. Переключить вид: `POST /set-quarter-view` (сессия `view_quarter`) или выбрать
   финансовый год (объединяет 4 квартала).  
2. Зафиксировать квартал: `POST /fix-quarter` → `app_settings.finalized_quarters`.  
3. Снять фиксацию: `POST /unfix-quarter`.  
4. Перейти на новый квартал: загрузить файл с `upload_quarter=следующий` → rollover.  
5. Архив: `/archive` → список всех финализированных кварталов с резюме.

**Затронутые таблицы:** `fp_rows`, `app_settings`, `import_log`.

---

### UC-12 · Управление пользователями

**Актор:** `owner`.  
**Путь:**  
1. `/admin/users` → создать / изменить / удалить пользователей.  
2. Роли: `owner` (полный доступ), `team_lead` (управление своими обязательствами),
   `viewer` (только просмотр).  
3. Telegram-привязка: `POST /admin/users/<id>/telegram_code` → генерирует код-ссылку;
   пользователь вводит `/link КОД` в боте.

**Затронутые таблицы:** `users`.

---

## 3. Архитектурные компоненты

### 3.1 Веб-приложение (`app/`)

| Компонент | Файл | Назначение |
|---|---|---|
| Flask app | `app/main.py` | 45+ маршрутов, бизнес-логика, сессии |
| БД-слой | `app/db.py` | `get_conn()`, `init_db()`, миграции, константы |
| Импортёр ФП | `app/importer.py` | Разбор Excel-выгрузки, diffing, rollover |
| Парсер ТЗ | `app/ts_parser.py` | Плоский Excel/xlsb, авто-поиск заголовков |
| STT | `app/stt.py` | faster-whisper, ленивая загрузка модели |
| Шаблоны | `app/templates/` | Jinja2, 16 HTML-файлов |
| Статика | `app/static/` | CSS, версионирование через `?v=mtime` |

### 3.2 Telegram-бот (`bot/`)

| Компонент | Файл | Назначение |
|---|---|---|
| Бот | `bot/fp_bot.py` | Команды, диалоги, polling, reminder-поток |
| HTTP-клиент | `bot/telegram_client.py` | Обёртка над Telegram Bot API (urllib) |

### 3.3 Python-библиотеки

| Библиотека | Версия | Зачем |
|---|---|---|
| Flask | ≥3.0, <4.0 | Web-фреймворк |
| openpyxl | ≥3.1, <4.0 | Чтение/запись `.xlsx` |
| faster-whisper | ≥1.0 | STT (Speech-to-Text, Whisper-модели) |
| pyxlsb | —¹ | Чтение `.xlsb` (устанавливается при деплое) |
| gunicorn | —¹ | Production WSGI-сервер (Docker) |
| werkzeug | транзитив | Безопасность паролей, `secure_filename` |

¹ — не в `requirements.txt`, устанавливается через `Dockerfile`.

### 3.4 Системные зависимости (runtime)

| Инструмент | Назначение |
|---|---|
| ffmpeg | Декодирование аудио для faster-whisper |
| libreoffice | Конвертация `.xlsb` → `.xlsx` (subprocess) |
| SQLite 3 | Встроенная в Python, WAL-режим |

### 3.5 Внешние сервисы

| Сервис | Протокол | Обязательность |
|---|---|---|
| **Ollama** (localhost:11434) | REST JSON | Опционально; fallback без него |
| **Telegram Bot API** | HTTPS polling | Только для бота; портал работает без него |

### 3.6 Развёртывание

- **macOS локально:** `python3 run.py` (Flask dev-сервер, порт 5001).
- **Docker:** `docker-compose up` → gunicorn + volume `fp_data` для `app/data/`.
- **БД:** `app/data/fp_portal.db` (или в Docker volume). Синхронизируется через iCloud Drive
  (retry при `authorization denied`).

---

## 4. Карта зависимостей продукта

### 4.1 Upstream — на чём зависит fp-portal

```
СЦ (Система СЦ)
  └─ выгружает Excel «Факт + Прогноз по РП»
        └─ загружается в /import → fp_rows

ПО (Производственный отдел / HR-система)
  └─ выгружает Excel «Отчёт распределение ресурсов»
        └─ загружается в /timesheets/import → ts_rows

Ollama (localhost)
  └─ модель qwen3:8b (или другая)
        └─ отвечает на вопросы чата (/chat, /chat/voice, Telegram)

Telegram Bot API (api.telegram.org)
  └─ доставляет сообщения боту и отправляет ответы пользователям
```

### 4.2 Downstream — что зависит от fp-portal

На данный момент **нет downstream систем**: fp-portal является конечной точкой сбора
и отображения данных. Экспорт — только вручную через `/timesheets/export` (CSV).

### 4.3 Внутренние события (event flow)

| Событие | Источник | Потребитель |
|---|---|---|
| `row_events` (new/zeroed/…) | `import_excel()` при загрузке ФП | Страница «Изменения» `/changes` |
| `import_log` | `import_excel()` | Страница «Настройки» `/import`, diff |
| `obligation_history` | смена статуса / переназначение | (аудит, пока только в БД) |
| Reminder (pending deadlines) | Фоновый поток бота (15 мин) | Telegram-сообщения пользователям |
| `pending_chat_action` | `answer_portal_query()` | Flask session → подтверждение действия |

### 4.4 API-контракты

#### Внутренний REST (браузер ↔ портал)

| Метод | Путь | Назначение |
|---|---|---|
| GET | `/` | Дашборд |
| GET | `/rows` | Детализация статей ФП |
| GET | `/rows/<id>/obl-panel` | Ленивая подгрузка панели обязательств |
| POST | `/rows/<id>/comment` | Сохранить комментарий |
| POST | `/rows/<id>/risk` | Установить уровень риска |
| POST | `/rows/<id>/obligations/add` | Добавить обязательство |
| POST | `/obligations/<id>/status` | Сменить статус обязательства |
| POST | `/obligations/<id>/reassign` | Переназначить |
| POST | `/obligations/<id>/delete` | Удалить |
| POST | `/chat` | Текстовый вопрос (JSON ответ) |
| POST | `/chat/voice` | Голосовой вопрос (multipart audio) |
| POST | `/import` | Загрузить файл ФП |
| POST | `/timesheets/import` | Загрузить файл ТЗ |
| GET | `/timesheets/compare` | Расходы vs Доходы |
| GET | `/timesheets/export` | Выгрузка ТЗ в CSV |
| GET | `/changes` | Журнал изменений ФП |
| POST | `/fix-quarter` | Зафиксировать квартал |
| POST | `/set-quarter-view` | Переключить видимый квартал |

#### Ollama REST (портал → Ollama)

```
POST http://localhost:11434/api/chat
Body: { "model": "qwen3:8b", "messages": [...], "stream": false,
        "think": false, "options": {"temperature": 0.2, "num_ctx": 8192} }
```

#### Telegram Bot API (бот → Telegram)

- Polling: `GET /bot{TOKEN}/getUpdates?offset=…&timeout=30`
- Отправка: `POST /bot{TOKEN}/sendMessage`
- Inline-кнопки: `POST /bot{TOKEN}/answerCallbackQuery`

---

## 5. Схема данных

### 5.1 Сущности и ключевые поля

#### `users` — пользователи портала

| Поле | Тип | Описание |
|---|---|---|
| `id` | INTEGER PK | |
| `username` | TEXT UNIQUE | Логин |
| `full_name` | TEXT | Отображаемое имя |
| `password_hash` | TEXT | pbkdf2:sha256 (werkzeug) |
| `role` | TEXT | `owner` / `team_lead` / `viewer` |
| `team_lead_name` | TEXT | Для role=team_lead: имя как в ФП |
| `telegram_chat_id` | TEXT | Привязанный Telegram-чат |
| `telegram_link_code` | TEXT | Временный код привязки |

#### `fp_rows` — строки финансового плана

| Поле | Тип | Описание |
|---|---|---|
| `id` | INTEGER PK | |
| `row_key` | TEXT UNIQUE | SHA-256 по 10 полям + quarter_label |
| `quarter_label` | TEXT | Формат `2026-Q3` |
| `month` | TEXT | Русское название месяца |
| `pc` | TEXT | Продуктовый центр |
| `section` | TEXT | Раздел ФП (Проекты / Заказные доработки / Лицензии / Докупки / Сопровождение / TaaS) |
| `client_name` | TEXT | Наименование клиента |
| `project_num` | TEXT | Номер проекта |
| `project_name` | TEXT | Наименование проекта |
| `project_manager` | TEXT | Менеджер проекта |
| `contract_num` | TEXT | Номер договора |
| `ez_num` | TEXT | Номер ЭЗ |
| `sdz_date` | TEXT | Дата СДЗ (ISO-8601) |
| `dpa_date` | TEXT | Дата ДПА (ISO-8601) — ключевая для «К получению» |
| `fdz_date` | TEXT | Дата фДЗ |
| `amount_0_100` | REAL | Сумма по 0-100, руб. |
| `crm_amount` | REAL | Сумма из CRM |
| `portfolio` | TEXT | `Факт` / `0-100` / `Возможности` |
| `is_active` | INTEGER | 1 = активна, 0 = деактивирована |
| `risk_level` | INTEGER | 0 = нет риска, 1/2/3 = Низкий/Средний/Высокий |
| `responsible_user_id` | INTEGER FK→users | Ответственный пользователь |
| `internal_comment` | TEXT | Внутреннее примечание |
| `last_import_id` | INTEGER | Последний импорт, в котором строка видна |

**Индексы:** `(is_active, quarter_label, month)` — покрывает основной WHERE на всех страницах.

#### `obligations` — обязательства по статьям ФП

| Поле | Тип | Описание |
|---|---|---|
| `id` | INTEGER PK | |
| `fp_row_id` | INTEGER FK→fp_rows | ON DELETE CASCADE |
| `title` | TEXT | Краткое название |
| `description` | TEXT | Подробности |
| `responsible_type` | TEXT | `owner` / `team_lead` |
| `responsible_name` | TEXT | Имя ответственного (из ALL_RESPONSIBLE) |
| `due_date` | TEXT | Срок (ISO-8601) |
| `status` | TEXT | `not_started` / `in_progress` / `done` / `blocked` |
| `created_by` | INTEGER FK→users | |
| `completed_at` | TEXT | Дата выполнения |

#### `obligation_history` — история изменений обязательств

| Поле | Тип | Описание |
|---|---|---|
| `obligation_id` | INTEGER FK→obligations | ON DELETE CASCADE |
| `user_id` | INTEGER FK→users | Кто изменил |
| `action` | TEXT | Описание действия |
| `details` | TEXT | JSON с деталями |

#### `import_log` — журнал загрузок ФП

| Поле | Тип | Описание |
|---|---|---|
| `id` | INTEGER PK | |
| `filename` | TEXT | Имя файла |
| `rows_total/new/updated/deactivated` | INTEGER | Статистика |
| `imported_by` | INTEGER FK→users | |
| `diff_json` | TEXT | JSON-срез изменений (для `/import?last=1`) |

#### `row_events` — постоянный журнал событий по строкам ФП

| Поле | Тип | Описание |
|---|---|---|
| `fp_row_id` | INTEGER FK→fp_rows | |
| `import_log_id` | INTEGER FK→import_log | |
| `event_type` | TEXT | `new` / `zeroed` / `deactivated` / `reactivated` / `amount_changed` / `field_changed` |
| `old_value / new_value` | TEXT | Старое и новое значение поля |
| `amount_before / amount_after` | REAL | Суммы до и после |
| `reviewed_at / reviewed_by` | TEXT/INT | Отметка просмотра |

#### `quarter_targets` — плановые цели по кварталам

| Поле | Тип | Описание |
|---|---|---|
| `period_label` | TEXT UNIQUE | Квартал (`2026-Q3`) |
| `target_amount` | REAL | Целевая сумма, руб. |

#### `app_settings` — ключ-значение настроек

| Ключ | Назначение |
|---|---|
| `current_quarter_label` | Текущий активный квартал |
| `finalized_quarters` | CSV-список зафиксированных кварталов |
| `ts_rp_filter` | Фильтр по РП для трудозатрат |
| `ts_pc_filter` | Фильтр по ПЦ для трудозатрат |
| `ts_daily_rate` | Ставка ₽/день для расчёта стоимости ТЗ |
| `row_key_v2` | Флаг миграции row_key |

#### `ts_imports` — загрузки файлов трудозатрат

| Поле | Тип | Описание |
|---|---|---|
| `id` | INTEGER PK | |
| `filename` | TEXT | Имя файла |
| `period_label` | TEXT | Период (метка месяца/квартала) |
| `rp_filter / pc_filter` | TEXT | Фильтры РП/ПЦ на момент загрузки |
| `file_type` | TEXT | `office` (сводный) / `detail` (до задачи) |

#### `ts_rows` — строки трудозатрат (плоская структура)

| Поле | Тип | Описание |
|---|---|---|
| `import_id` | INTEGER FK→ts_imports | ON DELETE CASCADE |
| `rp` | TEXT | Руководитель проекта |
| `rp_product` | TEXT | РП продукта (при `file_type=detail`) |
| `dept` | TEXT | Департамент |
| `division` | TEXT | Управление |
| `employee` | TEXT | Сотрудник |
| `project` | TEXT | Проект (может содержать номер в начале) |
| `task` | TEXT | Задача |
| `client` | TEXT | Клиент |
| `hours` | REAL | Трудозатраты, ч |

#### `my_employees` — справочник «моих» сотрудников

| Поле | Тип | Описание |
|---|---|---|
| `full_name` | TEXT UNIQUE | Полное имя сотрудника |

#### `saved_filters` — сохранённые фильтры пользователей

| Поле | Тип | Описание |
|---|---|---|
| `user_id` | INTEGER FK→users | ON DELETE CASCADE |
| `name` | TEXT | Название набора |
| `query_string` | TEXT | URL query string с параметрами фильтра |

### 5.2 Связи между сущностями

```
users ──────────────────────────────────────────────────────────────────┐
  │ responsible_user_id                                                  │
  ▼                                                                      │
fp_rows ──── obligations ──── obligation_history ──── users (user_id)   │
  │               │                                                      │
  │         created_by ──────────────────────────────────────────────────┘
  │
  ├── import_log ──── row_events ──── users (reviewed_by)
  │      │
  │   imported_by ─── users
  │
  └── row_events (fp_row_id)

ts_imports ──── ts_rows
     │
imported_by ─── users

users ──── saved_filters
users ──── quarter_targets (set_by)
```

### 5.3 Ограничения целостности

- `fp_rows.row_key` — UNIQUE, пересчитывается при миграции (флаг `row_key_v2`).
- `obligations` при удалении `fp_rows` — CASCADE DELETE.
- `obligation_history` при удалении `obligations` — CASCADE DELETE.
- `row_events` при удалении `fp_rows` — CASCADE DELETE.
- `ts_rows` при удалении `ts_imports` — CASCADE DELETE.
- `saved_filters` при удалении `users` — CASCADE DELETE.
- `PRAGMA foreign_keys = ON` — включается при каждом соединении.

---

## 6. Архитектурные ограничения для агентов

1. **Единый файл БД.** Портал и бот используют одну SQLite-базу. Нельзя разделять или
   переименовывать `fp_portal.db` без синхронного обновления обоих процессов.

2. **WAL-режим обязателен.** `PRAGMA journal_mode = WAL` устанавливается при каждом
   `get_conn()`. Бот читает из той же базы — WAL позволяет параллельное чтение без
   блокировки записей портала.

3. **iCloud-retry.** `get_conn()` делает 5 попыток с паузой 0.4 с при `authorization denied`
   (iCloud Drive фоновая синхронизация). Это поведение нельзя убирать.

4. **row_key — контрактный ключ.** Функция `row_key()` в `db.py` является контрактом
   идентификации строк ФП между загрузками. Изменение логики хэша требует миграции
   (как `row_key_v2`). Нельзя менять набор или порядок полей без явной миграции.

5. **Lazy-load обязательств.** Панель обязательств загружается AJAX-запросом
   (`/rows/<id>/obl-panel`), а не на старте страницы. `rows_list()` возвращает только
   `{row_id: count}`, не полные данные обязательств. Это ключевое для производительности
   при >200 строк.

6. **Ollama — опциональная зависимость.** Чат обязан работать без Ollama через
   `_answer_portal_query_local()`. При добавлении новых типов вопросов нужно обновить
   оба пути (Ollama и fallback).

7. **bot/token.txt — секрет.** Файл не должен попадать в git (включён в `.gitignore`).
   Бот читает токен из этого файла при старте. Нельзя логировать содержимое файла.

8. **Финансовый год ≠ календарный.** ФГ начинается 1 апреля. `calendar_to_fiscal()` —
   единственный источник истины для конвертации. Квартальные метки хранятся в
   календарном формате (`2026-Q3`), финансовая группировка вычисляется на лету.

9. **Деактивация строк — не удаление.** Строки ФП никогда не удаляются, только
   `is_active = 0`. Это позволяет хранить историю и rollover обязательств между
   кварталами.

10. **Роли и доступ.** `owner_required` / `login_required` — декораторы Flask. Маршруты
    записи (import, admin, set-quarter) требуют `owner`. Бот проверяет `telegram_chat_id`
    в таблице `users` — непривязанные пользователи получают только публичный ответ чата.

---

*Сгенерировано автоматически на основе кодовой базы `fp-portal` (commit: июль 2026).*
*При изменении архитектуры — обновить этот документ вместе с кодом.*
