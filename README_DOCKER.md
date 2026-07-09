# ФП-Контроль — запуск через Docker

## Что нужно установить

- [Docker Desktop](https://www.docker.com/products/docker-desktop/) — работает на macOS, Windows и Linux.  
  Это единственная зависимость. Python, SQLite и всё остальное уже внутри образа.

---

## Быстрый старт

### 1. Откройте терминал и перейдите в папку с проектом

```bash
cd путь/к/fp-portal
```

### 2. Соберите образ (один раз, занимает 3–5 минут)

```bash
docker compose build
```

### 3. Запустите

```bash
docker compose up -d
```

Портал будет доступен по адресу **http://localhost:5001**

### 4. Остановить

```bash
docker compose down
```

---

## Первый вход

При первом запуске база данных пустая. Войдите с помощью учётной записи, которую создаст руководитель (owner) через страницу **Пользователи**.

Если нужно создать первого пользователя-владельца вручную, выполните в терминале:

```bash
docker exec -it fp-portal python -c "
from app.db import get_conn
from werkzeug.security import generate_password_hash
conn = get_conn()
conn.execute(
    \"INSERT INTO users (username, full_name, password_hash, role) VALUES (?,?,?,?)\",
    ('admin', 'Администратор', generate_password_hash('admin123'), 'owner')
)
conn.commit()
print('Пользователь создан: логин admin, пароль admin123')
"
```

После входа сразу смените пароль в настройках.

---

## Данные

Все данные (база и загруженные файлы) хранятся в Docker-томе `fp_data`. Они **не теряются** при перезапуске или обновлении контейнера.

Посмотреть, где физически лежит том:
```bash
docker volume inspect fp-portal_fp_data
```

Сделать резервную копию базы:
```bash
docker cp fp-portal:/app/app/data/fp_portal.db ./backup_$(date +%Y%m%d).db
```

Восстановить из резервной копии:
```bash
docker cp ./backup_20260101.db fp-portal:/app/app/data/fp_portal.db
```

---

## Настройки (docker-compose.yml)

| Переменная | Описание |
|---|---|
| `FP_PORTAL_SECRET` | Секретный ключ сессий — замените на случайную строку |
| `OLLAMA_HOST` | Адрес Ollama для чата «Спросить портал» (опционально) |
| `OLLAMA_MODEL` | Модель Ollama (по умолчанию `qwen3:8b`) |

### Чат «Спросить портал»

Чат работает через [Ollama](https://ollama.com) — локальную нейросеть на вашем компьютере.  
Если Ollama не установлена — чат просто не отвечает, всё остальное работает нормально.

Чтобы включить: установите Ollama, загрузите модель (`ollama pull qwen3:8b`), затем перезапустите портал.

---

## Обновление

```bash
# Пересобрать образ с новым кодом
docker compose build

# Перезапустить (данные сохранятся)
docker compose up -d
```

---

## Порт

По умолчанию портал слушает на **5001**. Если этот порт занят, измените в `docker-compose.yml`:

```yaml
ports:
  - "8080:5001"   # будет доступен на http://localhost:8080
```

---

## Проблемы

**Порт занят:**
```bash
docker compose down
# Измените порт в docker-compose.yml и запустите снова
docker compose up -d
```

**Посмотреть логи:**
```bash
docker compose logs -f
```

**Войти внутрь контейнера:**
```bash
docker exec -it fp-portal sh
```
