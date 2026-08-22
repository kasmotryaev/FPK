from app.main import app, migrate_target_keys
from app.db import get_conn, init_db

# Шаблоны (.html) перечитываются на каждый запрос без перезапуска процесса.
app.config["TEMPLATES_AUTO_RELOAD"] = True

# Порт 5001 — порт 5000 занят ControlCenter (AirPlay Receiver) на macOS.
FP_PORT = 5001

if __name__ == "__main__":
    init_db()
    # Разовый перенос целей со старого ключа-перечня месяцев на метку квартала/финансового года.
    _conn = get_conn()
    _moved = migrate_target_keys(_conn)
    _conn.close()
    if _moved:
        print(f"Целей перенесено на новый ключ периода: {_moved}")
    print("=" * 60)
    print(f"ФП-Контроль запущен: http://127.0.0.1:{FP_PORT}")
    print("=" * 60)
    app.run(host="127.0.0.1", port=FP_PORT, debug=False, use_reloader=True)
