from app.main import app
from app.db import init_db

# Шаблоны (.html) перечитываются на каждый запрос без перезапуска процесса.
app.config["TEMPLATES_AUTO_RELOAD"] = True

# Порт 5001 — порт 5000 занят ControlCenter (AirPlay Receiver) на macOS.
FP_PORT = 5001

if __name__ == "__main__":
    init_db()
    print("=" * 60)
    print(f"ФП-Контроль запущен: http://127.0.0.1:{FP_PORT}")
    print("=" * 60)
    app.run(host="127.0.0.1", port=FP_PORT, debug=False, use_reloader=True)
