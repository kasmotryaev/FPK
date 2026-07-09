"""
Запуск Telegram-бота ФП-Контроль.

Использование:
  python3 run_bot.py

Токен бота берётся из переменной окружения FP_BOT_TOKEN,
либо из файла bot/token.txt (создайте его и вставьте туда токен от @BotFather).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from bot.fp_bot import main

if __name__ == "__main__":
    main()
