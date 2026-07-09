"""
Распознавание речи (Speech-to-Text) для голосовых сообщений в чате портала и в боте.

Работает локально, через faster-whisper (модель Whisper, которая выполняется на этом же
компьютере, без интернета и без ключей API) -- по тому же принципу, что и сам чат с
порталом отвечает через локальную Ollama, а не облачный ИИ.

Модель загружается лениво при первом голосовом сообщении и кэшируется в памяти процесса:
загрузка весов с диска занимает несколько секунд, и не имеет смысла делать это на каждое
сообщение. Если пакет faster-whisper не установлен (pip install faster-whisper) или модель
не удалось загрузить -- transcribe() поднимает SttError с понятным текстом, а не падает
с трудночитаемым traceback-ом; вызывающий код (main.py /chat/voice, bot/fp_bot.py) ловит
именно это исключение и показывает сообщение пользователю, не трогая остальной чат.
"""
import os
import threading

# Размер модели Whisper: tiny/base/small/medium/large-v3. "small" -- разумный баланс
# скорости и качества распознавания русской речи на CPU. Можно переопределить переменной
# окружения, например поставить "base" для скорости или "medium" для точности.
WHISPER_MODEL_SIZE = os.environ.get("FP_WHISPER_MODEL", "small")
WHISPER_DEVICE = os.environ.get("FP_WHISPER_DEVICE", "cpu")
WHISPER_COMPUTE_TYPE = os.environ.get("FP_WHISPER_COMPUTE_TYPE", "int8")

_model = None
_model_lock = threading.Lock()


class SttError(Exception):
    """Не удалось распознать голосовое сообщение (пакет/модель не установлены, файл
    повреждён, речь не разобрать и т.п.) -- текст исключения уже готов для показа
    пользователю как есть."""


def _get_model():
    global _model
    if _model is None:
        with _model_lock:
            if _model is None:
                try:
                    from faster_whisper import WhisperModel
                except ImportError as e:
                    raise SttError(
                        "Распознавание голоса не настроено на сервере: не установлен пакет "
                        "faster-whisper (pip install faster-whisper)."
                    ) from e
                try:
                    _model = WhisperModel(
                        WHISPER_MODEL_SIZE, device=WHISPER_DEVICE, compute_type=WHISPER_COMPUTE_TYPE,
                    )
                except Exception as e:
                    raise SttError(f"Не удалось загрузить модель распознавания речи: {e}") from e
    return _model


def transcribe(audio_path):
    """Распознаёт речь из аудиофайла (подходит почти любой формат -- ogg/opus из Telegram,
    webm/opus или mp4/aac из браузера, mp3, wav и т.п., декодирование берёт на себя
    faster-whisper) и возвращает обычный текст на русском.

    Поднимает SttError, если распознавание не удалось -- в этом случае вызывающий код
    должен показать текст ошибки пользователю и не пытаться передавать пустой вопрос
    дальше в answer_portal_query."""
    model = _get_model()
    try:
        segments, _info = model.transcribe(str(audio_path), language="ru", vad_filter=True)
        text = " ".join(seg.text.strip() for seg in segments).strip()
    except SttError:
        raise
    except Exception as e:
        raise SttError(f"Не удалось распознать голосовое сообщение: {e}") from e
    if not text:
        raise SttError("Не разобрал речь в сообщении -- похоже, оно пустое или слишком тихое. Попробуйте ещё раз.")
    return text
