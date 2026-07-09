"""
Минимальный клиент Telegram Bot API на стандартной библиотеке Python.
Не требует pip install requests / python-telegram-bot — работает где угодно,
где есть доступ в интернет к api.telegram.org.

Прокси для обхода блокировок (РФ и т.п.). api.telegram.org часто недоступен напрямую --
задайте переменную окружения FP_BOT_PROXY, и именно ЭТОТ клиент (только обращения к
Telegram) пойдёт через прокси. Локальные вызовы внутри портала/бота (например к Ollama
на localhost) НЕ затрагиваются -- прокси не патчит сокеты глобально на весь процесс.

Поддерживаются:
  FP_BOT_PROXY=http://host:port            (или https://host:port -- для HTTP(S)-прокси)
  FP_BOT_PROXY=http://user:pass@host:port  (с авторизацией)
  FP_BOT_PROXY=socks5://host:port          (для SOCKS5 -- нужен пакет PySocks: pip install pysocks)
  FP_BOT_PROXY=socks5://user:pass@host:port
"""
import json
import os
import ssl
import http.client
import urllib.request
import urllib.parse
import urllib.error


class TelegramError(Exception):
    pass


def _build_opener(proxy_url):
    """Строит urllib-opener, который ходит через прокси из FP_BOT_PROXY (если задана),
    либо обычный opener (эквивалент urllib.request.urlopen) если переменная не задана."""
    if not proxy_url:
        return urllib.request.build_opener()

    scheme = urllib.parse.urlsplit(proxy_url).scheme.lower()

    if scheme in ("http", "https"):
        return urllib.request.build_opener(
            urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
        )

    if scheme in ("socks5", "socks5h", "socks4"):
        try:
            import socks  # PySocks -- опционален, нужен только для socks5/socks4-прокси
        except ImportError as e:
            raise TelegramError(
                "FP_BOT_PROXY указывает на SOCKS-прокси, но не установлен пакет PySocks "
                "(pip install pysocks)."
            ) from e

        parts = urllib.parse.urlsplit(proxy_url)
        proxy_type = socks.SOCKS4 if scheme == "socks4" else socks.SOCKS5
        proxy_host, proxy_port = parts.hostname, parts.port or 1080
        proxy_user, proxy_pass = parts.username, parts.password

        def _connect_via_socks(conn):
            sock = socks.socksocket()
            sock.set_proxy(proxy_type, proxy_host, proxy_port, username=proxy_user, password=proxy_pass)
            sock.settimeout(conn.timeout)
            sock.connect((conn.host, conn.port))
            return sock

        class SocksHTTPConnection(http.client.HTTPConnection):
            def connect(self):
                self.sock = _connect_via_socks(self)

        class SocksHTTPSConnection(http.client.HTTPSConnection):
            def connect(self):
                sock = _connect_via_socks(self)
                context = self._context or ssl.create_default_context()
                self.sock = context.wrap_socket(sock, server_hostname=self.host)

        class SocksHTTPHandler(urllib.request.HTTPHandler):
            def http_open(self, req):
                return self.do_open(SocksHTTPConnection, req)

        class SocksHTTPSHandler(urllib.request.HTTPSHandler):
            def https_open(self, req):
                return self.do_open(SocksHTTPSConnection, req)

        return urllib.request.build_opener(SocksHTTPHandler(), SocksHTTPSHandler())

    raise TelegramError(
        f"FP_BOT_PROXY: неизвестная схема прокси «{scheme}» -- поддерживаются http(s):// и socks5://"
    )


class TelegramClient:
    def __init__(self, token, proxy=None):
        self.token = token
        self.base_url = f"https://api.telegram.org/bot{token}/"
        # Скачивание самого файла (голосовые и т.п.) идёт не через base_url, а через
        # отдельный "файловый" базовый URL -- так устроен Bot API.
        self.file_base_url = f"https://api.telegram.org/file/bot{token}/"
        # proxy: явный параметр для тестов/особых случаев, по умолчанию -- из окружения.
        self.proxy_url = proxy if proxy is not None else os.environ.get("FP_BOT_PROXY") or None
        self._opener = _build_opener(self.proxy_url)

    def _call(self, method, params=None, timeout=35):
        url = self.base_url + method
        data = json.dumps(params or {}).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        try:
            with self._opener.open(req, timeout=timeout) as resp:
                body = json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            body = json.loads(e.read().decode("utf-8"))
        except urllib.error.URLError as e:
            raise TelegramError(f"Сетевая ошибка при обращении к Telegram: {e}")
        if not body.get("ok"):
            raise TelegramError(f"Telegram API error: {body}")
        return body["result"]

    def get_me(self):
        return self._call("getMe")

    def get_updates(self, offset=None, timeout=30):
        params = {"timeout": timeout}
        if offset is not None:
            params["offset"] = offset
        return self._call("getUpdates", params, timeout=timeout + 10)

    def send_message(self, chat_id, text, reply_markup=None, parse_mode="HTML"):
        params = {"chat_id": chat_id, "text": text}
        # Telegram отвечает 400 Bad Request, если parse_mode передан явным null в JSON --
        # значение должно быть либо валидной строкой ("HTML"/"MarkdownV2"), либо параметр
        # должен отсутствовать вовсе (тогда сообщение шлётся как обычный текст без разметки).
        if parse_mode:
            params["parse_mode"] = parse_mode
        if reply_markup:
            params["reply_markup"] = reply_markup
        return self._call("sendMessage", params)

    def send_chat_action(self, chat_id, action="typing"):
        return self._call("sendChatAction", {"chat_id": chat_id, "action": action})

    def answer_callback_query(self, callback_query_id, text=None):
        params = {"callback_query_id": callback_query_id}
        if text:
            params["text"] = text
        return self._call("answerCallbackQuery", params)

    def edit_message_text(self, chat_id, message_id, text, reply_markup=None, parse_mode="HTML"):
        params = {"chat_id": chat_id, "message_id": message_id, "text": text}
        if parse_mode:
            params["parse_mode"] = parse_mode
        if reply_markup:
            params["reply_markup"] = reply_markup
        return self._call("editMessageText", params)

    def get_file(self, file_id):
        """Возвращает метаданные файла, включая file_path -- относительный путь, который
        нужно передать в download_file()."""
        return self._call("getFile", {"file_id": file_id})

    def download_file(self, file_path, timeout=35):
        """Скачивает содержимое файла (например голосового сообщения) и возвращает байты.
        file_path берётся из ответа get_file()."""
        url = self.file_base_url + file_path
        try:
            with self._opener.open(url, timeout=timeout) as resp:
                return resp.read()
        except urllib.error.URLError as e:
            raise TelegramError(f"Не удалось скачать файл из Telegram: {e}")


def inline_keyboard(rows):
    """rows: список списков {"text": ..., "callback_data": ...}"""
    return {"inline_keyboard": rows}
