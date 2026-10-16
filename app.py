# -*- coding: utf-8 -*-
"""
Журнал регистрации приказов по видам деятельности
Веб-приложение (Flask + SQLite) с ролевым доступом, загрузкой PDF-сканов,
отчётами (Excel/Word/печать), архивированием и восстановлением БД.
"""
import os
import io
import re
import sys
import ast
import json
import hashlib
import zlib
import shutil
import socket
import sqlite3
import logging
import tempfile
import threading
import time
import traceback
from logging.handlers import RotatingFileHandler
from datetime import date, datetime

from flask import (
    Flask, render_template, request, redirect, url_for, session, flash,
    send_file, abort, Response, g, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, Border, Side
except ImportError:
    Workbook = None
    load_workbook = None

try:
    from docx import Document as DocxDocument
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
except ImportError:
    DocxDocument = None

# Поддержка запуска из собранного PyInstaller-бинарника:
#  - ресурсы (templates/, static/) берутся из внутриархивного каталога _MEIPASS;
#  - рабочие данные (data/) — рядом с исполняемым файлом, чтобы обновление
#    программы не затрагивало базу, сканы и логи.
if getattr(sys, 'frozen', False):               # запуск EXE-сборки
    RESOURCE_DIR = sys._MEIPASS                  # распакованные ресурсы
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    RESOURCE_DIR = BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data', 'orders.db')
UPLOAD_DIR = os.path.join(BASE_DIR, 'data', 'scans')
ARCHIVE_DIR = os.path.join(BASE_DIR, 'data', 'archives')
VERSION_FILE = os.path.join(BASE_DIR, 'data', 'version.json')
LOG_DIR = os.path.join(BASE_DIR, 'data', 'logs')
LOG_FILE = os.path.join(LOG_DIR, 'app.log')

# Порт по умолчанию изменён на 8080, т.к. на сервере порт 5000 уже занят
# другой программой. Порт можно переопределить переменной окружения PORT
# или параметром командной строки: python3 app.py --port 8090
DEFAULT_PORT = 8080


def find_free_port(start):
    """Возвращает первый свободный TCP-порт, начиная с start (для 0.0.0.0)."""
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                s.bind(('0.0.0.0', port))
                return port
            except OSError:
                continue
    return None

APP_VERSION = '1.5.3'  # 1.5.0 — клиентское приложение; 1.5.1 — поддержка сборки PyInstaller
# 1.2.0 — поле «Примечание» во всех приказах + загрузка сотрудников и видов приказов из Excel
# 1.2.1 — куратору разрешено редактирование справочника «Виды приказов»
# 1.2.2 — панель загрузки приказов в журналы из Excel + подробное описание обновления программы
# 1.3.0 — роль «Начальник отдела» + обновление без выключения сервера
# 1.3.1 — отображение активных пользователей в меню (Администратор, Куратор, Начальник отдела)
# 1.3.2 — фильтрация мусора от HTTPS-подключений к HTTP-порту в логах
# 1.4.0 — REST API v1 (/api/v1/...) + клиентское десктопное приложение client_app.py
# 1.4.1 — очистка журнала логирования администратором (режимы: с копией / только копия / полная)
# 1.5.0 — настраиваемый префикс номера приказа по каждому виду приказа
#         (справочник «Виды приказов» -> колонка «Префикс номера», пример: 1-ОП)
# 1.5.1 — поддержка сборки PyInstaller (Windows / ALT Linux)
# 1.5.2 — фильтры-воронки и сортировка по столбцам в журналах; выделение и массовое
#         удаление приказов; отбор по периоду с сохранением для пользователя;
#         пакетная загрузка документов в сервисе; смена пароля пользователя администратором;
#         справочник «Причины» + поле «Причина» в журнале «Опека»; новая логика отчёта
#         «Распечатка журнала за период» (печать / Excel / Word, выборка по виду приказа,
#         для «Опеки» наименование + причина в скобках)
# 1.5.3 — загрузка документов из Excel перенесена в меню администратора «Сервис»
#         с разделением по видам приказов (Опека / Нуждаемость / Профилактика);
#         импорты справочников сотрудников и видов приказов — там же


def setup_logging():
    """Настройка журнала логирования: файл data/logs/app.log с ротацией
    (5 файлов по 2 МБ) + вывод в консоль."""
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger('orders_app')
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        fmt = logging.Formatter(
            '%(asctime)s | %(levelname)-8s | %(message)s', '%Y-%m-%d %H:%M:%S')
        fh = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000,
                                 backupCount=5, encoding='utf-8')
        fh.setFormatter(fmt)
        logger.addHandler(fh)
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)
    return logger


class _TlsNoiseFilter(logging.Filter):
    """Фильтр для журнала Werkzeug: подавляет «мусорные» записи 400 Bad request
    syntax/HTTP version, возникающие когда клиент обращается к HTTP-порту по
    протоколу HTTPS (TLS Client Hello). Такие обращения не являются ошибками
    приложения и только засоряют консоль и server.out.

    Дополнительно фильтр помечает текущий запрос признаком TLS-мусора — этот
    флаг использует log_request(), чтобы не писать такие обращения в файл
    data/logs/app.log."""

    def filter(self, record):
        try:
            msg = record.getMessage()
        except Exception:
            return True
        # Типичный признак: бинарные escape-последовательности \x16\x03\x01…
        if '400' in msg and ('\\x16' in msg or '\\x03\\x01' in msg):
            try:
                g = getattr(_tls_noise_ctx, 'g', None)
                if g is not None:
                    g['tls_noise'] = True
            except Exception:
                pass
            return False
        return True


# Хранилище контекста текущего запроса для фильтра (Flask app context 'g')
_tls_noise_ctx = threading.local()

log = setup_logging()
# Подключаем фильтр к штатному логгеру Werkzeug (консольный вывод запросов)
logging.getLogger('werkzeug').addFilter(_TlsNoiseFilter())


def audit(action, details='', user=None):
    """Запись события аудита в журнал логирования (кто и что сделал).

    user — необязательная строка с описанием пользователя (используется API,
    где авторизация идёт по Basic-auth без cookie-сессии)."""
    if user is None:
        user = session.get('user_id')
    if user:
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            r = conn.execute('SELECT full_name, login FROM employees '
                             'WHERE id=?', (user,)).fetchone()
            conn.close()
            who = f'{r["full_name"]} ({r["login"]})' if r else f'id={user}'
        except Exception:
            who = f'id={user}'
    else:
        who = 'не авторизован'
    ip = request.remote_addr if has_request_context_safe() else '-'
    log.info(f'[АУДИТ] {action} | пользователь: {who} | IP: {ip} | {details}')


def has_request_context_safe():
    try:
        from flask import has_request_context
        return has_request_context()
    except Exception:
        return False

JOURNAL_TYPES = ['opeka', 'nuzhdaemost', 'profilaktika']
JOURNAL_TITLES = {
    'opeka': 'Опека',
    'nuzhdaemost': 'Нуждаемость',
    'profilaktika': 'Профилактика',
}
JOURNAL_SUFFIX = {  # суффикс номера приказа по умолчанию (если префикс не задан)
    'opeka': '-оп',
    'nuzhdaemost': '-ну',
    'profilaktika': '-пр',
}
# Журнал, в котором у приказа есть поле «Причина» (данные — из справочника «Причины»)
REASON_JOURNAL = 'opeka'
REASONS_ENABLED = {j: (j == REASON_JOURNAL) for j in JOURNAL_TYPES}
# Начальные значения справочника «Причины» (заполняется также при импорте Excel)
REASON_DEMO = [
    'Заявление гражданина',
    'Выявление органом опеки и попечительства',
    'Направление медицинской организации',
    'Решение суда',
    'Заявление законного представителя',
]


def _get_or_create_reason(conn, name):
    """Идентификатор причины по наименованию из справочника «Причины».

    Пустое значение -> None. Неизвестное наименование регистрируется в
    справочнике автоматически (как это сделано для видов приказов).
    """
    name = _norm(name or '')
    if not name:
        return None
    row = conn.execute('SELECT id FROM reasons WHERE LOWER(TRIM(name))=?',
                       (name.lower(),)).fetchone()
    if row:
        return row['id']
    try:
        cur = conn.execute('INSERT INTO reasons(name) VALUES (?)', (name,))
        return cur.lastrowid
    except sqlite3.IntegrityError:
        row = conn.execute('SELECT id FROM reasons WHERE LOWER(TRIM(name))=?',
                           (name.lower(),)).fetchone()
        return row['id'] if row else None


def journal_prefix(conn, journal):
    """Префикс номера приказа для журнала (настройка в справочнике «Виды приказов»).

    Префиксы хранятся в таблице order_types в колонке prefix: строка с непустым
    prefix и есть настройка вида приказа. Значение приводится к виду «-ОП»:
    добавляется ведущий дефис, пробелы убираются. Пустой префикс -> используется
    суффикс журнала по умолчанию (JOURNAL_SUFFIX). Пример: вид «Опека»,
    префикс «ОП», номер 1 -> «1-ОП».
    """
    try:
        row = conn.execute(
            "SELECT prefix FROM order_types WHERE kind=? AND TRIM(COALESCE(prefix,''))<>''"
            ' ORDER BY id LIMIT 1', (journal,)).fetchone()
    except sqlite3.Error:   # колонки ещё нет (миграция не выполнена) — дефолт
        row = None
    if not row or not (row['prefix'] or '').strip():
        return JOURNAL_SUFFIX.get(journal, '')
    p = row['prefix'].strip().upper() if row['prefix'].strip().isascii() else row['prefix'].strip()
    if not p.startswith('-'):
        p = '-' + p
    return p
# Роль «Начальник отдела» добавлена в каждую группу доступа — она имеет
# доступ ко всем журналам, всем отчётам и всем справочникам (версия 1.3.0).
ROLE_ACCESS = {
    'opeka': ('Администратор', 'Куратор', 'Начальник отдела', 'Опека'),
    'nuzhdaemost': ('Администратор', 'Куратор', 'Начальник отдела', 'Нуждаемость'),
    'profilaktika': ('Администратор', 'Куратор', 'Начальник отдела', 'Профилактика'),
}
ALL_ROLES = ['Администратор', 'Начальник отдела', 'Куратор',
             'Опека', 'Нуждаемость', 'Профилактика']
# Роли с полным доступом к данным (все журналы + все отчёты + все справочники)
FULL_DATA_ROLES = ('Администратор', 'Начальник отдела')

app = Flask(__name__,
            static_folder=os.path.join(RESOURCE_DIR, 'static'),
            template_folder=os.path.join(RESOURCE_DIR, 'templates'))


def _load_or_create_secret():
    p = os.path.join(BASE_DIR, 'data', 'secret_key.txt')
    os.makedirs(os.path.dirname(p), exist_ok=True)
    if os.path.exists(p):
        with open(p) as f:
            return f.read().strip()
    key = os.urandom(24).hex()
    with open(p, 'w') as f:
        f.write(key)
    return key


app.secret_key = _load_or_create_secret()

# ------------------------------------------------------------- активные пользователи
# Реестр активных сессий: user_id -> {'sid': подпись сессии, 'last_seen': unix-time}.
# Заполняется при входе, обновляется на каждом запросе; запись удаляется при выходе.
ACTIVE_SESSIONS = {}
ACTIVE_SESSIONS_LOCK = threading.Lock()
# Сессия считается активной, если пользователь действовал не более N секунд назад
ACTIVE_TIMEOUT = int(os.environ.get('ACTIVE_TIMEOUT', '300'))  # 5 минут


def _session_signature(uid):
    """Подпись сессии пользователя из signed cookie (секрет известен серверу).
    Совпадение подписи означает, что cookie принадлежит реальной активной
    сессии данного пользователя (а не «протухшему» cookie после перезапуска)."""
    raw = request.cookies.get(app.config['SESSION_COOKIE_NAME'], '')
    if not raw:
        return ''
    try:
        data = app.session_serializer.loads(raw)
        if isinstance(data, dict) and data.get('user_id') == uid:
            return hashlib.sha1(
                json.dumps(data, sort_keys=True).encode('utf-8')).hexdigest()[:16]
    except Exception:
        pass
    return ''


# --- Хранение активных сессий в БД (переживает перезапуск/обновление сервера)
def _db_touch_throttled(key, min_interval=10.0):
    """True, если с момента последней записи в БД по ключу прошло >= min_interval сек.
    Защита от лишней нагрузки: UPDATE в active_sessions выполняется не на каждый
    HTTP-запрос, а не чаще одного раза в 10 секунд на сессию."""
    now = time.time()
    last = _DB_TOUCH_TS.get(key, 0)
    if now - last >= min_interval:
        _DB_TOUCH_TS[key] = now
        return True
    return False


_DB_TOUCH_TS = {}  # {(uid, sid): unix-time последней записи в БД}


def db_touch_active_session(uid, sid, ip):
    """Записать/обновить отметку активности сессии в таблице active_sessions."""
    try:
        conn = get_db()
        conn.execute(
            'INSERT INTO active_sessions(user_id, sid, last_seen, ip) '
            'VALUES(?,?,?,?) '
            'ON CONFLICT(user_id, sid) DO UPDATE SET last_seen=excluded.last_seen, '
            'ip=excluded.ip', (uid, sid or '', time.time(), ip or ''))
        conn.commit()
        conn.close()
    except Exception:
        pass  # реестр в памяти остаётся основным источником


def db_drop_active_session(uid, sid=None):
    """Удалить запись(и) о сессиях пользователя из таблицы БД."""
    try:
        conn = get_db()
        if sid is not None:
            conn.execute('DELETE FROM active_sessions WHERE user_id=? AND sid=?',
                         (uid, sid or ''))
        else:
            conn.execute('DELETE FROM active_sessions WHERE user_id=?', (uid,))
        conn.commit()
        conn.close()
    except Exception:
        pass


def db_cleanup_active_sessions(cutoff):
    """Удалить из таблицы БД устаревшие отметки активности."""
    try:
        conn = get_db()
        conn.execute('DELETE FROM active_sessions WHERE last_seen < ?', (cutoff,))
        conn.commit()
        conn.close()
    except Exception:
        pass


def db_load_active_sessions(cutoff):
    """Загрузить свежие сессии из таблицы БД: {(user_id, sid): {...}}.
    Используется при старте приложения, чтобы список «Онлайн» восстановился
    после перезапуска без потери информации об активных пользователях."""
    result = {}
    try:
        conn = get_db()
        for r in conn.execute('SELECT user_id, sid, last_seen, ip '
                              'FROM active_sessions WHERE last_seen >= ?',
                              (cutoff,)):
            result[(r['user_id'], r['sid'])] = {'last_seen': r['last_seen'],
                                                'ip': r['ip'] or ''}
        conn.close()
    except Exception:
        pass
    return result


def restore_active_sessions():
    """При старте процесса загрузить активные сессии из БД в память."""
    loaded = db_load_active_sessions(time.time() - ACTIVE_TIMEOUT)
    if loaded:
        with ACTIVE_SESSIONS_LOCK:
            ACTIVE_SESSIONS.update(loaded)
    return len(loaded)


def touch_active_session(uid):
    """Отметить сессию пользователя как активную (вызывается на каждом запросе).
    Ключ реестра — (user_id, подпись cookie-сессии), поэтому несколько устройств
    или вкладок одного сотрудника учитываются раздельно.
    Дополнительно запись дублируется в таблицу active_sessions (БД) — это даёт
    возможность видеть активных пользователей сразу после перезапуска сервера."""
    sid = _session_signature(uid)
    ip = request.remote_addr if has_request_context_safe() else ''
    now = time.time()
    with ACTIVE_SESSIONS_LOCK:
        ACTIVE_SESSIONS[(uid, sid)] = {'last_seen': now, 'ip': ip}
    # в БД пишем не на каждый запрос, а не чаще раза в 10 секунд на сессию
    if _db_touch_throttled((uid, sid)):
        db_touch_active_session(uid, sid, ip)


def drop_active_session(uid, sid=None):
    """Убрать сессию из реестра (выход / истечение срока)."""
    with ACTIVE_SESSIONS_LOCK:
        if sid is not None:
            ACTIVE_SESSIONS.pop((uid, sid), None)
            _DB_TOUCH_TS.pop((uid, sid), None)
        else:  # сбрасываем все сессии данного пользователя
            for key in [k for k in ACTIVE_SESSIONS if k[0] == uid]:
                del ACTIVE_SESSIONS[key]
                _DB_TOUCH_TS.pop(key, None)
    db_drop_active_session(uid, sid)


def cleanup_active_sessions():
    """Удалить из реестра записи, превысившие срок неактивности
    (как из памяти, так и из таблицы БД active_sessions)."""
    cutoff = time.time() - ACTIVE_TIMEOUT
    with ACTIVE_SESSIONS_LOCK:
        expired = [k for k, v in ACTIVE_SESSIONS.items()
                   if v['last_seen'] < cutoff]
        for k in expired:
            del ACTIVE_SESSIONS[k]
    db_cleanup_active_sessions(cutoff)


def get_active_users():
    """Список активных пользователей (свежие сессии за последние ACTIVE_TIMEOUT сек):
    [{'id','full_name','role_name','ip','seconds_ago'}] — сортировка по ФИО.
    Если у сотрудника несколько активных сессий, показывается самая свежая."""
    cleanup_active_sessions()
    with ACTIVE_SESSIONS_LOCK:
        items = [(k[0], dict(v)) for k, v in ACTIVE_SESSIONS.items()]
    if not items:
        return []
    # самая свежая сессия для каждого пользователя
    best = {}
    for uid, v in items:
        if uid not in best or v['last_seen'] > best[uid]['last_seen']:
            best[uid] = v
    ids = list(best.keys())
    conn = get_db()
    qmarks = ','.join('?' * len(ids))
    rows = {r['id']: r for r in conn.execute(
        f'SELECT e.id, e.full_name, r.name AS role_name FROM employees e '
        f'LEFT JOIN roles r ON r.id=e.role_id WHERE e.id IN ({qmarks})',
        ids)}
    conn.close()
    now = time.time()
    result = []
    for u, v in best.items():
        row = rows.get(u)
        if not row:
            continue
        result.append({'id': u, 'full_name': row['full_name'],
                       'role_name': row['role_name'] or '—',
                       'ip': v.get('ip', ''),
                       'seconds_ago': int(now - v['last_seen'])})
    result.sort(key=lambda x: x['full_name'])
    return result


def _looks_like_tls_handshake() -> bool:
    """Признак TLS Client Hello, отправленного на HTTP-порт.

    Браузер или сервис обращается по адресу https://, а приложение принимает
    только http://. Werkzeug не может разобрать зашифрованный «запрос» и пишет
    в консоль мусор с ошибкой 400 Bad request syntax. Функция позволяет
    распознать такие обращения и отфильтровать их из журнала.

    Источники признака (в порядке надёжности):
    1) флаг 'tls_noise', который выставляет _TlsNoiseFilter при подавлении
       мусорной записи werkzeug;
    2) первый байт REQUEST_METHOD == 0x16 (TLS ContentType=handshake).
    """
    try:
        if getattr(_tls_noise_ctx, 'g', None) and \
                _tls_noise_ctx.g.get('tls_noise'):
            return True
        method = request.environ.get('REQUEST_METHOD', '')
        first = method[:1] if isinstance(method, str) else b''
        if isinstance(first, str):
            first = first.encode('latin-1', 'ignore')[:1]
        return first == b'\x16'  # ContentType=22 (handshake) в протоколе TLS
    except Exception:
        return False


@app.before_request
def _mark_tls_noise_context():
    """Сохраняет ссылку на объект g текущего запроса в потокобезопасное
    хранилище, чтобы фильтр журнала Werkzeug мог пометить TLS-мусор."""
    try:
        from flask import g as _fg
        _tls_noise_ctx.g = _fg
    except Exception:
        pass


@app.after_request
def log_request(response):
    """Журнал логирования HTTP-запросов (каждый запрос)."""
    try:
        # Пропускаем запись в файл логов мусора от HTTPS-подключений к HTTP-порту
        if response.status_code == 400 and _looks_like_tls_handshake():
            return response
        uid = session.get('user_id')
        status = response.status_code
        lvl = logging.WARNING if status >= 400 else logging.INFO
        # Ошибки доступа (403) и 404 пишем отдельной строкой для аудита
        log.log(lvl, f'[HTTP] {request.method} {request.path} -> {status} '
                     f'(IP: {request.remote_addr}, user_id: {uid or "-"})')
        if status == 403:
            audit('ОТКАЗ В ДОСТУПЕ (403)', f'{request.method} {request.path}')
        # Неудачные попытки входа пишем в журнал отдельно
        if (request.endpoint == 'login' and request.method == 'POST'
                and status == 200):
            log.warning(f'[БЕЗОПАСНОСТЬ] Неудачная попытка входа: логин '
                        f"'{request.form.get('login','')}' от IP "
                        f"{request.remote_addr}")
    except Exception:
        pass
    return response

# ---------------------------------------------------------------- DB helpers

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL
);
CREATE TABLE IF NOT EXISTS employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name TEXT NOT NULL,
    position TEXT NOT NULL DEFAULT '',
    login TEXT UNIQUE,
    password_hash TEXT,
    pwd_changed_at TEXT,           -- дата последней самостоятельной смены пароля
    role_id INTEGER REFERENCES roles(id)
);
CREATE TABLE IF NOT EXISTS order_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL,              -- вид приказа (название журнала)
    name TEXT NOT NULL,              -- наименование приказа
    prefix TEXT NOT NULL DEFAULT '', -- префикс номера приказа для вида (напр. "ОП" -> 1-ОП)
    UNIQUE(kind, name)
);
CREATE TABLE IF NOT EXISTS reasons (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT UNIQUE NOT NULL        -- справочник «Причины» (заполняется при импорте из Excel)
);
CREATE TABLE IF NOT EXISTS orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    journal TEXT NOT NULL,           -- opeka / nuzhdaemost / profilaktika
    order_number TEXT NOT NULL,      -- номер приказа, напр. "12-оп"
    order_date TEXT NOT NULL,        -- ГГГГ-ММ-ДД
    order_type_id INTEGER REFERENCES order_types(id),
    person_fio TEXT NOT NULL DEFAULT '',   -- ФИО в отношении кого
    specialist_fio TEXT NOT NULL DEFAULT '', -- ФИО специалиста
    note TEXT NOT NULL DEFAULT '',         -- примечание (добавлено в 1.2.0)
    reason_id INTEGER REFERENCES reasons(id), -- причина приказа (поле журнала «Опека»)
    scan_path TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS counters (
    journal TEXT PRIMARY KEY,
    value INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS active_sessions (
    user_id INTEGER NOT NULL,
    sid TEXT NOT NULL DEFAULT '',      -- подпись cookie-сессии
    last_seen REAL NOT NULL,           -- unix-time последней активности
    ip TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (user_id, sid)
);                     -- реестр активных пользователей (переживает перезапуск)
CREATE TABLE IF NOT EXISTS user_prefs (
    user_id INTEGER NOT NULL REFERENCES employees(id),
    key TEXT NOT NULL,                 -- напр. journal_period.opeka
    value TEXT NOT NULL DEFAULT '',    -- сериализованный JSON
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (user_id, key)
);                     -- персональные настройки пользователей (сохраняются между сессиями)
"""


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    conn = get_db()
    conn.executescript(SCHEMA)
    # Миграция: добавляем колонку pwd_changed_at в существующих базах
    cols = [c['name'] for c in conn.execute('PRAGMA table_info(employees)')]
    if 'pwd_changed_at' not in cols:
        conn.execute('ALTER TABLE employees ADD COLUMN pwd_changed_at TEXT')
    # Миграция: добавляем колонку note (примечание) в существующих базах
    ocols = [c['name'] for c in conn.execute('PRAGMA table_info(orders)')]
    if 'note' not in ocols:
        conn.execute("ALTER TABLE orders ADD COLUMN note TEXT NOT NULL DEFAULT ''")
    # Миграция v1.5.0: колонка prefix (префикс номера приказа) в справочнике видов приказов
    tcols = [c['name'] for c in conn.execute('PRAGMA table_info(order_types)')]
    if 'prefix' not in tcols:
        conn.execute("ALTER TABLE order_types ADD COLUMN prefix TEXT NOT NULL DEFAULT ''")
    # Миграция: колонка reason_id (причина приказа, журнал «Опека») + справочник «Причины»
    if 'reason_id' not in ocols:
        conn.execute('ALTER TABLE orders ADD COLUMN reason_id INTEGER REFERENCES reasons(id)')
    # seed roles
    for r in ALL_ROLES:
        conn.execute('INSERT OR IGNORE INTO roles(name) VALUES (?)', (r,))
    # seed admin if no users
    cnt = conn.execute('SELECT COUNT(*) c FROM employees').fetchone()['c']
    if cnt == 0:
        role = conn.execute("SELECT id FROM roles WHERE name='Администратор'").fetchone()
        conn.execute(
            'INSERT INTO employees(full_name, position, login, password_hash, role_id) '
            'VALUES (?,?,?,?,?)',
            ('Системный администратор', 'Администратор', 'admin',
             generate_password_hash('admin'), role['id']))
    # seed order types
    seeded = conn.execute("SELECT COUNT(*) c FROM order_types").fetchone()['c']
    if seeded == 0:
        demo = {
            'opeka': ['Назначение опекуна', 'Прекращение опеки',
                      'Помещение под опеку', 'Отстранение опекуна'],
            'nuzhdaemost': ['Признание нуждающимся в социальном обслуживании',
                            'Снятие с учёта нуждающихся',
                            'Продление статуса нуждающегося'],
            'profilaktika': ['Постановка на профилактический учёт',
                             'Снятие с профилактического учёта',
                             'Проведение профилактического мероприятия'],
        }
        for kind, names in demo.items():
            for n in names:
                conn.execute('INSERT OR IGNORE INTO order_types(kind,name) VALUES (?,?)',
                             (kind, n))
    # seed справочника «Причины» (поле «Причина» журнала «Опека») — только при пустом
    # справочнике; значения также добавляются автоматически при импорте Excel
    rseeded = conn.execute('SELECT COUNT(*) c FROM reasons').fetchone()['c']
    if rseeded == 0:
        for rn in REASON_DEMO:
            conn.execute('INSERT OR IGNORE INTO reasons(name) VALUES (?)', (rn,))
    for j in JOURNAL_TYPES:
        conn.execute('INSERT OR IGNORE INTO counters(journal,value) VALUES (?,0)', (j,))
    conn.commit()
    conn.close()


def write_version_file():
    os.makedirs(os.path.dirname(VERSION_FILE), exist_ok=True)
    with open(VERSION_FILE, 'w', encoding='utf-8') as f:
        json.dump({'version': APP_VERSION,
                   'updated': datetime.now().isoformat(timespec='seconds')}, f)


init_db()
write_version_file()
# Восстанавливаем реестр активных пользователей из БД после перезапуска/обновления
_n_restored = restore_active_sessions()
if _n_restored:
    log.info(f'[HTTP] Восстановлено активных сессий из БД: {_n_restored}')

# ---------------------------------------------------------------- auth

def current_user():
    uid = session.get('user_id')
    if not uid:
        return None
    conn = get_db()
    u = conn.execute('SELECT e.*, r.name AS role_name FROM employees e '
                     'LEFT JOIN roles r ON r.id=e.role_id WHERE e.id=?', (uid,)).fetchone()
    conn.close()
    return u


def require_login():
    if not current_user():
        abort(redirect(url_for('login')))


def is_admin(user=None):
    user = user or current_user()
    return bool(user) and user['role_name'] == 'Администратор'


def is_full_access(user=None):
    """Полный доступ к данным: Администратор и Начальник отдела
    (все журналы, все отчёты, все справочники)."""
    user = user or current_user()
    return bool(user) and user['role_name'] in FULL_DATA_ROLES


def can_edit_employees(user=None):
    """Справочник «Сотрудники»: Администратор и Начальник отдела."""
    return is_full_access(user)


def can_edit_order_types(user=None):
    """Справочник «Виды приказов»: Администратор, Начальник отдела, Куратор."""
    user = user or current_user()
    return bool(user) and user['role_name'] in FULL_DATA_ROLES + ('Куратор',)


def can_view_reports(user=None):
    """Отчёты: Администратор, Начальник отдела и Куратор."""
    user = user or current_user()
    return bool(user) and user['role_name'] in FULL_DATA_ROLES + ('Куратор',)


def can_access_journal(user, journal):
    if not user:
        return False
    if user['role_name'] in FULL_DATA_ROLES:  # админ и начальник — полный доступ
        return True
    return user['role_name'] in ROLE_ACCESS[journal]


@app.before_request
def _guard():
    open_exceptions = ('login', 'logout', 'static')
    if request.endpoint in open_exceptions:
        return
    # REST API v1 (клиентское приложение): авторизация Basic выполняется
    # внутри самих api_*-обработчиков; публичный /api/v1/info открыт всем
    ep = request.endpoint or ''
    if request.path.startswith('/api/'):
        if ep in ('api_info', 'api_user'):
            return
        if ep == 'api_scans_upload':
            # пакетная загрузка документов из «Сервиса»: работает по cookie-сессии,
            # права (только Администратор) проверяет сам обработчик
            if not session.get('user_id'):
                return redirect(url_for('login'))
            touch_active_session(session['user_id'])
            return
        return          # дальнейшие права проверяет обработчик (401 вместо редиректа)
    if not session.get('user_id'):
        return redirect(url_for('login'))
    # отмечаем сессию пользователя активной (для счётчика «онлайн» в меню)
    touch_active_session(session['user_id'])


@app.context_processor
def inject_globals():
    user = current_user()
    journals = [j for j in JOURNAL_TYPES if user and can_access_journal(user, j)]
    # Активные пользователи — видны Администратору, Куратору и Начальнику отдела
    active_users = None
    if user and user['role_name'] in ('Администратор', 'Куратор',
                                      'Начальник отдела'):
        active_users = get_active_users()
    return dict(
        cu=user,
        accessible_journals=[(j, JOURNAL_TITLES[j]) for j in journals],
        all_journals=[(j, JOURNAL_TITLES[j]) for j in JOURNAL_TYPES],
        journal_titles=JOURNAL_TITLES,
        all_roles=ALL_ROLES,
        app_version=APP_VERSION,
        active_users=active_users,
        active_timeout=ACTIVE_TIMEOUT,
    )

# ---------------------------------------------------------------- login

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        login_name = request.form.get('login', '').strip()
        password = request.form.get('password', '')
        conn = get_db()
        u = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                         'LEFT JOIN roles r ON r.id=e.role_id '
                         'WHERE e.login=? AND e.password_hash IS NOT NULL',
                         (login_name,)).fetchone()
        conn.close()
        if u and check_password_hash(u['password_hash'], password):
            session['user_id'] = u['id']
            # При входе показываем подсказку о смене пароля, если сотрудник
            # ещё не менял пароль после его установки администратором.
            if not u['pwd_changed_at']:
                flash('Рекомендуем сменить пароль: перейдите в «Моя учётная запись».', 'warn')
            audit('ВХОД В СИСТЕМУ', f'роль: {u["role_name"]}')
            return redirect(url_for('index'))
        error = 'Неверный логин или пароль'
    if session.get('user_id'):
        return redirect(url_for('index'))
    return render_template('login.html', error=error)


@app.route('/logout')
def logout():
    uid = session.get('user_id')
    audit('ВЫХОД ИЗ СИСТЕМЫ')
    if uid:  # убираем пользователя из реестра активных сессий
        drop_active_session(uid, _session_signature(uid))
    session.clear()
    return redirect(url_for('login'))

# ---------------------------------------------------------------- main page

@app.route('/')
def index():
    user = current_user()
    conn = get_db()
    stats = {}
    for j in JOURNAL_TYPES:
        if can_access_journal(user, j):
            stats[j] = conn.execute(
                'SELECT COUNT(*) c FROM orders WHERE journal=?', (j,)).fetchone()['c']
    conn.close()
    return render_template('index.html', stats=stats)

# ---------------------------------------------------------------- справочник сотрудников

@app.route('/employees')
def employees():
    require_login()
    conn = get_db()
    rows = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                        'LEFT JOIN roles r ON r.id=e.role_id ORDER BY e.full_name').fetchall()
    conn.close()
    return render_template('employees.html', rows=rows)


@app.route('/employees/save', methods=['POST'])
def employee_save():
    user = current_user()
    # Администратор и Начальник отдела
    if not can_edit_employees(user):
        abort(403)
    eid = request.form.get('id', type=int)
    full_name = request.form.get('full_name', '').strip()
    position = request.form.get('position', '').strip()
    login_name = request.form.get('login', '').strip() or None
    role_name = request.form.get('role', 'Куратор')
    password = request.form.get('password', '')
    conn = get_db()
    role = conn.execute('SELECT id FROM roles WHERE name=?', (role_name,)).fetchone()
    role_id = role['id'] if role else None
    try:
        if eid:
            cur = conn.execute('SELECT * FROM employees WHERE id=?', (eid,)).fetchone()
            if not cur:
                abort(404)
            pw_hash = cur['password_hash']
            pwd_changed = cur['pwd_changed_at'] if 'pwd_changed_at' in cur.keys() else None
            if password:
                # администратор установил новый пароль — счётчик смены обнуляется,
                # сотруднику будет рекомендовано сменить его самостоятельно
                pw_hash = generate_password_hash(password)
                pwd_changed = None
            elif not login_name:
                pw_hash = None
            conn.execute(
                'UPDATE employees SET full_name=?, position=?, login=?, '
                'password_hash=?, pwd_changed_at=?, role_id=? WHERE id=?',
                (full_name, position, login_name, pw_hash, pwd_changed, role_id, eid))
            audit('ИЗМЕНЁН СОТРУДНИК', f'id={eid}, {full_name}')
        else:
            # новый сотрудник: если логин задан, но пароль не указан —
            # устанавливаем временный пароль '1234' (сотрудник сменит сам)
            if login_name and not password:
                password = '1234'
            pw_hash = generate_password_hash(password) if (login_name and password) else None
            conn.execute(
                'INSERT INTO employees(full_name, position, login, password_hash, role_id) '
                'VALUES (?,?,?,?,?)',
                (full_name, position, login_name, pw_hash, role_id))
        conn.commit()
        audit('СОТРУДНИК СОХРАНЁН', full_name)
        flash('Сотрудник сохранён', 'ok')
    except sqlite3.IntegrityError:
        conn.rollback()
        audit('ОШИБКА СОХРАНЕНИЯ СОТРУДНИКА', f'занятый логин: {login_name}')
        flash('Ошибка: такой логин уже используется', 'err')
    finally:
        conn.close()
    return redirect(url_for('employees'))


@app.route('/employees/delete/<int:eid>', methods=['POST'])
def employee_delete(eid):
    if not can_edit_employees():
        abort(403)
    conn = get_db()
    emp = conn.execute('SELECT * FROM employees WHERE id=?', (eid,)).fetchone()
    if emp and emp['id'] == session.get('user_id'):
        conn.close()
        flash('Нельзя удалить текущую учётную запись', 'err')
        return redirect(url_for('employees'))
    conn.execute('DELETE FROM employees WHERE id=?', (eid,))
    conn.commit()
    conn.close()
    audit('СОТРУДНИК УДАЛЁН', f'id={eid}')
    flash('Сотрудник удалён', 'ok')
    return redirect(url_for('employees'))

# ---------------------------------------------------------------- справочник видов приказов

@app.route('/order_types')
def order_types():
    require_login()
    conn = get_db()
    rows = conn.execute('SELECT * FROM order_types ORDER BY kind, name').fetchall()
    conn.close()
    return render_template('order_types.html', rows=rows)


@app.route('/order_types/save', methods=['POST'])
def order_type_save():
    user = current_user()
    # Разрешено Администратору и Куратору (куратор ведёт справочник видов приказов)
    if not can_edit_order_types(user):
        abort(403)
    tid = request.form.get('id', type=int)
    kind = request.form.get('kind', 'opeka')
    name = request.form.get('name', '').strip()
    if not name:
        flash('Укажите наименование приказа', 'err')
        return redirect(url_for('order_types'))
    conn = get_db()
    try:
        if tid:
            conn.execute('UPDATE order_types SET kind=?, name=? WHERE id=?', (kind, name, tid))
        else:
            conn.execute('INSERT INTO order_types(kind,name) VALUES (?,?)', (kind, name))
        conn.commit()
        # в этой функции переменная называется kind (вид приказа), а не journal —
        # исправлена ошибка NameError при аудите сохранения вида приказа
        audit('ВИД ПРИКАЗА СОХРАНЁН', f'{JOURNAL_TITLES.get(kind, kind)}: {name}')
        flash('Вид приказа сохранён', 'ok')
    except sqlite3.IntegrityError:
        conn.rollback()
        audit('ОШИБКА: ДУБЛЬ ВИДА ПРИКАЗА', f'{JOURNAL_TITLES.get(kind, kind)}: {name}')
        flash('Такой вид приказа уже существует', 'err')
    finally:
        conn.close()
    return redirect(url_for('order_types'))


@app.route('/order_types/delete/<int:tid>', methods=['POST'])
def order_type_delete(tid):
    # Разрешено Администратору и Куратору; удаление по-прежнему блокируется,
    # если вид приказа используется в документах журналов.
    if not can_edit_order_types():
        abort(403)
    conn = get_db()
    used = conn.execute('SELECT COUNT(*) c FROM orders WHERE order_type_id=?', (tid,)).fetchone()['c']
    if used:
        conn.close()
        flash('Нельзя удалить: вид приказа используется в документах', 'err')
    else:
        conn.execute('DELETE FROM order_types WHERE id=?', (tid,))
        conn.commit()
        conn.close()
        audit('ВИД ПРИКАЗА УДАЛЁН', f'id={tid}')
        flash('Удалено', 'ok')
    return redirect(url_for('order_types'))

# ------------------------------------------- импорт справочников из Excel (v1.2.0)

ROLE_ALIASES = {  # частые варианты написания ролей в таблицах Excel
    'админ': 'Администратор', 'admin': 'Администратор',
    'начальник': 'Начальник отдела', 'нач. отдела': 'Начальник отдела',
    'начальник отдела': 'Начальник отдела',
    'опекун': 'Опека', 'нуждающиеся': 'Нуждаемость',
    'профилактический учёт': 'Профилактика',
}


def _norm(s):
    """Строка ячейки Excel -> чистое значение (даты приводим к ГГГГ-ММ-ДД)."""
    if s is None:
        return ''
    if isinstance(s, datetime):
        return s.strftime('%Y-%m-%d')
    if hasattr(s, 'isoformat'):          # datetime.date / time
        return s.isoformat()
    return str(s).strip()


def _cell_map(header):
    """Заголовок листа -> {нормализованное имя столбца: номер столбца}.

    Нормализация: нижний регистр, без точек/скобок/дефисов — чтобы распознать
    «Ф.И.О. сотрудника», «ФИО сотрудника», «Наименование приказа» и т.п.
    """
    import re
    m = {}
    for i, c in enumerate(header):
        key = _norm(c).lower().replace('ё', 'е')
        key = re.sub(r'[^а-яa-z0-9]+', '', key)
        if key and key not in m:
            m[key] = i
    return m


def _pick(m, *variants):
    """Возвращает индекс столбца по первому совпавшему варианту заголовка."""
    for v in variants:
        if v in m:
            return m[v]
    return None


def _read_xlsx_rows(stream):
    """Открывает поток как XLSX и возвращает список строк первого листа."""
    wb = load_workbook(stream, data_only=True, read_only=True)
    ws = wb.active
    rows = [[_norm(c) for c in row] for row in ws.iter_rows(values_only=True)]
    wb.close()
    return [r for r in rows if any(r)]


def _service_redirect():
    """URL страницы «Сервис» с якорем раздела загрузки документов из Excel."""
    return url_for('service', _external=True) + '#excel-import'


def _redirect_back(default_endpoint, back_to_service=False, **defaults):
    """Вернуться на страницу, откуда был вызван импорт (например, «Сервис»),
    иначе — на указанный запасной адрес.

    При back_to_service=True возврат на собственную страницу справочника
    не выполняется (панели загрузки оттуда перенесены в сервис) — вместо
    этого пользователь попадает на «Сервис → Загрузка документов из Excel».
    """
    ref = request.referrer or ''
    try:
        same_site = urlparse(ref).netloc == urlparse(request.url).netloc
    except Exception:
        same_site = False
    ep = url_for(default_endpoint, **defaults)
    if ref and same_site and urlparse(ref).path == ep:
        if back_to_service:
            return redirect(_service_redirect())
        return redirect(url_for(default_endpoint, **defaults))
    if ref and same_site and urlparse(ref).path not in ('/', '/login'):
        return redirect(ref)
    if back_to_service:
        return redirect(_service_redirect())
    return redirect(url_for(default_endpoint, **defaults))


def _mem_xlsx(wb, filename):
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(
        buf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename)


@app.route('/employees/template')
def employees_template():
    """Шаблон Excel для загрузки сотрудников (заполните лист и импортируйте)."""
    _need_xlsx()
    require_login()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Сотрудники'
    ws.append(['ФИО сотрудника', 'Должность', 'Роль', 'Логин', 'Пароль'])
    ws.append(['Иванов Иван Иванович', 'Главный специалист', 'Куратор', 'ivanov', '1234'])
    ws.append(['Петрова Анна Сергеевна', 'Специалист', 'Опека', 'petrova', '1234'])
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip('ABCDE', (32, 26, 18, 16, 14)):
        ws.column_dimensions[col].width = w
    return _mem_xlsx(wb, 'шаблон_сотрудники.xlsx')


@app.route('/employees/import', methods=['POST'])
def employees_import():
    """Массовая загрузка сотрудников из файла Excel (.xlsx). Только Администратор.

    Формат: первая строка — заголовки (ФИО, Должность, Роль, Логин, Пароль).
    Существующие сотрудники (совпадение по логину или ФИО+должности) обновляются;
    при отсутствии пароля у нового пользователя создаётся пароль '1234'.
    """
    user = current_user()
    if not can_edit_employees(user):
        abort(403)
    _need_xlsx()
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'):
        flash('Выберите файл Excel в формате .xlsx', 'err')
        return _redirect_back('employees', back_to_service=True)
    try:
        rows = _read_xlsx_rows(io.BytesIO(f.read()))
    except Exception as e:
        flash(f'Не удалось прочитать файл: {e}', 'err')
        return _redirect_back('employees', back_to_service=True)
    if len(rows) < 2:
        flash('В файле нет данных (нужны заголовок и хотя бы одна строка)', 'err')
        return _redirect_back('employees', back_to_service=True)
    m = _cell_map(rows[0])
    ci_fio = _pick(m, 'фиосотрудника', 'фио', 'фамилияимяотчество', 'сотрудник')
    ci_pos = _pick(m, 'должность')
    ci_role = _pick(m, 'роль', 'rol')
    ci_login = _pick(m, 'логин', 'login', 'пользователь')
    ci_pw = _pick(m, 'пароль', 'password')
    if ci_fio is None:
        flash('В файле не найден столбец «ФИО сотрудника»', 'err')
        return _redirect_back('employees', back_to_service=True)
    conn = get_db()
    roles = {r['name']: r['id'] for r in conn.execute('SELECT id, name FROM roles')}
    created = updated = skipped = 0
    errors = []
    for n, row in enumerate(rows[1:], start=2):
        fio = row[ci_fio] if ci_fio < len(row) else ''
        if not fio:
            skipped += 1
            continue
        position = row[ci_pos] if ci_pos is not None and ci_pos < len(row) else ''
        role_raw = row[ci_role] if ci_role is not None and ci_role < len(row) else ''
        login_name = row[ci_login] if ci_login is not None and ci_login < len(row) else ''
        password = row[ci_pw] if ci_pw is not None and ci_pw < len(row) else ''
        role_name = ROLE_ALIASES.get(role_raw.lower(), role_raw)
        if role_name and role_name not in roles:
            role_name = 'Куратор'
            errors.append(f'строка {n}: роль «{role_raw}» неизвестна — назначен Куратор')
        role_id = roles.get(role_name or 'Куратор')
        ex = None
        if login_name:
            ex = conn.execute('SELECT * FROM employees WHERE login=?',
                              (login_name,)).fetchone()
        if not ex:
            ex = conn.execute('SELECT * FROM employees WHERE full_name=? AND position=?',
                              (fio, position)).fetchone()
        try:
            if ex:
                pw_hash = ex['password_hash']
                pwd_changed = ex['pwd_changed_at'] \
                    if 'pwd_changed_at' in ex.keys() else None
                if password:  # задан новый пароль — счётчик смены обнуляется
                    pw_hash = generate_password_hash(password)
                    pwd_changed = None
                conn.execute(
                    'UPDATE employees SET full_name=?, position=?, login=?, '
                    'password_hash=?, pwd_changed_at=?, role_id=? WHERE id=?',
                    (fio, position, login_name or ex['login'], pw_hash,
                     pwd_changed, role_id, ex['id']))
                updated += 1
            else:
                pw = password or '1234'  # временный пароль, сотрудник сменит сам
                conn.execute(
                    'INSERT INTO employees(full_name, position, login, '
                    'password_hash, role_id) VALUES (?,?,?,?,?)',
                    (fio, position, login_name or None,
                     generate_password_hash(pw), role_id))
                created += 1
        except sqlite3.IntegrityError:
            skipped += 1
            errors.append(f'строка {n}: логин «{login_name}» уже занят')
    conn.commit()
    conn.close()
    audit('ИМПОРТ СОТРУДНИКОВ ИЗ EXCEL',
          f'файл={f.filename}, создано={created}, обновлено={updated}, пропущено={skipped}')
    msg = f'Импорт завершён: создано — {created}, обновлено — {updated}, пропущено — {skipped}'
    if errors:
        msg += '. Замечания: ' + '; '.join(errors[:5])
    flash(msg, 'ok')
    return _redirect_back('employees', back_to_service=True)


@app.route('/order_types/template')
def order_types_template():
    """Шаблон Excel для загрузки видов приказов."""
    _need_xlsx()
    require_login()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Виды приказов'
    ws.append(['Вид приказа', 'Наименование приказа'])
    for kind_title in JOURNAL_TITLES.values():
        ws.append([kind_title, 'Например: Назначение опекуна'])
    for c in ws[1]:
        c.font = Font(bold=True)
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 45
    return _mem_xlsx(wb, 'шаблон_виды_приказов.xlsx')


@app.route('/order_types/import', methods=['POST'])
def order_types_import():
    """Массовая загрузка видов приказов из Excel (.xlsx). Администратор и Куратор.

    Столбец «Вид приказа» — название журнала (Опека / Нуждаемость / Профилактика),
    дубликаты пар (вид+наименование) пропускаются.
    """
    user = current_user()
    if not can_edit_order_types(user):
        abort(403)
    _need_xlsx()
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'):
        flash('Выберите файл Excel в формате .xlsx', 'err')
        return _redirect_back('order_types', back_to_service=True)
    try:
        rows = _read_xlsx_rows(io.BytesIO(f.read()))
    except Exception as e:
        flash(f'Не удалось прочитать файл: {e}', 'err')
        return _redirect_back('order_types', back_to_service=True)
    if len(rows) < 2:
        flash('В файле нет данных (нужны заголовок и хотя бы одна строка)', 'err')
        return _redirect_back('order_types', back_to_service=True)
    m = _cell_map(rows[0])
    ci_kind = _pick(m, 'видприказа', 'вид', 'журнал')
    ci_name = _pick(m, 'наименованиеприказа', 'наименование', 'приказ')
    if ci_kind is None or ci_name is None:
        flash('В файле не найдены столбцы «Вид приказа» и «Наименование приказа»', 'err')
        return _redirect_back('order_types', back_to_service=True)
    # карта: название журнала (в любом регистре) -> ключ журнала
    title_to_key = {t.lower(): k for k, t in JOURNAL_TITLES.items()}
    conn = get_db()
    created = skipped = 0
    errors = []
    for n, row in enumerate(rows[1:], start=2):
        kind_raw = row[ci_kind] if ci_kind < len(row) else ''
        name = row[ci_name] if ci_name < len(row) else ''
        key = title_to_key.get(kind_raw.lower())
        if not name or not key:
            skipped += 1
            if kind_raw and not key:
                errors.append(f'строка {n}: неизвестный вид приказа «{kind_raw}»')
            continue
        try:
            conn.execute('INSERT INTO order_types(kind,name) VALUES (?,?)', (key, name))
            created += 1
        except sqlite3.IntegrityError:
            skipped += 1
    conn.commit()
    conn.close()
    audit('ИМПОРТ ВИДОВ ПРИКАЗОВ ИЗ EXCEL',
          f'файл={f.filename}, создано={created}, пропущено={skipped}')
    msg = f'Импорт завершён: создано — {created}, пропущено — {skipped}'
    if errors:
        msg += '. Замечания: ' + '; '.join(errors[:5])
    flash(msg, 'ok')
    return _redirect_back('order_types', back_to_service=True)


# ---------------------------------------------------------------- журналы

JOURNAL_ALIASES = {  # частые варианты названия журнала в Excel (в нижнем регистре)
    'opeka': ('опека',),
    'nuzhdaemost': ('нуждаемость', 'нуждающиеся'),
    'profilaktika': ('профилактика', 'профилактический учет', 'профилактический учёт'),
}


def _journal_key(raw, title_to_key):
    """Название журнала из Excel -> ключ журнала ('opeka' и т.п.) или None."""
    s = _norm(raw).lower().replace('ё', 'е').strip()
    if not s:
        return None
    for k, variants in JOURNAL_ALIASES.items():
        if s in [v.replace('ё', 'е') for v in variants]:
            return k
    return title_to_key.get(s)


def _orders_xlsx_headers(journal):
    """Заголовки столбцов для шаблона/импорта приказов конкретного журнала.

    Для журнала «Опека» добавляется столбец «Причина» — данные берутся из
    справочника «Причины» (неизвестные значения регистрируются в нём).
    """
    if REASONS_ENABLED.get(journal):
        return ['№ приказа', 'Дата приказа', 'Вид приказа', 'Наименование приказа',
                'Ф.И.О. в отношении кого приказ', 'ФИО специалиста',
                'Причина', 'Примечание']
    return ['№ приказа', 'Дата приказа', 'Вид приказа', 'Наименование приказа',
            'Ф.И.О. в отношении кого приказ', 'ФИО специалиста', 'Примечание']


# SQL-выражение для числовой части номера приказа (до первого дефиса),
# используется для сортировки документов: сначала по дате, внутри даты - по номеру
_ORDER_NUM_SQL = ("CAST(NULLIF(SUBSTR({col}, 1, INSTR({col} || '-', '-')), '') AS INTEGER)")


def next_order_number(conn, journal):
    conn.execute('UPDATE counters SET value=value+1 WHERE journal=?', (journal,))
    row = conn.execute('SELECT value FROM counters WHERE journal=?', (journal,)).fetchone()
    if not row:
        conn.execute('INSERT INTO counters(journal,value) VALUES (?,1)', (journal,))
        row = conn.execute('SELECT value FROM counters WHERE journal=?', (journal,)).fetchone()
    return f"{row['value']}{journal_prefix(conn, journal)}"


def _valid_iso_date(s):
    """Проверка строки формата ГГГГ-ММ-ДД; None — некорректная дата."""
    try:
        d = date.fromisoformat((s or '').strip())
        return d.isoformat()
    except (ValueError, TypeError):
        return None


def pref_get(conn, user_id, key, default=None):
    """Персональная настройка пользователя (JSON в таблице user_prefs)."""
    row = conn.execute('SELECT value FROM user_prefs WHERE user_id=? AND key=?',
                       (user_id, key)).fetchone()
    if not row:
        return default
    try:
        return json.loads(row['value'])
    except (ValueError, TypeError):
        return default


def pref_set(conn, user_id, key, value):
    """Сохранить персональную настройку пользователя (сериализует в JSON)."""
    conn.execute('INSERT INTO user_prefs(user_id, key, value, updated_at) '
                 'VALUES (?,?,?,CURRENT_TIMESTAMP) '
                 'ON CONFLICT(user_id, key) DO UPDATE SET '
                 'value=excluded.value, updated_at=CURRENT_TIMESTAMP',
                 (user_id, key, json.dumps(value, ensure_ascii=False)))


def journal_period(conn, user_id, journal, args):
    """Период отображения журнала: из параметров запроса (?d1&?d2), иначе —
    сохранённый ранее для данного пользователя. Значения сохраняются в БД."""
    key = f'journal_period.{journal}'
    raw1 = (args.get('d1') or '').strip()
    raw2 = (args.get('d2') or '').strip()
    if 'd1' in args or 'd2' in args:   # форма периода отправлялась (в т.ч. сброс)
        d1 = _valid_iso_date(raw1)
        d2 = _valid_iso_date(raw2)
        pref_set(conn, user_id, key, {'d1': d1 or '', 'd2': d2 or ''})
        conn.commit()
    else:
        saved = pref_get(conn, user_id, key, {}) or {}
        d1 = _valid_iso_date(saved.get('d1'))
        d2 = _valid_iso_date(saved.get('d2'))
    return d1, d2


@app.route('/journal/<journal>')
def journal_list(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    q = request.args.get('q', '').strip()
    conn = get_db()
    # отбор по периоду (сохраняется отдельно для каждого пользователя)
    p1, p2 = journal_period(conn, user['id'], journal, request.args)
    sql = ('SELECT o.*, t.name AS type_name, rs.name AS reason_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id '
           'LEFT JOIN reasons rs ON rs.id=o.reason_id WHERE o.journal=?')
    params = [journal]
    if p1:
        sql += ' AND o.order_date>=?'
        params.append(p1)
    if p2:
        sql += ' AND o.order_date<=?'
        params.append(p2)
    if q:
        sql += ' AND (o.order_number LIKE ? OR o.person_fio LIKE ? OR o.specialist_fio LIKE ?)'
        like = f'%{q}%'
        params += [like, like, like]
    # сортировка документов: по дате, внутри даты - по номеру
    sql += (' ORDER BY o.order_date DESC, ' + _ORDER_NUM_SQL.format(col='o.order_number') +
            ' DESC, o.order_number DESC')
    rows = conn.execute(sql, params).fetchall()
    types = conn.execute('SELECT * FROM order_types WHERE kind=? ORDER BY name',
                         (journal,)).fetchall()
    # справочник «Причины» — используется только в журнале «Опека»
    reasons = conn.execute('SELECT * FROM reasons ORDER BY name').fetchall()
    today = date.today().isoformat()
    conn.close()
    return render_template('journal.html', journal=journal, title=JOURNAL_TITLES[journal],
                           rows=rows, types=types, reasons=reasons, today=today,
                           q=q, p1=p1, p2=p2, has_reason=REASONS_ENABLED.get(journal, False))


@app.route('/journal/<journal>/api/next_number')
def api_next_number(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    if not can_access_journal(current_user(), journal):
        abort(403)
    conn = get_db()
    num = next_order_number(conn, journal)
    conn.commit()
    conn.close()
    return {'number': num}


@app.route('/journal/<journal>/create', methods=['POST'])
def journal_create(journal):
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    conn = get_db()
    order_number = request.form.get('order_number', '').strip()
    if not order_number:
        order_number = next_order_number(conn, journal)
    else:
        # если номер введён без префикса вида приказа — добавляем его (v1.5.0): 1 -> 1-ОП
        if '-' not in order_number and order_number.split('-')[0].isdigit():
            order_number += journal_prefix(conn, journal)
        # синхронизация счётчика, если пользователь ввёл больший номер
        base = order_number.split('-')[0]
        if base.isdigit():
            conn.execute('UPDATE counters SET value=MAX(value,?) WHERE journal=?',
                         (int(base), journal))
    order_date = request.form.get('order_date') or date.today().isoformat()
    type_id = request.form.get('order_type_id', type=int)
    person_fio = request.form.get('person_fio', '').strip()
    specialist_fio = request.form.get('specialist_fio', '').strip()
    note = request.form.get('note', '').strip()          # примечание к приказу
    # причина приказа (журнал «Опека») — берётся из справочника «Причины»
    reason_id = None
    if REASONS_ENABLED.get(journal):
        rid = request.form.get('reason_id', type=int)
        if rid and conn.execute('SELECT 1 FROM reasons WHERE id=?', (rid,)).fetchone():
            reason_id = rid
        else:
            reason_id = _get_or_create_reason(conn, request.form.get('reason', ''))
    if type_id:
        ok = conn.execute('SELECT 1 FROM order_types WHERE id=? AND kind=?',
                          (type_id, journal)).fetchone()
        if not ok:
            type_id = None
    scan_rel = None
    f = request.files.get('scan')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext != '.pdf':
            conn.close()
            audit('ОТКАЗ: СКАН НЕ В ФОРМАТЕ PDF', f.filename)
            flash('Скан должен быть в формате PDF', 'err')
            return redirect(url_for('journal_list', journal=journal))
        fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(f.filename)}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        f.save(os.path.join(UPLOAD_DIR, fname))
        scan_rel = fname
    conn.execute(
        'INSERT INTO orders(journal, order_number, order_date, order_type_id, '
        'person_fio, specialist_fio, note, reason_id, scan_path) VALUES (?,?,?,?,?,?,?,?,?)',
        (journal, order_number, order_date, type_id, person_fio, specialist_fio,
         note, reason_id, scan_rel))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ СОЗДАН', f'журнал={JOURNAL_TITLES.get(journal)}, №={order_number}, по кому={person_fio}')
    flash('Приказ создан', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/journal/<journal>/template')
def journal_template(journal):
    """Шаблон Excel для загрузки приказов в журнал (заполните лист и импортируйте)."""
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    _need_xlsx()
    conn = get_db()
    types = conn.execute('SELECT name FROM order_types WHERE kind=? ORDER BY name',
                         (journal,)).fetchall()
    maxn = 0
    for r in conn.execute('SELECT order_number FROM orders WHERE journal=?', (journal,)):
        base = _norm(r['order_number']).split('-')[0]
        if base.isdigit():
            maxn = max(maxn, int(base))
    pfx = journal_prefix(conn, journal)   # префикс из справочника видов приказов (v1.5.0)
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = JOURNAL_TITLES[journal]
    ws.append(_orders_xlsx_headers(journal))
    # примеры заполнения
    ws.append([f'{maxn + 1}{pfx}', date.today().isoformat(),
               JOURNAL_TITLES[journal], types[0]['name'] if types else 'Наименование приказа',
               'Иванов Иван Иванович', 'Петров П. П.',
               *( ['Заявление гражданина'] if REASONS_ENABLED.get(journal) else [] ),
               'Примечание (необязательно)'])
    ws.append(['', '', JOURNAL_TITLES[journal],
               types[1]['name'] if len(types) > 1 else 'Другой приказ',
               'Сидорова Анна Петровна', '',
               *( ['Решение суда'] if REASONS_ENABLED.get(journal) else [] ),
               ''])
    for c in ws[1]:
        c.font = Font(bold=True)
    widths = (12, 14, 16, 40, 32, 26, 30, 30)
    for col, w in zip('ABCDEFGH', widths[:len(_orders_xlsx_headers(journal))]):
        ws.column_dimensions[col].width = w
    return _mem_xlsx(wb, f'шаблон_приказы_{journal}.xlsx')


@app.route('/journal/<journal>/import', methods=['POST'])
def journal_import(journal):
    """Массовая загрузка приказов в журнал из файла Excel (.xlsx).

    Формат: первая строка — заголовки (см. шаблон). Столбцы «№ приказа» и
    «Дата приказа» необязательны: номер присваивается автоматически (с суффиксом
    журнала), дата — текущая. Наименование приказа сверяется со справочником
    «Виды приказов» своего журнала; при отсутствии — вид приказа создаётся.
    Дубликаты (тот же номер либо дата+наименование+ФИО) пропускаются.
    Скан-коды PDF через Excel не загружаются — их можно приложить к приказу
    вручную после импорта.
    """
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    _need_xlsx()
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'):
        flash('Выберите файл Excel в формате .xlsx', 'err')
        return _redirect_back('journal_list', journal=journal)
    try:
        rows = _read_xlsx_rows(io.BytesIO(f.read()))
    except Exception as e:
        flash(f'Не удалось прочитать файл: {e}', 'err')
        return _redirect_back('journal_list', journal=journal)
    if len(rows) < 2:
        flash('В файле нет данных (нужны заголовок и хотя бы одна строка)', 'err')
        return _redirect_back('journal_list', journal=journal)
    m = _cell_map(rows[0])
    ci_num = _pick(m, '№приказа', 'нприказа', 'номерприказа', 'номер', 'num')
    ci_date = _pick(m, 'датаприказа', 'дата', 'date')
    ci_kind = _pick(m, 'видприказа', 'вид', 'журнал')
    ci_name = _pick(m, 'наименованиеприказа', 'наименование', 'приказ')
    ci_person = _pick(m, 'фиовотношениикогоприказ', 'фиовотношениикого',
                      'фио', 'фивоотношениикого')
    ci_spec = _pick(m, 'фиоспециалиста', 'специалист', 'фиоспеца')
    ci_note = _pick(m, 'примечание', 'комментарий', 'note')
    # столбец «Причина» — только для журнала «Опека» (справочник «Причины»)
    ci_reason = (_pick(m, 'причина', 'основание', 'reason')
                 if REASONS_ENABLED.get(journal) else None)
    if ci_name is None and ci_person is None:
        flash('В файле не найдены столбцы «Наименование приказа» / «Ф.И.О. в отношении кого»', 'err')
        return _redirect_back('journal_list', journal=journal)
    title_to_key = {t.lower(): k for k, t in JOURNAL_TITLES.items()}
    today = date.today().isoformat()
    conn = get_db()
    created = skipped = 0
    errors = []

    def cell(row, idx):
        return row[idx] if idx is not None and idx < len(row) else ''

    # существующие номера и «отпечатки» приказов журнала — для защиты от дублей
    existing_nums = set()
    existing_fp = set()
    for r in conn.execute('SELECT o.order_number, o.order_date, o.person_fio, '
                          't.name AS tname FROM orders o '
                          'LEFT JOIN order_types t ON t.id=o.order_type_id '
                          'WHERE o.journal=?', (journal,)):
        existing_nums.add(_norm(r['order_number']).lower())
        existing_fp.add((_norm(r['order_date']), _norm(r['tname']).lower(),
                         _norm(r['person_fio']).lower()))

    def add_order(num, dt, tid, person, spec, note, reason=None):
        nonlocal created, skipped
        fp = (dt, '', _norm(person).lower())
        type_name = ''
        if tid:
            tn = conn.execute('SELECT name FROM order_types WHERE id=?', (tid,)).fetchone()
            type_name = tn['name'] if tn else ''
            fp = (dt, _norm(type_name).lower(), _norm(person).lower())
        if num.lower() in existing_nums or fp in existing_fp:
            skipped += 1
            errors.append(f'дубликат пропущен: № {num} / {dt} / {person}')
            return
        conn.execute(
            'INSERT INTO orders(journal, order_number, order_date, order_type_id, '
            'person_fio, specialist_fio, note, reason_id, scan_path) VALUES (?,?,?,?,?,?,?,?,?)',
            (journal, num, dt, tid, person, spec, note, reason, None))
        existing_nums.add(num.lower())
        existing_fp.add(fp)
        base = num.split('-')[0]
        if base.isdigit():
            conn.execute('UPDATE counters SET value=MAX(value,?) WHERE journal=?',
                         (int(base), journal))
        created += 1

    for n, row in enumerate(rows[1:], start=2):
        kind_raw = cell(row, ci_kind)
        key = _journal_key(kind_raw, title_to_key) if kind_raw else journal
        if key and key != journal:      # строка относится к другому журналу
            skipped += 1
            errors.append(f'строка {n}: вид приказа «{kind_raw}» не относится к журналу '
                          f'«{JOURNAL_TITLES[journal]}»')
            continue
        person = cell(row, ci_person)
        name = cell(row, ci_name)
        if not person and not name:
            skipped += 1
            continue
        tid = None
        if name:
            trow = conn.execute('SELECT id FROM order_types WHERE kind=? AND '
                                'LOWER(TRIM(name))=?', (journal, _norm(name).lower())).fetchone()
            if trow:
                tid = trow['id']
            else:  # наименование ещё не есть в справочнике — регистрируем его
                try:
                    cur = conn.execute('INSERT INTO order_types(kind,name) VALUES (?,?)',
                                       (journal, _norm(name)))
                    tid = cur.lastrowid
                    errors.append(f'строка {n}: вид приказа «{name}» добавлен в справочник')
                except sqlite3.IntegrityError:
                    trow = conn.execute('SELECT id FROM order_types WHERE kind=? AND '
                                        'LOWER(TRIM(name))=?',
                                        (journal, _norm(name).lower())).fetchone()
                    tid = trow['id'] if trow else None
        dt = cell(row, ci_date) or today
        dnum = _norm(dt).replace('/', '-').replace('.', '-')
        parts = dnum.split('-')
        if len(parts) == 3:
            y, mo, da = parts[0], parts[1].zfill(2), parts[2].zfill(2)
            if len(y) == 4:
                dt = f'{y}-{mo}-{da}'
            elif len(y) <= 2:           # формат ДД-ММ-ГГ -> ГГГГ-ММ-ДД
                dt = f'20{y.zfill(2)}-{mo}-{da}'
        num = cell(row, ci_num)
        if num:
            num = _norm(num)
            if '-' not in num:          # добавляем префикс вида приказа: 15 -> 15-ОП
                num += journal_prefix(conn, journal)
        else:
            num = next_order_number(conn, journal)
        # причина из столбца «Причина» (журнал «Опека») — регистрируется в справочнике
        reason_id = (_get_or_create_reason(conn, cell(row, ci_reason))
                     if ci_reason is not None else None)
        add_order(num, dt, tid, person, cell(row, ci_spec), cell(row, ci_note),
                  reason_id)
    conn.commit()
    conn.close()
    audit('ИМПОРТ ПРИКАЗОВ ИЗ EXCEL',
          f'журнал={JOURNAL_TITLES.get(journal)}, файл={f.filename}, '
          f'создано={created}, пропущено={skipped}')
    msg = f'Импорт завершён: создано — {created}, пропущено — {skipped}'
    if errors:
        msg += '. Замечания: ' + '; '.join(errors[:8])
    flash(msg, 'ok')
    return _redirect_back('journal_list', journal=journal)


def _parse_order_filename(fname):
    """Разбор имени файла скана в реквизиты приказа.

    Формат: №приказа_ГГГГ-ММ-ДД_ФамилияИмяОтчество_Примечание.pdf
    Пример: 12-оп_2026-10-15_Иванов Иван Иванович_о назначении опекуна.pdf
    Необязательные части пропускаются; разбор максимально терпимый:
      - номер — первое поле до '_' (или вся строка, если полей нет);
      - дата — любое поле вида ГГГГ-ММ-ДД / ДД.ММ.ГГГГ / ДД-ММ-ГГГГ;
      - примечание — поле после ФИО (если полей больше трёх).
    Возвращает dict(number, date, person, note) — значения могут быть ''.
    """
    base = os.path.splitext(os.path.basename(fname))[0]
    parts = [p.strip() for p in base.split('_')]
    number = date_v = person = note = ''
    if parts and parts[0]:
        number = parts[0]
    for p in parts[1:]:
        d = _norm(p).replace('.', '-').replace('/', '-')
        m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', d) or \
            re.match(r'^(\d{1,2})-(\d{1,2})-(\d{2,4})$', d)
        if m:
            if len(m.group(1)) == 4:
                y, mo, da = m.group(1), m.group(2), m.group(3)
            else:
                y2 = m.group(3)
                y = ('20' + y2.zfill(2)) if len(y2) <= 2 else y2
                mo, da = m.group(1), m.group(2)
            date_v = f'{y}-{mo.zfill(2)}-{da.zfill(2)}'
            continue
        if not person:
            person = p
        elif not note:
            note = p
    return {'number': number, 'date': date_v, 'person': person, 'note': note}


def _find_type_id(conn, journal, name):
    """Ищет вид приказа (order_types) по наименованию для журнала; не найдено — создаёт."""
    if not name:
        return None
    r = conn.execute('SELECT id FROM order_types WHERE kind=? AND LOWER(TRIM(name))=?',
                     (journal, _norm(name).lower())).fetchone()
    if r:
        return r['id']
    try:
        cur = conn.execute('INSERT INTO order_types(kind,name) VALUES (?,?)',
                           (journal, _norm(name)))
        conn.commit()
        return cur.lastrowid
    except sqlite3.IntegrityError:
        r = conn.execute('SELECT id FROM order_types WHERE kind=? AND LOWER(TRIM(name))=?',
                         (journal, _norm(name).lower())).fetchone()
        return r['id'] if r else None


@app.route('/service/scans')
def scans_upload_page():
    """Страница пакетной загрузки PDF-сканов приказов (перенесена в сервис)."""
    if not is_admin():
        abort(403)
    conn = get_db()
    all_types = {}
    for j in JOURNAL_TYPES:
        all_types[j] = conn.execute(
            'SELECT id, name FROM order_types WHERE kind=? ORDER BY name', (j,)).fetchall()
    conn.close()
    return render_template('scans_upload.html', journals=JOURNAL_TYPES,
                           titles=JOURNAL_TITLES, types_by_journal=all_types)


@app.route('/api/scans/upload', methods=['POST'])
def api_scans_upload():
    """Пакетная загрузка сканов приказов (PDF) в журнал из архива ZIP или списка файлов.

    Обслуживает страницу «Сервис → Загрузка документов». Правила:
      - принимаются только PDF (и .zip с PDF внутри);
      - если имя файла совпадает с № существующего приказа без скана — скан
        прикрепляется к нему;
      - иначе создаётся новый приказ: реквизиты берутся из имени файла
        (№приказа_дата_ФИО_примечание.pdf); недостающий номер присваивается
        автоматически, дата — текущая;
      - дубликаты (та же пара дата+ФИО) пропускаются;
      - столбец «Наименование приказа» из реестров не используется:
        наименование всегда выбирается на странице загрузки.
    Ответ — JSON: {attached, created, skipped, errors[], message}.
    """
    if not is_admin():
        return jsonify({'error': 'Доступ запрещён: раздел доступен только Администратору'}), 403
    journal = request.form.get('journal', '')
    if journal not in JOURNAL_TYPES:
        return jsonify({'error': 'Не выбран журнал для загрузки'}), 400
    type_id = request.form.get('order_type_id', type=int)
    spec_fio = request.form.get('specialist_fio', '').strip()
    note_extra = request.form.get('note', '').strip()
    mode = request.form.get('mode', 'auto')   # auto | new
    files = request.files.getlist('files')
    single = request.files.get('file')
    if single and single.filename:
        files.append(single)
    if not files or not any(f.filename for f in files):
        return jsonify({'error': 'Выберите файлы PDF или ZIP-архив с файлами PDF'}), 400

    today = date.today().isoformat()
    workdir = tempfile.mkdtemp(prefix='scans_')
    items = []          # (имя_файла_без_расширения, путь_к_pdf_во_временном_каталоге)
    errors = []
    try:
        import zipfile
        for f in files:
            fn = f.filename or ''
            low = fn.lower()
            if not (low.endswith('.pdf') or low.endswith('.zip')):
                errors.append(f'«{os.path.basename(fn)}»: пропускается — не PDF/ZIP')
                continue
            tmp_path = os.path.join(workdir, secure_filename(fn) or 'f.bin')
            f.save(tmp_path)
            if low.endswith('.zip'):
                try:
                    with zipfile.ZipFile(tmp_path) as z:
                        bad = [n for n in z.namelist()
                               if not n.lower().endswith('.pdf') and not n.endswith('/')]
                        if bad:
                            errors.append(f'в архиве «{os.path.basename(fn)}» '\
                                          f'пропущено не-PDF файлов: {len(bad)}')
                        zdir = os.path.join(workdir, 'z_' + secure_filename(fn))
                        os.makedirs(zdir, exist_ok=True)
                        for n in z.namelist():
                            if not n.lower().endswith('.pdf') or n.startswith('__MACOSX'):
                                continue
                            out = os.path.join(zdir, secure_filename(os.path.basename(n)))
                            with z.open(n) as src, open(out, 'wb') as dst:
                                shutil.copyfileobj(src, dst)
                            items.append((os.path.splitext(os.path.basename(n))[0], out))
                except zipfile.BadZipFile:
                    errors.append(f'«{os.path.basename(fn)}»: повреждённый ZIP-архив')
            else:
                items.append((os.path.splitext(os.path.basename(fn))[0], tmp_path))

        attached = created = skipped = 0
        conn = get_db()
        existing_nums = {}
        existing_fp = set()
        for r in conn.execute('SELECT o.id, o.order_number, o.order_date, o.person_fio, '
                              'o.scan_path, t.name AS tname FROM orders o '
                              'LEFT JOIN order_types t ON t.id=o.order_type_id '
                              'WHERE o.journal=?', (journal,)):
            key = _norm(r['order_number']).lower()
            existing_nums.setdefault(key, []).append(r)
            existing_fp.add((_norm(r['order_date']), _norm(r['tname']).lower(),
                             _norm(r['person_fio']).lower()))
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        for stem, path in items:
            meta = _parse_order_filename(stem)
            num = meta['number']
            dt = meta['date'] or today
            person = meta['person']
            note = meta['note']
            if note_extra:
                note = (note + '; ' + note_extra) if note else note_extra
            has_suffix = '-' in num
            if not has_suffix and num and num.split('-')[0].isdigit():
                num += journal_prefix(conn, journal)   # префикс из справочника (v1.5.0)
            recs = existing_nums.get(_norm(num).lower(), []) if num else []
            if mode != 'new' and recs:
                rec = next((x for x in recs if not x['scan_path']), recs[0])
                if rec['scan_path']:
                    skipped += 1
                    errors.append(f'«{stem}»: у приказа № {num} уже есть скан — пропущен')
                    continue
                fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(os.path.basename(path))}"
                shutil.copy2(path, os.path.join(UPLOAD_DIR, fname))
                conn.execute('UPDATE orders SET scan_path=? WHERE id=?', (fname, rec['id']))
                conn.commit()
                attached += 1
                continue
            if mode == 'new' and num:
                num = next_order_number(conn, journal)
            elif num and recs:
                num = next_order_number(conn, journal)
            elif not num:
                num = next_order_number(conn, journal)
            fp = (dt, '', _norm(person).lower())
            if fp in existing_fp:
                skipped += 1
                errors.append(f'«{stem}»: дубликат ({dt} / {person}) — пропущен')
                continue
            fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(os.path.basename(path))}"
            shutil.copy2(path, os.path.join(UPLOAD_DIR, fname))
            conn.execute(
                'INSERT INTO orders(journal, order_number, order_date, order_type_id, '
                'person_fio, specialist_fio, note, scan_path) VALUES (?,?,?,?,?,?,?,?)',
                (journal, num, dt, type_id, person, spec_fio, note, fname))
            base = num.split('-')[0]
            if base.isdigit():
                conn.execute('UPDATE counters SET value=MAX(value,?) WHERE journal=?',
                             (int(base), journal))
            conn.commit()
            existing_nums.setdefault(_norm(num).lower(), []).append(
                {'id': None, 'order_number': num, 'order_date': dt,
                 'person_fio': person, 'scan_path': fname, 'tname': ''})
            existing_fp.add(fp)
            created += 1
        conn.close()
        audit('ЗАГРУЗКА СКАНОВ PDF',
              f'журнал={JOURNAL_TITLES.get(journal)}, приложено={attached}, '
              f'создано={created}, пропущено={skipped}')
        msg = (f'Загрузка завершена: сканов приложено — {attached}, '
               f'приказов создано — {created}, пропущено — {skipped}')
        return jsonify({'attached': attached, 'created': created,
                        'skipped': skipped, 'errors': errors[:20],
                        'message': msg})
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.route('/order/<int:oid>/update', methods=['POST'])
def order_update(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close()
        abort(404)
    journal = row['journal']
    if not can_access_journal(current_user(), journal):
        conn.close()
        abort(403)
    order_number = request.form.get('order_number', '').strip() or row['order_number']
    # если номер изменён на число без префикса вида приказа — добавляем префикс (v1.5.0)
    if '-' not in order_number and order_number.split('-')[0].isdigit():
        order_number += journal_prefix(conn, journal)
    order_date = request.form.get('order_date') or row['order_date']
    type_id = request.form.get('order_type_id', type=int)
    person_fio = request.form.get('person_fio', '').strip()
    specialist_fio = request.form.get('specialist_fio', '').strip()
    note = request.form.get('note', '').strip()          # примечание к приказу
    # причина приказа (журнал «Опека») — берётся из справочника «Причины»
    reason_id = row['reason_id'] if 'reason_id' in row.keys() else None
    if REASONS_ENABLED.get(journal):
        rid = request.form.get('reason_id', type=int)
        if rid:
            reason_id = rid if conn.execute('SELECT 1 FROM reasons WHERE id=?',
                                            (rid,)).fetchone() else None
        elif request.form.get('reason') is not None:
            reason_id = _get_or_create_reason(conn, request.form.get('reason'))
    scan_rel = row['scan_path']
    f = request.files.get('scan')
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext != '.pdf':
            conn.close()
            audit('ОТКАЗ: СКАН НЕ В ФОРМАТЕ PDF', f.filename)
            flash('Скан должен быть в формате PDF', 'err')
            return redirect(url_for('journal_list', journal=journal))
        fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(f.filename)}"
        os.makedirs(UPLOAD_DIR, exist_ok=True)
        f.save(os.path.join(UPLOAD_DIR, fname))
        if scan_rel and os.path.exists(os.path.join(UPLOAD_DIR, scan_rel)):
            os.remove(os.path.join(UPLOAD_DIR, scan_rel))
        scan_rel = fname
    conn.execute(
        'UPDATE orders SET order_number=?, order_date=?, order_type_id=?, '
        'person_fio=?, specialist_fio=?, note=?, reason_id=?, scan_path=? WHERE id=?',
        (order_number, order_date, type_id, person_fio, specialist_fio,
         note, reason_id, scan_rel, oid))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ ИЗМЕНЁН', f'id={oid}')
    flash('Изменения сохранены', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/order/<int:oid>/delete', methods=['POST'])
def order_delete(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close()
        abort(404)
    journal = row['journal']
    if not can_access_journal(current_user(), journal):
        conn.close()
        abort(403)
    if row['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, row['scan_path'])):
        os.remove(os.path.join(UPLOAD_DIR, row['scan_path']))
    conn.execute('DELETE FROM orders WHERE id=?', (oid,))
    conn.commit()
    conn.close()
    audit('ПРИКАЗ УДАЛЁН', f'id={oid}')
    flash('Приказ удалён', 'ok')
    return redirect(url_for('journal_list', journal=journal))


@app.route('/journal/<journal>/delete_bulk', methods=['POST'])
def journal_delete_bulk(journal):
    """Массовое удаление выделенных в журнале приказов (чекбоксы строк)."""
    if journal not in JOURNAL_TYPES:
        abort(404)
    user = current_user()
    if not can_access_journal(user, journal):
        abort(403)
    ids = []
    for v in request.form.getlist('ids'):
        try:
            ids.append(int(v))
        except (ValueError, TypeError):
            abort(400)
    if not ids:
        flash('Не выбран ни один приказ для удаления', 'err')
        return redirect(url_for('journal_list', journal=journal))
    conn = get_db()
    marks = ','.join('?' * len(ids))
    rows = conn.execute(f'SELECT id, scan_path FROM orders WHERE journal=? AND id IN ({marks})',
                        [journal] + ids).fetchall()
    deleted = 0
    for r in rows:
        if r['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, r['scan_path'])):
            try:
                os.remove(os.path.join(UPLOAD_DIR, r['scan_path']))
            except OSError:
                pass
    if rows:
        conn.execute(f'DELETE FROM orders WHERE journal=? AND id IN ({marks})',
                     [journal] + ids)
        conn.commit()
        deleted = len(rows)
    skipped = len(ids) - deleted
    conn.close()
    if deleted:
        audit('ПРИКАЗЫ УДАЛЕНЫ МАССОВО', f'журнал={journal} count={deleted} ids={ids}')
        msg = f'Удалено приказов: {deleted}'
        if skipped > 0:
            msg += f', пропущено (не найдены или из другого журнала): {skipped}'
        flash(msg, 'ok')
    else:
        flash('Удалять нечего: приказы не найдены', 'err')
    # возвращаемся с текущим фильтром по периоду и поиском
    qs = {}
    for a in ('d1', 'd2', 'q'):
        v = request.args.get(a, '').strip()
        if v:
            qs[a] = v
    return redirect(url_for('journal_list', journal=journal, **qs))


@app.route('/order/<int:oid>/scan')
def order_scan(oid):
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    conn.close()
    if not row or not row['scan_path']:
        abort(404)
    if not can_access_journal(current_user(), row['journal']):
        abort(403)
    path = os.path.join(UPLOAD_DIR, os.path.basename(row['scan_path']))
    if not os.path.exists(path):
        abort(404)
    return send_file(path, mimetype='application/pdf')

# ---------------------------------------------------------------- отчёты

def _report_period():
    d1 = request.args.get('d1') or date.today().replace(day=1).isoformat()
    d2 = request.args.get('d2') or date.today().isoformat()
    return d1, d2


def _query_report_rows(d1, d2, journal=None):
    conn = get_db()
    sql = ('SELECT o.*, t.name AS type_name, rs.name AS reason_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id '
           'LEFT JOIN reasons rs ON rs.id=o.reason_id '
           'WHERE o.order_date>=? AND o.order_date<=?')
    params = [d1, d2]
    if journal:
        sql += ' AND o.journal=?'
        params.append(journal)
    # сортировка документов: по дате, внутри даты - по номеру
    sql += (' ORDER BY o.order_date, ' + _ORDER_NUM_SQL.format(col='o.order_number') +
            ', o.order_number')
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


def _query_report_rows_sorted_asc(d1, d2, journal=None):
    """Строки отчёта «распечатка журнала»: сортировка по возрастанию —
    сначала дата приказа, внутри одной даты — номер (числовое сравнение:
    2 < 9 < 12 < 12а < 105). Используется в печати/выгрузке Excel и Word."""
    conn = get_db()
    sql = ('SELECT o.*, t.name AS type_name, rs.name AS reason_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id '
           'LEFT JOIN reasons rs ON rs.id=o.reason_id '
           'WHERE o.order_date>=? AND o.order_date<=?')
    params = [d1, d2]
    if journal:
        sql += ' AND o.journal=?'
        params.append(journal)
    sql += (' ORDER BY o.order_date ASC, ' + _ORDER_NUM_SQL.format(col='o.order_number') +
            ' ASC, o.order_number ASC')
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


@app.route('/reports')
def reports():
    user = current_user()
    if not can_view_reports(user):
        abort(403)
    return render_template('reports.html')


@app.route('/reports/by_period')
def report_by_period():
    user = current_user()
    if not can_view_reports(user):
        abort(403)
    d1, d2 = _report_period()
    rows = _query_report_rows(d1, d2)
    agg = {}
    for r in rows:
        agg[r['journal']] = agg.get(r['journal'], 0) + 1
    data = [(JOURNAL_TITLES[j], agg.get(j, 0)) for j in JOURNAL_TYPES]
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_simple(d1, d2, data, 'Приказы за период')
    if fmt == 'docx':
        return _export_docx_simple(d1, d2, data, 'Приказы за период')
    return render_template('report_table.html', title='Приказы за период',
                           d1=d1, d2=d2, headers=('Вид приказа', 'Количество приказов'),
                           rows=data, total=sum(x[1] for x in data),
                           base_url=url_for('report_by_period'))


@app.route('/reports/by_period_detail')
def report_by_period_detail():
    user = current_user()
    if not can_view_reports(user):
        abort(403)
    d1, d2 = _report_period()
    rows = _query_report_rows(d1, d2)
    matrix = {}
    col_names = set()
    for r in rows:
        kind = JOURNAL_TITLES[r['journal']]
        name = r['type_name'] or '(не указан)'
        col_names.add(name)
        matrix.setdefault(kind, {})
        matrix[kind][name] = matrix[kind].get(name, 0) + 1
    cols = sorted(col_names)
    table = []
    for kind in [JOURNAL_TITLES[j] for j in JOURNAL_TYPES]:
        if kind in matrix:
            table.append([kind] + [matrix[kind].get(c, 0) for c in cols])
    totals = [sum(row[i] for row in table) for i in range(1, len(cols) + 1)] if table else []
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_matrix(d1, d2, ('Вид приказа',) + tuple(cols), table, totals,
                                   'Приказы за период с расшифровкой')
    if fmt == 'docx':
        return _export_docx_matrix(d1, d2, ('Вид приказа',) + tuple(cols), table, totals,
                                   'Приказы за период с расшифровкой')
    return render_template('report_matrix.html', title='Приказы за период с расшифровкой',
                           d1=d1, d2=d2, cols=cols, table=table, totals=totals,
                           base_url=url_for('report_by_period_detail'))


@app.route('/reports/journal_print')
def report_journal_print():
    user = current_user()
    if not can_view_reports(user):
        abort(403)
    journal = request.args.get('journal', 'opeka')
    if journal not in JOURNAL_TYPES:
        abort(404)
    d1, d2 = _report_period()
    # выборка по виду приказа; сортировка: по дате, внутри даты — по номеру
    rows = _query_report_rows_sorted_asc(d1, d2, journal=journal)
    title = f'Журнал «{JOURNAL_TITLES[journal]}» за период'
    # для журнала «Опека»: наименование приказа = вид приказа + (причина из справочника);
    # для «Нуждаемость» и «Профилактика» — только наименование приказа
    with_reason = REASONS_ENABLED.get(journal, False)
    data = []
    for i, r in enumerate(rows, 1):
        # примечание — отдельным столбцом (по ТЗ)
        note = r['note'] if 'note' in r.keys() else ''
        name = r['type_name'] or ''
        reason = ''
        if with_reason and 'reason_name' in r.keys():
            reason = r['reason_name'] or ''
        if reason:
            name = f'{name} ({reason})' if name else f'({reason})'
        data.append((i, r['order_number'], r['order_date'], r['person_fio'],
                     name, r['specialist_fio'], note))
    headers = ('№ п/п', '№ приказа', 'Дата', 'Ф.И.О.',
               'Наименование приказа', 'ФИО специалиста', 'Примечание')
    fmt = request.args.get('fmt')
    if fmt == 'xlsx':
        return _export_xlsx_table(d1, d2, headers, data, title)
    if fmt == 'docx':
        return _export_docx_table(d1, d2, headers, data, title)
    return render_template('report_table.html', title=title, d1=d1, d2=d2,
                           headers=headers, rows=data, total=len(data),
                           base_url=url_for('report_journal_print'),
                           extra_args={'journal': journal})

# ============================================================================
# REST API v1 — используется клиентским десктопным приложением (client_app.py).
# Аутентификация: Basic (логин/пароль сотрудника) либо cookie-сессия веб-версии.
# Все ответы — JSON в кодировке UTF-8. Права доступа идентичны веб-интерфейсу.
# ============================================================================

API_JOURNAL_FIELDS = ('order_number', 'order_date', 'order_type_id',
                      'person_fio', 'specialist_fio', 'note')


def api_user():
    """Возвращает пользователя по Basic-auth или сессии; иначе 401."""
    auth = request.authorization
    if auth and auth.type == 'basic' and auth.username:
        conn = get_db()
        u = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                         'LEFT JOIN roles r ON r.id=e.role_id '
                         'WHERE e.login=? AND e.password_hash IS NOT NULL',
                         (auth.username,)).fetchone()
        conn.close()
        if u and check_password_hash(u['password_hash'], auth.password or ''):
            return u
        audit('НЕУДАЧНЫЙ ВХОД (API)', f'login={auth.username}')
        abort(Response(json.dumps({'error': 'Неверный логин или пароль'},
                                  ensure_ascii=False), status=401,
                       mimetype='application/json'))
    uid = session.get('user_id')
    if uid:
        u = current_user()
        if u:
            return u
    abort(Response(json.dumps({'error': 'Требуется авторизация'},
                              ensure_ascii=False), status=401,
                   mimetype='application/json'))


def api_json(obj, status=200):
    return Response(json.dumps(obj, ensure_ascii=False, default=str),
                    status=status, mimetype='application/json')


def _order_row_to_dict(r):
    keys = r.keys()
    d = {'id': r['id'], 'journal': r['journal'],
         'order_number': r['order_number'], 'order_date': r['order_date'],
         'order_type_id': r['order_type_id'],
         'order_type_name': r['type_name'] if 'type_name' in keys else None,
         'person_fio': r['person_fio'], 'specialist_fio': r['specialist_fio'],
         'note': r['note'] if 'note' in keys else '',
         'has_scan': bool(r['scan_path']),
         'created_at': r['created_at'] if 'created_at' in keys else ''}
    # причина приказа (поле журнала «Опека», справочник «Причины»)
    if 'reason_id' in keys:
        d['reason_id'] = r['reason_id']
    if 'reason_name' in keys:
        d['reason_name'] = r['reason_name']
    return d


@app.route('/api/v1/session')
def api_session():
    """Проверка учётных данных и прав для клиента: роль, журналы, отчёты."""
    u = api_user()
    journals = [{'key': j, 'title': JOURNAL_TITLES[j],
                 'can_edit': can_access_journal(u, j)}
                for j in JOURNAL_TYPES if can_access_journal(u, j)]
    return api_json({'ok': True, 'id': u['id'], 'full_name': u['full_name'],
                     'position': u['position'], 'role': u['role_name'],
                     'journals': journals,
                     'reports': can_view_reports(u),
                     'employees_dir': can_edit_employees(u),
                     'order_types_dir': can_edit_order_types(u),
                     'version': APP_VERSION})


@app.route('/api/v1/order_types', methods=['GET', 'POST'])
def api_order_types():
    u = api_user()
    if request.method == 'POST':                      # добавление/изменение вида
        if not can_edit_order_types(u):
            abort(403)
        data = request.get_json(silent=True) or {}
        kind = (data.get('kind') or '').strip()
        name = (data.get('name') or '').strip()
        prefix = _norm((data.get('prefix') or '')).strip()  # префикс номера приказа (v1.5.0)
        tid = data.get('id')
        if kind not in JOURNAL_TYPES or not name:
            return api_json({'error': 'Укажите журнал (вид приказа) и наименование'}, 400)
        conn = get_db()
        try:
            if tid:
                if not conn.execute('SELECT 1 FROM order_types WHERE id=?',
                                    (int(tid),)).fetchone():
                    conn.close(); abort(404)
                dup = conn.execute(
                    'SELECT 1 FROM order_types WHERE kind=? AND name=? AND id<>?',
                    (kind, name, int(tid))).fetchone()
                if dup:
                    conn.close()
                    return api_json({'error': 'Такой вид приказа уже существует'}, 400)
                conn.execute('UPDATE order_types SET kind=?, name=?, prefix=? WHERE id=?',
                             (kind, name, prefix, int(tid)))
                action = 'ВИД ПРИКАЗА ИЗМЕНЁН (API)'
            else:
                dup = conn.execute('SELECT 1 FROM order_types WHERE kind=? AND name=?',
                                   (kind, name)).fetchone()
                if dup:
                    conn.close()
                    return api_json({'error': 'Такой вид приказа уже существует'}, 400)
                conn.execute('INSERT INTO order_types(kind,name,prefix) VALUES (?,?,?)',
                             (kind, name, prefix))
                action = 'ВИД ПРИКАЗА СОХРАНЁН (API)'
            conn.commit()
        finally:
            conn.close()
        audit(action, f'{JOURNAL_TITLES.get(kind, kind)}: {name}',
              user=f"{u['full_name']} ({u['login']})")
        return api_json({'ok': True}, 201 if not tid else 200)
    journal = request.args.get('journal')
    # GET /api/v1/reasons — справочник «Причины» (поле журнала «Опека»)
    if request.args.get('reasons'):
        conn = get_db()
        rows = conn.execute('SELECT * FROM reasons ORDER BY name').fetchall()
        conn.close()
        return api_json([{'id': r['id'], 'name': r['name']} for r in rows])
    conn = get_db()
    if journal:
        if journal not in JOURNAL_TYPES or not can_access_journal(u, journal):
            conn.close(); abort(403)
        rows = conn.execute('SELECT * FROM order_types WHERE kind=? ORDER BY name',
                            (journal,)).fetchall()
    else:
        rows = conn.execute('SELECT * FROM order_types ORDER BY kind, name').fetchall()
    conn.close()
    return api_json([{'id': r['id'], 'kind': r['kind'],
                      'kind_title': JOURNAL_TITLES.get(r['kind'], r['kind']),
                      'name': r['name']} for r in rows])


@app.route('/api/v1/order_types/<int:tid>', methods=['DELETE'])
def api_order_type_delete(tid):
    """Удаление вида приказа (куратор и выше); запрещено, если вид используется."""
    u = api_user()
    if not can_edit_order_types(u):
        abort(403)
    conn = get_db()
    row = conn.execute('SELECT * FROM order_types WHERE id=?', (tid,)).fetchone()
    if not row:
        conn.close(); abort(404)
    used = conn.execute('SELECT COUNT(*) c FROM orders WHERE order_type_id=?',
                        (tid,)).fetchone()['c']
    if used:
        conn.close()
        return api_json({'error': f'Вид используется в приказах ({used} шт.), '
                                  'удаление запрещено'}, 400)
    conn.execute('DELETE FROM order_types WHERE id=?', (tid,))
    conn.commit(); conn.close()
    audit('ВИД ПРИКАЗА УДАЛЁН (API)',
          f'{JOURNAL_TITLES.get(row["kind"], row["kind"])}: {row["name"]}',
          user=f"{u['full_name']} ({u['login']})")
    return api_json({'ok': True})


@app.route('/api/v1/journal/<journal>/orders')
def api_orders(journal):
    u = api_user()
    if journal not in JOURNAL_TYPES or not can_access_journal(u, journal):
        abort(403)
    q = request.args.get('q', '').strip()
    d1 = request.args.get('d1'); d2 = request.args.get('d2')
    sql = ('SELECT o.*, t.name AS type_name, rs.name AS reason_name FROM orders o '
           'LEFT JOIN order_types t ON t.id=o.order_type_id '
           'LEFT JOIN reasons rs ON rs.id=o.reason_id WHERE o.journal=?')
    params = [journal]
    if q:
        sql += (' AND (o.order_number LIKE ? OR o.person_fio LIKE ? '
                'OR o.specialist_fio LIKE ? OR o.note LIKE ?)')
        params += [f'%{q}%'] * 4
    if d1:
        sql += ' AND o.order_date>=?'; params.append(d1)
    if d2:
        sql += ' AND o.order_date<=?'; params.append(d2)
    # сортировка документов: по дате, внутри даты - по номеру
    sql += (' ORDER BY o.order_date DESC, ' + _ORDER_NUM_SQL.format(col='o.order_number') +
            ' DESC, o.order_number DESC')
    conn = get_db()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return api_json([_order_row_to_dict(r) for r in rows])


def _api_reason_id(conn, journal, data):
    """reason_id для API-запроса: только в журнале «Опека» (справочник «Причины»).

    Принимается готовый reason_id или наименование reason (регистрируется в
    справочнике автоматически). Для остальных журналов — всегда None.
    """
    if not REASONS_ENABLED.get(journal):
        return None
    rid = data.get('reason_id')
    if rid:
        ok = conn.execute('SELECT 1 FROM reasons WHERE id=?', (int(rid),)).fetchone()
        if ok:
            return int(rid)
    return _get_or_create_reason(conn, data.get('reason') or '')


@app.route('/api/v1/journal/<journal>/orders', methods=['POST'])
def api_order_create(journal):
    u = api_user()
    if journal not in JOURNAL_TYPES or not can_access_journal(u, journal):
        abort(403)
    data = request.get_json(silent=True) or {}
    conn = get_db()
    order_number = (data.get('order_number') or '').strip() \
        or next_order_number(conn, journal)
    # номер без префикса вида приказа дополняем префиксом из справочника (v1.5.0)
    if '-' not in order_number and order_number.split('-')[0].isdigit():
        order_number += journal_prefix(conn, journal)
    base = order_number.split('-')[0]
    if base.isdigit():
        conn.execute('UPDATE counters SET value=MAX(value,?) WHERE journal=?',
                     (int(base), journal))
    order_date = (data.get('order_date') or '').strip() or date.today().isoformat()
    type_id = data.get('order_type_id')
    if type_id:
        ok = conn.execute('SELECT 1 FROM order_types WHERE id=? AND kind=?',
                          (int(type_id), journal)).fetchone()
        type_id = int(type_id) if ok else None
    cur = conn.execute(
        'INSERT INTO orders(journal, order_number, order_date, order_type_id, '
        'person_fio, specialist_fio, note, reason_id) VALUES (?,?,?,?,?,?,?,?)',
        (journal, order_number, order_date, type_id,
         (data.get('person_fio') or '').strip(),
         (data.get('specialist_fio') or '').strip(),
         (data.get('note') or '').strip(),
         _api_reason_id(conn, journal, data)))
    conn.commit()
    oid = cur.lastrowid
    conn.close()
    audit('ПРИКАЗ СОЗДАН (API)',
          f'журнал={JOURNAL_TITLES.get(journal)}, №={order_number}, по кому='
          f"{data.get('person_fio', '')}", user=f"{u['full_name']} ({u['login']})")
    return api_json({'ok': True, 'id': oid, 'order_number': order_number}, 201)


@app.route('/api/v1/orders/<int:oid>', methods=['GET', 'PUT', 'DELETE'])
def api_order_item(oid):
    u = api_user()
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close(); abort(404)
    journal = row['journal']
    if not can_access_journal(u, journal):
        conn.close(); abort(403)
    if request.method == 'GET':
        full = conn.execute('SELECT o.*, t.name AS type_name, rs.name AS reason_name '
                            'FROM orders o '
                            'LEFT JOIN order_types t ON t.id=o.order_type_id '
                            'LEFT JOIN reasons rs ON rs.id=o.reason_id '
                            'WHERE o.id=?', (oid,)).fetchone()
        conn.close()
        return api_json(_order_row_to_dict(full))
    if request.method == 'DELETE':
        if row['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, row['scan_path'])):
            os.remove(os.path.join(UPLOAD_DIR, row['scan_path']))
        conn.execute('DELETE FROM orders WHERE id=?', (oid,))
        conn.commit(); conn.close()
        audit('ПРИКАЗ УДАЛЁН (API)', f'id={oid}, №={row["order_number"]}', user=f"{u['full_name']} ({u['login']})")
        return api_json({'ok': True})
    # PUT — частичное обновление разрешённых полей
    data = request.get_json(silent=True) or {}
    upd = dict(row)
    for f in API_JOURNAL_FIELDS:
        if f in data:
            upd[f] = (data[f] or '').strip() if isinstance(data[f], str) else data[f]
    # причина приказа (журнал «Опека»): reason_id или наименование reason;
    # если поле не передано в запросе — остаётся без изменений
    upd['reason_id'] = row['reason_id'] if 'reason_id' in row.keys() else None
    if REASONS_ENABLED.get(journal) and ('reason_id' in data or 'reason' in data):
        upd['reason_id'] = _api_reason_id(conn, journal, data)
    if upd['order_type_id']:
        ok = conn.execute('SELECT 1 FROM order_types WHERE id=? AND kind=?',
                          (int(upd['order_type_id']), journal)).fetchone()
        if not ok:
            upd['order_type_id'] = None
    # номер без префикса вида приказа дополняем префиксом из справочника (v1.5.0)
    _num = str(upd['order_number'] or '')
    if '-' not in _num and _num.split('-')[0].isdigit():
        upd['order_number'] = _num + journal_prefix(conn, journal)
    conn.execute('UPDATE orders SET order_number=?, order_date=?, order_type_id=?, '
                 'person_fio=?, specialist_fio=?, note=?, reason_id=? WHERE id=?',
                 (upd['order_number'], upd['order_date'], upd['order_type_id'],
                  upd['person_fio'], upd['specialist_fio'], upd['note'],
                  upd['reason_id'], oid))
    conn.commit(); conn.close()
    audit('ПРИКАЗ ИЗМЕНЁН (API)', f'id={oid}, №={upd["order_number"]}', user=f"{u['full_name']} ({u['login']})")
    return api_json({'ok': True})


@app.route('/api/v1/orders/<int:oid>/scan', methods=['GET', 'POST', 'DELETE'])
def api_order_scan(oid):
    u = api_user()
    conn = get_db()
    row = conn.execute('SELECT * FROM orders WHERE id=?', (oid,)).fetchone()
    if not row:
        conn.close(); abort(404)
    if not can_access_journal(u, row['journal']):
        conn.close(); abort(403)
    if request.method == 'GET':
        if not row['scan_path']:
            conn.close(); abort(404)
        path = os.path.join(UPLOAD_DIR, os.path.basename(row['scan_path']))
        conn.close()
        if not os.path.exists(path):
            abort(404)
        return send_file(path, mimetype='application/pdf',
                         as_attachment=request.args.get('download') == '1',
                         download_name=f"скан_{row['order_number']}.pdf")
    if request.method == 'DELETE':
        if row['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, row['scan_path'])):
            os.remove(os.path.join(UPLOAD_DIR, row['scan_path']))
        conn.execute('UPDATE orders SET scan_path=NULL WHERE id=?', (oid,))
        conn.commit(); conn.close()
        audit('СКАН УДАЛЁН (API)', f'id={oid}', user=f"{u['full_name']} ({u['login']})")
        return api_json({'ok': True})
    # POST — загрузка скана (multipart: file=...); принимается только PDF
    f = request.files.get('file')
    if not f or not f.filename:
        conn.close()
        return api_json({'error': 'Файл не передан (поле file)'}, 400)
    ext = os.path.splitext(f.filename)[1].lower()
    if ext != '.pdf':
        conn.close()
        audit('ОТКАЗ: СКАН НЕ В ФОРМАТЕ PDF (API)', f.filename, user=f"{u['full_name']} ({u['login']})")
        return api_json({'error': 'Скан должен быть в формате PDF'}, 400)
    fname = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{secure_filename(f.filename)}"
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    f.save(os.path.join(UPLOAD_DIR, fname))
    if row['scan_path'] and os.path.exists(os.path.join(UPLOAD_DIR, row['scan_path'])):
        os.remove(os.path.join(UPLOAD_DIR, row['scan_path']))
    conn.execute('UPDATE orders SET scan_path=? WHERE id=?', (fname, oid))
    conn.commit(); conn.close()
    audit('СКАН ЗАГРУЖЕН (API)', f'id={oid}, файл={fname}', user=f"{u['full_name']} ({u['login']})")
    return api_json({'ok': True, 'scan': fname}, 201)


@app.route('/api/v1/reports/by_period')
def api_report_by_period():
    u = api_user()
    if not can_view_reports(u):
        abort(403)
    d1, d2 = _report_period()
    agg = {}
    for r in _query_report_rows(d1, d2):
        agg[r['journal']] = agg.get(r['journal'], 0) + 1
    return api_json({'d1': d1, 'd2': d2,
                     'rows': [{'kind': JOURNAL_TITLES[j], 'count': agg.get(j, 0)}
                              for j in JOURNAL_TYPES],
                     'total': sum(agg.values())})


@app.route('/api/v1/reports/by_period_detail')
def api_report_by_period_detail():
    u = api_user()
    if not can_view_reports(u):
        abort(403)
    d1, d2 = _report_period()
    matrix, cols = {}, set()
    for r in _query_report_rows(d1, d2):
        kind = JOURNAL_TITLES[r['journal']]
        name = r['type_name'] or '(не указан)'
        cols.add(name)
        matrix.setdefault(kind, {})
        matrix[kind][name] = matrix[kind].get(name, 0) + 1
    cols = sorted(cols)
    table = [[kind] + [matrix[kind].get(c, 0) for c in cols]
             for kind in [JOURNAL_TITLES[j] for j in JOURNAL_TYPES] if kind in matrix]
    totals = [sum(row[i] for row in table) for i in range(1, len(cols) + 1)] if table else []
    return api_json({'d1': d1, 'd2': d2, 'cols': cols, 'table': table,
                     'totals': totals})


@app.route('/api/v1/reports/journal_print')
def api_report_journal_print():
    u = api_user()
    if not can_view_reports(u):
        abort(403)
    journal = request.args.get('journal', 'opeka')
    if journal not in JOURNAL_TYPES:
        abort(404)
    d1, d2 = _report_period()
    rows = _query_report_rows_sorted_asc(d1, d2, journal=journal)
    # та же логика, что и в веб-отчёте: Опека — наименование + (причина),
    # Нуждаемость/Профилактика — только наименование приказа
    with_reason = REASONS_ENABLED.get(journal, False)
    data = []
    for i, r in enumerate(rows, 1):
        note = r['note'] if 'note' in r.keys() else ''
        name = r['type_name'] or ''
        reason = ''
        if with_reason and 'reason_name' in r.keys():
            reason = r['reason_name'] or ''
        if reason:
            name = f'{name} ({reason})' if name else f'({reason})'
        data.append({'n': i, 'number': r['order_number'], 'date': r['order_date'],
                     'fio': r['person_fio'], 'name': name,
                     'specialist': r['specialist_fio'], 'note': note})
    return api_json({'d1': d1, 'd2': d2, 'journal': journal,
                     'title': f'Журнал «{JOURNAL_TITLES[journal]}» за период',
                     'rows': data, 'total': len(data)})


@app.route('/api/v1/employees')
def api_employees():
    u = api_user()
    conn = get_db()
    rows = conn.execute('SELECT e.*, r.name role_name FROM employees e '
                        'LEFT JOIN roles r ON r.id=e.role_id '
                        'ORDER BY e.full_name').fetchall()
    conn.close()
    # список сотрудников виден всем вошедшим (используется в журналах как подсказка),
    # но пароли/логины — только тем, кто может редактировать справочник
    out = []
    for r in rows:
        item = {'id': r['id'], 'full_name': r['full_name'],
                'position': r['position'], 'role': r['role_name']}
        if can_edit_employees(u):
            item['login'] = r['login']
        out.append(item)
    return api_json(out)


@app.route('/api/v1/info')
def api_info():
    """Публичная информация о сервере (без авторизации)."""
    conn = get_db()
    counts = {JOURNAL_TITLES[j]: conn.execute(
        'SELECT COUNT(*) c FROM orders WHERE journal=?', (j,)).fetchone()['c']
        for j in JOURNAL_TYPES}
    conn.close()
    return api_json({'app': 'Журнал регистрации приказов', 'version': APP_VERSION,
                     'orders': counts})


@app.route('/api/v1/password', methods=['POST'])
def api_change_password():
    """Самостоятельная смена пароля сотрудником (используется клиентом)."""
    u = api_user()
    data = request.get_json(silent=True) or {}
    old = data.get('old_password') or ''
    new = data.get('new_password') or ''
    confirm = data.get('confirm_password') or ''
    conn = get_db()
    row = conn.execute('SELECT * FROM employees WHERE id=?', (u['id'],)).fetchone()
    if not row['password_hash'] or not check_password_hash(row['password_hash'], old):
        conn.close()
        audit('ПАРОЛЬ НЕ ИЗМЕНЁН (API)', 'неверно указан текущий пароль',
              user=f"{u['full_name']} ({u['login']})")
        return api_json({'error': 'Текущий пароль указан неверно'}, 400)
    if len(new) < 4:
        conn.close()
        return api_json({'error': 'Новый пароль слишком короткий (минимум 4 символа)'}, 400)
    if new != confirm:
        conn.close()
        return api_json({'error': 'Новый пароль и подтверждение не совпадают'}, 400)
    if new == old:
        conn.close()
        return api_json({'error': 'Новый пароль не должен совпадать с текущим'}, 400)
    conn.execute('UPDATE employees SET password_hash=?, pwd_changed_at=? WHERE id=?',
                 (generate_password_hash(new),
                  datetime.now().strftime('%Y-%m-%d %H:%M:%S'), u['id']))
    conn.commit()
    conn.close()
    audit('ПАРОЛЬ ИЗМЕНЁН (API)', 'клиентское приложение',
          user=f"{u['full_name']} ({u['login']})")
    return api_json({'ok': True})

# ---------------------------------------------------------------- export helpers

def _mem_download(stream, mimetype, filename):
    stream.seek(0)
    return send_file(stream, mimetype=mimetype, as_attachment=True, download_name=filename)


def _need_xlsx():
    if Workbook is None:
        abort(500, description='Не установлен модуль openpyxl (pip install openpyxl)')


def _need_docx():
    if DocxDocument is None:
        abort(500, description='Не установлен модуль python-docx (pip install python-docx)')


def _export_xlsx_simple(d1, d2, data, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(['Вид приказа', 'Количество приказов'])
    for row in data:
        ws.append(list(row))
    ws.append(['Итого', sum(r[1] for r in data)])
    for c in ws[1][0], ws[5][0], ws[5][1]:
        c.font = Font(bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_period.xlsx')


def _export_xlsx_matrix(d1, d2, headers, table, totals, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Отчёт'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(list(headers))
    for row in table:
        ws.append(list(row))
    if totals:
        ws.append(['Итого'] + list(totals))
    for cell in ws[len(table) + 4]:
        cell.font = Font(bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_period_detail.xlsx')


def _export_xlsx_table(d1, d2, headers, data, title):
    _need_xlsx()
    wb = Workbook()
    ws = wb.active
    ws.title = 'Журнал'
    ws.append([title])
    ws.append([f'Период: с {d1} по {d2}'])
    ws.append([])
    ws.append(list(headers))
    for row in data:
        ws.append(list(row))
    for cell in ws[4]:
        cell.font = Font(bold=True)
    widths = [6, 12, 12, 30, 45, 30, 30]
    for i, w in enumerate(widths[:len(headers)], 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         'report_journal.xlsx')


def _export_docx_simple(d1, d2, data, title):
    _need_docx()
    doc = DocxDocument()
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=2)
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    hdr[0].text, hdr[1].text = 'Вид приказа', 'Количество приказов'
    for row in data:
        cells = t.add_row().cells
        cells[0].text, cells[1].text = str(row[0]), str(row[1])
    cells = t.add_row().cells
    cells[0].text, cells[1].text = 'Итого', str(sum(r[1] for r in data))
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_period.docx')


def _export_docx_matrix(d1, d2, headers, table, totals, title):
    _need_docx()
    doc = DocxDocument()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = str(h)
    for row in table:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    if totals:
        cells = t.add_row().cells
        cells[0].text = 'Итого'
        for i, v in enumerate(totals):
            cells[i + 1].text = str(v)
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_period_detail.docx')


def _export_docx_table(d1, d2, headers, data, title):
    _need_docx()
    doc = DocxDocument()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    doc.add_heading(title, level=1)
    doc.add_paragraph(f'Период: с {d1} по {d2}')
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    for i, h in enumerate(headers):
        t.rows[0].cells[i].text = str(h)
    for row in data:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
    buf = io.BytesIO()
    doc.save(buf)
    return _mem_download(buf,
                         'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         'report_journal.docx')

# ---------------------------------------------------------------- сервис (архив/восстановление/обновление)

@app.route('/service')
def service():
    if not is_admin():
        abort(403)
    archives = []
    if os.path.isdir(ARCHIVE_DIR):
        for fn in sorted(os.listdir(ARCHIVE_DIR), reverse=True):
            if fn.endswith('.zip'):
                p = os.path.join(ARCHIVE_DIR, fn)
                archives.append({'name': fn,
                                 'size': f"{os.path.getsize(p)/1024:.0f} КБ",
                                 'date': datetime.fromtimestamp(os.path.getmtime(p))
                                         .strftime('%d.%m.%Y %H:%M')})
    # данные для раздела «Загрузка документов из Excel» (v1.5.3):
    # справочники + виды приказов (журналы) с количеством наименований
    conn = get_db()
    types_by_journal = {}
    for j in JOURNAL_TYPES:
        cnt = conn.execute('SELECT COUNT(*) c FROM order_types WHERE kind=?',
                           (j,)).fetchone()['c']
        types_by_journal[j] = {'title': JOURNAL_TITLES[j], 'count': cnt}
    employees_n = conn.execute('SELECT COUNT(*) c FROM employees').fetchone()['c']
    conn.close()
    return render_template('service.html', archives=archives,
                           journals=JOURNAL_TYPES, journal_titles=JOURNAL_TITLES,
                           types_by_journal=types_by_journal, employees_n=employees_n)


@app.route('/service/backup', methods=['POST'])
def service_backup():
    if not is_admin():
        abort(403)
    stamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    tmp = tempfile.mkdtemp(prefix='backup_')
    try:
        dbcopy = os.path.join(tmp, 'orders.db')
        src = get_db()
        dst = sqlite3.connect(dbcopy)
        src.backup(dst)
        src.close()
        dst.close()
        scans_dir = os.path.join(tmp, 'scans')
        if os.path.isdir(UPLOAD_DIR):
            shutil.copytree(UPLOAD_DIR, scans_dir)
        zipname = f'backup_{stamp}.zip'
        zpath = os.path.join(ARCHIVE_DIR, zipname)
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        shutil.make_archive(zpath[:-4], 'zip', tmp)
        audit('РЕЗЕРВНАЯ КОПИЯ СОЗДАНА', zipname)
        flash(f'Архив создан: {zipname}', 'ok')
    except Exception as e:
        flash(f'Ошибка создания архива: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))


@app.route('/service/download/<name>')
def service_download(name):
    if not is_admin():
        abort(403)
    p = os.path.join(ARCHIVE_DIR, os.path.basename(name))
    if not os.path.exists(p):
        abort(404)
    return send_file(p, as_attachment=True, download_name=os.path.basename(p))


@app.route('/service/delete_archive/<name>', methods=['POST'])
def service_delete_archive(name):
    if not is_admin():
        abort(403)
    p = os.path.join(ARCHIVE_DIR, os.path.basename(name))
    if os.path.exists(p):
        os.remove(p)
        audit('АРХИВ УДАЛЁН', name)
        flash('Архив удалён', 'ok')
    return redirect(url_for('service'))


@app.route('/service/restore', methods=['POST'])
def service_restore():
    if not is_admin():
        abort(403)
    f = request.files.get('archive')
    if not f or not f.filename.lower().endswith('.zip'):
        flash('Выберите ZIP-архив для восстановления', 'err')
        return redirect(url_for('service'))
    tmp = tempfile.mkdtemp(prefix='restore_')
    try:
        apath = os.path.join(tmp, 'upload.zip')
        f.save(apath)
        extract = os.path.join(tmp, 'ex')
        shutil.unpack_archive(apath, extract, 'zip')
        db_in = os.path.join(extract, 'orders.db')
        if not os.path.exists(db_in):
            raise FileNotFoundError('В архиве отсутствует orders.db')
        # safety: make current backup first
        stamp = datetime.now().strftime('pre_restore_%Y-%m-%d_%H-%M-%S')
        os.makedirs(ARCHIVE_DIR, exist_ok=True)
        cur_zip = os.path.join(ARCHIVE_DIR, f'backup_{stamp}.zip')
        pre = tempfile.mkdtemp(prefix='pre_')
        try:
            s = get_db(); d = sqlite3.connect(os.path.join(pre, 'orders.db'))
            s.backup(d); s.close(); d.close()
            if os.path.isdir(UPLOAD_DIR):
                shutil.copytree(UPLOAD_DIR, os.path.join(pre, 'scans'))
            shutil.make_archive(cur_zip[:-4], 'zip', pre)
        finally:
            shutil.rmtree(pre, ignore_errors=True)
        # restore db
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        shutil.copy(db_in, DB_PATH)
        # restore scans
        scans_in = os.path.join(extract, 'scans')
        if os.path.isdir(scans_in):
            if os.path.isdir(UPLOAD_DIR):
                shutil.rmtree(UPLOAD_DIR)
            shutil.copytree(scans_in, UPLOAD_DIR)
        audit('ВОССТАНОВЛЕНИЕ ИЗ АРХИВА', f.filename)
        flash('База данных восстановлена из архива', 'ok')
    except Exception as e:
        flash(f'Ошибка восстановления: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))


@app.route('/service/users')
def service_users():
    """Список пользователей для администрирования (смена паролей)."""
    if not is_admin():
        abort(403)
    conn = get_db()
    users = conn.execute('SELECT e.id, e.full_name, e.position, e.login, '
                         'e.password_hash, e.pwd_changed_at, r.name AS role_name '
                         'FROM employees e LEFT JOIN roles r ON r.id=e.role_id '
                         'ORDER BY r.name DESC, e.full_name').fetchall()
    conn.close()
    return jsonify([{'id': u['id'], 'full_name': u['full_name'],
                     'position': u['position'] or '', 'login': u['login'] or '',
                     'role': u['role_name'] or '',
                     'has_password': bool(u['password_hash']),
                     'pwd_changed_at': u['pwd_changed_at'] or ''} for u in users])


@app.route('/service/password', methods=['POST'])
def service_set_password():
    """Смена пароля пользователя администратором.

    Форма: user_id (или login), password. Пароль задаётся напрямую; после
    установки сотрудник при желании сменит его самостоятельно в профиле.
    Действие фиксируется в аудите.
    """
    if not is_admin():
        abort(403)
    uid = request.form.get('user_id', type=int)
    login_name = (request.form.get('login') or '').strip()
    new_password = request.form.get('password', '')
    if not uid and not login_name:
        flash('Не указан пользователь для смены пароля', 'err')
        return redirect(url_for('service'))
    if len(new_password) < 4:
        flash('Новый пароль слишком короткий (минимум 4 символа)', 'err')
        return redirect(url_for('service'))
    conn = get_db()
    if uid:
        emp = conn.execute('SELECT * FROM employees WHERE id=?', (uid,)).fetchone()
    else:
        emp = conn.execute('SELECT * FROM employees WHERE login=?', (login_name,)).fetchone()
    if not emp:
        conn.close()
        flash('Пользователь не найден', 'err')
        return redirect(url_for('service'))
    conn.execute('UPDATE employees SET password_hash=?, pwd_changed_at=NULL WHERE id=?',
                 (generate_password_hash(new_password), emp['id']))
    conn.commit()
    conn.close()
    audit('ПАРОЛЬ ИЗМЕНЁН АДМИНИСТРАТОРОМ',
          f"пользователь={emp['full_name']} (id={emp['id']}, "
          f"логин={emp['login'] or '—'})")
    flash(f"Пароль изменён пользователю «{emp['full_name']}»", 'ok')
    return redirect(url_for('service'))


RESTART_PENDING_FILE = os.path.join(BASE_DIR, 'data', 'restart_pending')


def _touch_restart_flag():
    """Создаёт флаг-файл, сигнализирующий обёртке run.sh о необходимости
    мягкого перезапуска приложения без выключения сервера."""
    try:
        os.makedirs(os.path.dirname(RESTART_PENDING_FILE), exist_ok=True)
        with open(RESTART_PENDING_FILE, 'w') as fh:
            fh.write(datetime.now().isoformat(timespec='seconds'))
    except OSError:
        pass


def _find_app_pid():
    """Находит PID процесса python, запустившего app.py (для завершения
    с последующим автоматическим перезапуском обёрткой run.sh)."""
    pid = None
    try:
        import psutil
        me = os.getpid()
        for p in psutil.process_iter(['pid', 'cmdline']):
            try:
                cmd = p.info['cmdline'] or []
                if any(str(c).endswith('app.py') for c in cmd):
                    if p.info['pid'] == me:
                        return me
                    pid = pid or p.info['pid']
            except Exception:
                continue
    except ImportError:
        pass
    if pid is None:  # запасной вариант через ps
        try:
            import subprocess
            out = subprocess.run(['ps', '-eo', 'pid,args'], capture_output=True,
                                 text=True).stdout
            for line in out.splitlines():
                if 'app.py' in line and 'python' in line:
                    cand = int(line.split()[0])
                    if cand != os.getpid():
                        pid = cand
                        break
        except Exception:
            pass
    return pid


@app.route('/service/update', methods=['POST'])
def service_update():
    """Обновление приложения из ZIP-архива новой версии (только Администратор).

    Работает БЕЗ ВЫКЛЮЧЕНИЯ СЕРВЕРА: файлы заменяются «на лету», затем
    приложение мягко перезапускается (см. UPDATE_HOT.md и обёртку run.sh).
    Параметры: soft=1 — только применить файлы, перезапуск выполнить вручную;
    по умолчанию выполняется автоматический мягкий перезапуск."""
    if not is_admin():
        abort(403)
    f = request.files.get('package')
    if not f or not f.filename.lower().endswith('.zip'):
        flash('Загрузите ZIP-архив обновления', 'err')
        return redirect(url_for('service'))
    soft = request.form.get('soft') == '1'
    tmp = tempfile.mkdtemp(prefix='update_')
    try:
        upath = os.path.join(tmp, 'pkg.zip')
        f.save(upath)
        ex = os.path.join(tmp, 'ex')
        shutil.unpack_archive(upath, ex, 'zip')
        applied = 0
        for root, dirs, files in os.walk(ex):
            rel = os.path.relpath(root, ex)
            if rel.startswith('data'):
                continue  # не трогаем базу, сканы и логи — данные сохраняются
            target = os.path.normpath(os.path.join(BASE_DIR, rel))
            os.makedirs(target, exist_ok=True)
            for fn in files:
                if fn in ('orders.db', 'secret_key.txt', 'version.json'):
                    continue
                srcf = os.path.join(root, fn)
                dstf = os.path.join(target, fn)
                if os.path.commonpath([dstf, BASE_DIR]) != BASE_DIR:
                    continue
                shutil.copy2(srcf, dstf)
                applied += 1
        write_version_file()
        audit('ОБНОВЛЕНИЕ ПРОГРАММЫ ПРИМЕНЕНО', f'файлов: {applied}')
        if not soft:
            # Мягкий перезапуск без простоя сервера: файл-флаг + самозавершение
            # процесса; обёртка run.sh сразу поднимет новую версию кода.
            _touch_restart_flag()
            audit('МЯГКИЙ ПЕРЕЗАПУСК', 'инициирован после применения обновления')
            flash(f'Обновление применено: файлов — {applied}. '
                  f'Сервер автоматически перезапускается на новой версии '
                  f'(без потери данных).', 'ok')
            def _delayed_exit():
                time.sleep(1.5)  # даём ответу браузера (flash) уйти полностью
                pid = _find_app_pid()
                if pid == os.getpid() or pid is None:
                    os._exit(0)
                else:
                    import signal
                    os.kill(pid, signal.SIGTERM)
            threading.Thread(target=_delayed_exit, daemon=True).start()
        else:
            flash(f'Обновление применено: файлов — {applied}. '
                  f'Перезапустите сервер вручную (или снимите флажок мягкого режима).',
                  'ok')
    except Exception as e:
        flash(f'Ошибка обновления: {e}', 'err')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return redirect(url_for('service'))


# ------------------------------------------------- активные пользователи (API)

@app.route('/api/active_users')
def api_active_users():
    """JSON со списком активных пользователей для меню «👥 Онлайн».
    Доступ: Администратор, Куратор, Начальник отдела."""
    user = current_user()
    if not user or user['role_name'] not in ('Администратор', 'Куратор',
                                             'Начальник отдела'):
        abort(403)
    users = get_active_users()
    return Response(json.dumps({'timeout': ACTIVE_TIMEOUT, 'users': users},
                                ensure_ascii=False),
                    mimetype='application/json')

# ---------------------------------------------------------------- profile / password

@app.route('/profile', methods=['GET', 'POST'])
def profile():
    """Личная страница сотрудника: информация + самостоятельная смена пароля."""
    user = current_user()
    if request.method == 'POST':
        old = request.form.get('old_password', '')
        new = request.form.get('new_password', '')
        confirm = request.form.get('confirm_password', '')
        conn = get_db()
        u = conn.execute('SELECT * FROM employees WHERE id=?', (user['id'],)).fetchone()
        if not u['password_hash'] or not check_password_hash(u['password_hash'], old):
            conn.close()
            audit('ПАРОЛЬ НЕ ИЗМЕНЁН', 'неверно указан текущий пароль')
            flash('Текущий пароль указан неверно', 'err')
            return redirect(url_for('profile'))
        if len(new) < 4:
            conn.close()
            flash('Новый пароль слишком короткий (минимум 4 символа)', 'err')
            return redirect(url_for('profile'))
        if new != confirm:
            conn.close()
            flash('Новый пароль и подтверждение не совпадают', 'err')
            return redirect(url_for('profile'))
        if new == old:
            conn.close()
            flash('Новый пароль не должен совпадать с текущим', 'err')
            return redirect(url_for('profile'))
        conn.execute(
            'UPDATE employees SET password_hash=?, pwd_changed_at=? WHERE id=?',
            (generate_password_hash(new),
             datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user['id']))
        conn.commit()
        conn.close()
        audit('ПАРОЛЬ ИЗМЕНЁН', 'собственный профиль (сотрудник сменил пароль сам)')
        flash('Пароль успешно изменён', 'ok')
        return redirect(url_for('profile'))
    return render_template('profile.html')


@app.errorhandler(403)
def err403(e):
    return render_template('error.html', code=403, msg='Доступ запрещён. Недостаточно прав.'), 403


@app.errorhandler(404)
def err404(e):
    return render_template('error.html', code=404, msg='Страница не найдена.'), 404


@app.errorhandler(Exception)
def err500(e):
    """Любое необработанное исключение пишется в журнал логирования."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e  # 403/404 и т.п. отдаём штатными обработчиками
    log.error(f'НЕОБРАБОТАННАЯ ОШИБКА при {request.method} {request.path}: '
              f'{e}\n{traceback.format_exc()}')
    return render_template('error.html', code=500,
                           msg='Внутренняя ошибка сервера. Подробности — в журнале логов.'), 500

# ------------------------------------------------------------- просмотр логов

LOG_TAIL_MAX = 2000  # максимум строк, выводимых на страницу журнала


def _read_log_tail(n):
    """Читает последние n строк основного файла лога (с учётом ротации)."""
    try:
        with open(LOG_FILE, encoding='utf-8', errors='replace') as f:
            lines = f.readlines()[-n:]
        return [l.rstrip('\n') for l in lines]
    except FileNotFoundError:
        return []


@app.route('/logs')
def logs_view():
    """Просмотр журнала логирования — только Администратор."""
    if not is_admin():
        abort(403)
    try:
        n = min(int(request.args.get('n', 200)), LOG_TAIL_MAX)
    except ValueError:
        n = 200
    lines = _read_log_tail(n)
    files = sorted([fn for fn in os.listdir(LOG_DIR) if fn.startswith('app.log')]
                   ) if os.path.isdir(LOG_DIR) else []
    return render_template('logs.html', lines=lines, n=n, files=files,
                           log_file=LOG_FILE)


@app.route('/logs/download')
def logs_download():
    """Скачивание файла журнала — только Администратор."""
    if not is_admin():
        abort(403)
    name = os.path.basename(request.args.get('file', 'app.log'))
    p = os.path.join(LOG_DIR, name)
    if not os.path.isfile(p):
        abort(404)
    return send_file(p, as_attachment=True, download_name=name)


@app.route('/logs/clear', methods=['POST'])
def logs_clear():
    """Очистка журнала логирования — только Администратор.

    Режимы (параметр mode):
      keep   — создать резервную копию app.log.<дата_время>, затем очистить
               текущий файл и удалить архивные файлы ротации (по умолчанию);
      backup — только создать резервную копию, ничего не удаляя;
      full   — полностью очистить текущий файл и удалить все архивные
               файлы ротации (без создания копии).

    Перед очисткой в журнал записывается событие аудита, которое вместе с
    копией попадает в резервный файл — след действия администратора теряется.
    """
    if not is_admin():
        abort(403)
    mode = request.form.get('mode', 'keep')
    audit('ЖУРНАЛ ОЧИЩАЕТСЯ', f'режим: {mode}')

    # 1. Резервная копия текущего файла журнала (если он существует)
    copied = None
    if mode in ('keep', 'backup') and os.path.isfile(LOG_FILE):
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        copied = os.path.join(LOG_DIR, f'app.log.{ts}.bak')
        try:
            with open(LOG_FILE, encoding='utf-8', errors='replace') as src, \
                    open(copied, 'w', encoding='utf-8') as dst:
                dst.write(src.read())
        except OSError as e:
            flash(f'Ошибка создания резервной копии журнала: {e}', 'err')
            return redirect(url_for('logs_view'))

    # 2. Удаляем архивные файлы ротации (app.log.1 … app.log.5)
    removed = 0
    if mode in ('keep', 'full'):
        for fn in sorted(os.listdir(LOG_DIR)):
            if re.fullmatch(r'app\.log\.\d+', fn):
                try:
                    os.remove(os.path.join(LOG_DIR, fn))
                    removed += 1
                except OSError:
                    pass

    # 3. Обрезаем основной файл «на лету» (без остановки сервера).
    #    Файл открываем на запись с усечением; потоковый handler продолжает
    #    писать в тот же файловый дескриптор, поэтому после truncate нужно
    #    принудительно выставить позицию записи в начало, иначе следующая
    #    строка лога появится на прежнем смещении («дырка» из нулей).
    if mode in ('keep', 'full'):
        try:
            for h in logging.getLogger('orders_app').handlers:
                if isinstance(h, RotatingFileHandler):
                    h.stream.flush()
                    os.truncate(LOG_FILE, 0)
                    h.stream.seek(0)
                    break
            else:
                with open(LOG_FILE, 'w', encoding='utf-8'):
                    pass
        except OSError as e:
            flash(f'Ошибка очистки файла журнала: {e}', 'err')
            return redirect(url_for('logs_view'))

    log.info(f'[АУДИТ] ЖУРНАЛ ОЧИЩЕН пользователем id='
                f"{session.get('user_id')} (IP: {request.remote_addr}); "
                f"режим: {mode}; "
                f"копия: {os.path.basename(copied) if copied else 'нет'}; "
                f"удалено архивов ротации: {removed}")

    if mode == 'backup':
        flash(f'Резервная копия журнала создана: '
              f'{os.path.basename(copied)} (файл app.log сохранён)', 'ok')
    elif copied:
        flash(f'Журнал очищен. Резервная копия: '
              f'{os.path.basename(copied)}; удалено архивных файлов: {removed}',
              'ok')
    else:
        flash(f'Журнал полностью очищен (удалено архивных файлов: {removed})',
              'ok')
    return redirect(url_for('logs_view'))


if __name__ == '__main__':
    import sys
    host = os.environ.get('HOST', '0.0.0.0')
    port = None
    # Порт можно задать: аргументом --port, переменной окружения PORT
    # или значением по умолчанию DEFAULT_PORT (8080, т.к. 5000 занят другой
    # программой на этом сервере). Если порт занят — автоматически
    # подбирается ближайший свободный.
    args = sys.argv[1:]
    if '--port' in args:
        try:
            port = int(args[args.index('--port') + 1])
        except (IndexError, ValueError):
            print('Неверное значение --port'); sys.exit(1)
    elif os.environ.get('PORT'):
        port = int(os.environ['PORT'])
    else:
        port = DEFAULT_PORT
    requested_port = port
    port = find_free_port(port)
    if port is None:
        print(f'Не найден свободный порт начиная с {requested_port}. '
              f'Запустите с другим портом: python3 app.py --port <порт>')
        sys.exit(1)
    if port != requested_port:
        print(f'Порт {requested_port} занят другой программой, '
              f'использую свободный порт {port}')
        log.warning(f'Запрошенный порт {requested_port} занят, '
                    f'приложение запущено на порту {port}')
    log.info(f'Сервер запущен: http://{host}:{port} '
             f'(версия {APP_VERSION}, логи: {LOG_FILE})')
    print(f'Сервер запущен: http://{host}:{port}  (логин: admin / admin)')
    print(f'Журнал логирования: {LOG_FILE}')
    app.run(host=host, port=port, debug=False)
