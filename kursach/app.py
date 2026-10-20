"""
Автоматизация жизненного цикла реквизита и технического оборудования
(на примере Астраханской государственной филармонии).

Функции: учет, выдача/возврат, бронирование, инвентаризация, контроль состояния,
планирование закупок, отчеты (Excel/PDF), уведомления, резервное копирование, журнал действий.

Запуск: pip install -r requirements.txt && python app.py  ->  http://127.0.0.1:5000
Первый вход: admin / admin (смените пароль в разделе «Профиль»).
Переменные окружения: см. README.md
"""
import hmac
import io
import math
import os
import re
import secrets
import smtplib
import sqlite3
import threading
import time
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from functools import wraps
from xml.sax.saxutils import escape as xml_escape

from flask import (Flask, g, request, redirect, url_for, session, render_template, flash,
                   abort, send_file, send_from_directory, has_request_context)
from jinja2 import DictLoader
from markupsafe import Markup
from werkzeug.security import generate_password_hash, check_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "philharmonia.db"))
BACKUP_DIR = os.environ.get("BACKUP_DIR", os.path.join(BASE_DIR, "backups"))
BACKUP_INTERVAL_HOURS = float(os.environ.get("BACKUP_INTERVAL_HOURS", "24"))  # 0 — выключить
BACKUP_KEEP_AUTO = int(os.environ.get("BACKUP_KEEP_AUTO", "14"))
REMIND_DAYS = int(os.environ.get("REMIND_DAYS", "3"))       # за сколько дней предупреждать о возврате
DEFAULT_ISSUE_DAYS = int(os.environ.get("DEFAULT_ISSUE_DAYS", "7"))
CHECK_PERIOD_DAYS = 180  # межремонтный/проверочный интервал
INV_PREFIX = "АГФ-"      # префикс инвентарных номеров

OK_STATE = "Исправно"
SYSTEM_STATES = ["Исправно", "Требует ремонта", "В ремонте", "Списано"]  # на них опирается логика
DEFAULT_CATEGORIES = ["Музыкальные инструменты", "Звук", "Свет", "Сценическое оборудование",
                      "Реквизит", "Костюмы", "Прочее"]
DEFAULT_LOCATIONS = ["Большой зал", "Склад звука", "Склад сцены", "Склад света",
                     "Аппаратная", "Склад реквизита"]
AVAIL_STATUSES = ["В наличии", "Выдан частично", "Выдан", "В ремонте", "Списан", "Отсутствует"]
NO_EVENT = "Хозяйственные нужды (без мероприятия)"

ROLES = {"admin": "Администратор", "responsible": "Ответственный", "user": "Пользователь"}
STAFF = {"admin", "responsible"}
PERMS = {  # право -> роли
    "view": set(ROLES),
    "edit_items": STAFF,      # карточки, акты проверки
    "operations": STAFF,      # выдача/возврат, брони, инвентаризация, закупки
    "dictionaries": STAFF,
    "reports": STAFF,
    "admin": {"admin"},       # пользователи, журнал, резервные копии
}


def _secret_key():
    """Ключ подписи сессий: из окружения, иначе генерируется один раз и хранится в .secret_key."""
    env = os.environ.get("SECRET_KEY")
    if env:
        return env
    path = os.path.join(BASE_DIR, ".secret_key")
    try:
        with open(path, encoding="ascii") as fh:
            val = fh.read().strip()
            if val:
                return val
    except OSError:
        pass
    val = secrets.token_hex(32)
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="ascii") as fh:
            fh.write(val)
    except OSError:
        pass
    return val


app = Flask(__name__)
app.secret_key = _secret_key()
app.config.update(SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_HTTPONLY=True,
                  PERMANENT_SESSION_LIFETIME=timedelta(hours=8))
ADMIN_LOGIN = os.environ.get("ADMIN_LOGIN", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "admin")

# ---------------------------------------------------------------- БД
SCHEMA = """
CREATE TABLE IF NOT EXISTS items(
  id INTEGER PRIMARY KEY AUTOINCREMENT, inv_no TEXT UNIQUE, name TEXT NOT NULL,
  category TEXT, location TEXT, qty INTEGER DEFAULT 1, min_qty INTEGER DEFAULT 0,
  status TEXT DEFAULT 'Исправно', price REAL DEFAULT 0, bought TEXT,
  life_years REAL DEFAULT 5, next_check TEXT);
CREATE TABLE IF NOT EXISTS bookings(
  id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER, qty INTEGER,
  event TEXT, start TEXT, end TEXT, person TEXT);
CREATE TABLE IF NOT EXISTS checks(
  id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER, day TEXT,
  status TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS purchases(
  id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER, name TEXT, qty INTEGER,
  reason TEXT, status TEXT DEFAULT 'Запланирована', created TEXT);

CREATE TABLE IF NOT EXISTS users(
  id INTEGER PRIMARY KEY AUTOINCREMENT, login TEXT UNIQUE NOT NULL COLLATE NOCASE,
  pw_hash TEXT NOT NULL, full_name TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'user',
  employee_id INTEGER, active INTEGER NOT NULL DEFAULT 1, created TEXT, last_login TEXT);
CREATE TABLE IF NOT EXISTS categories(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS states(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS locations(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL);
CREATE TABLE IF NOT EXISTS employees(
  id INTEGER PRIMARY KEY AUTOINCREMENT, full_name TEXT NOT NULL, position TEXT, phone TEXT,
  email TEXT, active INTEGER NOT NULL DEFAULT 1);
CREATE TABLE IF NOT EXISTS events(
  id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, event_date TEXT, place TEXT, note TEXT);
CREATE TABLE IF NOT EXISTS issues(
  id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER NOT NULL, employee_id INTEGER NOT NULL,
  event_id INTEGER, qty INTEGER NOT NULL, issued_at TEXT NOT NULL, due_date TEXT NOT NULL,
  issued_by INTEGER, issue_note TEXT, returned_at TEXT, return_state TEXT, return_note TEXT,
  received_by INTEGER);
CREATE TABLE IF NOT EXISTS movements(
  id INTEGER PRIMARY KEY AUTOINCREMENT, item_id INTEGER NOT NULL, ts TEXT NOT NULL, kind TEXT NOT NULL,
  qty INTEGER, employee_id INTEGER, event_id INTEGER, user_id INTEGER, details TEXT);
CREATE TABLE IF NOT EXISTS audit_log(
  id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT NOT NULL, user_id INTEGER, login TEXT,
  action TEXT NOT NULL, entity TEXT, entity_id TEXT, details TEXT, ip TEXT);
CREATE TABLE IF NOT EXISTS inventories(
  id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, created TEXT, finished TEXT,
  status TEXT DEFAULT 'Идёт', scope TEXT, created_by INTEGER, applied INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS inventory_lines(
  id INTEGER PRIMARY KEY AUTOINCREMENT, inv_id INTEGER NOT NULL, item_id INTEGER NOT NULL,
  expected INTEGER NOT NULL, issued INTEGER DEFAULT 0, actual INTEGER, note TEXT);
CREATE INDEX IF NOT EXISTS ix_issues_item ON issues(item_id);
CREATE INDEX IF NOT EXISTS ix_issues_open ON issues(returned_at, due_date);
CREATE INDEX IF NOT EXISTS ix_mov_item ON movements(item_id);
CREATE INDEX IF NOT EXISTS ix_audit_ts ON audit_log(ts);
CREATE INDEX IF NOT EXISTS ix_invl_inv ON inventory_lines(inv_id);
"""
# новые столбцы для баз, созданных прежней версией приложения
MIGRATIONS = [("items", "descr", "TEXT"), ("items", "serial", "TEXT"),
              ("bookings", "event_id", "INTEGER"), ("bookings", "employee_id", "INTEGER")]


def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB, timeout=15)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_):
    d = g.pop("db", None)
    if d:
        d.close()


def q(sql, args=(), one=False):
    cur = db().execute(sql, args)
    rows = [dict(r) for r in cur.fetchall()]
    return (rows[0] if rows else None) if one else rows


def x(sql, args=()):
    """Выполнить без commit (для составных операций; завершать db().commit())."""
    return db().execute(sql, args)


def run(sql, args=()):
    db().execute(sql, args)
    db().commit()


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today():
    return date.today().isoformat()


def init_db():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    con = sqlite3.connect(DB, timeout=15)
    con.executescript(SCHEMA)
    for table, col, ddl in MIGRATIONS:
        cols = [r[1] for r in con.execute(f"PRAGMA table_info({table})")]
        if col not in cols:
            con.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ddl}")

    # справочники: стартовые значения + всё, что уже используется в карточках
    for s in SYSTEM_STATES:
        con.execute("INSERT OR IGNORE INTO states(name) VALUES(?)", (s,))
    for tbl, col, defaults in (("categories", "category", DEFAULT_CATEGORIES),
                               ("locations", "location", DEFAULT_LOCATIONS),
                               ("states", "status", [])):
        if not con.execute(f"SELECT 1 FROM {tbl}").fetchone() or tbl == "states":
            for n in defaults:
                con.execute(f"INSERT OR IGNORE INTO {tbl}(name) VALUES(?)", (n,))
        for (n,) in con.execute(f"SELECT DISTINCT {col} FROM items WHERE {col} IS NOT NULL AND {col}<>''").fetchall():
            con.execute(f"INSERT OR IGNORE INTO {tbl}(name) VALUES(?)", (n,))
    if not con.execute("SELECT 1 FROM events").fetchone():
        con.execute("INSERT INTO events(name, note) VALUES(?,?)",
                    (NO_EVENT, "Служебная запись для операций вне мероприятий"))

    fresh_demo = False
    if not con.execute("SELECT 1 FROM items").fetchone():
        fresh_demo = True
        _seed_demo(con)
    if not con.execute("SELECT 1 FROM users").fetchone():
        con.execute("INSERT INTO users(login,pw_hash,full_name,role,created) VALUES(?,?,?,?,?)",
                    (ADMIN_LOGIN, generate_password_hash(ADMIN_PASS), "Администратор", "admin", now()))
        if fresh_demo:  # демонстрационные учётные записи только для новой демо-базы
            emp = con.execute("SELECT id FROM employees ORDER BY id").fetchall()
            con.execute("INSERT INTO users(login,pw_hash,full_name,role,employee_id,created) VALUES(?,?,?,?,?,?)",
                        ("manager", generate_password_hash("manager"), "Кузнецова А. П.", "responsible",
                         emp[0][0] if emp else None, now()))
            con.execute("INSERT INTO users(login,pw_hash,full_name,role,employee_id,created) VALUES(?,?,?,?,?,?)",
                        ("user", generate_password_hash("user"), "Иванов С. В.", "user",
                         emp[1][0] if len(emp) > 1 else None, now()))
    con.commit()
    con.close()


def _seed_demo(con):
    t = date.today()
    demo = [
        ("АГФ-001", "Рояль концертный Steinway", "Музыкальные инструменты", "Большой зал", 1, 1, "Исправно", 6500000, 12, 30),
        ("АГФ-002", "Микрофон Shure SM58", "Звук", "Склад звука", 12, 8, "Исправно", 9500, 6, 40),
        ("АГФ-003", "Пульт музыкантский", "Сценическое оборудование", "Склад сцены", 40, 30, "Исправно", 3200, 10, 90),
        ("АГФ-004", "Прожектор LED PAR 64", "Свет", "Склад света", 6, 10, "Требует ремонта", 14000, 5, -10),
        ("АГФ-005", "Микшерный пульт Yamaha", "Звук", "Аппаратная", 1, 1, "Исправно", 210000, 7, -5),
        ("АГФ-006", "Театральный реквизит: кресло XIX в.", "Реквизит", "Склад реквизита", 2, 2, "Исправно", 25000, 15, 120),
    ]
    ids = []
    for i, (inv, n, c, loc, qty, mq, st, pr, life, chk) in enumerate(demo):
        bought = (t - timedelta(days=365 * (life - 0.5 if i == 3 else 2))).isoformat()
        cur = con.execute("INSERT INTO items(inv_no,name,category,location,qty,min_qty,status,price,"
                          "bought,life_years,next_check) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                          (inv, n, c, loc, qty, mq, st, pr, bought, life, (t + timedelta(days=chk)).isoformat()))
        ids.append(cur.lastrowid)
        con.execute("INSERT INTO movements(item_id,ts,kind,qty,details) VALUES(?,?,?,?,?)",
                    (cur.lastrowid, now(), "Постановка на учет", qty, f"Инв. № {inv} (начальные данные)"))
    emps = [("Кузнецова Анна Петровна", "Заведующая реквизитом", "+7 (8512) 00-00-01", "kuznetsova@example.org"),
            ("Иванов Сергей Викторович", "Звукорежиссёр", "+7 (8512) 00-00-02", "ivanov@example.org"),
            ("Сидоров Павел Андреевич", "Художник по свету", "+7 (8512) 00-00-03", "sidorov@example.org")]
    eids = [con.execute("INSERT INTO employees(full_name,position,phone,email) VALUES(?,?,?,?)", e).lastrowid
            for e in emps]
    evs = [("Концерт симфонического оркестра", (t - timedelta(days=3)).isoformat(), "Большой зал"),
           ("Гала-концерт «Астраханская осень»", (t + timedelta(days=5)).isoformat(), "Большой зал")]
    evids = [con.execute("INSERT INTO events(name,event_date,place) VALUES(?,?,?)", e).lastrowid for e in evs]
    # демонстрация уведомлений: одна просроченная выдача и одна со скорым сроком
    for item_i, emp, ev, qty, out_days, due_days in ((1, eids[1], evids[0], 2, 5, -1), (2, eids[0], evids[1], 10, 1, 2)):
        iid = ids[item_i]
        issued_at, due = (t - timedelta(days=out_days)).isoformat(), (t + timedelta(days=due_days)).isoformat()
        con.execute("INSERT INTO issues(item_id,employee_id,event_id,qty,issued_at,due_date,issue_note) "
                    "VALUES(?,?,?,?,?,?,?)", (iid, emp, ev, qty, issued_at, due, "Демонстрационная выдача"))
        con.execute("INSERT INTO movements(item_id,ts,kind,qty,employee_id,event_id,details) VALUES(?,?,?,?,?,?,?)",
                    (iid, issued_at + " 10:00:00", "Выдача", qty, emp, ev, f"Срок возврата {due}"))


# ---------------------------------------------------------------- вспомогательное
def to_int(v, default=None):
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return default


def to_float(v, default=None):
    try:
        r = float(str(v).replace(",", ".").strip())
        return r if math.isfinite(r) else default
    except (TypeError, ValueError):
        return default


def valid_date(v):
    try:
        date.fromisoformat(v or "")
        return True
    except ValueError:
        return False


def names(table):
    return [r["name"] for r in q(f"SELECT name FROM {table} ORDER BY name")]


def active_employees():
    return q("SELECT id, full_name, position FROM employees WHERE active=1 ORDER BY full_name")


def event_choices():
    return q("SELECT id, name, event_date FROM events ORDER BY COALESCE(event_date,'9999') DESC, name")


def wear(item):
    """Износ в % по сроку службы."""
    try:
        age = (date.today() - date.fromisoformat(item["bought"])).days / 365
        return round(min(100, age / item["life_years"] * 100))
    except Exception:
        return 0


def stage(item):
    """Этап жизненного цикла."""
    if item["status"] == "Списано":
        return "Списание"
    if item["status"] in ("Требует ремонта", "В ремонте"):
        return "Ремонт"
    return "Эксплуатация"


def avail_status(it):
    """Статус объекта (наличие), вычисляется автоматически."""
    issued = it.get("issued", 0) or 0
    if it["status"] == "Списано":
        return "Списан"
    if it["status"] == "В ремонте":
        return "В ремонте"
    if it["qty"] <= 0:
        return "Отсутствует"
    if issued >= it["qty"]:
        return "Выдан"
    if issued > 0:
        return "Выдан частично"
    return "В наличии"


# ---------------------------------------------------------------- 3D-модели инструментов
INSTR_CATEGORY_HINT = "инструмент"  # категории, в названии которых есть это слово, считаются музыкальными
INSTR3D_TITLES = {
    "grand": "рояль", "upright": "пианино", "synth": "синтезатор", "violin": "скрипка", "viola": "альт",
    "cello": "виолончель", "bass": "контрабас", "guitar": "гитара", "electric": "электрогитара",
    "bassg": "бас-гитара", "balalaika": "балалайка", "domra": "домра", "mandolin": "мандолина", "harp": "арфа",
    "flute": "флейта", "oboe": "гобой", "clarinet": "кларнет", "bassoon": "фагот", "sax": "саксофон",
    "trumpet": "труба", "trombone": "тромбон", "horn": "валторна", "tuba": "туба", "snare": "малый барабан",
    "bassdrum": "большой барабан", "timpani": "литавры", "cymbals": "тарелки", "xylo": "ксилофон / маримба",
    "accordion": "аккордеон / баян", "organ": "орган", "note": "музыкальный инструмент"}
# порядок важен: первое совпадение побеждает (синтезатор раньше пианино, бас-гитара раньше гитары и т. д.)
INSTR3D_RULES = [(k, re.compile(p)) for k, p in (
    ("synth", r"синтезатор|цифров\w*\s+(?:пианино|фортепиано|рояль)|электропиано|клавишн"),
    ("grand", r"роял"),
    ("upright", r"пианин|фортепиано"),
    ("bassg", r"бас[\s-]*гитар|гитар\w*[\s-]+бас"),
    ("electric", r"электро[\s-]*гитар"),
    ("guitar", r"гитар"),
    ("balalaika", r"балалайк"),
    ("domra", r"домр"),
    ("mandolin", r"мандолин"),
    ("violin", r"скрипк"),
    ("cello", r"виолончел"),
    ("bass", r"контрабас"),
    ("sax", r"саксофон"),
    ("viola", r"(?<![а-я])альт(?![а-я])|(?<![а-я])виол(?:а|ы|у|е|ой)(?![а-я])"),
    ("harp", r"арф"),
    ("flute", r"флейт|пикколо"),
    ("oboe", r"гобо"),
    ("clarinet", r"кларнет"),
    ("bassoon", r"фагот"),
    ("trombone", r"тромбон"),
    ("horn", r"валторн|(?<![а-я])горн"),
    ("tuba", r"(?<![а-я])туб(?![а-я]*ус)|геликон|эуфониум|басов\w*\s+труб"),
    ("trumpet", r"труб(?!ч)|корнет|флюгельгорн"),
    ("bassdrum", r"(?:бас|больш)\w*[\s-]+барабан|барабан\w*[\s-]+(?:бас|больш)"),
    ("timpani", r"литавр|тимпан"),
    ("snare", r"барабан"),
    ("cymbals", r"тарелк"),
    ("xylo", r"ксилофон|маримб|вибрафон|металлофон|глокеншпиль"),
    ("accordion", r"аккордеон|баян|гармон"),
    ("organ", r"(?<![а-я])орган(?![а-я])"),
)]
# принадлежности («чехол для скрипки», «струны», «стойка»…) — это не инструмент, модель не показываем
INSTR3D_ACCESSORY = re.compile(r"^\s*(?:чехол|кейс|футляр|струн|смыч|канифол|медиатор|палоч|мундштук|трост|сурдин|ремен|ремн|"
                               r"подставк|стойк|пюпитр|пульт|тюнер|метроном|педал|крышк|накладк|кожа|набор|комплект|запчаст|аксессуар)")


def instrument_kind(name, category=""):
    """Тип 3D-модели по названию карточки (None — модель не нужна). Музыкальная категория без опознанного
    названия получает универсальную модель «нота»."""
    n = (name or "").lower().replace("ё", "е")
    if INSTR3D_ACCESSORY.match(n) or re.search(r"\s(?:для|под)\s", n):
        return None
    for kind, rx in INSTR3D_RULES:
        if rx.search(n):
            return kind
    if INSTR_CATEGORY_HINT in (category or "").lower():
        return "note"
    return None



def enrich(rows):
    for r in rows:
        r["wear"] = wear(r)
        r["stage"] = stage(r)
        r["overdue"] = bool(r["next_check"] and r["next_check"] < date.today().isoformat()
                            and r["status"] != "Списано")
        r["issued"] = r.get("issued") or 0
        r["free"] = r["qty"] - r["issued"]
        r["avail"] = avail_status(r)
        r["m3d"] = instrument_kind(r.get("name"), r.get("category"))
        r["m3d_title"] = INSTR3D_TITLES.get(r["m3d"], "")
    return rows


ITEM_SQL = ("SELECT i.*, COALESCE((SELECT SUM(s.qty) FROM issues s WHERE s.item_id=i.id "
            "AND s.returned_at IS NULL),0) AS issued FROM items i")


def get_item(iid):
    r = q(ITEM_SQL + " WHERE i.id=?", (iid,), one=True)
    return enrich([r])[0] if r else None


def next_inv_no():
    mx = 0
    pat = re.compile(re.escape(INV_PREFIX) + r"(\d+)")
    for r in q("SELECT inv_no FROM items"):
        m = pat.fullmatch(r["inv_no"] or "")
        if m:
            mx = max(mx, int(m.group(1)))
    return f"{INV_PREFIX}{mx + 1:03d}"


def free_qty(item_id, start, end):
    """Свободно на период: всего − брони − невозвращённые выдачи, пересекающиеся с периодом."""
    item = q("SELECT qty FROM items WHERE id=?", (item_id,), one=True)
    booked = q("SELECT COALESCE(SUM(qty),0) s FROM bookings WHERE item_id=? AND start<=? AND end>=?",
               (item_id, end, start), one=True)["s"]
    out = q("SELECT COALESCE(SUM(qty),0) s FROM issues WHERE item_id=? AND returned_at IS NULL "
            "AND issued_at<=? AND MAX(due_date,?)>=?", (item_id, end, today(), start), one=True)["s"]
    return item["qty"] - booked - out


def suggest_purchases():
    """Автоматическое формирование плана закупок."""
    added = 0
    for it in enrich(q(ITEM_SQL)):
        reason = None
        if it["status"] == "Списано":
            reason = "Объект списан — требуется замена"
        elif it["wear"] >= 100:
            reason = "Исчерпан срок службы"
        elif it["qty"] < it["min_qty"]:
            reason = f"Остаток {it['qty']} ниже минимального {it['min_qty']}"
        elif it["status"] == "Требует ремонта" and it["wear"] >= 80:
            reason = "Ремонт нецелесообразен (износ ≥ 80%)"
        if reason and not q("SELECT 1 FROM purchases WHERE item_id=? AND status IN "
                            "('Запланирована','Заказана')", (it["id"],), one=True):
            need = max(1, it["min_qty"] - it["qty"]) if it["qty"] < it["min_qty"] else 1
            x("INSERT INTO purchases(item_id,name,qty,reason,created) VALUES(?,?,?,?,?)",
              (it["id"], it["name"], need, reason, date.today().isoformat()))
            added += 1
    db().commit()
    return added


# ---------------------------------------------------------------- журнал и история
def write_audit(con, user_id, login, action, entity="", entity_id="", details="", ip=""):
    con.execute("INSERT INTO audit_log(ts,user_id,login,action,entity,entity_id,details,ip) "
                "VALUES(?,?,?,?,?,?,?,?)", (now(), user_id, login, action, entity, str(entity_id), details, ip))


def audit(action, entity="", entity_id="", details=""):
    """Запись в журнал действий (commit — на вызывающей стороне)."""
    u = g.get("user") if has_request_context() else None
    ip = request.remote_addr if has_request_context() else ""
    write_audit(db(), u["id"] if u else None, u["login"] if u else None, action, entity, entity_id, details, ip)


def log_move(item_id, kind, qty=None, details="", employee_id=None, event_id=None):
    """Запись в историю движения объекта (commit — на вызывающей стороне)."""
    u = g.get("user")
    x("INSERT INTO movements(item_id,ts,kind,qty,employee_id,event_id,user_id,details) VALUES(?,?,?,?,?,?,?,?)",
      (item_id, now(), kind, qty, employee_id, event_id, u["id"] if u else None, details))


# ---------------------------------------------------------------- аутентификация и права
FAILS = {}  # (ip, login) -> [время неудачных попыток]


def throttled(key):
    t = time.time()
    FAILS[key] = [v for v in FAILS.get(key, []) if t - v < 300]
    return len(FAILS[key]) >= 5


def csrf_token():
    if "csrf" not in session:
        session["csrf"] = secrets.token_hex(16)
    return session["csrf"]


def csrf_input():
    return Markup('<input type="hidden" name="_csrf" value="%s">' % csrf_token())


@app.before_request
def load_user():
    g.user = None
    uid = session.get("uid")
    if uid:
        u = q("SELECT * FROM users WHERE id=? AND active=1", (uid,), one=True)
        if u:
            g.user = u
        else:
            session.clear()
    if request.method == "POST":
        sent, real = request.form.get("_csrf", ""), session.get("csrf", "")
        if not real or not hmac.compare_digest(sent.encode(), real.encode()):
            abort(400)


def can(perm):
    return bool(g.get("user")) and g.user["role"] in PERMS[perm]


def login_required(f):
    @wraps(f)
    def w(*a, **k):
        if not g.user:
            return redirect(url_for("login", next=request.full_path.rstrip("?")))
        return f(*a, **k)
    return w


def need(perm):
    def deco(f):
        @wraps(f)
        def w(*a, **k):
            if not g.user:
                return redirect(url_for("login", next=request.full_path.rstrip("?")))
            if not can(perm):
                abort(403)
            return f(*a, **k)
        return w
    return deco


def scope_employee():
    """Для обычного пользователя — id связанного сотрудника (видит только свои выдачи); для staff — None."""
    if can("operations"):
        return None
    return g.user["employee_id"] or -1


def validate_password(p):
    return None if len(p or "") >= 6 else "Пароль должен содержать не менее 6 символов"


@app.errorhandler(403)
def e403(_):
    return render_template("error.html", code=403, text="Недостаточно прав для этого действия"), 403


@app.errorhandler(404)
def e404(_):
    return render_template("error.html", code=404, text="Страница не найдена"), 404


@app.errorhandler(400)
def e400(_):
    return render_template("error.html", code=400,
                           text="Сессия устарела или запрос некорректен. Обновите страницу и повторите действие"), 400


# ---------------------------------------------------------------- вход / профиль
def safe_next(n):
    return n if n and n.startswith("/") and not n.startswith("//") and "\\" not in n else url_for("dashboard")


@app.route("/assets/instrument3d.js")
def instrument3d_js():
    """Автономный WebGL-просмотрщик 3D-моделей (без внешних библиотек и интернета)."""
    resp = app.response_class(INSTRUMENT3D_JS, mimetype="application/javascript")
    resp.headers["Cache-Control"] = "public, max-age=3600"
    return resp


@app.route("/login", methods=["GET", "POST"])
def login():
    if g.user:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        lg, pw = request.form.get("login", "").strip(), request.form.get("password", "")
        key = (request.remote_addr, lg.lower())
        if throttled(key):
            audit("Вход заблокирован (слишком много попыток)", "users", "", f"логин: {lg}")
            db().commit()
            flash("Слишком много неудачных попыток. Повторите через несколько минут", "danger")
            return render_template("login.html"), 429
        u = q("SELECT * FROM users WHERE login=?", (lg,), one=True)
        ok = check_password_hash(u["pw_hash"] if u else generate_password_hash("x"), pw)
        if u and ok and u["active"]:
            session.clear()
            session["uid"] = u["id"]
            session.permanent = True
            FAILS.pop(key, None)
            g.user = u
            x("UPDATE users SET last_login=? WHERE id=?", (now(), u["id"]))
            audit("Вход в систему", "users", u["id"])
            db().commit()
            return redirect(safe_next(request.args.get("next")))
        FAILS.setdefault(key, []).append(time.time())
        audit("Неудачная попытка входа", "users", "", f"логин: {lg}")
        db().commit()
        flash("Неверный логин или пароль", "danger")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
def logout():
    if g.user:
        audit("Выход из системы", "users", g.user["id"])
        db().commit()
    session.clear()
    return redirect(url_for("login"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    if request.method == "POST":
        f = request.form
        err = None
        if not check_password_hash(g.user["pw_hash"], f.get("old", "")):
            err = "Текущий пароль указан неверно"
        elif f.get("new") != f.get("new2"):
            err = "Новый пароль и подтверждение не совпадают"
        else:
            err = validate_password(f.get("new"))
        if err:
            flash(err, "danger")
        else:
            x("UPDATE users SET pw_hash=? WHERE id=?", (generate_password_hash(f["new"]), g.user["id"]))
            audit("Смена пароля", "users", g.user["id"])
            db().commit()
            flash("Пароль изменён", "success")
            return redirect(url_for("profile"))
    emp = q("SELECT * FROM employees WHERE id=?", (g.user["employee_id"],), one=True) if g.user["employee_id"] else None
    return render_template("profile.html", emp=emp)


# ---------------------------------------------------------------- уведомления
ISSUE_SQL = ("SELECT s.*, i.name AS item_name, i.inv_no, e.full_name AS emp_name, e.email AS emp_email, "
             "ev.name AS event_name, u.full_name AS issued_by_name, ur.full_name AS received_by_name "
             "FROM issues s JOIN items i ON i.id=s.item_id JOIN employees e ON e.id=s.employee_id "
             "LEFT JOIN events ev ON ev.id=s.event_id LEFT JOIN users u ON u.id=s.issued_by "
             "LEFT JOIN users ur ON ur.id=s.received_by")


def decorate_issue(r):
    if r["returned_at"]:
        r["days"], r["flag"] = None, "returned"
    else:
        r["days"] = (date.fromisoformat(r["due_date"]) - date.today()).days
        r["flag"] = "overdue" if r["days"] < 0 else "soon" if r["days"] <= REMIND_DAYS else "ok"
    return r


def issue_alerts(employee_id=None):
    """Открытые выдачи с просроченным или скорым сроком возврата."""
    limit = (date.today() + timedelta(days=REMIND_DAYS)).isoformat()
    sql, args = ISSUE_SQL + " WHERE s.returned_at IS NULL AND s.due_date<=?", [limit]
    if employee_id is not None:
        sql += " AND s.employee_id=?"
        args.append(employee_id)
    return [decorate_issue(r) for r in q(sql + " ORDER BY s.due_date, s.id", args)]


@app.context_processor
def inject():
    n = 0
    if g.get("user"):
        n = len(issue_alerts(scope_employee()))
        if can("operations"):
            n += q("SELECT COUNT(*) c FROM items WHERE next_check<? AND status<>'Списано'", (today(),), one=True)["c"]
    return dict(user=g.get("user"), can=can, csrf_input=csrf_input, ROLES=ROLES, today=today(),
                notif_count=n, REMIND_DAYS=REMIND_DAYS)


@app.route("/notifications")
@login_required
def notifications():
    alerts = issue_alerts(scope_employee())
    staff = can("operations")
    items = enrich(q(ITEM_SQL)) if staff else []
    return render_template("notifications.html",
                           overdue=[a for a in alerts if a["flag"] == "overdue"],
                           soon=[a for a in alerts if a["flag"] != "overdue"],
                           checks=[i for i in items if i["overdue"]],
                           low=[i for i in items if i["qty"] < i["min_qty"] and i["status"] != "Списано"],
                           smtp_on=bool(os.environ.get("SMTP_HOST")))


def send_reminders():
    """E-mail-напоминания сотрудникам. Возвращает (отправлено, без e-mail, ошибка)."""
    host = os.environ.get("SMTP_HOST")
    if not host:
        return 0, 0, "SMTP не настроен (переменная окружения SMTP_HOST)"
    by_emp = {}
    for a in issue_alerts():
        by_emp.setdefault((a["employee_id"], a["emp_name"], a["emp_email"]), []).append(a)
    sent = skipped = 0
    try:
        srv = smtplib.SMTP(host, int(os.environ.get("SMTP_PORT", "25")), timeout=20)
        if os.environ.get("SMTP_TLS", "0") == "1":
            srv.starttls()
        if os.environ.get("SMTP_USER"):
            srv.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASSWORD", ""))
        for (_, name, email), rows in by_emp.items():
            if not email:
                skipped += 1
                continue
            lines = []
            for a in rows:
                state = (f"ПРОСРОЧЕНО на {-a['days']} дн." if a["days"] < 0
                         else "срок возврата сегодня" if a["days"] == 0 else f"осталось дней: {a['days']}")
                lines.append(f"— {a['item_name']} ({a['inv_no']}), {a['qty']} шт., вернуть до {a['due_date']} — {state}")
            msg = EmailMessage()
            msg["Subject"] = "Напоминание о возврате имущества"
            msg["From"] = os.environ.get("SMTP_FROM", "inventory@localhost")
            msg["To"] = email
            msg.set_content(f"Здравствуйте, {name}!\n\nСроки возврата закреплённого за вами имущества:\n\n"
                            + "\n".join(lines) + "\n\nАвтоматическое сообщение системы учета.")
            srv.send_message(msg)
            sent += 1
        srv.quit()
    except (OSError, smtplib.SMTPException) as e:
        return sent, skipped, f"Ошибка отправки: {e}"
    return sent, skipped, None


@app.route("/notifications/send", methods=["POST"])
@need("operations")
def notifications_send():
    sent, skipped, err = send_reminders()
    audit("Отправка напоминаний", "notifications", "", f"отправлено: {sent}, без e-mail: {skipped}" + (f"; {err}" if err else ""))
    db().commit()
    if err:
        flash(err, "danger")
    else:
        flash(f"Отправлено писем: {sent}" + (f"; у {skipped} сотрудников не указан e-mail" if skipped else ""), "success")
    return redirect(url_for("notifications"))


# ---------------------------------------------------------------- главная
@app.route("/")
@login_required
def dashboard():
    items = enrich(q(ITEM_SQL))
    emp = scope_employee()
    alerts = issue_alerts(emp)
    stats = {
        "total": len(items),
        "ok": sum(i["status"] == "Исправно" for i in items),
        "repair": sum(i["stage"] == "Ремонт" for i in items),
        "overdue": [i for i in items if i["overdue"]],
        "worn": [i for i in items if i["wear"] >= 80 and i["status"] != "Списано"],
        "value": sum(i["price"] * i["qty"] for i in items if i["status"] != "Списано"),
        "issued": sum(i["issued"] for i in items),
        "late": [a for a in alerts if a["flag"] == "overdue"],
        "soon": [a for a in alerts if a["flag"] != "overdue"],
    }
    upcoming = q("SELECT b.*, i.name FROM bookings b JOIN items i ON i.id=b.item_id "
                 "WHERE b.end>=? ORDER BY b.start LIMIT 8", (today(),))
    return render_template("dashboard.html", s=stats, upcoming=upcoming)


# ---------------------------------------------------------------- имущество
FIELD_LABELS = {"inv_no": "Инв. №", "name": "Название", "category": "Категория", "location": "Место хранения",
                "qty": "Количество", "min_qty": "Мин. запас", "price": "Цена", "bought": "Дата покупки",
                "life_years": "Срок службы", "descr": "Описание", "serial": "Серийный №"}


def search_items(a):
    """Поиск и фильтрация: текст, категория, место, статус (наличие), состояние."""
    s = (a.get("s") or "").strip().casefold()
    sql, args, w = ITEM_SQL, [], []
    for key, col in (("cat", "category"), ("loc", "location"), ("state", "status")):
        if a.get(key):
            w.append(f"i.{col}=?")
            args.append(a[key])
    rows = enrich(q(sql + (" WHERE " + " AND ".join(w) if w else "") + " ORDER BY i.name", args))
    if a.get("status"):
        rows = [r for r in rows if r["avail"] == a["status"]]
    if s:
        rows = [r for r in rows if s in " ".join(str(r.get(k) or "") for k in ("name", "inv_no", "serial", "descr")).casefold()]
    return rows


def item_filters_ctx():
    return dict(cats=names("categories"), locs=names("locations"), states=names("states"), avail=AVAIL_STATUSES)


@app.route("/items")
@login_required
def items():
    a = request.args
    return render_template("items.html", items=search_items(a), a=a, **item_filters_ctx())


def parse_item(f, min_total=1, issued=0):
    """Проверка формы карточки. Возвращает (данные, ошибки)."""
    err, d = [], {}
    d["inv_no"] = f.get("inv_no", "").strip()
    d["name"] = f.get("name", "").strip()
    d["descr"] = f.get("descr", "").strip()
    d["serial"] = f.get("serial", "").strip()
    if not d["name"]:
        err.append("Укажите название")
    d["category"] = f.get("category", "").strip()
    if d["category"] not in names("categories"):
        err.append("Выберите категорию из справочника")
    d["location"] = f.get("location", "").strip()
    if d["location"] not in names("locations"):
        err.append("Выберите место хранения из справочника")
    d["qty"] = to_int(f.get("qty"))
    if d["qty"] is None or d["qty"] < min_total:
        err.append(f"Количество должно быть целым числом не меньше {min_total}"
                   + (f" (сейчас выдано {issued} шт.)" if issued and min_total == issued else ""))
    d["min_qty"] = to_int(f.get("min_qty"), 0)
    if d["min_qty"] < 0:
        err.append("Минимальный запас не может быть отрицательным")
    d["price"] = to_float(f.get("price"), None)
    if d["price"] is None or d["price"] < 0:
        err.append("Цена указана некорректно")
    d["bought"] = f.get("bought", "").strip() or today()
    if not valid_date(d["bought"]):
        err.append("Дата покупки указана некорректно")
    d["life_years"] = to_float(f.get("life_years"), None)
    if d["life_years"] is None or d["life_years"] <= 0:
        err.append("Срок службы должен быть положительным числом")
    return d, err


@app.route("/items/new", methods=["GET", "POST"])
@need("edit_items")
def item_new():
    form = request.form if request.method == "POST" else {"qty": 1, "min_qty": 0, "price": 0, "life_years": 5,
                                                         "bought": today(), "state": OK_STATE}
    if request.method == "POST":
        d, err = parse_item(form, 1)
        state = form.get("state", OK_STATE)
        if state not in names("states"):
            err.append("Выберите состояние из справочника")
        if not err:
            auto = not d["inv_no"]
            for _ in range(5 if auto else 1):
                inv = d["inv_no"] or next_inv_no()
                try:
                    cur = x("INSERT INTO items(inv_no,name,category,location,qty,min_qty,status,price,bought,"
                            "life_years,next_check,descr,serial) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (inv, d["name"], d["category"], d["location"], d["qty"], d["min_qty"], state,
                             d["price"], d["bought"], d["life_years"],
                             (date.today() + timedelta(days=CHECK_PERIOD_DAYS)).isoformat(), d["descr"], d["serial"]))
                    break
                except sqlite3.IntegrityError:
                    inv = None
            if inv is None:
                err.append("Инвентарный номер уже существует")
            else:
                log_move(cur.lastrowid, "Постановка на учет", d["qty"],
                         f"Инв. № {inv}" + (" (присвоен автоматически)" if auto else "") + f"; состояние: {state}")
                audit("Добавлена карточка", "items", cur.lastrowid, f"{inv} {d['name']}, {d['qty']} шт.")
                db().commit()
                flash(f"Объект поставлен на учет. Инвентарный номер: {inv}", "success")
                return redirect(url_for("item", iid=cur.lastrowid))
        for e in err:
            flash(e, "danger")
    return render_template("item_form.html", it=None, f=form, next_inv=next_inv_no(), **item_filters_ctx())


@app.route("/items/<int:iid>/edit", methods=["GET", "POST"])
@need("edit_items")
def item_edit(iid):
    it = get_item(iid)
    if not it:
        abort(404)
    form = request.form if request.method == "POST" else it
    if request.method == "POST":
        d, err = parse_item(form, max(it["issued"], 0), it["issued"])
        if can("admin") and d["inv_no"]:
            pass
        else:
            d["inv_no"] = it["inv_no"]  # менять инв. номер может только администратор
        if not err:
            try:
                x("UPDATE items SET inv_no=?,name=?,category=?,location=?,qty=?,min_qty=?,price=?,bought=?,"
                  "life_years=?,descr=?,serial=? WHERE id=?",
                  (d["inv_no"], d["name"], d["category"], d["location"], d["qty"], d["min_qty"], d["price"],
                   d["bought"], d["life_years"], d["descr"], d["serial"], iid))
            except sqlite3.IntegrityError:
                err.append("Инвентарный номер уже существует")
        if not err:
            diff = [f"{FIELD_LABELS[k]}: {it.get(k) or '—'} → {d[k] or '—'}" for k in FIELD_LABELS
                    if str(it.get(k) if it.get(k) is not None else "") != str(d[k] if d[k] is not None else "")]
            if diff:
                kind = "Перемещение" if it["location"] != d["location"] else "Изменение карточки"
                log_move(iid, kind, d["qty"], "; ".join(diff))
                audit("Изменена карточка", "items", iid, "; ".join(diff))
            db().commit()
            flash("Карточка сохранена", "success")
            return redirect(url_for("item", iid=iid))
        db().rollback()
        for e in err:
            flash(e, "danger")
    return render_template("item_form.html", it=it, f=form, next_inv=None, **item_filters_ctx())


MOVE_SQL = ("SELECT m.*, i.name AS item_name, i.inv_no, e.full_name AS emp_name, ev.name AS event_name, "
            "u.full_name AS user_name FROM movements m JOIN items i ON i.id=m.item_id "
            "LEFT JOIN employees e ON e.id=m.employee_id LEFT JOIN events ev ON ev.id=m.event_id "
            "LEFT JOIN users u ON u.id=m.user_id")


@app.route("/items/<int:iid>")
@login_required
def item(iid):
    it = get_item(iid)
    if not it:
        return redirect(url_for("items"))
    staff = can("operations")
    emp = scope_employee()
    issues = [decorate_issue(r) for r in q(ISSUE_SQL + " WHERE s.item_id=?" + ("" if staff else " AND s.employee_id=?")
                                           + " ORDER BY s.id DESC", (iid,) if staff else (iid, emp))]
    return render_template("item.html", it=it, states=names("states"), issues=issues,
                           moves=q(MOVE_SQL + " WHERE m.item_id=? ORDER BY m.id DESC", (iid,)) if staff else [],
                           checks=q("SELECT * FROM checks WHERE item_id=? ORDER BY id DESC", (iid,)),
                           bookings=q("SELECT * FROM bookings WHERE item_id=? ORDER BY start DESC", (iid,)))


@app.route("/items/<int:iid>/check", methods=["POST"])
@need("edit_items")
def item_check(iid):
    it = get_item(iid)
    if not it:
        abort(404)
    st, note = request.form.get("status", ""), request.form.get("note", "").strip()
    if st not in names("states") or not note:
        flash("Выберите состояние и укажите комментарий", "danger")
    elif st == "Списано" and it["issued"]:
        flash(f"Нельзя списать: не возвращено {it['issued']} шт. Сначала оформите возврат", "danger")
    else:
        x("INSERT INTO checks(item_id,day,status,note) VALUES(?,?,?,?)", (iid, today(), st, note))
        x("UPDATE items SET status=?, next_check=? WHERE id=?",
          (st, (date.today() + timedelta(days=CHECK_PERIOD_DAYS)).isoformat(), iid))
        if st != it["status"]:
            log_move(iid, "Смена состояния", None, f"{it['status']} → {st}. {note}")
        else:
            log_move(iid, "Проверка состояния", None, f"{st}. {note}")
        audit("Акт проверки состояния", "items", iid, f"{it['status']} → {st}: {note}")
        db().commit()
        flash("Состояние обновлено", "success")
    return redirect(url_for("item", iid=iid))


# ---------------------------------------------------------------- выдача и возврат
@app.route("/issues")
@login_required
def issues():
    a = request.args
    tab = a.get("tab", "open")
    w, args = [], []
    emp = scope_employee()
    if emp is not None:
        w.append("s.employee_id=?")
        args.append(emp)
    elif a.get("emp"):
        w.append("s.employee_id=?")
        args.append(to_int(a["emp"], -1))
    if a.get("event"):
        w.append("s.event_id=?")
        args.append(to_int(a["event"], -1))
    if tab in ("open", "overdue"):
        w.append("s.returned_at IS NULL")
    if tab == "overdue":
        w.append("s.due_date<?")
        args.append(today())
    if tab == "closed":
        w.append("s.returned_at IS NOT NULL")
    rows = [decorate_issue(r) for r in q(ISSUE_SQL + (" WHERE " + " AND ".join(w) if w else "")
                                          + " ORDER BY s.returned_at IS NOT NULL, s.due_date, s.id DESC", args)]
    s = (a.get("s") or "").strip().casefold()
    if s:
        rows = [r for r in rows if s in f"{r['item_name']} {r['inv_no']} {r['emp_name']}".casefold()]
    return render_template("issues.html", rows=rows, tab=tab, a=a, emps=q("SELECT id, full_name FROM employees ORDER BY full_name"),
                           events=event_choices())


@app.route("/issues/new", methods=["GET", "POST"])
@need("operations")
def issue_new():
    avail = [i for i in enrich(q(ITEM_SQL + " ORDER BY i.name")) if i["status"] == OK_STATE and i["free"] > 0]
    form = request.form if request.method == "POST" else {
        "item_id": request.args.get("item_id", ""), "event_id": request.args.get("event_id", ""),
        "employee_id": request.args.get("employee_id", ""), "qty": request.args.get("qty", 1),
        "due_date": (date.today() + timedelta(days=DEFAULT_ISSUE_DAYS)).isoformat()}
    if request.method == "POST":
        f, err = request.form, []
        it = next((i for i in avail if str(i["id"]) == f.get("item_id")), None)
        emp = q("SELECT * FROM employees WHERE id=? AND active=1", (to_int(f.get("employee_id"), -1),), one=True)
        ev = q("SELECT * FROM events WHERE id=?", (to_int(f.get("event_id"), -1),), one=True)
        qty, due = to_int(f.get("qty")), f.get("due_date", "")
        if not it:
            err.append("Выберите объект: он должен быть исправен и доступен для выдачи")
        elif qty is None or not 1 <= qty <= it["free"]:
            err.append(f"Количество должно быть от 1 до {it['free']} (свободно сейчас)")
        if not emp:
            err.append("Выберите сотрудника из справочника")
        if not ev:
            err.append("Выберите мероприятие (для хозяйственных нужд есть отдельная запись)")
        if not valid_date(due) or due < today():
            err.append("Срок возврата указан некорректно (не раньше сегодняшней даты)")
        if not err:
            cur = x("INSERT INTO issues(item_id,employee_id,event_id,qty,issued_at,due_date,issued_by,issue_note) "
                    "VALUES(?,?,?,?,?,?,?,?)", (it["id"], emp["id"], ev["id"], qty, today(), due,
                                               g.user["id"], f.get("note", "").strip()))
            log_move(it["id"], "Выдача", qty, f"Срок возврата {due}. {f.get('note', '').strip()}".strip(),
                     emp["id"], ev["id"])
            audit("Выдача имущества", "issues", cur.lastrowid, f"{it['inv_no']} {it['name']} ×{qty} → {emp['full_name']}")
            db().commit()
            flash(f"Выдано: {it['name']} × {qty} — {emp['full_name']}", "success")
            return redirect(url_for("issues"))
        for e in err:
            flash(e, "danger")
    return render_template("issue_new.html", f=form, avail=avail, emps=active_employees(), events=event_choices())


@app.route("/issues/<int:sid>/return", methods=["GET", "POST"])
@need("operations")
def issue_return(sid):
    s = q(ISSUE_SQL + " WHERE s.id=?", (sid,), one=True)
    if not s or s["returned_at"]:
        flash("Выдача не найдена или уже закрыта", "warning")
        return redirect(url_for("issues"))
    decorate_issue(s)
    if request.method == "POST":
        f, err = request.form, []
        rq, state, note = to_int(f.get("qty")), f.get("state", ""), f.get("note", "").strip()
        if rq is None or not 1 <= rq <= s["qty"]:
            err.append(f"Возвращаемое количество — от 1 до {s['qty']}")
        if state not in names("states"):
            err.append("Выберите состояние при возврате")
        elif state != OK_STATE and not note:
            err.append("Если состояние отличается от «Исправно», опишите повреждение в комментарии")
        if not err:
            it = get_item(s["item_id"])
            if rq < s["qty"]:  # частичный возврат: закрываем часть записи, остаток остаётся открытым
                x("UPDATE issues SET qty=qty-? WHERE id=?", (rq, sid))
                x("INSERT INTO issues(item_id,employee_id,event_id,qty,issued_at,due_date,issued_by,issue_note,"
                  "returned_at,return_state,return_note,received_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                  (s["item_id"], s["employee_id"], s["event_id"], rq, s["issued_at"], s["due_date"],
                   s["issued_by"], s["issue_note"], today(), state, note, g.user["id"]))
            else:
                x("UPDATE issues SET returned_at=?, return_state=?, return_note=?, received_by=? WHERE id=?",
                  (today(), state, note, g.user["id"], sid))
            extra = ""
            if state != OK_STATE:
                if rq == it["qty"]:  # вернули всю партию — состояние карточки меняется
                    x("INSERT INTO checks(item_id,day,status,note) VALUES(?,?,?,?)",
                      (it["id"], today(), state, f"Возврат от {s['emp_name']}: {note}"))
                    x("UPDATE items SET status=? WHERE id=?", (state, it["id"]))
                    log_move(it["id"], "Смена состояния", None, f"{it['status']} → {state} (при возврате). {note}",
                             s["employee_id"], s["event_id"])
                else:
                    extra = f" (в карточке {it['qty']} шт.: состояние карточки не менялось)"
                    flash("Состояние зафиксировано в истории, но в карточке с несколькими единицами оно не изменено. "
                          "При необходимости уменьшите количество или заведите отдельную карточку на повреждённые единицы.",
                          "warning")
            log_move(s["item_id"], "Возврат", rq, f"Состояние: {state}. {note}".strip() + extra,
                     s["employee_id"], s["event_id"])
            audit("Возврат имущества", "issues", sid, f"{s['inv_no']} ×{rq}, состояние: {state}")
            db().commit()
            flash("Возврат оформлен", "success")
            return redirect(url_for("issues"))
        for e in err:
            flash(e, "danger")
    return render_template("issue_return.html", s=s, states=names("states"), f=request.form)


# ---------------------------------------------------------------- бронирование
@app.route("/bookings", methods=["GET", "POST"])
@login_required
def bookings():
    if request.method == "POST":
        if not can("operations"):
            abort(403)
        f = request.form
        it = q("SELECT * FROM items WHERE id=?", (to_int(f.get("item_id"), -1),), one=True)
        ev = q("SELECT * FROM events WHERE id=?", (to_int(f.get("event_id"), -1),), one=True)
        emp = q("SELECT * FROM employees WHERE id=? AND active=1", (to_int(f.get("employee_id"), -1),), one=True)
        qty = to_int(f.get("qty"))
        if not (it and ev and emp and qty and qty > 0 and valid_date(f.get("start")) and valid_date(f.get("end"))):
            flash("Заполните все поля корректно", "danger")
        elif f["end"] < f["start"]:
            flash("Дата окончания раньше даты начала", "danger")
        elif it["status"] != "Исправно":
            flash(f"Нельзя забронировать: состояние «{it['status']}»", "danger")
        elif free_qty(it["id"], f["start"], f["end"]) < qty:
            flash("Недостаточно свободных единиц на выбранные даты (конфликт бронирования)", "danger")
        else:
            cur = x("INSERT INTO bookings(item_id,qty,event,start,end,person,event_id,employee_id) VALUES(?,?,?,?,?,?,?,?)",
                    (it["id"], qty, ev["name"], f["start"], f["end"], emp["full_name"], ev["id"], emp["id"]))
            audit("Создана бронь", "bookings", cur.lastrowid, f"{it['inv_no']} ×{qty}, {f['start']}…{f['end']}, {ev['name']}")
            db().commit()
            flash("Бронь создана", "success")
        return redirect(url_for("bookings"))
    rows = q("SELECT b.*, i.name, i.inv_no FROM bookings b JOIN items i ON i.id=b.item_id ORDER BY b.start DESC")
    return render_template("bookings.html", rows=rows, events=event_choices(), emps=active_employees(),
                           items=q("SELECT id,name,inv_no,qty FROM items WHERE status='Исправно' ORDER BY name"))


@app.route("/bookings/<int:bid>/delete", methods=["POST"])
@need("operations")
def booking_delete(bid):
    b = q("SELECT b.*, i.inv_no FROM bookings b JOIN items i ON i.id=b.item_id WHERE b.id=?", (bid,), one=True)
    if b:
        x("DELETE FROM bookings WHERE id=?", (bid,))
        audit("Отменена бронь", "bookings", bid, f"{b['inv_no']} ×{b['qty']}, {b['event']}")
        db().commit()
        flash("Бронь отменена", "info")
    return redirect(url_for("bookings"))


# ---------------------------------------------------------------- закупки
@app.route("/purchases", methods=["GET", "POST"])
@need("operations")
def purchases():
    if request.method == "POST":
        action = request.form.get("action")
        if action == "auto":
            n = suggest_purchases()
            audit("Формирование плана закупок", "purchases", "", f"добавлено позиций: {n}")
            db().commit()
            flash(f"Добавлено позиций в план закупок: {n}", "info")
        elif action == "status":
            pid, st = to_int(request.form.get("id"), -1), request.form.get("status")
            p = q("SELECT * FROM purchases WHERE id=?", (pid,), one=True)
            if p and st in ("Запланирована", "Заказана", "Получена", "Отменена") and p["status"] != st:
                if p["status"] == "Получена":
                    flash("Закупка уже оформлена как полученная — статус не меняется", "warning")
                else:
                    x("UPDATE purchases SET status=? WHERE id=?", (st, pid))
                    if st == "Получена" and p["item_id"]:  # закрытие цикла: возврат в эксплуатацию
                        x("UPDATE items SET qty=qty+?, status='Исправно' WHERE id=?", (p["qty"], p["item_id"]))
                        log_move(p["item_id"], "Поступление", p["qty"], f"Закупка №{pid}: {p['reason']}")
                    audit("Изменён статус закупки", "purchases", pid, f"{p['status']} → {st}")
                    db().commit()
        return redirect(url_for("purchases"))
    return render_template("purchases.html", rows=q("SELECT * FROM purchases ORDER BY id DESC"),
                           states=["Запланирована", "Заказана", "Получена", "Отменена"])


# ---------------------------------------------------------------- справочники
DICTS = {
    "categories": dict(title="Категории", table="categories", cols=[("name", "Название", "text", True)],
                       text_refs=[("items", "category")], id_refs=[], protected=[]),
    "states": dict(title="Состояния", table="states", cols=[("name", "Название", "text", True)],
                   text_refs=[("items", "status"), ("checks", "status"), ("issues", "return_state")],
                   id_refs=[], protected=SYSTEM_STATES),
    "locations": dict(title="Места хранения", table="locations", cols=[("name", "Название", "text", True)],
                      text_refs=[("items", "location")], id_refs=[], protected=[]),
    "employees": dict(title="Сотрудники", table="employees",
                      cols=[("full_name", "ФИО", "text", True), ("position", "Должность", "text", False),
                            ("phone", "Телефон", "text", False), ("email", "E-mail", "email", False),
                            ("active", "Работает", "bool", False)],
                      text_refs=[], protected=[],
                      id_refs=[("issues", "employee_id"), ("bookings", "employee_id"),
                               ("users", "employee_id"), ("movements", "employee_id")]),
    "events": dict(title="Мероприятия", table="events",
                   cols=[("name", "Название", "text", True), ("event_date", "Дата", "date", False),
                         ("place", "Площадка", "text", False), ("note", "Примечание", "text", False)],
                   text_refs=[], protected=[NO_EVENT],
                   id_refs=[("issues", "event_id"), ("bookings", "event_id"), ("movements", "event_id")]),
}


def dict_or_404(name):
    if name not in DICTS:
        abort(404)
    return DICTS[name]


def dict_usage(d, row):
    key = d["cols"][0][0]
    n = sum(q(f"SELECT COUNT(*) c FROM {t} WHERE {c}=?", (row[key],), one=True)["c"] for t, c in d["text_refs"])
    return n + sum(q(f"SELECT COUNT(*) c FROM {t} WHERE {c}=?", (row["id"],), one=True)["c"] for t, c in d["id_refs"])


def dict_values(d, f):
    vals, err = {}, []
    for col, label, typ, req in d["cols"]:
        v = f.get(col, "").strip()
        if typ == "bool":
            v = 1 if f.get(col) else 0
        elif req and not v:
            err.append(f"Поле «{label}» обязательно")
        elif typ == "date" and v and not valid_date(v):
            err.append(f"Поле «{label}»: некорректная дата")
        elif typ == "email" and v and not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", v):
            err.append(f"Поле «{label}»: некорректный e-mail")
        vals[col] = v
    return vals, err


@app.route("/dict/<name>", methods=["GET", "POST"])
@need("dictionaries")
def dictionary(name):
    d = dict_or_404(name)
    if request.method == "POST":
        vals, err = dict_values(d, request.form)
        if not err:
            cols = list(vals)
            try:
                cur = x(f"INSERT INTO {d['table']}({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                        [vals[c] for c in cols])
                audit("Справочник: добавлена запись", d["table"], cur.lastrowid, f"{d['title']}: {vals[cols[0]]}")
                db().commit()
                flash("Запись добавлена", "success")
                return redirect(url_for("dictionary", name=name))
            except sqlite3.IntegrityError:
                err.append("Такая запись уже существует")
        for e in err:
            flash(e, "danger")
    order = d["cols"][0][0]
    rows = q(f"SELECT * FROM {d['table']} ORDER BY " + ("COALESCE(event_date,'') DESC, name" if name == "events" else order))
    for r in rows:
        r["used"] = dict_usage(d, r)
        r["locked"] = r[order] in d["protected"]
    return render_template("dict.html", d=d, name=name, rows=rows, dicts=DICTS, order=order,
                           f=request.form if request.method == "POST" else {"active": "1"})


@app.route("/dict/<name>/<int:rid>/edit", methods=["GET", "POST"])
@need("dictionaries")
def dictionary_edit(name, rid):
    d = dict_or_404(name)
    row = q(f"SELECT * FROM {d['table']} WHERE id=?", (rid,), one=True)
    if not row:
        abort(404)
    key = d["cols"][0][0]
    locked = row[key] in d["protected"]
    if request.method == "POST":
        vals, err = dict_values(d, request.form)
        if locked:
            err.append("Системная запись: изменение и удаление запрещены, на неё опирается логика приложения")
        if not err:
            try:
                x(f"UPDATE {d['table']} SET " + ",".join(f"{c}=?" for c in vals) + " WHERE id=?", [*vals.values(), rid])
                if vals[key] != row[key]:  # переименование распространяем на карточки/акты
                    for t, c in d["text_refs"]:
                        x(f"UPDATE {t} SET {c}=? WHERE {c}=?", (vals[key], row[key]))
                audit("Справочник: изменена запись", d["table"], rid, f"{d['title']}: {row[key]} → {vals[key]}")
                db().commit()
                flash("Запись сохранена", "success")
                return redirect(url_for("dictionary", name=name))
            except sqlite3.IntegrityError:
                db().rollback()
                err.append("Такая запись уже существует")
        for e in err:
            flash(e, "danger")
    return render_template("dict_edit.html", d=d, name=name, row=row, locked=locked,
                           f=request.form if request.method == "POST" else row)


@app.route("/dict/<name>/<int:rid>/delete", methods=["POST"])
@need("dictionaries")
def dictionary_delete(name, rid):
    d = dict_or_404(name)
    row = q(f"SELECT * FROM {d['table']} WHERE id=?", (rid,), one=True)
    if row:
        key = d["cols"][0][0]
        if row[key] in d["protected"]:
            flash("Системную запись удалить нельзя", "danger")
        elif dict_usage(d, row):
            flash("Запись используется в учетных данных — удалить нельзя"
                  + (" (сотрудника можно отметить как неработающего)" if name == "employees" else ""), "danger")
        else:
            x(f"DELETE FROM {d['table']} WHERE id=?", (rid,))
            audit("Справочник: удалена запись", d["table"], rid, f"{d['title']}: {row[key]}")
            db().commit()
            flash("Запись удалена", "info")
    return redirect(url_for("dictionary", name=name))


# ---------------------------------------------------------------- инвентаризация
LINE_SQL = ("SELECT l.*, i.name, i.inv_no, i.category, i.location FROM inventory_lines l "
            "JOIN items i ON i.id=l.item_id WHERE l.inv_id=? ORDER BY i.location, i.name")


@app.route("/inventory", methods=["GET", "POST"])
@need("operations")
def inventory():
    if request.method == "POST":
        f = request.form
        loc, cat = f.get("location", ""), f.get("category", "")
        w, args = ["status<>'Списано'"], []
        for col, v in (("location", loc), ("category", cat)):
            if v:
                w.append(f"{col}=?")
                args.append(v)
        items_ = q(ITEM_SQL + " WHERE " + " AND ".join("i." + c for c in w), args)
        title = f.get("title", "").strip() or f"Инвентаризация от {today()}"
        scope = ", ".join(filter(None, [f"место: {loc}" if loc else "", f"категория: {cat}" if cat else ""])) or "все имущество"
        if not items_:
            flash("В выбранной области нет имущества", "danger")
        else:
            cur = x("INSERT INTO inventories(title,created,scope,created_by) VALUES(?,?,?,?)",
                    (title, now(), scope, g.user["id"]))
            for i in items_:
                x("INSERT INTO inventory_lines(inv_id,item_id,expected,issued) VALUES(?,?,?,?)",
                  (cur.lastrowid, i["id"], i["qty"] - i["issued"], i["issued"]))
            audit("Начата инвентаризация", "inventories", cur.lastrowid, f"{title} ({scope}), позиций: {len(items_)}")
            db().commit()
            return redirect(url_for("inventory_view", inv_id=cur.lastrowid))
    rows = q("SELECT v.*, (SELECT COUNT(*) FROM inventory_lines WHERE inv_id=v.id) n, "
             "(SELECT COUNT(*) FROM inventory_lines WHERE inv_id=v.id AND actual IS NOT NULL) done, "
             "(SELECT COUNT(*) FROM inventory_lines WHERE inv_id=v.id AND actual IS NOT NULL AND actual<>expected) diff "
             "FROM inventories v ORDER BY v.id DESC")
    return render_template("inventory.html", rows=rows, cats=names("categories"), locs=names("locations"))


@app.route("/inventory/<int:inv_id>", methods=["GET", "POST"])
@need("operations")
def inventory_view(inv_id):
    inv = q("SELECT * FROM inventories WHERE id=?", (inv_id,), one=True)
    if not inv:
        abort(404)
    lines = q(LINE_SQL, (inv_id,))
    if request.method == "POST" and inv["status"] == "Идёт":
        f, err = request.form, []
        for l in lines:
            raw = f.get(f"a{l['id']}", "").strip()
            l["new_actual"] = None if raw == "" else to_int(raw)
            if raw != "" and (l["new_actual"] is None or l["new_actual"] < 0):
                err.append(f"«{l['name']}»: фактическое количество — целое число ≥ 0")
            l["new_note"] = f.get(f"n{l['id']}", "").strip()
        action = f.get("action", "save")
        if not err and action != "save" and any(l["new_actual"] is None for l in lines):
            err.append(f"Для завершения заполните фактическое количество по всем позициям "
                       f"(не заполнено: {sum(l['new_actual'] is None for l in lines)})")
        if err:
            for e in err[:5]:
                flash(e, "danger")
        else:
            for l in lines:
                x("UPDATE inventory_lines SET actual=?, note=? WHERE id=?", (l["new_actual"], l["new_note"], l["id"]))
            if action == "save":
                audit("Сохранены результаты инвентаризации", "inventories", inv_id)
                db().commit()
                flash("Результаты сохранены", "success")
            else:
                apply = action == "finish_apply"
                fixed = 0
                for l in lines:
                    diff = l["new_actual"] - l["expected"]
                    if diff and apply:
                        x("UPDATE items SET qty=qty+? WHERE id=?", (diff, l["item_id"]))
                        log_move(l["item_id"], "Инвентаризация", diff,
                                 f"«{inv['title']}»: по учёту {l['expected']}, фактически {l['new_actual']} "
                                 f"→ корректировка {diff:+d}")
                        fixed += 1
                    elif diff:
                        log_move(l["item_id"], "Инвентаризация (расхождение)", diff,
                                 f"«{inv['title']}»: по учёту {l['expected']}, фактически {l['new_actual']}; без корректировки")
                    else:
                        log_move(l["item_id"], "Инвентаризация (сверка)", l["new_actual"], f"«{inv['title']}»: расхождений нет")
                x("UPDATE inventories SET status='Завершена', finished=?, applied=? WHERE id=?",
                  (now(), 1 if apply else 0, inv_id))
                audit("Завершена инвентаризация", "inventories", inv_id,
                      f"{inv['title']}; корректировок: {fixed}" if apply else f"{inv['title']}; без корректировки")
                db().commit()
                flash("Инвентаризация завершена" + (f", скорректировано позиций: {fixed}" if apply else ""), "success")
            return redirect(url_for("inventory_view", inv_id=inv_id))
        db().rollback()
    inv = q("SELECT * FROM inventories WHERE id=?", (inv_id,), one=True)
    lines = q(LINE_SQL, (inv_id,))
    return render_template("inventory_view.html", inv=inv, lines=lines)


# ---------------------------------------------------------------- отчеты
def fmt_value(v, kind):
    if v is None or v == "":
        return ""
    if kind == "int":
        return f"{int(v):,}".replace(",", "\u00a0")
    if kind == "money":
        return f"{float(v):,.2f}".replace(",", "\u00a0").replace(".", ",")
    return str(v)


def finish_rep(rep, a):
    rep["disp"] = [[fmt_value(v, k) for v, k in zip(r, rep["kinds"])] for r in rep["rows"]]
    rep["totals_disp"] = [fmt_value(v, k) for v, k in zip(rep["totals"], rep["kinds"])] if rep.get("totals") else None
    rep.setdefault("vals", dict(a))
    parts = []
    for name, label, typ, opts in rep["filters"]:
        v = rep["vals"].get(name)
        if v:
            if typ == "emp":
                v = next((o["full_name"] for o in opts if str(o["id"]) == v), v)
            parts.append(f"{label}: {'да' if typ == 'check' else v}")
    rep["subtitle"] = "; ".join(parts) if parts else "без ограничений"
    return rep


def rep_stock(a):
    rows = search_items(a)
    data = [[r["inv_no"], r["name"], r["category"], r["location"], r["status"], r["avail"], r["qty"], r["issued"],
             r["free"], r["price"], r["price"] * r["qty"]] for r in rows]
    tot = ["", "Итого", "", "", "", "", sum(r[6] for r in data), sum(r[7] for r in data), sum(r[8] for r in data),
           None, sum(r[10] for r in data)]
    c = item_filters_ctx()
    return finish_rep(dict(
        title="Наличие имущества", headers=["Инв. №", "Наименование", "Категория", "Место хранения", "Состояние",
                                            "Статус", "Всего", "Выдано", "В наличии", "Цена, ₽", "Стоимость, ₽"],
        kinds=["text"] * 6 + ["int"] * 3 + ["money"] * 2, widths=[8, 22, 14, 13, 12, 12, 7, 8, 9, 13, 15],
        rows=data, totals=tot,
        filters=[("cat", "Категория", "select", c["cats"]), ("loc", "Место хранения", "select", c["locs"]),
                 ("status", "Статус", "select", c["avail"]), ("state", "Состояние", "select", c["states"]),
                 ("s", "Поиск", "text", None)]), a)


def rep_movements(a):
    df = a["date_from"] if "date_from" in a else (date.today() - timedelta(days=30)).isoformat()
    dt = a.get("date_to", "")
    w, args = [], []
    if valid_date(df):
        w.append("substr(m.ts,1,10)>=?")
        args.append(df)
    if valid_date(dt):
        w.append("substr(m.ts,1,10)<=?")
        args.append(dt)
    for key, col in (("kind", "m.kind"), ("cat", "i.category")):
        if a.get(key):
            w.append(f"{col}=?")
            args.append(a[key])
    rows = q(MOVE_SQL + (" WHERE " + " AND ".join(w) if w else "") + " ORDER BY m.id DESC LIMIT 5000", args)
    s = (a.get("s") or "").strip().casefold()
    if s:
        rows = [r for r in rows if s in f"{r['item_name']} {r['inv_no']}".casefold()]
    data = [[r["ts"][:16], r["inv_no"], r["item_name"], r["kind"], r["qty"], r["emp_name"] or "", r["event_name"] or "",
             r["user_name"] or "", r["details"] or ""] for r in rows]
    kinds = [r["kind"] for r in q("SELECT DISTINCT kind FROM movements ORDER BY kind")]
    return finish_rep(dict(
        title="Движение имущества", headers=["Дата и время", "Инв. №", "Наименование", "Операция", "Кол-во",
                                             "Сотрудник", "Мероприятие", "Выполнил", "Примечание"],
        kinds=["text"] * 4 + ["int"] + ["text"] * 4, widths=[15, 8, 24, 16, 6, 18, 20, 16, 34], rows=data,
        totals=["", "", f"Операций: {len(data)}", "", None, "", "", "", ""],
        vals={**a, "date_from": df},
        filters=[("date_from", "С даты", "date", None), ("date_to", "По дату", "date", None),
                 ("kind", "Операция", "select", kinds), ("cat", "Категория", "select", names("categories")),
                 ("s", "Поиск", "text", None)]), a)


def rep_debts(a):
    w, args = ["s.returned_at IS NULL"], []
    if a.get("emp"):
        w.append("s.employee_id=?")
        args.append(to_int(a["emp"], -1))
    if a.get("overdue"):
        w.append("s.due_date<?")
        args.append(today())
    rows = [decorate_issue(r) for r in q(ISSUE_SQL + " WHERE " + " AND ".join(w) + " ORDER BY e.full_name, s.due_date", args)]
    data = [[r["emp_name"], r["inv_no"], r["item_name"], r["qty"], r["event_name"] or "", r["issued_at"],
             r["due_date"], max(0, -r["days"]), r["issued_by_name"] or ""] for r in rows]
    return finish_rep(dict(
        title="Задолженности по возврату имущества",
        headers=["Сотрудник", "Инв. №", "Наименование", "Кол-во", "Мероприятие", "Выдано", "Вернуть до",
                 "Просрочка, дн.", "Выдал"],
        kinds=["text"] * 3 + ["int"] + ["text"] * 3 + ["int", "text"], widths=[22, 8, 25, 7, 24, 11, 11, 11, 17],
        rows=data, totals=["Итого", "", f"Записей: {len(data)}", sum(r[3] for r in data), "", "", "", None, ""],
        filters=[("emp", "Сотрудник", "emp", q("SELECT id, full_name FROM employees ORDER BY full_name")),
                 ("overdue", "Только просроченные", "check", None)]), a)


def rep_categories(a):
    agg = {n: dict(pos=0, qty=0, out=0, ok=0, rep=0, off=0, val=0.0) for n in names("categories")}
    for i in enrich(q(ITEM_SQL)):
        c = agg.setdefault(i["category"] or "(без категории)", dict(pos=0, qty=0, out=0, ok=0, rep=0, off=0, val=0.0))
        c["pos"] += 1
        if i["status"] == "Списано":
            c["off"] += 1
            continue
        c["qty"] += i["qty"]
        c["out"] += i["issued"]
        c["val"] += i["price"] * i["qty"]
        c["rep" if i["stage"] == "Ремонт" else "ok"] += 1
    data = [[n, c["pos"], c["qty"], c["out"], c["qty"] - c["out"], c["ok"], c["rep"], c["off"], c["val"]]
            for n, c in sorted(agg.items())]
    tot = ["Итого"] + [sum(r[i] for r in data) for i in range(1, 9)]
    return finish_rep(dict(
        title="Сводка по категориям", headers=["Категория", "Позиций", "Единиц (без списанных)", "Выдано", "В наличии",
                                               "Исправно, поз.", "Ремонт, поз.", "Списано, поз.", "Стоимость, ₽"],
        kinds=["text"] + ["int"] * 7 + ["money"], widths=[26, 9, 14, 9, 10, 10, 10, 10, 17], rows=data, totals=tot,
        filters=[]), a)


def rep_inventory(inv_id):
    inv = q("SELECT * FROM inventories WHERE id=?", (inv_id,), one=True)
    if not inv:
        abort(404)
    lines = q(LINE_SQL, (inv_id,))
    data = [[l["inv_no"], l["name"], l["location"], l["expected"], l["issued"], l["actual"],
             (l["actual"] - l["expected"]) if l["actual"] is not None else None, l["note"] or ""] for l in lines]
    diff = sum(1 for r in data if r[6])
    rep = dict(
        title=f"Сличительная ведомость: {inv['title']}",
        headers=["Инв. №", "Наименование", "Место хранения", "По учёту (на месте)", "Выдано (у сотрудников)",
                 "Фактически", "Расхождение", "Примечание"],
        kinds=["text"] * 3 + ["int"] * 4 + ["text"], widths=[9, 30, 16, 12, 12, 11, 11, 30], rows=data,
        totals=["", f"Позиций: {len(data)}, с расхождениями: {diff}", "", None, None, None, None, ""], filters=[])
    finish_rep(rep, {})
    rep["subtitle"] = f"Область: {inv['scope']}; начата {inv['created'][:16]}; статус: {inv['status']}"
    return rep


REPORTS = {
    "stock": dict(title="Наличие имущества", fn=rep_stock, hint="Остатки, выдано и свободно, стоимость. Фильтры: категория, место, статус, состояние."),
    "movements": dict(title="Движение имущества", fn=rep_movements, hint="Постановка на учёт, выдачи, возвраты, перемещения, корректировки за период."),
    "debts": dict(title="Задолженности", fn=rep_debts, hint="Невозвращённое имущество по сотрудникам, просрочка в днях."),
    "categories": dict(title="По категориям", fn=rep_categories, hint="Сводка: количество, выдано, состояние, стоимость фонда."),
}


_FONT = None


def pdf_fonts():
    global _FONT
    if _FONT:
        return _FONT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.pdfmetrics import registerFontFamily
    from reportlab.pdfbase.ttfonts import TTFont
    fd = os.path.join(BASE_DIR, "fonts")
    cands = [(os.path.join(fd, "DejaVuSans.ttf"), os.path.join(fd, "DejaVuSans-Bold.ttf")),
             ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
             ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"),
             (r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\arialbd.ttf"),
             ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf")]
    for reg, bold in cands:
        if os.path.exists(reg):
            pdfmetrics.registerFont(TTFont("AppFont", reg))
            pdfmetrics.registerFont(TTFont("AppFont-Bold", bold if os.path.exists(bold) else reg))
            registerFontFamily("AppFont", normal="AppFont", bold="AppFont-Bold", italic="AppFont", boldItalic="AppFont-Bold")
            _FONT = ("AppFont", "AppFont-Bold")
            return _FONT
    raise RuntimeError("Не найден шрифт с кириллицей для PDF. Поместите DejaVuSans.ttf и DejaVuSans-Bold.ttf в папку fonts/")


def export_xlsx(rep):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = re.sub(r"[\\/*?:\[\]]", " ", rep["title"])[:31]
    n = len(rep["headers"])
    ws.cell(1, 1, rep["title"]).font = Font(bold=True, size=14)
    ws.cell(2, 1, f"Фильтры: {rep['subtitle']}").font = Font(italic=True, color="555555")
    ws.cell(3, 1, f"Сформировано: {now()[:16]}" + (f", {g.user['full_name']}" if g.get("user") else "")).font = Font(color="555555")
    hr = 5
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c, h in enumerate(rep["headers"], 1):
        cell = ws.cell(hr, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="5B1A2B")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border

    def put(r, vals, bold=False):
        for c, (v, k) in enumerate(zip(vals, rep["kinds"]), 1):
            cell = ws.cell(r, c, v)
            if isinstance(v, str) and v.startswith(("=", "+", "-", "@")):
                cell.data_type = "s"  # текст пользователя не должен выполняться как формула
            cell.border = border
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right" if k in ("int", "money") else "left")
            if k == "int":
                cell.number_format = "#,##0"
            elif k == "money":
                cell.number_format = "#,##0.00"
            if bold:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="F1E4E8")
    for i, row in enumerate(rep["rows"]):
        put(hr + 1 + i, row)
    last = hr + len(rep["rows"])
    if rep.get("totals"):
        put(last + 1, rep["totals"], True)
    for c, wd in enumerate(rep["widths"], 1):
        ws.column_dimensions[get_column_letter(c)].width = wd + 2
    ws.freeze_panes = ws.cell(hr + 1, 1)
    if rep["rows"]:
        ws.auto_filter.ref = f"A{hr}:{get_column_letter(n)}{last}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def export_pdf(rep):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    reg, bold = pdf_fonts()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=10 * mm, rightMargin=10 * mm,
                            topMargin=10 * mm, bottomMargin=14 * mm, title=rep["title"], author="Система учета АГФ")
    st = ParagraphStyle("c", fontName=reg, fontSize=7.5, leading=9.5)
    stn = ParagraphStyle("n", parent=st, alignment=2)
    sth = ParagraphStyle("h", parent=st, fontName=bold, textColor=colors.white, alignment=1)
    stb = ParagraphStyle("b", parent=st, fontName=bold)
    stbn = ParagraphStyle("bn", parent=stb, alignment=2)

    def cell(v, kind, head=False, tot=False):
        s = xml_escape(str(v)).replace("\n", "<br/>")
        if head:
            return Paragraph(s, sth)
        num = kind in ("int", "money")
        return Paragraph(s, (stbn if num else stb) if tot else (stn if num else st))

    width = doc.width
    wsum = sum(rep["widths"])
    colw = [width * w / wsum for w in rep["widths"]]
    data = [[cell(h, "text", head=True) for h in rep["headers"]]]
    data += [[cell(v, k) for v, k in zip(r, rep["kinds"])] for r in rep["disp"]]
    if rep.get("totals_disp"):
        data.append([cell(v, k, tot=True) for v, k in zip(rep["totals_disp"], rep["kinds"])])
    style = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#5B1A2B")),
             ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BBBBBB")), ("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F2F3")]),
             ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    if rep.get("totals_disp"):
        style.append(("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#F1E4E8")))
    tbl = Table(data, colWidths=colw, repeatRows=1)
    tbl.setStyle(TableStyle(style))
    who = g.user["full_name"] if g.get("user") else ""
    head = [Paragraph(xml_escape(rep["title"]), ParagraphStyle("t", fontName=bold, fontSize=14, leading=18)),
            Paragraph(xml_escape(f"Фильтры: {rep['subtitle']}"), ParagraphStyle("s", parent=st, fontSize=8.5, textColor=colors.HexColor("#555555"))),
            Spacer(1, 4 * mm)]

    def footer(canv, d):
        canv.saveState()
        canv.setFont(reg, 7)
        canv.setFillColor(colors.HexColor("#666666"))
        canv.drawString(10 * mm, 7 * mm, f"Астраханская государственная филармония · сформировано {now()[:16]} {who}")
        canv.drawRightString(landscape(A4)[0] - 10 * mm, 7 * mm, f"Страница {d.page}")
        canv.restoreState()
    doc.build(head + [tbl], onFirstPage=footer, onLaterPages=footer)
    buf.seek(0)
    return buf


def export_response(rep, fmt, key):
    try:
        buf = export_xlsx(rep) if fmt == "xlsx" else export_pdf(rep)
    except RuntimeError as e:
        flash(str(e), "danger")
        return None
    audit("Экспорт отчёта", "reports", key, f"{rep['title']} → {fmt.upper()}; фильтры: {rep['subtitle']}")
    db().commit()
    mime = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if fmt == "xlsx" else "application/pdf")
    return send_file(buf, mimetype=mime, as_attachment=True, download_name=f"report_{key}_{today()}.{fmt}")


@app.route("/reports")
@need("reports")
def reports():
    return render_template("reports.html", reports=REPORTS)


@app.route("/reports/<key>")
@need("reports")
def report(key):
    if key not in REPORTS:
        abort(404)
    rep = REPORTS[key]["fn"](request.args)
    fmt = request.args.get("fmt")
    if fmt in ("xlsx", "pdf"):
        resp = export_response(rep, fmt, key)
        if resp:
            return resp
    qs = {k: v for k, v in request.args.items() if k != "fmt"}
    return render_template("report.html", rep=rep, key=key, qs=qs, reports=REPORTS)


@app.route("/inventory/<int:inv_id>/export")
@need("reports")
def inventory_export(inv_id):
    fmt = request.args.get("fmt", "xlsx")
    if fmt not in ("xlsx", "pdf"):
        abort(400)
    resp = export_response(rep_inventory(inv_id), fmt, f"inventory_{inv_id}")
    return resp or redirect(url_for("inventory_view", inv_id=inv_id))


# ---------------------------------------------------------------- администрирование: пользователи
LOGIN_RE = re.compile(r"[A-Za-z0-9_.\-]{3,32}")


def other_admins(uid):
    return q("SELECT COUNT(*) c FROM users WHERE role='admin' AND active=1 AND id<>?", (uid,), one=True)["c"]


@app.route("/admin/users", methods=["GET", "POST"])
@need("admin")
def admin_users():
    if request.method == "POST":
        f, err = request.form, []
        lg, pw, role = f.get("login", "").strip(), f.get("password", ""), f.get("role", "")
        emp = to_int(f.get("employee_id")) if f.get("employee_id") else None
        if not LOGIN_RE.fullmatch(lg):
            err.append("Логин: 3–32 символа, латиница, цифры и . _ -")
        if not f.get("full_name", "").strip():
            err.append("Укажите ФИО")
        if role not in ROLES:
            err.append("Выберите роль")
        pe = validate_password(pw)
        if pe:
            err.append(pe)
        if emp and not q("SELECT 1 FROM employees WHERE id=?", (emp,), one=True):
            err.append("Сотрудник не найден")
        if not err:
            try:
                cur = x("INSERT INTO users(login,pw_hash,full_name,role,employee_id,created) VALUES(?,?,?,?,?,?)",
                        (lg, generate_password_hash(pw), f["full_name"].strip(), role, emp, now()))
                audit("Создан пользователь", "users", cur.lastrowid, f"{lg}, роль: {ROLES[role]}")
                db().commit()
                flash("Пользователь создан", "success")
                return redirect(url_for("admin_users"))
            except sqlite3.IntegrityError:
                err.append("Пользователь с таким логином уже существует")
        for e in err:
            flash(e, "danger")
    rows = q("SELECT u.*, e.full_name AS emp_name FROM users u LEFT JOIN employees e ON e.id=u.employee_id ORDER BY u.id")
    return render_template("admin_users.html", rows=rows, emps=q("SELECT id, full_name FROM employees ORDER BY full_name"),
                           f=request.form if request.method == "POST" else {})


@app.route("/admin/users/<int:uid>", methods=["GET", "POST"])
@need("admin")
def admin_user_edit(uid):
    u = q("SELECT * FROM users WHERE id=?", (uid,), one=True)
    if not u:
        abort(404)
    if request.method == "POST":
        f, err = request.form, []
        role, active = f.get("role", ""), 1 if f.get("active") else 0
        emp = to_int(f.get("employee_id")) if f.get("employee_id") else None
        if role not in ROLES:
            err.append("Выберите роль")
        if not f.get("full_name", "").strip():
            err.append("Укажите ФИО")
        if uid == g.user["id"] and not active:
            err.append("Нельзя отключить собственную учётную запись")
        if (role != "admin" or not active) and u["role"] == "admin" and u["active"] and not other_admins(uid):
            err.append("Должен остаться хотя бы один активный администратор")
        if f.get("password") and validate_password(f["password"]):
            err.append(validate_password(f["password"]))
        if not err:
            x("UPDATE users SET full_name=?, role=?, employee_id=?, active=? WHERE id=?",
              (f["full_name"].strip(), role, emp, active, uid))
            changes = []
            if role != u["role"]:
                changes.append(f"роль: {ROLES[u['role']]} → {ROLES[role]}")
            if active != u["active"]:
                changes.append("доступ " + ("включён" if active else "отключён"))
            if f.get("password"):
                x("UPDATE users SET pw_hash=? WHERE id=?", (generate_password_hash(f["password"]), uid))
                changes.append("сброшен пароль")
            audit("Изменён пользователь", "users", uid, f"{u['login']}: " + ("; ".join(changes) or "данные профиля"))
            db().commit()
            flash("Сохранено", "success")
            return redirect(url_for("admin_users"))
        for e in err:
            flash(e, "danger")
    return render_template("admin_user_edit.html", u=u, f=request.form if request.method == "POST" else u,
                           emps=q("SELECT id, full_name FROM employees ORDER BY full_name"))


# ---------------------------------------------------------------- резервное копирование
BACKUP_RE = re.compile(r"philharmonia_\d{8}_\d{6}(?:_[a-z]+)?\.db")


def make_backup(tag=""):
    """Согласованная копия БД средствами SQLite (безопасна при работающем приложении)."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    name = f"philharmonia_{datetime.now():%Y%m%d_%H%M%S}{tag}.db"
    path = os.path.join(BACKUP_DIR, name)
    src, dst = sqlite3.connect(DB, timeout=15), sqlite3.connect(path)
    try:
        src.backup(dst)
        ok = dst.execute("PRAGMA integrity_check").fetchone()[0]
    finally:
        dst.close()
        src.close()
    if ok != "ok":
        os.remove(path)
        raise RuntimeError("Копия не прошла проверку целостности")
    return name


def prune_backups():
    auto = sorted(f for f in os.listdir(BACKUP_DIR) if BACKUP_RE.fullmatch(f) and f.endswith("_auto.db"))
    for f in auto[:-BACKUP_KEEP_AUTO] if BACKUP_KEEP_AUTO > 0 else []:
        os.remove(os.path.join(BACKUP_DIR, f))


def list_backups():
    out = []
    if os.path.isdir(BACKUP_DIR):
        for f in sorted(os.listdir(BACKUP_DIR), reverse=True):
            if BACKUP_RE.fullmatch(f):
                st = os.stat(os.path.join(BACKUP_DIR, f))
                out.append(dict(name=f, size=st.st_size, ts=datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                                kind="авто" if f.endswith("_auto.db") else "перед восстановлением" if f.endswith("_prerestore.db") else "ручная"))
    return out


def _auto_backup_loop():
    while True:
        try:
            files = [os.path.getmtime(os.path.join(BACKUP_DIR, b["name"])) for b in list_backups()]
            if not files or time.time() - max(files) >= BACKUP_INTERVAL_HOURS * 3600:
                name = make_backup("_auto")
                prune_backups()
                con = sqlite3.connect(DB, timeout=15)
                write_audit(con, None, "system", "Автоматическое резервное копирование", "backup", name)
                con.commit()
                con.close()
        except Exception as e:  # поток не должен падать
            print("Ошибка автоматического резервного копирования:", e)
        time.sleep(600)


def start_background():
    if BACKUP_INTERVAL_HOURS > 0:
        threading.Thread(target=_auto_backup_loop, daemon=True, name="auto-backup").start()


@app.route("/admin/backup")
@need("admin")
def admin_backup():
    return render_template("admin_backup.html", rows=list_backups(), auto_h=BACKUP_INTERVAL_HOURS, keep=BACKUP_KEEP_AUTO)


@app.route("/admin/backup/create", methods=["POST"])
@need("admin")
def backup_create():
    try:
        name = make_backup()
        audit("Создана резервная копия БД", "backup", name)
        db().commit()
        flash(f"Резервная копия создана: {name}", "success")
    except (OSError, sqlite3.Error, RuntimeError) as e:
        flash(f"Не удалось создать копию: {e}", "danger")
    return redirect(url_for("admin_backup"))


def _backup_name():
    name = os.path.basename(request.form.get("name") or request.args.get("name") or "")
    if not BACKUP_RE.fullmatch(name) or not os.path.isfile(os.path.join(BACKUP_DIR, name)):
        abort(404)
    return name


@app.route("/admin/backup/download")
@need("admin")
def backup_download():
    name = _backup_name()
    audit("Скачана резервная копия БД", "backup", name)
    db().commit()
    return send_from_directory(BACKUP_DIR, name, as_attachment=True)


@app.route("/admin/backup/delete", methods=["POST"])
@need("admin")
def backup_delete():
    name = _backup_name()
    os.remove(os.path.join(BACKUP_DIR, name))
    audit("Удалена резервная копия БД", "backup", name)
    db().commit()
    flash("Копия удалена", "info")
    return redirect(url_for("admin_backup"))


@app.route("/admin/backup/restore", methods=["POST"])
@need("admin")
def backup_restore():
    name = _backup_name()
    path = os.path.join(BACKUP_DIR, name)
    try:  # проверяем, что файл — действительно наша база
        chk = sqlite3.connect(path)
        try:
            ok = chk.execute("PRAGMA integrity_check").fetchone()[0] == "ok"
            chk.execute("SELECT COUNT(*) FROM items").fetchone()
        finally:
            chk.close()
        if not ok:
            raise sqlite3.DatabaseError("integrity_check")
    except sqlite3.Error:
        flash("Файл копии повреждён или не является базой системы", "danger")
        return redirect(url_for("admin_backup"))
    uid, login_ = g.user["id"], g.user["login"]
    try:
        pre = make_backup("_prerestore")
        d = g.pop("db", None)
        if d:
            d.close()
        src, dst = sqlite3.connect(path), sqlite3.connect(DB, timeout=15)
        try:
            src.backup(dst)
        finally:
            src.close()
            dst.close()
        init_db()  # докатываем схему, если копия сделана старой версией
        con = sqlite3.connect(DB, timeout=15)
        # журнал действий не откатывается: переносим записи, сделанные после создания восстанавливаемой копии
        con.execute("ATTACH DATABASE ? AS pre", (os.path.join(BACKUP_DIR, pre),))
        con.execute("INSERT INTO audit_log(ts,user_id,login,action,entity,entity_id,details,ip) "
                    "SELECT p.ts,p.user_id,p.login,p.action,p.entity,p.entity_id,p.details,p.ip FROM pre.audit_log p "
                    "WHERE p.ts>=(SELECT COALESCE(MAX(ts),'') FROM main.audit_log) AND NOT EXISTS ("
                    "SELECT 1 FROM main.audit_log m WHERE m.ts=p.ts AND m.action=p.action AND "
                    "IFNULL(m.login,'')=IFNULL(p.login,'') AND IFNULL(m.entity_id,'')=IFNULL(p.entity_id,'') AND "
                    "IFNULL(m.details,'')=IFNULL(p.details,'')) ORDER BY p.id")
        con.commit()
        con.execute("DETACH DATABASE pre")
        write_audit(con, uid, login_, "Восстановление БД из копии", "backup", name,
                    f"предварительная копия текущего состояния: {pre}", request.remote_addr)
        con.commit()
        con.close()
        flash(f"База восстановлена из копии {name}. Состояние до восстановления сохранено: {pre}", "success")
    except (OSError, sqlite3.Error, RuntimeError) as e:
        flash(f"Восстановление не выполнено: {e}", "danger")
    return redirect(url_for("admin_backup"))


# ---------------------------------------------------------------- журнал действий
@app.route("/admin/audit")
@need("admin")
def admin_audit():
    a = request.args
    w, args = [], []
    for key, col in (("login", "login"), ("action", "action")):
        if a.get(key):
            w.append(f"{col}=?")
            args.append(a[key])
    if valid_date(a.get("date_from", "")):
        w.append("substr(ts,1,10)>=?")
        args.append(a["date_from"])
    if valid_date(a.get("date_to", "")):
        w.append("substr(ts,1,10)<=?")
        args.append(a["date_to"])
    if a.get("s"):
        w.append("(details LIKE ? OR entity_id LIKE ?)")
        args += [f"%{a['s']}%"] * 2
    where = " WHERE " + " AND ".join(w) if w else ""
    per, page = 100, max(1, to_int(a.get("page"), 1))
    total = q("SELECT COUNT(*) c FROM audit_log" + where, args, one=True)["c"]
    rows = q("SELECT * FROM audit_log" + where + " ORDER BY id DESC LIMIT ? OFFSET ?", args + [per, (page - 1) * per])
    qs = {k: v for k, v in a.items() if k != "page" and v}
    return render_template("admin_audit.html", rows=rows, a=a, page=page, pages=max(1, -(-total // per)), total=total, qs=qs,
                           logins=[r["login"] for r in q("SELECT DISTINCT login FROM audit_log WHERE login IS NOT NULL ORDER BY login")],
                           actions=[r["action"] for r in q("SELECT DISTINCT action FROM audit_log ORDER BY action")])


# ---------------------------------------------------------------- шаблоны
T = {}
T["macros.html"] = """{% macro state_badge(s) %}<span class="badge bg-{{ 'success' if s=='Исправно' else 'warning text-dark' if 'ремонт' in s else 'dark' if s=='Списано' else 'secondary' }}">{{ s }}</span>{% endmacro %}
{% macro avail_badge(s) %}<span class="badge bg-{{ {'В наличии':'success','Выдан частично':'primary','Выдан':'info text-dark','В ремонте':'warning text-dark','Списан':'dark','Отсутствует':'danger'}.get(s,'secondary') }}">{{ s }}</span>{% endmacro %}
{% macro sel(name, options, current, blank=None) %}<select class="form-select" name="{{ name }}">{% if blank is not none %}<option value="">{{ blank }}</option>{% endif %}{% for o in options %}<option {{ 'selected' if o==current }}>{{ o }}</option>{% endfor %}</select>{% endmacro %}
{% macro due_badge(r) %}{% if r.flag=='overdue' %}<span class="badge bg-danger">просрочено на {{ -r.days }} дн.</span>{% elif r.flag=='soon' %}<span class="badge bg-warning text-dark">{{ 'сегодня' if r.days==0 else 'через ' ~ r.days ~ ' дн.' }}</span>{% elif r.flag=='returned' %}<span class="badge bg-success">возвращено {{ r.returned_at }}</span>{% else %}<span class="badge bg-light text-dark border">в срок</span>{% endif %}{% endmacro %}
{% macro field(c, f) %}{% set col=c[0] %}{% if c[2]=='bool' %}<div class="form-check pt-2"><input type="checkbox" class="form-check-input" name="{{ col }}" value="1" id="f_{{ col }}" {{ 'checked' if f.get(col) }}><label class="form-check-label" for="f_{{ col }}">{{ c[1] }}</label></div>{% else %}<input class="form-control" name="{{ col }}" type="{{ 'date' if c[2]=='date' else 'email' if c[2]=='email' else 'text' }}" value="{{ f.get(col) or '' }}" placeholder="{{ c[1] }}" title="{{ c[1] }}" {{ 'required' if c[3] }}>{% endif %}{% endmacro %}"""

T["base.html"] = """<!doctype html><html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>АГФ — Учет оборудования</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>.navbar{background:#5b1a2b}.card{border-radius:.75rem}.table td,.table th{vertical-align:middle}
.table-responsive{border-radius:.75rem}</style></head><body class="bg-light">
{% if user %}<nav class="navbar navbar-expand-lg navbar-dark mb-4"><div class="container">
<a class="navbar-brand" href="/">🎼 Астраханская филармония</a>
<button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#nv"><span class="navbar-toggler-icon"></span></button>
<div class="collapse navbar-collapse" id="nv"><div class="navbar-nav me-auto">
<a class="nav-link" href="/items">Оборудование</a>
<a class="nav-link" href="/issues">{{ 'Выдачи' if can('operations') else 'Мои выдачи' }}</a>
<a class="nav-link" href="/bookings">Бронирование</a>
{% if can('operations') %}<a class="nav-link" href="/inventory">Инвентаризация</a><a class="nav-link" href="/purchases">Закупки</a>{% endif %}
{% if can('reports') %}<a class="nav-link" href="/reports">Отчёты</a>{% endif %}
{% if can('dictionaries') %}<a class="nav-link" href="/dict/categories">Справочники</a>{% endif %}
{% if can('admin') %}<div class="nav-item dropdown"><a class="nav-link dropdown-toggle" href="#" data-bs-toggle="dropdown">Администрирование</a>
<ul class="dropdown-menu"><li><a class="dropdown-item" href="/admin/users">Пользователи</a></li>
<li><a class="dropdown-item" href="/admin/audit">Журнал действий</a></li>
<li><a class="dropdown-item" href="/admin/backup">Резервные копии</a></li></ul></div>{% endif %}
</div>
<a class="nav-link text-white me-3" href="/notifications" title="Уведомления">🔔{% if notif_count %} <span class="badge bg-danger">{{ notif_count }}</span>{% endif %}</a>
<span class="navbar-text me-3 small"><a href="/profile" class="text-white text-decoration-none">{{ user.full_name }}</a> · {{ ROLES[user.role] }}</span>
<form method="post" action="/logout" class="d-inline">{{ csrf_input() }}<button class="btn btn-outline-light btn-sm">Выйти</button></form>
</div></div></nav>{% endif %}
<div class="container pb-5">{% with m=get_flashed_messages(with_categories=true) %}{% for c,t in m %}
<div class="alert alert-{{c}}">{{t}}</div>{% endfor %}{% endwith %}{% block c %}{% endblock %}</div>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script></body></html>"""

T["error.html"] = """{% extends 'base.html' %}{% block c %}<div class="text-center mt-5"><h1 class="display-4">{{ code }}</h1>
<p class="lead">{{ text }}</p><a class="btn btn-danger" href="/">На главную</a></div>{% endblock %}"""

T["login.html"] = """{% extends 'base.html' %}{% block c %}
<div class="row justify-content-center mt-5"><div class="col-md-4"><div class="card shadow-sm p-4">
<h4 class="mb-3">Вход в систему</h4><form method="post">{{ csrf_input() }}
<input class="form-control mb-2" name="login" placeholder="Логин" required autofocus autocomplete="username">
<input class="form-control mb-3" type="password" name="password" placeholder="Пароль" required autocomplete="current-password">
<button class="btn btn-danger w-100">Войти</button></form></div></div></div>{% endblock %}"""

T["profile.html"] = """{% extends 'base.html' %}{% block c %}
<h3>Профиль</h3><div class="row g-3 mt-1"><div class="col-md-5"><div class="card p-3">
<p class="mb-1">Логин: <b>{{ user.login }}</b></p><p class="mb-1">ФИО: <b>{{ user.full_name }}</b></p>
<p class="mb-1">Роль: <b>{{ ROLES[user.role] }}</b></p>
<p class="mb-0">Сотрудник: <b>{{ emp.full_name if emp else '— не привязан —' }}</b></p></div></div>
<div class="col-md-5"><div class="card p-3"><h6>Смена пароля</h6><form method="post">{{ csrf_input() }}
<input class="form-control mb-2" type="password" name="old" placeholder="Текущий пароль" required autocomplete="current-password">
<input class="form-control mb-2" type="password" name="new" placeholder="Новый пароль (от 6 символов)" required autocomplete="new-password">
<input class="form-control mb-2" type="password" name="new2" placeholder="Повторите новый пароль" required autocomplete="new-password">
<button class="btn btn-danger">Сменить пароль</button></form></div></div></div>{% endblock %}"""

T["dashboard.html"] = """{% extends 'base.html' %}{% from 'macros.html' import due_badge %}{% block c %}
<h3 class="mb-3">Сводка</h3><div class="row g-3 mb-4">
{% set cards=[('Единиц учета',s.total,'primary'),('Исправно',s.ok,'success'),('Ремонт',s.repair,'warning'),('Выдано, шт.',s.issued,'info')] %}
{% if can('operations') %}{% set cards=cards+[('Стоимость фонда, ₽','{:,.0f}'.format(s.value),'secondary')] %}{% endif %}
{% for l,v,c in cards %}<div class="col-6 col-md"><div class="card p-3 border-{{c}}"><div class="text-muted small">{{l}}</div>
<div class="fs-3">{{v}}</div></div></div>{% endfor %}</div>
<div class="row g-3"><div class="col-md-6"><div class="card p-3 border-danger"><h6>🔴 Просроченные возвраты</h6>
{% for r in s.late %}<div class="border-bottom py-1"><a href="/items/{{r.item_id}}">{{r.item_name}}</a> × {{r.qty}} — {{r.emp_name}} {{ due_badge(r) }}</div>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div>
<div class="col-md-6"><div class="card p-3 border-warning"><h6>🟡 Скоро срок возврата ({{REMIND_DAYS}} дн.)</h6>
{% for r in s.soon %}<div class="border-bottom py-1"><a href="/items/{{r.item_id}}">{{r.item_name}}</a> × {{r.qty}} — {{r.emp_name}}, до {{r.due_date}} {{ due_badge(r) }}</div>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div>
{% if can('operations') %}<div class="col-md-4"><div class="card p-3"><h6>⏰ Просрочена проверка</h6>
{% for i in s.overdue %}<a href="/items/{{i.id}}">{{i.name}}</a> <small class="text-danger">до {{i.next_check}}</small><br>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div>
<div class="col-md-4"><div class="card p-3"><h6>📉 Износ ≥ 80%</h6>
{% for i in s.worn %}<a href="/items/{{i.id}}">{{i.name}}</a> <small>{{i.wear}}%</small><br>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div>{% endif %}
<div class="col-md-4"><div class="card p-3"><h6>📅 Ближайшие брони</h6>
{% for b in upcoming %}<small>{{b.start}} — {{b.name}} ({{b.event}})</small><br>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div></div>{% endblock %}"""

T["notifications.html"] = """{% extends 'base.html' %}{% from 'macros.html' import due_badge %}{% block c %}
<div class="d-flex justify-content-between align-items-center mb-3"><h3>Уведомления</h3>
{% if can('operations') %}<form method="post" action="/notifications/send">{{ csrf_input() }}
<button class="btn btn-outline-danger btn-sm" {{ '' if smtp_on else 'title="SMTP не настроен: задайте SMTP_HOST"' }}>✉ Разослать напоминания сотрудникам</button></form>{% endif %}</div>
<div class="card p-3 mb-3 border-danger"><h6>Просроченные возвраты ({{ overdue|length }})</h6>
{% for r in overdue %}<div class="border-bottom py-1"><a href="/items/{{r.item_id}}">{{r.item_name}}</a> ({{r.inv_no}}) × {{r.qty}} — <b>{{r.emp_name}}</b>,
мероприятие: {{r.event_name or '—'}}, срок {{r.due_date}} {{ due_badge(r) }}
{% if can('operations') %}<a class="btn btn-sm btn-outline-secondary ms-2" href="/issues/{{r.id}}/return">Вернуть</a>{% endif %}</div>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div>
<div class="card p-3 mb-3 border-warning"><h6>Срок возврата в ближайшие {{REMIND_DAYS}} дн. ({{ soon|length }})</h6>
{% for r in soon %}<div class="border-bottom py-1"><a href="/items/{{r.item_id}}">{{r.item_name}}</a> × {{r.qty}} — <b>{{r.emp_name}}</b>, до {{r.due_date}} {{ due_badge(r) }}</div>
{% else %}<span class="text-muted">Нет</span>{% endfor %}</div>
{% if can('operations') %}<div class="row g-3"><div class="col-md-6"><div class="card p-3"><h6>Просрочена проверка состояния ({{ checks|length }})</h6>
{% for i in checks %}<a href="/items/{{i.id}}">{{i.name}}</a> <small class="text-danger">до {{i.next_check}}</small><br>{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div>
<div class="col-md-6"><div class="card p-3"><h6>Остаток ниже минимального ({{ low|length }})</h6>
{% for i in low %}<a href="/items/{{i.id}}">{{i.name}}</a> <small>{{i.qty}} из {{i.min_qty}}</small><br>{% else %}<span class="text-muted">Нет</span>{% endfor %}</div></div></div>{% endif %}
{% endblock %}"""

T["items.html"] = """{% extends 'base.html' %}{% from 'macros.html' import state_badge, avail_badge, sel %}{% block c %}
<div class="d-flex justify-content-between align-items-center"><h3>Учет оборудования и реквизита</h3>
<div>{% if can('reports') %}<a class="btn btn-outline-secondary btn-sm" href="{{ url_for('report', key='stock', **a.to_dict()) }}">Отчёт по выборке</a>{% endif %}
{% if can('edit_items') %}<a class="btn btn-danger btn-sm" href="/items/new">+ Новая карточка</a>{% endif %}</div></div>
<form class="row g-2 my-3"><div class="col-md-3"><input class="form-control" name="s" value="{{ a.get('s','') }}" placeholder="Название / инв. № / серийный №"></div>
<div class="col-md-2">{{ sel('cat', cats, a.get('cat'), 'Все категории') }}</div>
<div class="col-md-2">{{ sel('loc', locs, a.get('loc'), 'Все места') }}</div>
<div class="col-md-2">{{ sel('status', avail, a.get('status'), 'Любой статус') }}</div>
<div class="col-md-2">{{ sel('state', states, a.get('state'), 'Любое состояние') }}</div>
<div class="col-md-1"><button class="btn btn-secondary w-100">Найти</button></div></form>
<div class="card mb-3 table-responsive"><table class="table table-hover mb-0 align-middle"><thead><tr><th>Инв. №</th><th>Название</th>
<th>Категория</th><th>Место</th><th>Кол-во</th><th>Состояние</th><th>Статус</th><th>Износ</th></tr></thead><tbody>
{% for i in items %}<tr><td>{{i.inv_no}}</td><td><a href="/items/{{i.id}}">{{i.name}}</a>{% if i.m3d %} <span class="badge bg-secondary" title="В карточке есть 3D-модель">3D</span>{% endif %}</td><td>{{i.category}}</td><td>{{i.location}}</td>
<td>{{i.free}} / {{i.qty}}</td><td>{{ state_badge(i.status) }}{% if i.overdue %} <span class="badge bg-danger">проверка</span>{% endif %}</td>
<td>{{ avail_badge(i.avail) }}</td>
<td style="width:130px"><div class="progress"><div class="progress-bar {{'bg-danger' if i.wear>=80 else ''}}" style="width:{{i.wear}}%">{{i.wear}}%</div></div></td></tr>
{% else %}<tr><td colspan="8" class="text-muted">Ничего не найдено</td></tr>{% endfor %}</tbody></table></div>
<small class="text-muted">Найдено: {{ items|length }}. «Кол-во»: доступно / всего.</small>{% endblock %}"""

T["item_form.html"] = """{% extends 'base.html' %}{% from 'macros.html' import sel %}{% block c %}
<a href="{{ '/items/%d'|format(it.id) if it else '/items' }}">← Назад</a>
<h3 class="mt-2">{{ 'Редактирование: ' ~ it.name if it else 'Постановка на учет' }}</h3>
<div class="card p-3 mt-3"><form method="post" class="row g-3">{{ csrf_input() }}
<div class="col-md-3"><label class="form-label">Инвентарный номер</label>
<input class="form-control" name="inv_no" value="{{ f.get('inv_no') or '' }}" {{ 'readonly' if it and not can('admin') }}
placeholder="{{ 'Авто: ' ~ next_inv if next_inv else '' }}"><div class="form-text">{{ 'Оставьте пустым — присвоится автоматически' if not it else ('Меняет только администратор' if not can('admin') else '') }}</div></div>
<div class="col-md-9"><label class="form-label">Название *</label><input class="form-control" name="name" value="{{ f.get('name') or '' }}" required></div>
<div class="col-md-4"><label class="form-label">Категория *</label>{{ sel('category', cats, f.get('category')) }}</div>
<div class="col-md-4"><label class="form-label">Место хранения *</label>{{ sel('location', locs, f.get('location')) }}</div>
{% if not it %}<div class="col-md-4"><label class="form-label">Состояние</label>{{ sel('state', states, f.get('state')) }}</div>
{% else %}<div class="col-md-4"><label class="form-label">Серийный №</label><input class="form-control" name="serial" value="{{ f.get('serial') or '' }}"></div>{% endif %}
<div class="col-md-2"><label class="form-label">Количество</label><input class="form-control" type="number" min="0" name="qty" value="{{ f.get('qty') }}"></div>
<div class="col-md-2"><label class="form-label">Мин. запас</label><input class="form-control" type="number" min="0" name="min_qty" value="{{ f.get('min_qty') }}"></div>
<div class="col-md-2"><label class="form-label">Цена, ₽</label><input class="form-control" type="number" step="0.01" min="0" name="price" value="{{ f.get('price') }}"></div>
<div class="col-md-3"><label class="form-label">Дата покупки</label><input class="form-control" type="date" name="bought" value="{{ f.get('bought') or '' }}"></div>
<div class="col-md-3"><label class="form-label">Срок службы, лет</label><input class="form-control" type="number" step="0.5" min="0.5" name="life_years" value="{{ f.get('life_years') }}"></div>
{% if not it %}<div class="col-md-6"><label class="form-label">Серийный №</label><input class="form-control" name="serial" value="{{ f.get('serial') or '' }}"></div>{% endif %}
<div class="col-12"><label class="form-label">Описание</label><textarea class="form-control" name="descr" rows="2">{{ f.get('descr') or '' }}</textarea></div>
<div class="col-12"><button class="btn btn-danger">Сохранить</button></div></form></div>{% endblock %}"""

T["item.html"] = """{% extends 'base.html' %}{% from 'macros.html' import state_badge, avail_badge, sel, due_badge %}{% block c %}
<a href="/items">← К списку</a><h3 class="mt-2">{{it.name}} <small class="text-muted">{{it.inv_no}}</small></h3>
<div class="row g-3"><div class="col-md-5">
{% if it.m3d %}<div class="card p-2 mb-3"><div id="m3d" data-kind="{{ it.m3d }}" style="height:340px"></div>
<div class="small text-muted px-1 pt-2">3D-модель: {{ it.m3d_title }}. Обобщённая модель типа инструмента, не точная копия экземпляра.</div></div>
<script src="/assets/instrument3d.js"></script><script>(function(){var h=document.getElementById('m3d');Instrument3D.mount(h,h.dataset.kind);})();</script>{% endif %}
<div class="card p-3">
<p class="mb-1">Статус: {{ avail_badge(it.avail) }} · Состояние: {{ state_badge(it.status) }}</p>
<p class="mb-1">Категория: <b>{{it.category}}</b></p><p class="mb-1">Место: <b>{{it.location}}</b></p>
<p class="mb-1">Количество: <b>{{it.qty}}</b>, выдано <b>{{it.issued}}</b>, доступно <b>{{it.free}}</b> (мин. запас {{it.min_qty}})</p>
{% if it.serial %}<p class="mb-1">Серийный №: <b>{{it.serial}}</b></p>{% endif %}
<p class="mb-1">Куплено: <b>{{it.bought}}</b>, {{'{:,.0f}'.format(it.price)}} ₽</p>
<p class="mb-1">Срок службы: <b>{{it.life_years}}</b> лет, износ <b>{{it.wear}}%</b></p>
<p class="mb-1">Этап цикла: <b>{{it.stage}}</b></p>
<p class="mb-1">След. проверка: <b class="{{'text-danger' if it.overdue}}">{{it.next_check}}</b></p>
{% if it.descr %}<p class="mb-1 text-muted">{{it.descr}}</p>{% endif %}
<div class="mt-2">{% if can('edit_items') %}<a class="btn btn-sm btn-outline-secondary" href="/items/{{it.id}}/edit">✎ Редактировать</a> {% endif %}
{% if can('operations') and it.status=='Исправно' and it.free>0 %}<a class="btn btn-sm btn-danger" href="/issues/new?item_id={{it.id}}">Выдать</a>{% endif %}</div></div>
{% if can('edit_items') %}<div class="card p-3 mt-3"><h6>Акт проверки состояния</h6><form method="post" action="/items/{{it.id}}/check">{{ csrf_input() }}
{{ sel('status', states, it.status) }}<input class="form-control my-2" name="note" placeholder="Комментарий / причина" required>
<button class="btn btn-danger">Сохранить</button></form></div>{% endif %}</div>
<div class="col-md-7"><ul class="nav nav-tabs">
{% if can('operations') %}<li class="nav-item"><button class="nav-link active" data-bs-toggle="tab" data-bs-target="#t-hist">История движения</button></li>{% endif %}
<li class="nav-item"><button class="nav-link {{ '' if can('operations') else 'active' }}" data-bs-toggle="tab" data-bs-target="#t-iss">Выдачи</button></li>
<li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#t-chk">Состояние</button></li>
<li class="nav-item"><button class="nav-link" data-bs-toggle="tab" data-bs-target="#t-bk">Брони</button></li></ul>
<div class="tab-content card p-3 border-top-0 rounded-top-0">
{% if can('operations') %}<div class="tab-pane fade show active" id="t-hist"><div class="table-responsive"><table class="table table-sm mb-0"><thead><tr><th>Дата</th><th>Операция</th><th>Кол-во</th><th>Сотрудник / мероприятие</th><th>Выполнил</th><th>Детали</th></tr></thead>
{% for m in moves %}<tr><td class="text-nowrap">{{m.ts[:16]}}</td><td>{{m.kind}}</td><td>{{ m.qty if m.qty is not none else '' }}</td>
<td>{{m.emp_name or ''}}{% if m.event_name %}<br><small class="text-muted">{{m.event_name}}</small>{% endif %}</td><td>{{m.user_name or ''}}</td><td><small>{{m.details or ''}}</small></td></tr>
{% else %}<tr><td colspan="6" class="text-muted">Записей нет</td></tr>{% endfor %}</table></div></div>{% endif %}
<div class="tab-pane fade {{ '' if can('operations') else 'show active' }}" id="t-iss">
{% for r in issues %}<div class="border-bottom py-1">{{r.issued_at}} → {{r.emp_name}} × {{r.qty}}, {{r.event_name or '—'}}; срок {{r.due_date}} {{ due_badge(r) }}
{% if r.returned_at %}<small class="text-muted">({{r.return_state}}{{ ': ' ~ r.return_note if r.return_note }})</small>{% elif can('operations') %}<a class="btn btn-sm btn-outline-secondary ms-1" href="/issues/{{r.id}}/return">Вернуть</a>{% endif %}</div>
{% else %}<span class="text-muted">Выдач нет</span>{% endfor %}</div>
<div class="tab-pane fade" id="t-chk">{% for c in checks %}<div class="border-bottom py-1"><b>{{c.day}}</b> — {{c.status}}: {{c.note}}</div>
{% else %}<span class="text-muted">Записей нет</span>{% endfor %}</div>
<div class="tab-pane fade" id="t-bk">{% for b in bookings %}<div class="border-bottom py-1">{{b.start}} … {{b.end}} · {{b.event}} · {{b.qty}} шт. · {{b.person}}</div>
{% else %}<span class="text-muted">Броней нет</span>{% endfor %}</div></div></div></div>{% endblock %}"""

T["issues.html"] = """{% extends 'base.html' %}{% from 'macros.html' import due_badge %}{% block c %}
<div class="d-flex justify-content-between align-items-center"><h3>{{ 'Выдача и возврат' if can('operations') else 'Мои выдачи' }}</h3>
{% if can('operations') %}<a class="btn btn-danger btn-sm" href="/issues/new">+ Выдать</a>{% endif %}</div>
{% if not can('operations') and not user.employee_id %}<div class="alert alert-info mt-3">Ваша учётная запись не привязана к сотруднику — обратитесь к администратору.</div>{% endif %}
<ul class="nav nav-pills my-3">{% for k,l in [('open','На руках'),('overdue','Просроченные'),('closed','Возвращённые'),('all','Все')] %}
<li class="nav-item"><a class="nav-link {{ 'active bg-danger' if tab==k }}" href="?tab={{k}}">{{l}}</a></li>{% endfor %}</ul>
{% if can('operations') %}<form class="row g-2 mb-3"><input type="hidden" name="tab" value="{{tab}}">
<div class="col-md-4"><input class="form-control" name="s" value="{{ a.get('s','') }}" placeholder="Объект / инв. № / сотрудник"></div>
<div class="col-md-3"><select class="form-select" name="emp"><option value="">Все сотрудники</option>{% for e in emps %}<option value="{{e.id}}" {{ 'selected' if a.get('emp')==e.id|string }}>{{e.full_name}}</option>{% endfor %}</select></div>
<div class="col-md-3"><select class="form-select" name="event"><option value="">Все мероприятия</option>{% for e in events %}<option value="{{e.id}}" {{ 'selected' if a.get('event')==e.id|string }}>{{e.name}}</option>{% endfor %}</select></div>
<div class="col-md-2"><button class="btn btn-secondary w-100">Найти</button></div></form>{% endif %}
<div class="card table-responsive"><table class="table mb-0"><thead><tr><th>Объект</th><th>Кол-во</th><th>Сотрудник</th><th>Мероприятие</th><th>Выдано</th><th>Срок</th><th>Состояние</th><th></th></tr></thead><tbody>
{% for r in rows %}<tr><td><a href="/items/{{r.item_id}}">{{r.item_name}}</a><br><small class="text-muted">{{r.inv_no}}</small></td><td>{{r.qty}}</td><td>{{r.emp_name}}</td><td>{{r.event_name or '—'}}</td>
<td>{{r.issued_at}}<br><small class="text-muted">{{r.issued_by_name or ''}}</small></td><td>{{r.due_date}}<br>{{ due_badge(r) }}</td>
<td>{% if r.returned_at %}{{r.return_state}}{% if r.return_note %}<br><small class="text-muted">{{r.return_note}}</small>{% endif %}{% else %}—{% endif %}</td>
<td>{% if not r.returned_at and can('operations') %}<a class="btn btn-sm btn-outline-danger" href="/issues/{{r.id}}/return">Принять</a>{% endif %}</td></tr>
{% else %}<tr><td colspan="8" class="text-muted">Записей нет</td></tr>{% endfor %}</tbody></table></div>{% endblock %}"""

T["issue_new.html"] = """{% extends 'base.html' %}{% block c %}<a href="/issues">← К выдачам</a><h3 class="mt-2">Выдача имущества</h3>
<div class="card p-3 mt-3"><form method="post" class="row g-3">{{ csrf_input() }}
<div class="col-md-6"><label class="form-label">Объект *</label><select class="form-select" name="item_id" required><option value="">— выберите —</option>
{% for i in avail %}<option value="{{i.id}}" {{ 'selected' if f.get('item_id')==i.id|string }}>{{i.name}} ({{i.inv_no}}) — свободно {{i.free}}</option>{% endfor %}</select></div>
<div class="col-md-2"><label class="form-label">Количество *</label><input class="form-control" type="number" min="1" name="qty" value="{{ f.get('qty') or 1 }}" required></div>
<div class="col-md-4"><label class="form-label">Вернуть до *</label><input class="form-control" type="date" name="due_date" min="{{today}}" value="{{ f.get('due_date') }}" required></div>
<div class="col-md-6"><label class="form-label">Сотрудник (ответственный) *</label><select class="form-select" name="employee_id" required><option value="">— выберите —</option>
{% for e in emps %}<option value="{{e.id}}" {{ 'selected' if f.get('employee_id')==e.id|string }}>{{e.full_name}}{{ ' — ' ~ e.position if e.position }}</option>{% endfor %}</select></div>
<div class="col-md-6"><label class="form-label">Мероприятие *</label><select class="form-select" name="event_id" required><option value="">— выберите —</option>
{% for e in events %}<option value="{{e.id}}" {{ 'selected' if f.get('event_id')==e.id|string }}>{{e.name}}{{ ' (' ~ e.event_date ~ ')' if e.event_date }}</option>{% endfor %}</select>
<div class="form-text"><a href="/dict/events">Добавить мероприятие</a></div></div>
<div class="col-12"><label class="form-label">Комментарий</label><input class="form-control" name="note" value="{{ f.get('note') or '' }}"></div>
<div class="col-12"><button class="btn btn-danger">Оформить выдачу</button></div></form></div>{% endblock %}"""

T["issue_return.html"] = """{% extends 'base.html' %}{% from 'macros.html' import due_badge %}{% block c %}<a href="/issues">← К выдачам</a><h3 class="mt-2">Приём имущества</h3>
<div class="card p-3 mt-3 mb-3"><b>{{s.item_name}}</b> ({{s.inv_no}}) × {{s.qty}}<br>Сотрудник: {{s.emp_name}} · мероприятие: {{s.event_name or '—'}}<br>
Выдано {{s.issued_at}}, вернуть до {{s.due_date}} {{ due_badge(s) }}</div>
<div class="card p-3"><form method="post" class="row g-3">{{ csrf_input() }}
<div class="col-md-3"><label class="form-label">Возвращается, шт.</label><input class="form-control" type="number" min="1" max="{{s.qty}}" name="qty" value="{{ f.get('qty') or s.qty }}" required>
<div class="form-text">Меньше {{s.qty}} — частичный возврат</div></div>
<div class="col-md-4"><label class="form-label">Состояние при возврате *</label><select class="form-select" name="state">
{% for st in states %}<option {{ 'selected' if (f.get('state') or 'Исправно')==st }}>{{st}}</option>{% endfor %}</select></div>
<div class="col-12"><label class="form-label">Комментарий (обязателен, если не «Исправно»)</label><input class="form-control" name="note" value="{{ f.get('note') or '' }}"></div>
<div class="col-12"><button class="btn btn-danger">Принять</button></div></form></div>{% endblock %}"""

T["bookings.html"] = """{% extends 'base.html' %}{% block c %}
<h3>Бронирование оборудования</h3>{% if can('operations') %}<div class="card p-3 my-3"><form method="post" class="row g-2">{{ csrf_input() }}
<div class="col-md-4"><select class="form-select" name="item_id">{% for i in items %}
<option value="{{i.id}}">{{i.name}} ({{i.inv_no}}, всего {{i.qty}})</option>{% endfor %}</select></div>
<div class="col-md-1"><input class="form-control" type="number" min="1" name="qty" value="1" title="Количество"></div>
<div class="col-md-4"><select class="form-select" name="event_id" title="Мероприятие">{% for e in events %}<option value="{{e.id}}">{{e.name}}</option>{% endfor %}</select></div>
<div class="col-md-3"><select class="form-select" name="employee_id" title="Ответственный">{% for e in emps %}<option value="{{e.id}}">{{e.full_name}}</option>{% endfor %}</select></div>
<div class="col-md-3"><input class="form-control" type="date" name="start" value="{{today}}" required></div>
<div class="col-md-3"><input class="form-control" type="date" name="end" value="{{today}}" required></div>
<div class="col-md-2"><button class="btn btn-danger w-100">Забронировать</button></div></form></div>{% endif %}
<div class="card mt-3 table-responsive"><table class="table mb-0 align-middle"><thead><tr><th>Оборудование</th><th>Кол-во</th><th>Мероприятие</th>
<th>Период</th><th>Ответственный</th><th></th></tr></thead><tbody>{% for b in rows %}
<tr><td>{{b.name}}</td><td>{{b.qty}}</td><td>{{b.event}}</td><td>{{b.start}} — {{b.end}}</td><td>{{b.person}}</td>
<td class="text-nowrap">{% if can('operations') %}<a class="btn btn-sm btn-outline-secondary" href="/issues/new?item_id={{b.item_id}}&qty={{b.qty}}{{ '&event_id=%d'|format(b.event_id) if b.event_id }}{{ '&employee_id=%d'|format(b.employee_id) if b.employee_id }}">Выдать</a>
<form method="post" action="/bookings/{{b.id}}/delete" class="d-inline">{{ csrf_input() }}<button class="btn btn-sm btn-outline-danger">Отменить</button></form>{% endif %}</td></tr>
{% else %}<tr><td colspan="6" class="text-muted">Броней нет</td></tr>{% endfor %}</tbody></table></div>{% endblock %}"""

T["purchases.html"] = """{% extends 'base.html' %}{% block c %}
<div class="d-flex justify-content-between align-items-center mb-3"><h3>План закупок</h3>
<form method="post">{{ csrf_input() }}<input type="hidden" name="action" value="auto">
<button class="btn btn-danger">⚙ Сформировать автоматически</button></form></div>
<p class="text-muted">Система предлагает закупки по критериям: списание, исчерпан срок службы, остаток ниже минимума, нецелесообразный ремонт.</p>
<div class="card table-responsive"><table class="table mb-0 align-middle"><thead><tr><th>Дата</th><th>Позиция</th><th>Кол-во</th><th>Основание</th><th>Статус</th></tr></thead><tbody>
{% for p in rows %}<tr><td>{{p.created}}</td><td>{{p.name}}</td><td>{{p.qty}}</td><td>{{p.reason}}</td><td>
<form method="post" class="d-flex gap-1">{{ csrf_input() }}<input type="hidden" name="action" value="status"><input type="hidden" name="id" value="{{p.id}}">
<select class="form-select form-select-sm" name="status" {{ 'disabled' if p.status=='Получена' }}>{% for s in states %}<option {{'selected' if s==p.status}}>{{s}}</option>{% endfor %}</select>
{% if p.status!='Получена' %}<button class="btn btn-sm btn-outline-secondary">OK</button>{% endif %}</form></td></tr>
{% else %}<tr><td colspan="5" class="text-muted">План пуст — нажмите «Сформировать»</td></tr>{% endfor %}</tbody></table></div>{% endblock %}"""

T["dict.html"] = """{% extends 'base.html' %}{% from 'macros.html' import field %}{% block c %}
<h3 class="mb-3">Справочники</h3><ul class="nav nav-pills mb-3">{% for k,v in dicts.items() %}
<li class="nav-item"><a class="nav-link {{ 'active bg-danger' if k==name }}" href="/dict/{{k}}">{{v.title}}</a></li>{% endfor %}</ul>
<div class="card p-3 mb-3"><h6>Добавить: {{d.title|lower}}</h6><form method="post" class="row g-2 align-items-center">{{ csrf_input() }}
{% for c in d.cols %}<div class="col-md">{{ field(c, f) }}</div>{% endfor %}<div class="col-md-auto"><button class="btn btn-danger">Добавить</button></div></form></div>
<div class="card table-responsive"><table class="table mb-0 align-middle"><thead><tr>{% for c in d.cols %}<th>{{c[1]}}</th>{% endfor %}<th>Используется</th><th></th></tr></thead><tbody>
{% for r in rows %}<tr>{% for c in d.cols %}<td>{% if c[2]=='bool' %}{{ 'Да' if r[c[0]] else 'Нет' }}{% else %}{{ r[c[0]] or '' }}{% endif %}</td>{% endfor %}
<td>{{ r.used }}{% if r.locked %} <span class="badge bg-secondary">системная</span>{% endif %}</td>
<td class="text-nowrap"><a class="btn btn-sm btn-outline-secondary" href="/dict/{{name}}/{{r.id}}/edit">✎</a>
<form method="post" action="/dict/{{name}}/{{r.id}}/delete" class="d-inline" onsubmit="return confirm('Удалить запись?')">{{ csrf_input() }}
<button class="btn btn-sm btn-outline-danger" {{ 'disabled' if r.used or r.locked }}>✕</button></form></td></tr>
{% else %}<tr><td colspan="9" class="text-muted">Записей нет</td></tr>{% endfor %}</tbody></table></div>
{% if name in ('categories','states','locations') %}<small class="text-muted">При переименовании значение обновляется во всех карточках и актах.</small>{% endif %}{% endblock %}"""

T["dict_edit.html"] = """{% extends 'base.html' %}{% from 'macros.html' import field %}{% block c %}
<a href="/dict/{{name}}">← {{d.title}}</a><h3 class="mt-2">Редактирование записи</h3>
{% if locked %}<div class="alert alert-warning">Системная запись: на неё опирается логика приложения, изменение запрещено.</div>{% endif %}
<div class="card p-3 mt-3"><form method="post" class="row g-3">{{ csrf_input() }}
{% for c in d.cols %}<div class="col-md-6"><label class="form-label">{{c[1]}}</label>{{ field(c, f) }}</div>{% endfor %}
<div class="col-12"><button class="btn btn-danger" {{ 'disabled' if locked }}>Сохранить</button></div></form></div>{% endblock %}"""

T["inventory.html"] = """{% extends 'base.html' %}{% from 'macros.html' import sel %}{% block c %}
<h3>Инвентаризация</h3><div class="card p-3 my-3"><h6>Начать инвентаризацию</h6><form method="post" class="row g-2">{{ csrf_input() }}
<div class="col-md-4"><input class="form-control" name="title" placeholder="Название (по умолчанию — «Инвентаризация от …»)"></div>
<div class="col-md-3">{{ sel('location', locs, '', 'Все места хранения') }}</div><div class="col-md-3">{{ sel('category', cats, '', 'Все категории') }}</div>
<div class="col-md-2"><button class="btn btn-danger w-100">Начать</button></div></form>
<div class="form-text">В опись попадает всё имущество области, кроме списанного. «По учёту» = количество за вычетом выданного сотрудникам.</div></div>
<div class="card table-responsive"><table class="table mb-0"><thead><tr><th>Название</th><th>Область</th><th>Начата</th><th>Заполнено</th><th>Расхождений</th><th>Статус</th></tr></thead><tbody>
{% for r in rows %}<tr><td><a href="/inventory/{{r.id}}">{{r.title}}</a></td><td>{{r.scope}}</td><td>{{r.created[:16]}}</td><td>{{r.done}} / {{r.n}}</td>
<td>{{ r.diff if r.done else '—' }}</td><td><span class="badge bg-{{ 'success' if r.status=='Завершена' else 'warning text-dark' }}">{{r.status}}</span></td></tr>
{% else %}<tr><td colspan="6" class="text-muted">Инвентаризаций не было</td></tr>{% endfor %}</tbody></table></div>{% endblock %}"""

T["inventory_view.html"] = """{% extends 'base.html' %}{% block c %}
<a href="/inventory">← Инвентаризации</a>
<div class="d-flex justify-content-between align-items-center mt-2"><h3>{{inv.title}}</h3>
<div>{% if can('reports') %}<a class="btn btn-sm btn-outline-secondary" href="/inventory/{{inv.id}}/export?fmt=xlsx">Excel</a> <a class="btn btn-sm btn-outline-secondary" href="/inventory/{{inv.id}}/export?fmt=pdf">PDF</a>{% endif %}</div></div>
<p class="text-muted">Область: {{inv.scope}} · статус: <b>{{inv.status}}</b>{% if inv.finished %} · завершена {{inv.finished[:16]}}{{ ', корректировки применены' if inv.applied else ', без корректировки учёта' }}{% endif %}</p>
{% set open = inv.status=='Идёт' %}
<form method="post">{{ csrf_input() }}<div class="card table-responsive"><table class="table table-sm mb-0 align-middle"><thead><tr><th>Инв. №</th><th>Название</th><th>Место</th>
<th>По учёту</th><th>Выдано</th><th style="width:120px">Фактически</th><th>Расхождение</th><th>Примечание</th></tr></thead><tbody>
{% for l in lines %}{% set d = (l.actual - l.expected) if l.actual is not none else none %}
<tr class="{{ 'table-danger' if d and d<0 else 'table-warning' if d and d>0 }}"><td>{{l.inv_no}}</td><td>{{l.name}}</td><td>{{l.location}}</td><td>{{l.expected}}</td><td>{{l.issued}}</td>
<td>{% if open %}<input class="form-control form-control-sm" type="number" min="0" name="a{{l.id}}" value="{{ l.actual if l.actual is not none else '' }}">{% else %}{{ l.actual }}{% endif %}</td>
<td>{{ '%+d'|format(d) if d is not none else '' }}</td>
<td>{% if open %}<input class="form-control form-control-sm" name="n{{l.id}}" value="{{ l.note or '' }}">{% else %}{{ l.note or '' }}{% endif %}</td></tr>{% endfor %}</tbody></table></div>
{% if open %}<div class="mt-3 d-flex gap-2 flex-wrap"><button class="btn btn-secondary" name="action" value="save">Сохранить ход работы</button>
<button class="btn btn-danger" name="action" value="finish_apply" onclick="return confirm('Завершить и скорректировать количество в карточках по факту?')">Завершить и скорректировать учёт</button>
<button class="btn btn-outline-danger" name="action" value="finish_keep" onclick="return confirm('Завершить без изменения количества в карточках?')">Завершить без корректировки</button></div>{% endif %}</form>{% endblock %}"""

T["reports.html"] = """{% extends 'base.html' %}{% block c %}<h3 class="mb-3">Отчёты</h3><div class="row g-3">
{% for k,r in reports.items() %}<div class="col-md-6"><div class="card p-3 h-100"><h5>{{r.title}}</h5><p class="text-muted">{{r.hint}}</p>
<div><a class="btn btn-danger btn-sm" href="/reports/{{k}}">Открыть</a>
<a class="btn btn-outline-secondary btn-sm" href="/reports/{{k}}?fmt=xlsx">Excel</a> <a class="btn btn-outline-secondary btn-sm" href="/reports/{{k}}?fmt=pdf">PDF</a></div></div></div>{% endfor %}</div>{% endblock %}"""

T["report.html"] = """{% extends 'base.html' %}{% block c %}
<a href="/reports">← Отчёты</a><h3 class="mt-2">{{rep.title}}</h3>
{% if rep.filters %}<form class="row g-2 my-2">{% for name,label,typ,opts in rep.filters %}<div class="col-md-2 col-6"><label class="form-label small mb-0">{{label}}</label>
{% if typ=='select' %}<select class="form-select form-select-sm" name="{{name}}"><option value="">Все</option>{% for o in opts %}<option {{ 'selected' if rep.vals.get(name)==o }}>{{o}}</option>{% endfor %}</select>
{% elif typ=='emp' %}<select class="form-select form-select-sm" name="{{name}}"><option value="">Все</option>{% for o in opts %}<option value="{{o.id}}" {{ 'selected' if rep.vals.get(name)==o.id|string }}>{{o.full_name}}</option>{% endfor %}</select>
{% elif typ=='date' %}<input class="form-control form-control-sm" type="date" name="{{name}}" value="{{ rep.vals.get(name) or '' }}">
{% elif typ=='check' %}<div class="form-check pt-1"><input class="form-check-input" type="checkbox" name="{{name}}" value="1" {{ 'checked' if rep.vals.get(name) }}></div>
{% else %}<input class="form-control form-control-sm" name="{{name}}" value="{{ rep.vals.get(name) or '' }}">{% endif %}</div>{% endfor %}
<div class="col-auto align-self-end"><button class="btn btn-secondary btn-sm">Показать</button></div></form>{% endif %}
<div class="my-2"><a class="btn btn-success btn-sm" href="{{ url_for('report', key=key, fmt='xlsx', **qs) }}">⬇ Excel</a>
<a class="btn btn-danger btn-sm" href="{{ url_for('report', key=key, fmt='pdf', **qs) }}">⬇ PDF</a> <small class="text-muted ms-2">Фильтры: {{rep.subtitle}}</small></div>
<div class="card table-responsive"><table class="table table-sm table-hover mb-0"><thead class="table-light"><tr>{% for h in rep.headers %}<th>{{h}}</th>{% endfor %}</tr></thead><tbody>
{% for r in rep.disp %}<tr>{% for v in r %}<td class="{{ 'text-end' if rep.kinds[loop.index0] in ('int','money') }}">{{v}}</td>{% endfor %}</tr>
{% else %}<tr><td colspan="{{rep.headers|length}}" class="text-muted">Нет данных</td></tr>{% endfor %}</tbody>
{% if rep.totals_disp %}<tfoot class="table-secondary fw-bold"><tr>{% for v in rep.totals_disp %}<td class="{{ 'text-end' if rep.kinds[loop.index0] in ('int','money') }}">{{v}}</td>{% endfor %}</tr></tfoot>{% endif %}</table></div>{% endblock %}"""

T["admin_users.html"] = """{% extends 'base.html' %}{% block c %}<h3>Пользователи</h3>
<div class="card p-3 my-3"><h6>Новый пользователь</h6><form method="post" class="row g-2">{{ csrf_input() }}
<div class="col-md-2"><input class="form-control" name="login" placeholder="Логин" value="{{ f.get('login','') }}" required></div>
<div class="col-md-3"><input class="form-control" name="full_name" placeholder="ФИО" value="{{ f.get('full_name','') }}" required></div>
<div class="col-md-2"><select class="form-select" name="role">{% for k,v in ROLES.items() %}<option value="{{k}}" {{ 'selected' if f.get('role')==k }}>{{v}}</option>{% endfor %}</select></div>
<div class="col-md-3"><select class="form-select" name="employee_id"><option value="">Сотрудник: не привязан</option>{% for e in emps %}<option value="{{e.id}}">{{e.full_name}}</option>{% endfor %}</select></div>
<div class="col-md-2"><input class="form-control" type="password" name="password" placeholder="Пароль" required autocomplete="new-password"></div>
<div class="col-12"><button class="btn btn-danger">Создать</button> <small class="text-muted ms-2">Пользователь, привязанный к сотруднику, видит в разделе «Мои выдачи» своё имущество.</small></div></form></div>
<div class="card table-responsive"><table class="table mb-0 align-middle"><thead><tr><th>Логин</th><th>ФИО</th><th>Роль</th><th>Сотрудник</th><th>Последний вход</th><th>Доступ</th><th></th></tr></thead><tbody>
{% for u in rows %}<tr><td>{{u.login}}</td><td>{{u.full_name}}</td><td>{{ROLES[u.role]}}</td><td>{{u.emp_name or '—'}}</td><td>{{u.last_login or '—'}}</td>
<td><span class="badge bg-{{ 'success' if u.active else 'secondary' }}">{{ 'активен' if u.active else 'отключён' }}</span></td>
<td><a class="btn btn-sm btn-outline-secondary" href="/admin/users/{{u.id}}">✎</a></td></tr>{% endfor %}</tbody></table></div>
<div class="card p-3 mt-3"><h6>Права ролей</h6><ul class="mb-0 small"><li><b>Администратор</b> — всё, включая пользователей, журнал действий и резервные копии.</li>
<li><b>Ответственный</b> — карточки, выдача и возврат, брони, инвентаризация, закупки, справочники, отчёты и экспорт.</li>
<li><b>Пользователь</b> — просмотр каталога, поиск, свои выдачи и уведомления.</li></ul></div>{% endblock %}"""

T["admin_user_edit.html"] = """{% extends 'base.html' %}{% block c %}<a href="/admin/users">← Пользователи</a><h3 class="mt-2">Пользователь {{u.login}}</h3>
<div class="card p-3 mt-3"><form method="post" class="row g-3">{{ csrf_input() }}
<div class="col-md-5"><label class="form-label">ФИО</label><input class="form-control" name="full_name" value="{{ f.get('full_name','') }}" required></div>
<div class="col-md-3"><label class="form-label">Роль</label><select class="form-select" name="role">{% for k,v in ROLES.items() %}<option value="{{k}}" {{ 'selected' if f.get('role')==k }}>{{v}}</option>{% endfor %}</select></div>
<div class="col-md-4"><label class="form-label">Сотрудник</label><select class="form-select" name="employee_id"><option value="">— не привязан —</option>
{% for e in emps %}<option value="{{e.id}}" {{ 'selected' if f.get('employee_id')|string==e.id|string }}>{{e.full_name}}</option>{% endfor %}</select></div>
<div class="col-md-5"><label class="form-label">Новый пароль (оставьте пустым, чтобы не менять)</label><input class="form-control" type="password" name="password" autocomplete="new-password"></div>
<div class="col-md-4 pt-4"><div class="form-check"><input class="form-check-input" type="checkbox" name="active" value="1" id="act" {{ 'checked' if f.get('active') }}><label class="form-check-label" for="act">Доступ разрешён</label></div></div>
<div class="col-12"><button class="btn btn-danger">Сохранить</button></div></form></div>{% endblock %}"""

T["admin_backup.html"] = """{% extends 'base.html' %}{% block c %}<div class="d-flex justify-content-between align-items-center"><h3>Резервные копии</h3>
<form method="post" action="/admin/backup/create">{{ csrf_input() }}<button class="btn btn-danger">+ Создать копию сейчас</button></form></div>
<p class="text-muted mt-2">Автокопирование: {{ 'каждые %s ч, хранится последних %s авто-копий'|format(auto_h|round(1), keep) if auto_h>0 else 'отключено' }}
(работает, пока запущено приложение). Копии лежат в папке <code>backups/</code>; перед восстановлением состояние сохраняется отдельной копией.</p>
<div class="card table-responsive"><table class="table mb-0 align-middle"><thead><tr><th>Файл</th><th>Тип</th><th>Создана</th><th>Размер</th><th></th></tr></thead><tbody>
{% for b in rows %}<tr><td><code>{{b.name}}</code></td><td>{{b.kind}}</td><td>{{b.ts}}</td><td>{{ (b.size/1024)|round(0)|int }} КБ</td><td class="text-nowrap">
<a class="btn btn-sm btn-outline-secondary" href="/admin/backup/download?name={{b.name}}">Скачать</a>
<form method="post" action="/admin/backup/restore" class="d-inline" onsubmit="return confirm('Восстановить базу из этой копии? Текущие данные будут заменены (их копия сохранится).')">{{ csrf_input() }}<input type="hidden" name="name" value="{{b.name}}"><button class="btn btn-sm btn-outline-danger">Восстановить</button></form>
<form method="post" action="/admin/backup/delete" class="d-inline" onsubmit="return confirm('Удалить копию?')">{{ csrf_input() }}<input type="hidden" name="name" value="{{b.name}}"><button class="btn btn-sm btn-outline-secondary">✕</button></form></td></tr>
{% else %}<tr><td colspan="5" class="text-muted">Копий пока нет</td></tr>{% endfor %}</tbody></table></div>{% endblock %}"""

T["admin_audit.html"] = """{% extends 'base.html' %}{% block c %}<h3>Журнал действий</h3>
<form class="row g-2 my-2"><div class="col-md-2"><select class="form-select form-select-sm" name="login"><option value="">Все пользователи</option>{% for l in logins %}<option {{ 'selected' if a.get('login')==l }}>{{l}}</option>{% endfor %}</select></div>
<div class="col-md-3"><select class="form-select form-select-sm" name="action"><option value="">Все действия</option>{% for l in actions %}<option {{ 'selected' if a.get('action')==l }}>{{l}}</option>{% endfor %}</select></div>
<div class="col-md-2"><input class="form-control form-control-sm" type="date" name="date_from" value="{{ a.get('date_from','') }}"></div>
<div class="col-md-2"><input class="form-control form-control-sm" type="date" name="date_to" value="{{ a.get('date_to','') }}"></div>
<div class="col-md-2"><input class="form-control form-control-sm" name="s" value="{{ a.get('s','') }}" placeholder="Текст в деталях"></div>
<div class="col-md-1"><button class="btn btn-secondary btn-sm w-100">Найти</button></div></form>
<div class="card table-responsive"><table class="table table-sm mb-0"><thead><tr><th>Время</th><th>Пользователь</th><th>Действие</th><th>Объект</th><th>Детали</th><th>IP</th></tr></thead><tbody>
{% for r in rows %}<tr><td class="text-nowrap">{{r.ts}}</td><td>{{r.login or '—'}}</td><td>{{r.action}}</td><td>{{r.entity or ''}} {{r.entity_id or ''}}</td><td><small>{{r.details or ''}}</small></td><td>{{r.ip or ''}}</td></tr>
{% else %}<tr><td colspan="6" class="text-muted">Записей нет</td></tr>{% endfor %}</tbody></table></div>
<div class="d-flex justify-content-between mt-2"><small class="text-muted">Записей: {{total}}, страница {{page}} из {{pages}}</small><div>
{% if page>1 %}<a class="btn btn-sm btn-outline-secondary" href="{{ url_for('admin_audit', page=page-1, **qs) }}">←</a>{% endif %}
{% if page<pages %}<a class="btn btn-sm btn-outline-secondary" href="{{ url_for('admin_audit', page=page+1, **qs) }}">→</a>{% endif %}</div></div>{% endblock %}"""

# ---------------------------------------------------------------- 3D-просмотрщик (JavaScript)
INSTRUMENT3D_JS = r'''/* instrument3d.js — мини-движок WebGL и процедурные 3D-модели музыкальных инструментов.
   Без внешних библиотек и интернета: геометрия строится кодом, рендер — один WebGL-шейдер. */
(function (global) {
'use strict';
var PI = Math.PI, MODELS = {};

/* ---------- матрицы (column-major) ---------- */
function m4id() { return [1,0,0,0, 0,1,0,0, 0,0,1,0, 0,0,0,1]; }
function m4mul(a, b) {
  var o = new Array(16), c, r, k, s;
  for (c = 0; c < 4; c++) for (r = 0; r < 4; r++) {
    s = 0; for (k = 0; k < 4; k++) s += a[k*4+r] * b[c*4+k];
    o[c*4+r] = s;
  }
  return o;
}
function rotX(a) { var c = Math.cos(a), s = Math.sin(a); return [1,0,0,0, 0,c,s,0, 0,-s,c,0, 0,0,0,1]; }
function rotY(a) { var c = Math.cos(a), s = Math.sin(a); return [c,0,-s,0, 0,1,0,0, s,0,c,0, 0,0,0,1]; }
function rotZ(a) { var c = Math.cos(a), s = Math.sin(a); return [c,s,0,0, -s,c,0,0, 0,0,1,0, 0,0,0,1]; }
var D2R = PI / 180;
function xform(o) {            // T * Rz * Ry * Rx * S ; углы в градусах
  var s = o.s == null ? [1,1,1] : (typeof o.s === 'number' ? [o.s,o.s,o.s] : o.s);
  var r = o.r || [0,0,0], p = o.p || [0,0,0];
  var m = [s[0],0,0,0, 0,s[1],0,0, 0,0,s[2],0, 0,0,0,1];
  m = m4mul(rotX(r[0]*D2R), m); m = m4mul(rotY(r[1]*D2R), m); m = m4mul(rotZ(r[2]*D2R), m);
  m[12] = p[0]; m[13] = p[1]; m[14] = p[2];
  return m;
}
function cross(a, b) { return [a[1]*b[2]-a[2]*b[1], a[2]*b[0]-a[0]*b[2], a[0]*b[1]-a[1]*b[0]]; }
function sub(a, b) { return [a[0]-b[0], a[1]-b[1], a[2]-b[2]]; }
function dot(a, b) { return a[0]*b[0]+a[1]*b[1]+a[2]*b[2]; }
function norm(a) { var l = Math.sqrt(dot(a,a)) || 1; return [a[0]/l, a[1]/l, a[2]/l]; }
function perspective(fovy, asp, n, f) {
  var t = 1 / Math.tan(fovy / 2);
  return [t/asp,0,0,0, 0,t,0,0, 0,0,(f+n)/(n-f),-1, 0,0,2*f*n/(n-f),0];
}
function lookAt(e, c, u) {
  var z = norm(sub(e, c)), x = norm(cross(u, z)), y = cross(z, x);
  return [x[0],y[0],z[0],0, x[1],y[1],z[1],0, x[2],y[2],z[2],0, -dot(x,e),-dot(y,e),-dot(z,e),1];
}
function hex(h) {
  h = h.replace('#', ''); if (h.length === 3) h = h[0]+h[0]+h[1]+h[1]+h[2]+h[2];
  var n = parseInt(h, 16);
  return [Math.pow((n>>16&255)/255, 2.2), Math.pow((n>>8&255)/255, 2.2), Math.pow((n&255)/255, 2.2)];
}

/* ---------- построитель сцены ---------- */
function Builder() { this.P = []; this.N = []; this.C = []; this.M = []; this.I = []; this.st = [m4id()]; }
Builder.prototype.push = function (o) { this.st.push(m4mul(this.st[this.st.length-1], xform(o))); return this; };
Builder.prototype.pop = function () { this.st.pop(); return this; };
// o: p (позиция), r (углы, °), s (масштаб), c (цвет #hex), g (глянец 0..1, ≥.8 — металл), w (текстура дерева 0..1)
Builder.prototype.add = function (g, o) {
  o = o || {};
  var M = this.st[this.st.length-1];
  if (o.p || o.r || o.s != null) M = m4mul(M, xform(o));
  var c0 = [M[0],M[1],M[2]], c1 = [M[4],M[5],M[6]], c2 = [M[8],M[9],M[10]];
  var r0 = cross(c1, c2), r1 = cross(c2, c0), r2 = cross(c0, c1);
  var col = hex(o.c || '#888888'), gl = o.g == null ? 0.3 : o.g, gr = o.w || 0;
  var base = this.P.length / 3, i, x, y, z, nx, ny, nz, ax, ay, az, l;
  for (i = 0; i < g.p.length; i += 3) {
    x = g.p[i]; y = g.p[i+1]; z = g.p[i+2];
    this.P.push(M[0]*x+M[4]*y+M[8]*z+M[12], M[1]*x+M[5]*y+M[9]*z+M[13], M[2]*x+M[6]*y+M[10]*z+M[14]);
    nx = g.n[i]; ny = g.n[i+1]; nz = g.n[i+2];
    ax = nx*r0[0]+ny*r1[0]+nz*r2[0]; ay = nx*r0[1]+ny*r1[1]+nz*r2[1]; az = nx*r0[2]+ny*r1[2]+nz*r2[2];
    l = Math.sqrt(ax*ax+ay*ay+az*az) || 1;
    this.N.push(ax/l, ay/l, az/l);
    this.C.push(col[0], col[1], col[2]); this.M.push(gl, gr);
  }
  for (i = 0; i < g.i.length; i++) this.I.push(g.i[i] + base);
  return this;
};

/* ---------- геометрия ---------- */
function grid(nu, nv, fn, wrap) {
  var W = nu + 1, p = [], idx = [], n = new Array((nv+1)*W*3), i, j, k, q;
  for (j = 0; j <= nv; j++) for (i = 0; i <= nu; i++) { q = fn(wrap ? i % nu : i, j); p.push(q[0], q[1], q[2]); }
  for (k = 0; k < n.length; k++) n[k] = 0;
  for (j = 0; j < nv; j++) for (i = 0; i < nu; i++) {
    var a = j*W+i, b = a+1, c = a+W, d = c+1;
    idx.push(a,c,b, b,c,d);
  }
  for (k = 0; k < idx.length; k += 3) {
    var ia = idx[k]*3, ib = idx[k+1]*3, ic = idx[k+2]*3;
    var e1 = [p[ib]-p[ia], p[ib+1]-p[ia+1], p[ib+2]-p[ia+2]], e2 = [p[ic]-p[ia], p[ic+1]-p[ia+1], p[ic+2]-p[ia+2]];
    var f = cross(e1, e2);
    for (q = 0; q < 3; q++) { n[ia+q] += f[q]; n[ib+q] += f[q]; n[ic+q] += f[q]; }
  }
  if (wrap) for (j = 0; j <= nv; j++) {
    var s0 = (j*W)*3, s1 = (j*W+nu)*3;
    for (q = 0; q < 3; q++) { var sm = n[s0+q] + n[s1+q]; n[s0+q] = sm; n[s1+q] = sm; }
  }
  [0, nv].forEach(function (row) {   // полюса (все точки ряда совпадают)
    var a0 = row*W*3, b0 = (row*W + (nu>>1))*3;
    if (Math.abs(p[a0]-p[b0]) + Math.abs(p[a0+1]-p[b0+1]) + Math.abs(p[a0+2]-p[b0+2]) < 1e-7 && nv > 1) {
      var adj = (row === 0 ? 1 : nv-1) * W, sx = 0, sy = 0, sz = 0, t;
      for (t = 0; t < W; t++) { sx += n[(adj+t)*3]; sy += n[(adj+t)*3+1]; sz += n[(adj+t)*3+2]; }
      for (t = 0; t < W; t++) { n[(row*W+t)*3] = sx; n[(row*W+t)*3+1] = sy; n[(row*W+t)*3+2] = sz; }
    }
  });
  for (k = 0; k < n.length; k += 3) {
    var l = Math.sqrt(n[k]*n[k]+n[k+1]*n[k+1]+n[k+2]*n[k+2]);
    if (l < 1e-12) { n[k] = 0; n[k+1] = 1; n[k+2] = 0; } else { n[k] /= l; n[k+1] /= l; n[k+2] /= l; }
  }
  return { p: p, n: n, i: idx };
}
var G = {};
G.grid = grid;
G.lathe = function (prof, seg) {            // профиль [[r,y],...] вращается вокруг оси Y; дубли точек дают острое ребро
  seg = seg || 36;
  return grid(seg, prof.length - 1, function (i, j) {
    var a = 2*PI*i/seg; return [prof[j][0]*Math.cos(a), prof[j][1], prof[j][0]*Math.sin(a)];
  }, true);
};
G.cyl = function (r, h, seg) { return G.lathe([[0,-h/2],[r,-h/2],[r,-h/2],[r,h/2],[r,h/2],[0,h/2]], seg || 24); };
G.cone = function (r1, r2, h, seg) { return G.lathe([[0,-h/2],[r1,-h/2],[r1,-h/2],[r2,h/2],[r2,h/2],[0,h/2]], seg || 24); };
G.disc = function (r, seg) { return G.lathe([[0,0],[r,0]], seg || 32); };
G.ring = function (r1, r2, seg) { return G.lathe([[r1,0],[r2,0]], seg || 32); };
G.sphere = function (r, seg) {
  seg = seg || 20; var pr = [], k; for (k = 0; k <= seg/2; k++) { var a = PI*k/(seg/2); pr.push([r*Math.sin(a), -r*Math.cos(a)]); }
  return G.lathe(pr, seg);
};
G.torus = function (R, r, s1, s2) {
  s1 = s1 || 32; s2 = s2 || 12;
  return grid(s1, s2, function (i, j) {
    var a = 2*PI*i/s1, b = 2*PI*j/s2, rr = R + r*Math.cos(b);
    return [rr*Math.cos(a), r*Math.sin(b), rr*Math.sin(a)];
  }, true);
};
G.box = function (w, h, d) {
  var x = w/2, y = h/2, z = d/2, p = [], n = [], idx = [];
  var F = [ [[0,0,1],[-x,-y,z],[x,-y,z],[x,y,z],[-x,y,z]], [[0,0,-1],[x,-y,-z],[-x,-y,-z],[-x,y,-z],[x,y,-z]],
            [[1,0,0],[x,-y,z],[x,-y,-z],[x,y,-z],[x,y,z]], [[-1,0,0],[-x,-y,-z],[-x,-y,z],[-x,y,z],[-x,y,-z]],
            [[0,1,0],[-x,y,z],[x,y,z],[x,y,-z],[-x,y,-z]], [[0,-1,0],[-x,-y,-z],[x,-y,-z],[x,-y,z],[-x,-y,z]] ];
  F.forEach(function (f) {
    var b = p.length / 3, k;
    for (k = 1; k <= 4; k++) { p.push(f[k][0], f[k][1], f[k][2]); n.push(f[0][0], f[0][1], f[0][2]); }
    idx.push(b, b+1, b+2, b, b+2, b+3);
  });
  return { p: p, n: n, i: idx };
};
// путь: массив точек [x,y,z]; r — число или f(t), t∈[0,1]; ell — сжатие по второй оси сечения
G.tube = function (path, r, seg, opt) {
  opt = opt || {}; seg = seg || 10;
  var ell = opt.ell == null ? 1 : opt.ell, cap = opt.cap !== false, K = path.length - 1, k, i;
  var L = [0]; for (k = 1; k <= K; k++) L.push(L[k-1] + Math.sqrt(dot(sub(path[k], path[k-1]), sub(path[k], path[k-1]))));
  var tot = L[K] || 1, rf = typeof r === 'function' ? r : function () { return r; };
  var Tn = [], Nn = [], Bn = [];
  for (k = 0; k <= K; k++) {
    var t = norm(sub(path[Math.min(k+1, K)], path[Math.max(k-1, 0)]));
    var nrm;
    if (k === 0) { var ref = Math.abs(t[1]) < 0.9 ? [0,1,0] : [1,0,0]; var d0 = dot(ref, t); nrm = norm([ref[0]-t[0]*d0, ref[1]-t[1]*d0, ref[2]-t[2]*d0]); }
    else { var pv = Nn[k-1], d1 = dot(pv, t); nrm = norm([pv[0]-t[0]*d1, pv[1]-t[1]*d1, pv[2]-t[2]*d1]); }
    Tn.push(t); Nn.push(nrm); Bn.push(cross(t, nrm));
  }
  var rows = [];
  if (cap) rows.push([0,0]); rows.push([0,1]); if (cap) rows.push([0,1]);
  for (k = 1; k < K; k++) rows.push([k,1]);
  rows.push([K,1]); if (cap) rows.push([K,1], [K,0]);
  return grid(seg, rows.length - 1, function (i, j) {
    var kk = rows[j][0], sc = rows[j][1], a = 2*PI*i/seg, rad = rf(L[kk]/tot) * sc;
    var c = Math.cos(a)*rad, s = Math.sin(a)*rad*ell, P = path[kk];
    return [P[0]+c*Nn[kk][0]+s*Bn[kk][0], P[1]+c*Nn[kk][1]+s*Bn[kk][1], P[2]+c*Nn[kk][2]+s*Bn[kk][2]];
  }, true);
};
// плоский контур poly [[a,b]] с выдавливанием; axis 'z': (a,b,z∈±d/2), 'y': (a,y∈±d/2,b)
G.extrude = function (poly, d, axis) {
  var n = poly.length, i, cx = 0, cy = 0;
  function pt(a, b, t) { return axis === 'y' ? [a, t, b] : [a, b, t]; }
  var wall = grid(n, 1, function (i, j) { var q = poly[i % n]; return pt(q[0], q[1], j ? d/2 : -d/2); }, true);
  poly.forEach(function (q) { cx += q[0]; cy += q[1]; }); cx /= n; cy /= n;
  var p = wall.p.slice(), nn = wall.n.slice(), idx = wall.i.slice();
  [-1, 1].forEach(function (sg) {
    var base = p.length / 3, c = pt(cx, cy, sg*d/2), nv = axis === 'y' ? [0,sg,0] : [0,0,sg];
    p.push(c[0], c[1], c[2]); nn.push(nv[0], nv[1], nv[2]);
    for (i = 0; i < n; i++) { var v = pt(poly[i][0], poly[i][1], sg*d/2); p.push(v[0], v[1], v[2]); nn.push(nv[0], nv[1], nv[2]); }
    for (i = 0; i < n; i++) idx.push(base, base+1+i, base+1+(i+1)%n);
  });
  return { p: p, n: nn, i: idx };
};
// «пухлое» тело из контура: outline [[x,y]] (замкнутый), prof [[s,z]] — масштаб контура и высота по оси Z
G.body = function (outline, prof) {
  var n = outline.length;
  return grid(n, prof.length - 1, function (i, j) { var q = outline[i], s = prof[j][0]; return [q[0]*s, q[1]*s, prof[j][1]]; }, true);
};

/* ---------- кривые ---------- */
function catmull(pts, n, closed) {
  var out = [], N = pts.length, segs = closed ? N : N - 1, s, k, d, D = pts[0].length;
  for (s = 0; s < segs; s++) {
    var p0 = pts[closed ? (s-1+N) % N : Math.max(s-1, 0)], p1 = pts[s], p2 = pts[(s+1) % N], p3 = pts[closed ? (s+2) % N : Math.min(s+2, N-1)];
    for (k = 0; k < n; k++) {
      var t = k/n, t2 = t*t, t3 = t2*t, q = [];
      for (d = 0; d < D; d++) q.push(0.5*((2*p1[d]) + (-p0[d]+p2[d])*t + (2*p0[d]-5*p1[d]+4*p2[d]-p3[d])*t2 + (-p0[d]+3*p1[d]-3*p2[d]+p3[d])*t3));
      out.push(q);
    }
  }
  if (!closed) out.push(pts[N-1].slice());
  return out;
}
function roundPath(pts, rad, n) {         // ломаная со скруглёнными углами
  n = n || 8; var out = [pts[0].slice()], k, i, D = pts[0].length;
  for (k = 1; k < pts.length - 1; k++) {
    var a = pts[k-1], v = pts[k], b = pts[k+1], d1 = 0, d2 = 0, d;
    for (d = 0; d < D; d++) { d1 += (v[d]-a[d])*(v[d]-a[d]); d2 += (b[d]-v[d])*(b[d]-v[d]); }
    d1 = Math.sqrt(d1); d2 = Math.sqrt(d2);
    var r = Math.min(rad, d1/2, d2/2), pi = [], po = [];
    for (d = 0; d < D; d++) { pi.push(v[d] + (a[d]-v[d])/d1*r); po.push(v[d] + (b[d]-v[d])/d2*r); }
    for (i = 0; i <= n; i++) { var t = i/n, q = []; for (d = 0; d < D; d++) q.push((1-t)*(1-t)*pi[d] + 2*(1-t)*t*v[d] + t*t*po[d]); out.push(q); }
  }
  out.push(pts[pts.length-1].slice()); return out;
}
function mirrorHalf(half, per) {          // правая половина (сверху вниз, x=0 на концах) -> замкнутый гладкий контур
  var full = half.concat(half.slice(1, -1).reverse().map(function (p) { return [-p[0], p[1]]; }));
  return catmull(full, per || 6, true);
}
function polyScale(poly, k, cx, cy) { cx = cx || 0; cy = cy || 0; return poly.map(function (p) { return [cx + (p[0]-cx)*k, cy + (p[1]-cy)*k]; }); }
function ellipse(rx, ry, n) { var o = [], i; for (i = 0; i < (n||40); i++) { var a = 2*PI*i/(n||40); o.push([rx*Math.cos(a), ry*Math.sin(a)]); } return o; }
function outlineScale(outline, x, y) {    // масштаб s точки относительно контура (для посадки деталей на выпуклую деку)
  var a = Math.atan2(y, x), best = 1e9, bl = 1, i;
  for (i = 0; i < outline.length; i++) {
    var da = Math.abs(Math.atan2(Math.sin(Math.atan2(outline[i][1], outline[i][0]) - a), Math.cos(Math.atan2(outline[i][1], outline[i][0]) - a)));
    if (da < best) { best = da; bl = Math.sqrt(outline[i][0]*outline[i][0] + outline[i][1]*outline[i][1]); }
  }
  return Math.min(1, Math.sqrt(x*x + y*y) / bl);
}
function rod(b, p0, p1, r, o) { return b.add(G.tube([p0, p1], r, o && o.seg || 8), o || {}); }
function lerp(a, b, t) { return a + (b - a) * t; }

/* ---------- просмотрщик ---------- */
var VS = 'attribute vec3 a_pos,a_nrm,a_col;attribute vec2 a_mat;uniform mat4 u_mvp,u_view;' +
  'varying vec3 v_n,v_wn,v_col,v_vp,v_wp;varying vec2 v_mat;' +
  'void main(){v_n=mat3(u_view)*a_nrm;v_wn=a_nrm;v_vp=(u_view*vec4(a_pos,1.)).xyz;v_wp=a_pos;v_col=a_col;v_mat=a_mat;gl_Position=u_mvp*vec4(a_pos,1.);}';
var FS = '#ifdef GL_FRAGMENT_PRECISION_HIGH\nprecision highp float;\n#else\nprecision mediump float;\n#endif\n' +
  'varying vec3 v_n,v_wn,v_col,v_vp,v_wp;varying vec2 v_mat;uniform float u_gs;' +
  'void main(){vec3 N=normalize(v_n),V=normalize(-v_vp);if(dot(N,V)<0.)N=-N;' +
  'vec3 col=v_col;float g=v_mat.x;' +
  'if(v_mat.y>0.){float k=sin(v_wp.y*u_gs*38.+sin(v_wp.x*u_gs*9.+v_wp.z*u_gs*7.)*2.2)*.5+.5;col*=1.-v_mat.y*.22*k;}' +
  'vec3 L1=normalize(vec3(.45,.75,.6)),L2=normalize(vec3(-.7,.15,.5)),L3=normalize(vec3(0.,.35,-.9));' +
  'float amb=mix(.32,.62,v_wn.y*.5+.5);' +
  'vec3 lit=col*(amb+max(dot(N,L1),0.)*.8+max(dot(N,L2),0.)*.26+max(dot(N,L3),0.)*.2);' +
  'vec3 H=normalize(L1+V);float sp=pow(max(dot(N,H),0.),mix(18.,110.,g))*(.05+.75*g);' +
  'vec3 R=reflect(-V,N);float env=smoothstep(-.7,.9,R.y)*.85+.15+smoothstep(.45,.6,R.y)*smoothstep(.95,.7,R.y)*.5;' +
  'float m=smoothstep(.72,1.,g);vec3 metal=col*env*1.35+vec3(sp);' +
  'vec3 c=mix(lit+vec3(sp)*.6,metal,m);' +
  'c+=pow(1.-max(dot(N,V),0.),3.)*(.04+.1*g);' +
  'c=c/(1.+c*.22);gl_FragColor=vec4(pow(c,vec3(1./2.2)),1.);}';
var SVS = 'attribute vec2 a_p;uniform mat4 u_mvp;varying vec2 v_u;void main(){v_u=a_p;gl_Position=u_mvp*vec4(a_p.x,0.,a_p.y,1.);}';
var SFS = 'precision mediump float;varying vec2 v_u;void main(){float d=length(v_u);gl_FragColor=vec4(0.,0.,0.,.42*pow(clamp(1.-d,0.,1.),1.7));}';

function compile(gl, vs, fs) {
  function sh(t, s) { var o = gl.createShader(t); gl.shaderSource(o, s); gl.compileShader(o);
    if (!gl.getShaderParameter(o, gl.COMPILE_STATUS)) throw new Error(gl.getShaderInfoLog(o)); return o; }
  var p = gl.createProgram(); gl.attachShader(p, sh(gl.VERTEX_SHADER, vs)); gl.attachShader(p, sh(gl.FRAGMENT_SHADER, fs));
  gl.linkProgram(p); if (!gl.getProgramParameter(p, gl.LINK_STATUS)) throw new Error(gl.getProgramInfoLog(p)); return p;
}
function bounds(b) {
  var mn = [1e9,1e9,1e9], mx = [-1e9,-1e9,-1e9], i, k;
  for (i = 0; i < b.P.length; i += 3) for (k = 0; k < 3; k++) { mn[k] = Math.min(mn[k], b.P[i+k]); mx[k] = Math.max(mx[k], b.P[i+k]); }
  return { mn: mn, mx: mx };
}
function build(kind) {
  var def = MODELS[kind] || MODELS.note, b = def.build();
  return { def: def, b: b, bb: bounds(b) };
}
var CSS_DONE = false;
function injectCss() {
  if (CSS_DONE) return; CSS_DONE = true;
  var s = document.createElement('style');
  s.textContent = '.i3d{position:relative;width:100%;height:100%;min-height:260px;border-radius:.6rem;overflow:hidden;' +
    'background:radial-gradient(ellipse at 50% 35%,#fbf7f0 0%,#e9e1d3 62%,#d6cbb8 100%);user-select:none;-webkit-user-select:none}' +
    '.i3d canvas{width:100%;height:100%;display:block;touch-action:none;cursor:grab}.i3d canvas:active{cursor:grabbing}' +
    '.i3d-bar{position:absolute;right:8px;top:8px;display:flex;gap:6px}' +
    '.i3d-bar button{border:0;border-radius:.4rem;background:rgba(91,26,43,.88);color:#fff;font:12px/1 system-ui,sans-serif;padding:6px 9px;cursor:pointer}' +
    '.i3d-bar button.off{background:rgba(60,60,60,.55)}' +
    '.i3d-hint{position:absolute;left:10px;bottom:7px;font:11px system-ui,sans-serif;color:#6b5e4b;pointer-events:none}' +
    '.i3d-msg{display:flex;align-items:center;justify-content:center;height:100%;padding:1rem;text-align:center;color:#6b5e4b;font:14px system-ui,sans-serif}';
  document.head.appendChild(s);
}

function mount(host, kind, opts) {
  opts = opts || {}; injectCss();
  host.classList.add('i3d'); host.innerHTML = '';
  var canvas = document.createElement('canvas'); host.appendChild(canvas);
  var gl = null;
  try { gl = canvas.getContext('webgl', { antialias: true, alpha: true, preserveDrawingBuffer: !!opts.preserve }) ||
             canvas.getContext('experimental-webgl', { antialias: true, alpha: true, preserveDrawingBuffer: !!opts.preserve }); } catch (e) {}
  var prog, sprog;
  try { if (gl) { prog = compile(gl, VS, FS); sprog = compile(gl, SVS, SFS); } } catch (e) { gl = null; if (global.console) console.error(e); }
  if (!gl) { host.innerHTML = '<div class="i3d-msg">3D-просмотр недоступен: браузер или устройство не поддерживает WebGL.</div>'; return null; }
  var uint = gl.getExtension('OES_element_index_uint');
  var V = { yaw: 0, pitch: 0, dist: 1, auto: !opts.still && !(global.matchMedia && matchMedia('(prefers-reduced-motion: reduce)').matches),
            visible: true, raf: 0, last: 0, drag: false }, M = null, vbo = gl.createBuffer(), ibo = gl.createBuffer(),
      sbo = gl.createBuffer(), icount = 0, ptr = {}, pdist = 0;
  var sdata = [0,0], si; for (si = 0; si <= 40; si++) sdata.push(Math.cos(2*PI*si/40), Math.sin(2*PI*si/40));
  gl.bindBuffer(gl.ARRAY_BUFFER, sbo); gl.bufferData(gl.ARRAY_BUFFER, new Float32Array(sdata), gl.STATIC_DRAW);

  function setKind(k) {
    M = build(k); var b = M.b, n = b.P.length / 3, f = new Float32Array(n * 11), i;
    for (i = 0; i < n; i++) {
      f[i*11] = b.P[i*3]; f[i*11+1] = b.P[i*3+1]; f[i*11+2] = b.P[i*3+2];
      f[i*11+3] = b.N[i*3]; f[i*11+4] = b.N[i*3+1]; f[i*11+5] = b.N[i*3+2];
      f[i*11+6] = b.C[i*3]; f[i*11+7] = b.C[i*3+1]; f[i*11+8] = b.C[i*3+2];
      f[i*11+9] = b.M[i*2]; f[i*11+10] = b.M[i*2+1];
    }
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo); gl.bufferData(gl.ARRAY_BUFFER, f, gl.STATIC_DRAW);
    var ix = (n > 65535 && uint) ? new Uint32Array(b.I) : new Uint16Array(b.I);
    M.itype = ix instanceof Uint32Array ? gl.UNSIGNED_INT : gl.UNSIGNED_SHORT;
    gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo); gl.bufferData(gl.ELEMENT_ARRAY_BUFFER, ix, gl.STATIC_DRAW); icount = b.I.length;
    var mn = M.bb.mn, mx = M.bb.mx;
    M.c = [(mn[0]+mx[0])/2, (mn[1]+mx[1])/2, (mn[2]+mx[2])/2];
    var dm = [mx[0]-mn[0], mx[1]-mn[1], mx[2]-mn[2]];
    M.size = Math.max(dm[0], dm[1], dm[2]);
    M.rad = 0.5 * Math.sqrt(M.size*M.size + 0.55*(dm[0]*dm[0] + dm[1]*dm[1] + dm[2]*dm[2] - M.size*M.size));
    M.floor = mn[1] - (M.def.float ? M.size * 0.08 : 0.0005);
    reset();
  }
  function fitDist() {
    var w = canvas.clientWidth || 300, h = canvas.clientHeight || 300, fov = 32 * D2R;
    var half = Math.min(fov / 2, Math.atan(Math.tan(fov / 2) * w / h));
    return M.rad / Math.sin(half) * 1.07;
  }
  function reset() { V.yaw = (M.def.yaw != null ? M.def.yaw : -32) * D2R; V.pitch = (M.def.pitch != null ? M.def.pitch : 14) * D2R; V.dist = 1; kick(); }
  function resize() {
    var dpr = Math.min(global.devicePixelRatio || 1, 2), w = Math.max(1, Math.round(canvas.clientWidth * dpr)), h = Math.max(1, Math.round(canvas.clientHeight * dpr));
    if (canvas.width !== w || canvas.height !== h) { canvas.width = w; canvas.height = h; }
  }
  function draw() {
    if (!M) return; resize(); gl.viewport(0, 0, canvas.width, canvas.height);
    gl.clearColor(0, 0, 0, 0); gl.clear(gl.COLOR_BUFFER_BIT | gl.DEPTH_BUFFER_BIT);
    var d = fitDist() * V.dist, cp = Math.cos(V.pitch);
    var eye = [M.c[0] + d*cp*Math.sin(V.yaw), M.c[1] + d*Math.sin(V.pitch), M.c[2] + d*cp*Math.cos(V.yaw)];
    var view = lookAt(eye, M.c, [0,1,0]), proj = perspective(32*D2R, canvas.width / canvas.height, d*0.04, d*8), mvp = m4mul(proj, view);
    gl.enable(gl.DEPTH_TEST);
    // тень
    var mn = M.bb.mn, mx = M.bb.mx, sm = m4mul(mvp, xform({ p: [M.c[0], M.floor, M.c[2]], s: [(mx[0]-mn[0])*0.62 + M.size*0.06, 1, (mx[2]-mn[2])*0.62 + M.size*0.06] }));
    gl.useProgram(sprog); gl.enable(gl.BLEND); gl.blendFunc(gl.SRC_ALPHA, gl.ONE_MINUS_SRC_ALPHA); gl.depthMask(false);
    gl.uniformMatrix4fv(gl.getUniformLocation(sprog, 'u_mvp'), false, new Float32Array(sm));
    gl.bindBuffer(gl.ARRAY_BUFFER, sbo); var ap = gl.getAttribLocation(sprog, 'a_p');
    gl.enableVertexAttribArray(ap); gl.vertexAttribPointer(ap, 2, gl.FLOAT, false, 8, 0); gl.drawArrays(gl.TRIANGLE_FAN, 0, 42);
    gl.disableVertexAttribArray(ap); gl.depthMask(true); gl.disable(gl.BLEND);
    // модель
    gl.useProgram(prog);
    gl.uniformMatrix4fv(gl.getUniformLocation(prog, 'u_mvp'), false, new Float32Array(mvp));
    gl.uniformMatrix4fv(gl.getUniformLocation(prog, 'u_view'), false, new Float32Array(view));
    gl.uniform1f(gl.getUniformLocation(prog, 'u_gs'), 1 / M.size);
    gl.bindBuffer(gl.ARRAY_BUFFER, vbo); gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER, ibo);
    [['a_pos',3,0],['a_nrm',3,3],['a_col',3,6],['a_mat',2,9]].forEach(function (a) {
      var l = gl.getAttribLocation(prog, a[0]); gl.enableVertexAttribArray(l); gl.vertexAttribPointer(l, a[1], gl.FLOAT, false, 44, a[2]*4);
    });
    gl.drawElements(gl.TRIANGLES, icount, M.itype, 0);
  }
  function frame(t) {
    V.raf = 0; if (!V.visible) return;
    var dt = V.last ? Math.min(0.05, (t - V.last) / 1000) : 0; V.last = t;
    if (V.auto && !V.drag) V.yaw -= dt * 0.45;
    draw();
    if (V.auto && !V.drag) V.raf = requestAnimationFrame(frame); else V.last = 0;
  }
  function kick() { if (!V.raf) V.raf = requestAnimationFrame(frame); }

  // управление
  function clampV() { V.pitch = Math.max(0.03, Math.min(1.35, V.pitch)); V.dist = Math.max(0.45, Math.min(2.4, V.dist)); }
  canvas.addEventListener('pointerdown', function (e) {
    ptr[e.pointerId] = [e.clientX, e.clientY]; V.drag = true; try { canvas.setPointerCapture(e.pointerId); } catch (x) {}
    var ks = Object.keys(ptr); if (ks.length === 2) pdist = Math.hypot(ptr[ks[0]][0]-ptr[ks[1]][0], ptr[ks[0]][1]-ptr[ks[1]][1]);
  });
  canvas.addEventListener('pointermove', function (e) {
    if (!ptr[e.pointerId]) return;
    var ks = Object.keys(ptr), o = ptr[e.pointerId];
    if (ks.length >= 2) {
      ptr[e.pointerId] = [e.clientX, e.clientY];
      var nd = Math.hypot(ptr[ks[0]][0]-ptr[ks[1]][0], ptr[ks[0]][1]-ptr[ks[1]][1]);
      if (pdist > 0 && nd > 0) V.dist *= pdist / nd; pdist = nd;
    } else { V.yaw -= (e.clientX - o[0]) * 0.008; V.pitch += (e.clientY - o[1]) * 0.006; ptr[e.pointerId] = [e.clientX, e.clientY]; }
    clampV(); kick();
  });
  function up(e) { delete ptr[e.pointerId]; if (!Object.keys(ptr).length) { V.drag = false; kick(); } }
  canvas.addEventListener('pointerup', up); canvas.addEventListener('pointercancel', up);
  canvas.addEventListener('wheel', function (e) { e.preventDefault(); V.dist *= Math.exp(e.deltaY * 0.0012); clampV(); kick(); }, { passive: false });
  canvas.addEventListener('dblclick', function () { reset(); });
  canvas.addEventListener('webglcontextlost', function (e) { e.preventDefault(); });

  // кнопки
  var bar = document.createElement('div'); bar.className = 'i3d-bar';
  var bR = document.createElement('button'), bZ = document.createElement('button');
  bR.type = bZ.type = 'button'; bR.textContent = '⟳ Вращение'; bZ.textContent = '↺ Сброс';
  function paintR() { bR.className = V.auto ? '' : 'off'; }
  bR.onclick = function () { V.auto = !V.auto; paintR(); kick(); };
  bZ.onclick = function () { reset(); };
  paintR(); bar.appendChild(bR); bar.appendChild(bZ); host.appendChild(bar);
  if (!opts.still) { var hint = document.createElement('div'); hint.className = 'i3d-hint'; hint.textContent = 'Тяните — вращение, колесо — масштаб'; host.appendChild(hint); }

  if (global.ResizeObserver) new ResizeObserver(function () { kick(); }).observe(host);
  if (global.IntersectionObserver) new IntersectionObserver(function (en) { V.visible = en[0].isIntersecting; if (V.visible) kick(); }).observe(host);

  setKind(kind);
  return {
    setKind: setKind,
    setView: function (yaw, pitch, dist) { V.auto = false; V.yaw = yaw * D2R; V.pitch = pitch * D2R; V.dist = dist || 1; draw(); },
    canvas: canvas,
    destroy: function () { V.visible = false; if (V.raf) cancelAnimationFrame(V.raf); var x = gl.getExtension('WEBGL_lose_context'); if (x) x.loseContext(); host.innerHTML = ''; }
  };
}

global.Instrument3D = { mount: mount, build: build, models: MODELS, kinds: function () { return Object.keys(MODELS); },
  title: function (k) { return (MODELS[k] || MODELS.note).title; } };

/* ---------- палитра ---------- */
var K = { black: '#141418', gloss: '#0c0c10', ebony: '#1c1511', ivory: '#f4f0e4', maple: '#d6a96a', spruce: '#e2c48c', varnish: '#a04f1c',
  rosewood: '#4b2519', mahogany: '#6f301b', brass: '#d9aa38', gold: '#d8b04a', silver: '#cdd2d9', steel: '#9ea4ac', copper: '#b9652f',
  red: '#8f1c26', cream: '#efe6cf', gray: '#6f6f76', dark: '#2a2622' };

function keyboard(b, o) {                 // клавиатура: белые клавиши по оси X, длина клавиш по Z
  var n = o.n, pw = o.w / n, L = o.len, h = o.h || 0.022, bh = o.bh || 0.016, seq = [0,1,2,3,4,5,6], si = 'CDEFGAB'.indexOf(o.start || 'C'), i;
  for (i = 0; i < n; i++) {
    var x = o.x - o.w/2 + pw*(i + 0.5);
    b.add(G.box(pw*0.94, h, L), { p: [x, o.y + h/2, o.z + L/2], c: K.ivory, g: 0.45 });
    var letter = 'CDEFGAB'[(si + i) % 7];
    if (i < n - 1 && 'CDFGA'.indexOf(letter) >= 0)
      b.add(G.box(pw*0.58, bh, L*0.62), { p: [x + pw/2, o.y + h + bh/2, o.z + L*0.31], c: K.black, g: 0.6 });
  }
}
function legTaper(b, x, z, y0, y1, r0, r1, c) {
  b.add(G.lathe([[0,y0],[r1,y0],[r1,y0],[r0,y1],[r0,y1],[0,y1]], 20), { p: [x, 0, z], c: c || K.gloss, g: 0.7 });
  b.add(G.cyl(r1*1.1, 0.04, 16), { p: [x, y0 + 0.02, z], c: K.brass, g: 0.95 });
}
function bench(b, z, w) {
  b.add(G.box(w, 0.07, 0.38), { p: [0, 0.5, z], c: '#2b1c1c', g: 0.35 });
  b.add(G.box(w + 0.02, 0.025, 0.4), { p: [0, 0.455, z], c: K.gloss, g: 0.7 });
  [[-1,-1],[1,-1],[-1,1],[1,1]].forEach(function (s) { b.add(G.box(0.045, 0.43, 0.045), { p: [s[0]*(w/2-0.05), 0.22, z + s[1]*0.15], c: K.gloss, g: 0.7 }); });
}

/* ---------- рояль ---------- */
MODELS.grand = { title: 'Рояль', yaw: -38, pitch: 20, build: function () {
  var b = new Builder();
  var curve = catmull([[0.74,-0.40],[0.78,-0.95],[0.62,-1.75],[0.25,-2.3],[-0.2,-2.56],[-0.55,-2.6],[-0.71,-2.52]], 7, false);
  var poly = [[-0.74,0.0],[0.74,0.0]].concat(curve).concat([[-0.74,-2.42]]);
  var cx = 0, cz = -1.3, y0 = 0.64, hCase = 0.30;
  b.add(G.extrude(poly, hCase, 'y'), { p: [0, y0 + hCase/2, 0], c: K.gloss, g: 0.75 });
  b.add(G.extrude(polyScale(poly, 0.985, cx, cz), 0.02, 'y'), { p: [0, y0 + 0.14, 0], c: K.spruce, g: 0.2, w: 0.4 });
  b.add(G.extrude(polyScale(poly, 0.93, cx, cz), 0.014, 'y'), { p: [0, y0 + 0.22, 0], c: K.gold, g: 0.9 });
  var n = 44, i;
  for (i = 0; i < n; i++) {
    var t = i / (n - 1), c = curve[Math.floor((1 - t) * (curve.length - 1))];
    var ex = cx + (c[0]-cx)*0.9, ez = cz + (c[1]-cz)*0.92;
    rod(b, [-0.62 + 1.24*t, y0 + 0.24, -0.2], [ex, y0 + 0.24, ez], 0.0022, { c: K.silver, g: 0.95, seg: 5 });
  }
  // клавиатура
  b.add(G.box(1.46, 0.10, 0.30), { p: [0, 0.67, 0.12], c: K.gloss, g: 0.7 });
  b.add(G.box(1.50, 0.12, 0.035), { p: [0, 0.66, 0.285], c: K.gloss, g: 0.7 });
  keyboard(b, { x: 0, y: 0.72, z: -0.03, n: 52, start: 'A', w: 1.22, len: 0.26 });
  [-0.7, 0.7].forEach(function (x) { b.add(G.box(0.09, hCase, 0.4), { p: [x, y0 + hCase/2, 0.01], c: K.gloss, g: 0.75 }); });
  b.add(G.box(0.8, 0.3, 0.015), { p: [0, 1.03, -0.07], r: [-14, 0, 0], c: K.gloss, g: 0.75 });
  // крышка
  var lid = 31;
  b.push({ p: [-0.74, y0 + hCase + 0.03, 0], r: [0, 0, lid] });
  b.add(G.extrude(poly, 0.045, 'y'), { p: [0.74, 0.0225, 0], c: K.gloss, g: 0.8 });
  b.pop();
  rod(b, [0.52, y0 + hCase, -0.95], [0.52, y0 + hCase + 0.74, -0.95], 0.013, { c: K.rosewood, g: 0.5 });
  // ноги, лира, педали
  legTaper(b, -0.64, 0.04, 0.0, y0, 0.065, 0.036);
  legTaper(b, 0.64, 0.04, 0.0, y0, 0.065, 0.036);
  legTaper(b, -0.3, -2.15, 0.0, y0, 0.065, 0.036);
  [-0.07, 0.07].forEach(function (x) { b.add(G.cyl(0.022, 0.5, 14), { p: [x, 0.38, 0.28], c: K.gloss, g: 0.7 }); });
  b.add(G.box(0.26, 0.1, 0.14), { p: [0, 0.14, 0.28], c: K.gloss, g: 0.7 });
  [-0.07, 0, 0.07].forEach(function (x) { b.add(G.box(0.04, 0.012, 0.14), { p: [x, 0.04, 0.34], c: K.brass, g: 0.95 }); });
  bench(b, 0.95, 0.9);
  return b;
} };

/* ---------- пианино ---------- */
MODELS.upright = { title: 'Пианино', yaw: -30, pitch: 14, build: function () {
  var b = new Builder();
  b.add(G.box(1.5, 1.08, 0.58), { p: [0, 0.68, 0], c: K.gloss, g: 0.75 });
  b.add(G.box(1.56, 0.035, 0.64), { p: [0, 1.235, 0.02], c: K.gloss, g: 0.8 });
  b.add(G.box(1.54, 0.12, 0.62), { p: [0, 0.1, 0], c: K.gloss, g: 0.7 });
  b.add(G.box(1.46, 0.09, 0.32), { p: [0, 0.745, 0.35], c: K.gloss, g: 0.7 });
  b.add(G.box(1.5, 0.11, 0.04), { p: [0, 0.69, 0.53], c: K.gloss, g: 0.7 });
  keyboard(b, { x: 0, y: 0.79, z: 0.22, n: 52, start: 'A', w: 1.22, len: 0.25 });
  [-0.78, 0.78].forEach(function (x) { b.add(G.box(0.05, 0.62, 0.42), { p: [x, 0.78, 0.31], c: K.gloss, g: 0.75 }); });
  b.add(G.box(1.2, 0.34, 0.015), { p: [0, 1.0, 0.27], r: [-12, 0, 0], c: K.gloss, g: 0.8 });
  b.add(G.box(1.1, 0.4, 0.01), { p: [0, 1.0, 0.288], r: [-12, 0, 0], c: '#2a2a31', g: 0.3 });
  b.add(G.box(1.2, 0.025, 0.05), { p: [0, 0.83, 0.3], c: K.gloss, g: 0.7 });
  [-0.5, 0.5].forEach(function (x) { b.add(G.box(0.12, 0.05, 0.5), { p: [x, 0.025, 0.05], c: K.gloss, g: 0.7 }); });
  [-0.07, 0, 0.07].forEach(function (x) { b.add(G.box(0.045, 0.012, 0.13), { p: [x, 0.05, 0.4], c: K.brass, g: 0.95 }); });
  [-0.07, 0, 0.07].forEach(function (x) { b.add(G.box(0.025, 0.12, 0.02), { p: [x, 0.11, 0.31], c: K.brass, g: 0.95 }); });
  bench(b, 0.98, 0.8);
  return b;
} };

/* ---------- синтезатор / цифровое пианино ---------- */
MODELS.synth = { title: 'Синтезатор', yaw: -30, pitch: 22, build: function () {
  var b = new Builder();
  b.add(G.box(1.36, 0.1, 0.4), { p: [0, 0.9, 0], c: K.black, g: 0.5 });
  b.add(G.box(1.3, 0.012, 0.11), { p: [0, 0.953, -0.12], c: '#2a2d36', g: 0.6 });
  var i; for (i = 0; i < 16; i++) b.add(G.cyl(0.011, 0.012, 12), { p: [-0.55 + i*0.073, 0.965, -0.1 + (i % 2) * 0.04], c: i % 3 ? K.silver : K.red, g: 0.8 });
  b.add(G.box(0.2, 0.004, 0.07), { p: [0.52, 0.957, -0.16], c: '#7fd0e8', g: 0.9 });
  keyboard(b, { x: 0, y: 0.87, z: 0.02, n: 36, start: 'C', w: 1.26, len: 0.19, h: 0.022 });
  [-1, 1].forEach(function (s) {
    rod(b, [s*0.45, 0.85, 0], [-s*0.45, 0.04, 0], 0.016, { c: K.steel, g: 0.9 });
    b.add(G.box(0.07, 0.03, 0.34), { p: [-s*0.45, 0.02, 0], c: K.black, g: 0.4 });
  });
  return b;
} };

/* ---------- запасная модель: нота ---------- */
MODELS.note = { title: 'Музыкальный инструмент', yaw: -20, pitch: 10, build: function () {
  var b = new Builder();
  b.add(G.cyl(0.42, 0.05, 48), { p: [0, 0.025, 0], c: '#5b1a2b', g: 0.55 });
  b.add(G.torus(0.42, 0.012, 48, 8), { p: [0, 0.05, 0], c: K.gold, g: 0.95 });
  b.add(G.sphere(1, 24), { p: [-0.1, 0.33, 0], r: [0, 0, 22], s: [0.17, 0.11, 0.1], c: K.gold, g: 0.9 });
  b.add(G.cyl(0.018, 0.78, 14), { p: [0.062, 0.74, 0], c: K.gold, g: 0.9 });
  var flag = catmull([[0.062,1.11,0],[0.15,1.0,0],[0.24,0.9,0],[0.2,0.78,0]], 8, false);
  b.add(G.tube(flag, function (t) { return 0.03 * (1 - 0.6*t); }, 10, { ell: 0.6 }), { c: K.gold, g: 0.9 });
  return b;
} };

/* ---------- смычковые ---------- */
var VIOLIN_HALF = [[0,.5],[.09,.495],[.185,.455],[.232,.37],[.235,.27],[.215,.17],[.17,.09],[.157,.04],[.17,-.01],[.225,-.08],[.285,-.17],[.295,-.28],[.265,-.4],[.19,-.475],[.1,-.497],[0,-.5]];
function bowed(P) {
  var b = new Builder(), H = P.H, a = P.a, outline = mirrorHalf(VIOLIN_HALF, 6), k, s;
  function zt(s) { return H/2 + a*(1 - Math.pow(s, 2.2)); }
  var prof = [];
  for (k = 0; k <= 12; k++) { s = k/12; prof.push([s, zt(s)]); }
  prof.push([1, -H/2]);
  for (k = 11; k >= 0; k--) { s = k/12; prof.push([s, -(H/2 + a*0.75*(1 - Math.pow(s, 2.2)))]); }
  b.add(G.body(outline, prof), { c: P.c || K.varnish, g: 0.6, w: 0.35 });
  var yN = 0.5 + P.neck, yE = yN - P.fbL, yb = P.yb, zn = H/2 + 0.028;
  var zb = zt(outlineScale(outline, 0, yb)), zbTop = zb + P.bh;
  function zs(y) { return zn + (yN - y)/(yN - yb)*(zbTop - zn); }
  // гриф
  var nw = P.nw, ew = P.ew, tf = 0.022, topN = zn - 0.006, topE = zs(yE) - P.clr, yM = (yN + yE)/2;
  var ang = Math.atan2(topE - topN, yN - yE) / D2R;
  b.add(G.extrude([[-nw/2, yN-yM],[nw/2, yN-yM],[ew/2, yE-yM],[-ew/2, yE-yM]], tf, 'z'),
        { p: [0, yM, (topN + topE)/2 - tf/2], r: [-ang, 0, 0], c: K.ebony, g: 0.65 });
  rod(b, [0, yN, zn - 0.035], [0, 0.47, H/2 - 0.004], 0.03, { c: K.maple, g: 0.5, w: 0.2 });
  b.add(G.box(nw, 0.008, 0.012), { p: [0, yN, zn - 0.003], c: K.ivory, g: 0.3 });
  // колковая коробка, колки, завиток
  var zc = zn - 0.02;
  b.add(G.box(0.085, P.peg, 0.055), { p: [0, yN + P.peg/2, zc], c: K.maple, g: 0.5, w: 0.2 });
  for (k = 0; k < 4; k++) {
    var sd = k % 2 ? 1 : -1, yy = yN + 0.035 + k*(P.peg - 0.07)/3;
    rod(b, [0, yy, zc], [sd*0.092, yy, zc], 0.011, { c: K.ebony, g: 0.6 });
    b.add(G.cyl(0.026, 0.02, 14), { p: [sd*0.105, yy, zc], r: [0, 0, 90], c: K.ebony, g: 0.6 });
  }
  var yc = yN + P.peg + 0.03, sp = [];
  for (k = 0; k <= 40; k++) { var t = k/40, an = t*2.6*PI, R = 0.05*(1 - 0.78*t); sp.push([0, yc + R*Math.cos(an), zc - R*Math.sin(an)]); }
  b.add(G.tube(sp, function (t) { return 0.034*(1 - 0.55*t); }, 10), { c: K.maple, g: 0.5, w: 0.2 });
  // подставка, струны, подбородник, струнодержатель
  b.add(G.extrude([[-.058,0],[.058,0],[.052,.02],[.045,.06],[.04,P.bh],[-.04,P.bh],[-.045,.06],[-.052,.02]], 0.008, 'y'),
        { p: [0, yb, zb], c: '#e6bd80', g: 0.4 });
  var ytail = -0.2, ztail = zt(outlineScale(outline, 0, ytail)) + 0.022;
  for (k = 0; k < 4; k++) {
    var q = k - 1.5;
    b.add(G.tube([[q*0.0105, yN, zn + 0.004], [q*0.036, yb, zbTop], [q*0.02, ytail, ztail]], P.sr, 6),
          { c: k === 3 ? '#d9c36a' : K.silver, g: 0.95 });
  }
  var ytc = (ytail - 0.46)/2, ztc = zt(outlineScale(outline, 0, ytc)) + 0.012;
  b.add(G.extrude([[-.032,ytail-ytc],[.032,ytail-ytc],[.05,-.46-ytc],[-.05,-.46-ytc]], 0.016, 'z'), { p: [0, ytc, ztc], c: K.ebony, g: 0.7 });
  rod(b, [0, -0.45, ztc], [0, -0.503, 0.0], 0.006, { c: '#d8d2c0', g: 0.3 });
  b.add(G.cyl(0.017, 0.035, 14), { p: [0, -0.505, 0], c: K.ebony, g: 0.6 });
  [-1, 1].forEach(function (sg) {   // эфы
    var pts = [[0, .085],[.009, .045],[.015, 0],[.009, -.045],[0, -.085]].map(function (p) {
      var x = sg*(0.1 + p[0]*1.0), y = yb + p[1] + 0.005; return [x, y, zt(outlineScale(outline, x, y)) + 0.003];
    });
    b.add(G.tube(pts, 0.0048, 6, { ell: 0.35 }), { c: '#140b06', g: 0.3 });
  });
  if (P.chin) {
    var cz = zt(outlineScale(outline, -0.15, -0.36)) + 0.03;
    b.add(G.extrude(ellipse(0.13, 0.07, 32), 0.04, 'z'), { p: [-0.15, -0.36, cz], r: [0, 0, 20], c: K.ebony, g: 0.7 });
  }
  if (P.pin) {
    rod(b, [0, -0.5, 0], [0, -0.5 - P.pin, 0], 0.011, { c: K.steel, g: 0.95 });
    b.add(G.cone(0.011, 0.002, 0.04, 8), { p: [0, -0.5 - P.pin - 0.02, 0], r: [180, 0, 0], c: K.steel, g: 0.95 });
  }
  return b;
}
var BW = { H: .085, a: .045, neck: .37, fbL: .76, peg: .17, yb: -.06, bh: .093, clr: .03, nw: .07, ew: .12, sr: .003, chin: true };
function ext(o, p) { var r = {}, k; for (k in o) r[k] = o[k]; for (k in p) r[k] = p[k]; return r; }
MODELS.violin = { title: 'Скрипка', yaw: -28, pitch: 8, float: true, build: function () { return bowed(ext(BW, {})); } };
MODELS.viola = { title: 'Альт', yaw: -28, pitch: 8, float: true, build: function () { return bowed(ext(BW, { H: .09, c: '#8c3f17' })); } };
MODELS.cello = { title: 'Виолончель', yaw: -28, pitch: 8, build: function () {
  return bowed(ext(BW, { H: .16, a: .05, neck: .42, fbL: .92, peg: .16, yb: -.05, bh: .115, nw: .075, ew: .13, sr: .0036, chin: false, pin: .3, c: '#8f3d14' })); } };
MODELS.bass = { title: 'Контрабас', yaw: -28, pitch: 8, build: function () {
  return bowed(ext(BW, { H: .2, a: .05, neck: .45, fbL: .9, peg: .15, yb: -.05, bh: .12, nw: .075, ew: .14, sr: .004, chin: false, pin: .3, c: '#7a3512' })); } };

/* ---------- щипковые ---------- */
var GUITAR_HALF = [[0,.5],[.12,.497],[.22,.465],[.285,.4],[.297,.32],[.275,.22],[.25,.12],[.24,.05],[.262,-.02],[.32,-.1],[.37,-.2],[.385,-.3],[.355,-.4],[.27,-.48],[.15,-.497],[0,-.5]];
var ELEC_HALF = [[0,.47],[.07,.475],[.12,.5],[.18,.5],[.21,.44],[.2,.36],[.24,.27],[.31,.2],[.35,.1],[.34,-.02],[.33,-.1],[.35,-.2],[.35,-.33],[.3,-.44],[.18,-.5],[0,-.5]];
var PEAR_HALF = [[0,.5],[.2,.46],[.33,.32],[.4,.12],[.4,-.1],[.33,-.3],[.2,-.45],[0,-.5]];
var BAL_R = [[0,.485],[.07,.37],[.138,.2],[.23,0],[.322,-.2],[.4,-.38],[.44,-.47],[.38,-.5],[.2,-.5]];
var BAL_PTS = BAL_R.concat([[0,-.5]]).concat(BAL_R.slice(1).reverse().map(function (p) { return [-p[0], p[1]]; }));

function plucked(P) {
  var b = new Builder(), H = P.H, k, outline = P.outline;
  var top = P.body === 'solid' ? null : true;
  var prof;
  if (P.body === 'solid') prof = [[0,H/2],[.9,H/2],[.97,H/2-.008],[1,H/2-.03],[1,-H/2+.03],[.97,-H/2+.008],[.9,-H/2],[0,-H/2]];
  else if (P.body === 'bowl') { prof = [[0,0],[1,0]]; for (k = 0; k <= 14; k++) { var s = 1 - k/14; prof.push([s, -P.bowl*Math.sqrt(Math.max(0, 1 - s*s))]); } }
  else prof = [[0,H/2],[.92,H/2],[.985,H/2-.01],[1,H/2-.03],[1,-H/2+.03],[.985,-H/2+.01],[.92,-H/2],[0,-H/2]];
  b.add(G.body(outline, prof), { c: P.bodyC, g: 0.6, w: 0.35 });
  var zTop = P.body === 'bowl' ? 0 : H/2;
  if (P.topC) b.add(G.extrude(polyScale(outline, 0.992), 0.004, 'z'), { p: [0, 0, zTop + 0.0015], c: P.topC, g: 0.35, w: 0.3 });
  if (P.hole) {
    var hy = P.hole[0], hr = P.hole[1];
    b.add(G.disc(hr, 40), { p: [0, hy, zTop + 0.004], r: [90, 0, 0], c: '#0b0706', g: 0.2 });
    b.add(G.ring(hr, hr + 0.012, 40), { p: [0, hy, zTop + 0.0045], r: [90, 0, 0], c: P.rosC || K.rosewood, g: 0.5 });
    b.add(G.ring(hr + 0.016, hr + 0.021, 40), { p: [0, hy, zTop + 0.0045], r: [90, 0, 0], c: K.cream, g: 0.5 });
  }
  var yb = P.yb, yN = yb + P.sc, nutW = P.nutW, zfb = zTop + 0.008, zbr = zTop + P.bh, nF = P.frets;
  var yE = yN - P.sc*(1 - Math.pow(2, -nF/12)), yM = (yN + yE)/2;
  function wAt(y) { return nutW + (yN - y)*P.taper; }
  b.add(G.extrude([[-nutW/2, yN-yM],[nutW/2, yN-yM],[wAt(yE)/2, yE-yM],[-wAt(yE)/2, yE-yM]], 0.022, 'z'), { p: [0, yM, zfb - 0.011], c: P.fbC, g: 0.6 });
  rod(b, [0, yN, zfb - 0.05], [0, 0.42, zfb - 0.05], 0.05, { c: P.neckC, g: 0.5, w: 0.3, ell: 0.8 });
  b.add(G.tube([[0, yN, zfb - 0.05], [0, 0.42, zfb - 0.05]], 0.05, 12, { ell: 0.75 }), { c: P.neckC, g: 0.5, w: 0.3 });
  for (k = 1; k <= nF; k++) {
    var fy = yN - P.sc*(1 - Math.pow(2, -k/12));
    b.add(G.box(wAt(fy) + 0.006, 0.006, 0.006), { p: [0, fy, zfb + 0.002], c: K.silver, g: 0.95 });
    if ([3,5,7,9,15,17].indexOf(k) >= 0) {
      var fy0 = yN - P.sc*(1 - Math.pow(2, -(k-0.5)/12));
      b.add(G.disc(0.008, 12), { p: [0, fy0, zfb + 0.0012], r: [90, 0, 0], c: K.cream, g: 0.4 });
    }
  }
  b.add(G.box(nutW, 0.008, 0.014), { p: [0, yN, zfb + 0.004], c: K.cream, g: 0.3 });
  // колки
  var hl = P.hl, hw0 = nutW*1.15, hw1 = P.hw;
  b.push({ p: [0, yN, zfb - 0.03], r: [-12, 0, 0] });
  b.add(G.extrude([[-hw0/2,0],[hw0/2,0],[hw1/2,hl*.65],[hw1/2*.9,hl],[-hw1/2*.9,hl],[-hw1/2,hl*.65]], 0.034, 'z'), { c: P.hdC, g: 0.6 });
  var tl = P.tun.l || 0, tr = P.tun.r || 0;
  for (k = 0; k < Math.max(tl, tr); k++) {
    var ty = P.tun.face ? 0.045 + k*(hl - 0.07)/Math.max(1, Math.max(tl, tr) - 1) : 0.045 + k*(hl - 0.06)/Math.max(1, Math.max(tl, tr) - 1);
    [[-1, tl], [1, tr]].forEach(function (sd) {
      if (k >= sd[1]) return; var sg = sd[0];
      if (P.tun.face) {
        b.add(G.cyl(0.009, 0.05, 10), { p: [sg*0.022, ty, 0.0], r: [90, 0, 0], c: K.steel, g: 0.95 });
        b.add(G.cyl(0.013, 0.008, 12), { p: [sg*0.022, ty, 0.03], r: [90, 0, 0], c: K.silver, g: 0.95 });
      } else {
        b.add(G.cyl(0.007, 0.05, 10), { p: [sg*0.04, ty, 0.012], r: [90, 0, 0], c: K.steel, g: 0.95 });
        rod(b, [sg*hw1*0.45, ty, 0], [sg*(hw1/2 + 0.04), ty, 0], 0.007, { c: K.steel, g: 0.95 });
        b.add(G.cyl(0.02, 0.01, 14), { p: [sg*(hw1/2 + 0.045), ty, 0], r: [0, 0, 90], c: K.cream, g: 0.5 });
      }
    });
  }
  b.pop();
  // мост и струны
  if (P.body === 'solid') {
    b.add(G.box(0.2, 0.09, 0.012), { p: [0, yb, zTop + 0.006], c: K.silver, g: 0.95 });
    b.add(G.extrude(outline.map(function (p) { return [p[0]*0.8, Math.min(p[1]*0.8, 0.3)]; }), 0.004, 'z'), { p: [0, -0.02, zTop + 0.002], c: '#efe9dd', g: 0.45 });
    [0.14, 0.02, -0.1].forEach(function (y, i) { if (P.ns === 4 && i === 1) return; b.add(G.box(0.19, 0.05, 0.014), { p: [0, y, zTop + 0.01], c: i % 2 ? '#e9e0cc' : '#f4f1e8', g: 0.4 }); });
    [[0.2,-0.3],[0.12,-0.37]].forEach(function (p) { b.add(G.cyl(0.026, 0.03, 16), { p: [p[0], p[1], zTop + 0.016], r: [90, 0, 0], c: K.black, g: 0.7 }); });
  } else {
    b.add(G.box(P.bsp + 0.04, 0.035, P.bh - 0.003), { p: [0, yb, zTop + (P.bh - 0.003)/2], c: K.rosewood, g: 0.55 });
    b.add(G.box(P.bsp, 0.006, 0.01), { p: [0, yb + 0.012, zbr + 0.004], c: K.cream, g: 0.3 });
  }
  var ns = P.ns, q;
  for (k = 0; k < ns; k++) {
    q = k - (ns - 1)/2;
    b.add(G.tube([[q*nutW*0.78/Math.max(ns-1, 1), yN, zfb + 0.009], [q*P.bsp/Math.max(ns-1, 1), yb, zbr + 0.004]], P.sr, 6), { c: K.silver, g: 0.95 });
  }
  return b;
}
MODELS.guitar = { title: 'Гитара', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: mirrorHalf(GUITAR_HALF, 6), body: 'acoustic', H: .2, bodyC: '#5b2c18', topC: '#e3c68c', neckC: '#6a3a22', fbC: '#2a1710', hdC: '#2c1a12',
    yb: -.17, sc: 1.35, nutW: .105, taper: .06, frets: 19, ns: 6, bsp: .17, bh: .05, hl: .27, hw: .15, tun: { l: 3, r: 3 }, sr: .0022, hole: [.1, .085] }); } };
MODELS.electric = { title: 'Электрогитара', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: mirrorHalf(ELEC_HALF, 6), body: 'solid', H: .1, bodyC: '#a3202b', neckC: '#dcb27a', fbC: '#2a1710', hdC: '#dcb27a',
    yb: -.2, sc: 1.44, nutW: .095, taper: .045, frets: 21, ns: 6, bsp: .15, bh: .035, hl: .38, hw: .1, tun: { r: 6, face: 1 }, sr: .002 }); } };
MODELS.bassg = { title: 'Бас-гитара', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: mirrorHalf(ELEC_HALF, 6), body: 'solid', H: .1, bodyC: '#1b1b20', neckC: '#d2a56c', fbC: '#2a1710', hdC: '#d2a56c',
    yb: -.2, sc: 1.9, nutW: .1, taper: .04, frets: 20, ns: 4, bsp: .13, bh: .035, hl: .4, hw: .1, tun: { r: 4, face: 1 }, sr: .003 }); } };
MODELS.balalaika = { title: 'Балалайка', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: catmull(BAL_PTS, 5, true), body: 'flat', H: .16, bodyC: '#b5763a', topC: '#e2b373', neckC: '#2a1a10', fbC: '#1e130c', hdC: '#2a1a10',
    yb: -.3, sc: 1.62, nutW: .075, taper: .05, frets: 15, ns: 3, bsp: .1, bh: .04, hl: .17, hw: .11, tun: { r: 3 }, sr: .0025, hole: [.14, .07], rosC: K.dark }); } };
MODELS.domra = { title: 'Домра', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: ellipse(0.5, 0.5, 56), body: 'bowl', bowl: .36, H: 0, bodyC: '#b36a2c', topC: '#e5c07f', neckC: '#2a1a10', fbC: '#1e130c', hdC: '#2a1a10',
    yb: -.3, sc: 1.55, nutW: .09, taper: .05, frets: 15, ns: 3, bsp: .11, bh: .04, hl: .17, hw: .12, tun: { r: 3 }, sr: .0025, hole: [.05, .09], rosC: K.dark }); } };
MODELS.mandolin = { title: 'Мандолина', yaw: -28, pitch: 8, float: true, build: function () {
  return plucked({ outline: mirrorHalf(PEAR_HALF, 6), body: 'bowl', bowl: .3, H: 0, bodyC: '#9c4f20', topC: '#e0bb78', neckC: '#3a2214', fbC: '#1e130c', hdC: '#2a1a10',
    yb: -.3, sc: 1.4, nutW: .1, taper: .05, frets: 17, ns: 4, bsp: .12, bh: .04, hl: .2, hw: .13, tun: { l: 2, r: 2 }, sr: .0022, hole: [.1, .08], rosC: K.dark }); } };

/* ---------- арфа ---------- */
MODELS.harp = { title: 'Арфа', yaw: -40, pitch: 8, build: function () {
  var b = new Builder(), k, i;
  var np = catmull([[-0.42,1.80],[-0.25,1.75],[0.0,1.66],[0.28,1.56],[0.52,1.46]], 10, false);
  function yneck(x) { for (var j = 1; j < np.length; j++) if (np[j][0] >= x) { var t = (x - np[j-1][0]) / (np[j][0] - np[j-1][0] || 1); return lerp(np[j-1][1], np[j][1], t); } return np[np.length-1][1]; }
  function ysb(x) { return 0.26 + (x + 0.22) * 1.05; }
  b.add(G.extrude(ellipse(0.56, 0.27, 48).map(function (p) { return [p[0] + 0.1, p[1]]; }), 0.11, 'y'), { p: [0, 0.055, 0], c: K.mahogany, g: 0.7 });
  b.add(G.torus(0.3, 0.012, 40, 8), { p: [0.1, 0.11, 0], s: [2.2, 1, 1], c: K.gold, g: 0.95 });
  b.add(G.lathe([[0,0.11],[.07,.11],[.07,.11],[.05,.2],[.042,.5],[.042,1.6],[.05,1.72],[.066,1.76],[.066,1.79],[0,1.79]], 24), { p: [-0.42, 0, 0], c: K.gold, g: 0.9 });
  [0.3, 0.9, 1.4, 1.64].forEach(function (y) { b.add(G.torus(0.05, 0.009, 24, 8), { p: [-0.42, y, 0], c: '#b98f2b', g: 0.95 }); });
  b.add(G.tube(np.map(function (p) { return [p[0], p[1], 0]; }), 0.035, 12, { ell: 0.8 }), { c: K.mahogany, g: 0.7 });
  for (i = 0; i <= 6; i++) { var nx = -0.12 + i*0.09; b.add(G.cyl(0.036, 0.012, 20), { p: [nx, yneck(nx), 0.034], r: [90, 0, 0], c: K.gold, g: 0.95 }); }
  var sb = []; for (i = 0; i <= 20; i++) { var x = -0.24 + i*0.039; sb.push([x, ysb(x) - 0.03, 0]); }
  b.add(G.tube(sb, function (t) { return lerp(0.17, 0.07, t); }, 20, { ell: 1.7 }), { c: K.mahogany, g: 0.7 });
  var n = 47;
  for (i = 0; i < n; i++) {
    var sx = -0.22 + 0.72*i/(n-1), c = i % 7 === 0 ? '#b8242c' : i % 7 === 3 ? '#1f1f24' : '#efe9db';
    rod(b, [sx, yneck(sx) - 0.02, 0.0], [sx, ysb(sx), 0.0], 0.0024, { c: c, g: 0.5, seg: 5 });
  }
  for (i = 0; i < 7; i++) b.add(G.box(0.045, 0.014, 0.12), { p: [-0.35 + i*0.07, 0.12, 0.26], c: K.gold, g: 0.95 });
  return b;
} };

/* ---------- деревянные духовые ---------- */
function toneRing(b, y, R, ang, r, c) {      // тональное отверстие с металлическим кольцом на стволе (вертикальный)
  var x = R * Math.sin(ang), z = R * Math.cos(ang);
  b.add(G.cyl(r*1.45, 0.0035, 14), { p: [x, y, z], r: [90, 0, -ang/D2R + 0], c: c || K.silver, g: 0.95 });
  b.add(G.cyl(r, 0.0042, 14), { p: [x*1.0006, y, z*1.0006], r: [90, 0, -ang/D2R], c: '#050505', g: 0.2 });
}
function keyCup(b, y, R, ang, r, c) {
  var x = (R + 0.004) * Math.sin(ang), z = (R + 0.004) * Math.cos(ang);
  b.add(G.cyl(r, 0.006, 12), { p: [x, y, z], r: [90, 0, -ang/D2R], c: c || K.silver, g: 0.95 });
}
function singleReed(o) {              // кларнет / гобой — вертикально, раструбом вниз на пол
  var b = new Builder(), k, y;
  var R = function (y) {
    if (y < o.bellH) return o.rb + o.ba * Math.exp(-y / o.bs);
    if (y > o.top - 0.05) return o.rb * 0.97 * (o.barrel || 1.25);
    return lerp(o.rb, o.rt, (y - o.bellH) / (o.top - o.bellH));
  };
  var pr = [[0, 0]]; for (k = 0; k <= 70; k++) { y = o.top * k / 70; pr.push([R(y), y]); } pr.push([0, o.top]);
  pr[0] = [R(0), 0]; pr.unshift([0, 0], [R(0), 0]);
  b.add(G.lathe(pr, 28), { c: o.wood || '#17120f', g: 0.5, w: 0 });
  o.bands.forEach(function (y) { b.add(G.cyl(R(y) + 0.0012, 0.012, 24), { p: [0, y, 0], c: K.silver, g: 0.95 }); });
  var rr = o.holes; for (k = 0; k < rr.length; k++) toneRing(b, rr[k], R(rr[k]), 0, o.hr);
  o.rods.forEach(function (r) { [-1, 1].forEach(function (sg) {
    var a = sg * 0.65; rod(b, [(R(r[0]) + 0.002) * Math.sin(a), r[0], (R(r[0]) + 0.002) * Math.cos(a)], [(R(r[1]) + 0.002) * Math.sin(a), r[1], (R(r[1]) + 0.002) * Math.cos(a)], 0.0013, { c: K.silver, g: 0.95, seg: 5 }); }); });
  o.cups.forEach(function (c, i) { keyCup(b, c, R(c), (i % 2 ? 1 : -1) * 0.95, o.hr * 1.1); });
  return b;
}
MODELS.clarinet = { title: 'Кларнет', yaw: -25, pitch: 6, build: function () {
  var b = singleReed({ top: 0.63, bellH: 0.14, rb: 0.0155, ba: 0.0345, bs: 0.045, rt: 0.0144, barrel: 1.25, hr: 0.0043,
    bands: [0.14, 0.4, 0.57], holes: [0.2, 0.245, 0.29, 0.44, 0.485, 0.53], rods: [[0.16, 0.38], [0.42, 0.56]], cups: [0.17, 0.33, 0.36, 0.41, 0.55, 0.51] });
  b.add(G.lathe([[0,.62],[.0185,.62],[.0185,.62],[.0175,.64],[.0145,.665],[.011,.69],[.0075,.705],[0,.705]], 24), { c: '#17120f', g: 0.6 });
  b.add(G.box(0.0135, 0.07, 0.003), { p: [0, 0.665, 0.0108], c: '#d9bb78', g: 0.4 });
  b.add(G.cyl(0.0172, 0.014, 20), { p: [0, 0.645, 0], c: K.silver, g: 0.95 });
  return b;
} };
MODELS.oboe = { title: 'Гобой', yaw: -25, pitch: 6, build: function () {
  var b = singleReed({ top: 0.52, bellH: 0.1, rb: 0.0105, ba: 0.02, bs: 0.03, rt: 0.0075, barrel: 1.05, hr: 0.0033, wood: '#1b1410',
    bands: [0.1, 0.3, 0.46], holes: [0.17, 0.2, 0.23, 0.34, 0.37, 0.4], rods: [[0.12, 0.3], [0.33, 0.48]], cups: [0.14, 0.26, 0.29, 0.35, 0.43, 0.49, 0.22] });
  rod(b, [0, 0.52, 0], [0, 0.585, 0], 0.0042, { c: K.silver, g: 0.95 });
  b.add(G.lathe([[0.0045,.56],[.0075,.575],[.0075,.575],[.0068,.615],[.0048,.645],[.0048,.645],[0.0016,.646]], 14), { c: '#d6b676', g: 0.3 });
  b.add(G.cyl(0.0079, 0.008, 14), { p: [0, 0.585, 0], c: '#3a2a18', g: 0.3 });
  return b;
} };
MODELS.bassoon = { title: 'Фагот', yaw: -28, pitch: 8, build: function () {
  var b = new Builder(), M = '#8f4f22', k;
  b.add(G.cyl(0.082, 0.22, 28), { p: [0, 0.11, 0], c: M, g: 0.65, w: 0.2 });
  b.add(G.cyl(0.085, 0.012, 28), { p: [0, 0.006, 0], c: K.silver, g: 0.95 });
  b.add(G.lathe([[0,.18],[.034,.18],[.034,.18],[.032,.92],[.032,.92],[.04,1.0],[.07,1.18],[.08,1.3],[.08,1.3],[.074,1.3],[0,1.3]], 28), { p: [0.045, 0, 0], c: M, g: 0.65, w: 0.2 });
  b.add(G.torus(0.08, 0.005, 28, 8), { p: [0.045, 1.3, 0], c: K.silver, g: 0.95 });
  b.add(G.lathe([[0,.18],[.028,.18],[.028,.18],[.02,.96],[.02,.96],[0,.96]], 24), { p: [-0.045, 0, 0], c: M, g: 0.65, w: 0.2 });
  var bp = catmull([[-.045,.95,0],[-.045,1.08,.01],[-.037,1.2,.08],[-.02,1.26,.17]], 8, false);
  b.add(G.tube(bp, 0.0058, 8), { c: K.gold, g: 0.95 });
  rod(b, [-0.02, 1.26, 0.17], [-0.01, 1.3, 0.215], 0.0045, { c: '#d6b676', g: 0.3 });
  [0.28, 0.5, 0.72, 0.94, 1.1].forEach(function (y) { b.add(G.cyl(0.036, 0.012, 24), { p: [0.045, y, 0], c: K.silver, g: 0.95 }); });
  [0.3, 0.45, 0.6, 0.75].forEach(function (y) { b.add(G.cyl(0.0225, 0.012, 24), { p: [-0.045, y, 0], c: K.silver, g: 0.95 }); });
  for (k = 0; k < 13; k++) {
    var y = 0.26 + k * 0.07, sg = k % 2 ? 1 : -1;
    b.add(G.cyl(0.012, 0.007, 12), { p: [sg * 0.045 + sg * 0.012, y, 0.034], r: [90, 0, 0], c: K.silver, g: 0.95 });
    b.add(G.cyl(0.0042, 0.004, 10), { p: [sg * 0.045, y + 0.02, 0.0345], r: [90, 0, 0], c: '#050505', g: 0.2 });
  }
  rod(b, [0.045, 0.3, 0.036], [0.045, 1.0, 0.036], 0.0016, { c: K.silver, g: 0.95, seg: 5 });
  rod(b, [-0.045, 0.3, 0.024], [-0.045, 0.9, 0.024], 0.0016, { c: K.silver, g: 0.95, seg: 5 });
  return b;
} };
MODELS.sax = { title: 'Саксофон', yaw: -18, pitch: 8, build: function () {
  var b = new Builder(), k;
  var dense = catmull([[0,.98,0],[.005,.85,0],[.01,.6,0],[.02,.35,0],[.05,.15,0],[.11,.06,0],[.2,.1,0],[.255,.22,0],[.275,.38,0]], 10, false);
  var L = [0]; for (k = 1; k < dense.length; k++) L.push(L[k-1] + Math.hypot(dense[k][0]-dense[k-1][0], dense[k][1]-dense[k-1][1]));
  var tot = L[L.length-1];
  function rf(t) { return 0.017 + 0.03 * t + 0.075 * Math.pow(Math.max(0, (t - 0.84) / 0.16), 2.4); }
  b.add(G.tube(dense, rf, 18, { cap: false }), { c: '#d6a73a', g: 0.95 });
  b.add(G.tube([[0,.98,0],[-.012,1.03,0],[-.05,1.07,0],[-.1,1.08,0]], function (t) { return 0.0125 - 0.003*t; }, 10), { c: '#d6a73a', g: 0.95 });
  b.add(G.lathe([[0.0105,0],[0.0135,0.03],[0.0155,0.07],[0.011,0.1],[0,0.105]], 20), { p: [-0.1, 1.085, 0], r: [0, 0, 90], c: '#15110e', g: 0.6 });
  b.add(G.cyl(0.0165, 0.016, 18), { p: [-0.13, 1.087, 0], r: [0, 0, 90], c: K.silver, g: 0.95 });
  function at(y) { for (var j = 0; j < dense.length; j++) if (dense[j][1] <= y && dense[j][0] < 0.08) return { p: dense[j], t: L[j] / tot }; return { p: dense[dense.length-1], t: 1 }; }
  [0.9, 0.82, 0.74, 0.66, 0.58, 0.5, 0.42, 0.34, 0.27].forEach(function (y, i) {
    var a = at(y), r = rf(a.t);
    b.add(G.cyl(r * 0.5, 0.008, 14), { p: [a.p[0], a.p[1], r + 0.004], r: [90, 0, 0], c: '#d6a73a', g: 0.95 });
    b.add(G.cyl(0.0075, 0.006, 12), { p: [a.p[0], a.p[1] + 0.01, r + 0.016], r: [90, 0, 0], c: '#f2eee4', g: 0.4 });
    if (i % 2 === 0) rod(b, [a.p[0], a.p[1], r + 0.004], [a.p[0], a.p[1] - 0.04, r + 0.004], 0.003, { c: '#d6a73a', g: 0.95, seg: 6 });
  });
  rod(b, [0.01, 0.9, 0.034], [0.03, 0.3, 0.05], 0.0022, { c: '#d6a73a', g: 0.95, seg: 6 });
  var e = dense[dense.length-1], pn = dense[dense.length-3], tg = norm([e[0]-pn[0], e[1]-pn[1], 0]);
  var ang = Math.atan2(-tg[0], tg[1]) / D2R;
  b.add(G.torus(rf(1), 0.004, 28, 8), { p: [e[0], e[1], 0], r: [0, 0, ang], c: '#d6a73a', g: 0.95 });
  return b;
} };
MODELS.flute = { title: 'Флейта', yaw: -20, pitch: 14, float: true, build: function () {
  var b = new Builder(), k;
  b.push({ p: [-0.335, 0, 0], r: [0, 0, -90] });
  var R = 0.0095;
  b.add(G.lathe([[0,0],[.007,.003],[.0105,.012],[.0105,.012],[.0098,.19],[.0098,.19],[R,.67],[R,.67],[0,.67]], 24), { c: K.silver, g: 1 });
  b.add(G.cyl(0.0112, 0.012, 22), { p: [0, 0.006, 0], c: K.silver, g: 1 });
  b.add(G.extrude(ellipse(0.0125, 0.019, 24), 0.003, 'z'), { p: [0, 0.07, 0.0102], c: K.silver, g: 1 });
  b.add(G.extrude(ellipse(0.0065, 0.0095, 20), 0.001, 'z'), { p: [0, 0.07, 0.0118], c: '#040404', g: 0.2 });
  [0.19, 0.44, 0.55].forEach(function (y) { b.add(G.cyl(0.0108, 0.007, 22), { p: [0, y, 0], c: K.silver, g: 1 }); });
  [0.235, 0.275, 0.315, 0.36, 0.4, 0.46, 0.5, 0.58, 0.615].forEach(function (y, i) {
    b.add(G.cyl(0.0078, 0.005, 14), { p: [0, y, R + 0.002], r: [90, 0, 0], c: K.silver, g: 1 });
    b.add(G.cyl(0.0035, 0.0055, 10), { p: [0, y, R + 0.0022], r: [90, 0, 0], c: '#0a0a0a', g: 0.2 });
    if (i % 3 === 0) rod(b, [0.0125, y, 0.0075], [0, y, R + 0.002], 0.0012, { c: K.silver, g: 1, seg: 5 });
  });
  [-1, 1].forEach(function (sg) { rod(b, [sg * 0.0115, 0.2, 0.004], [sg * 0.0115, 0.64, 0.004], 0.0013, { c: K.silver, g: 1, seg: 5 }); });
  b.pop();
  return b;
} };

/* ---------- медные духовые ---------- */
var BRASS = '#d9ac3a';
function bellLathe(b, p, rot, len, r0, r1, pw, c) {
  var pr = [], k; for (k = 0; k <= 36; k++) { var t = k/36; pr.push([r0 + (r1 - r0) * Math.pow(t, pw), len * t]); }
  b.add(G.lathe(pr, 36), { p: p, r: rot, c: c || BRASS, g: 0.95 });
  b.add(G.torus(r1, r1 * 0.045, 40, 8), { p: [p[0] + (rot[2] === -90 ? len : 0), p[1] + (rot[2] === 0 ? len : 0), p[2]], r: rot[2] === -90 ? [0, 0, 90] : [0, 0, 0], c: c || BRASS, g: 0.95 });
}
function mouthpiece(b, p, rot) {
  b.add(G.lathe([[0.0068,0],[0.0095,0.012],[0.0145,0.028],[0.0155,0.04],[0,0.04]], 20), { p: p, r: rot, c: K.silver, g: 0.95 });
}
function valveBlock(b, xs, z, y0) {
  xs.forEach(function (x) {
    b.add(G.cyl(0.0095, 0.1, 16), { p: [x, y0 + 0.01, z], c: BRASS, g: 0.95 });
    b.add(G.cyl(0.0115, 0.012, 16), { p: [x, y0 + 0.062, z], c: K.silver, g: 0.95 });
    b.add(G.cyl(0.0115, 0.012, 16), { p: [x, y0 - 0.04, z], c: K.silver, g: 0.95 });
    rod(b, [x, y0 + 0.068, z], [x, y0 + 0.095, z], 0.0035, { c: K.silver, g: 0.95 });
    b.add(G.cyl(0.0105, 0.008, 16), { p: [x, y0 + 0.098, z], c: '#f2eee4', g: 0.4 });
  });
}
function xz(pts) { return pts.map(function (p) { return [p[0], 0, p[1]]; }); }
MODELS.trumpet = { title: 'Труба', yaw: -30, pitch: 28, float: true, build: function () {
  var b = new Builder(), r = 0.0072;
  b.add(G.tube(xz(roundPath([[-0.4,0.03],[0.10,0.03],[0.10,-0.015],[-0.30,-0.015],[-0.30,-0.06],[0.03,-0.06]], 0.0225, 8)), r, 10), { c: BRASS, g: 0.95 });
  b.add(G.tube(xz(roundPath([[0.12,0.03],[0.2,0.03],[0.2,0.065],[0.04,0.065]], 0.02, 8)), 0.0062, 10), { c: BRASS, g: 0.95 });
  bellLathe(b, [0.03, 0, -0.06], [0, 0, -90], 0.37, 0.0075, 0.062, 3.4);
  mouthpiece(b, [-0.4, 0, 0.03], [0, 0, 90]);
  valveBlock(b, [-0.04, 0.0, 0.04], 0.03, 0.0);
  b.add(G.box(0.012, 0.012, 0.055), { p: [-0.2, 0, -0.0375], c: BRASS, g: 0.95 });
  return b;
} };
MODELS.trombone = { title: 'Тромбон', yaw: -30, pitch: 26, float: true, build: function () {
  var b = new Builder();
  b.add(G.tube(xz(roundPath([[-0.1,0],[0.55,0],[0.55,0.05],[-0.1,0.05]], 0.025, 8)), 0.0075, 10), { c: K.silver, g: 0.98 });
  b.add(G.tube(xz(roundPath([[-0.1,0],[-0.3,0],[-0.3,-0.09],[0.05,-0.09]], 0.04, 8)), 0.0072, 10), { c: BRASS, g: 0.95 });
  b.add(G.tube(xz([[-0.1,0.05],[-0.36,0.05]]), 0.0072, 10), { c: BRASS, g: 0.95 });
  bellLathe(b, [0.05, 0, -0.09], [0, 0, -90], 0.47, 0.0075, 0.1, 3.6);
  mouthpiece(b, [-0.36, 0, 0.05], [0, 0, 90]);
  [0.1, 0.5].forEach(function (x) { b.add(G.box(0.012, 0.014, 0.062), { p: [x, 0, 0.025], c: BRASS, g: 0.95 }); });
  b.add(G.cyl(0.012, 0.07, 14), { p: [0.3, 0, 0.025], r: [0, 0, 90], c: BRASS, g: 0.95 });
  return b;
} };
MODELS.horn = { title: 'Валторна', yaw: -22, pitch: 12, build: function () {
  var b = new Builder(), cx = 0, cy = 0.22, sp = [], k;
  for (k = 0; k <= 120; k++) { var t = k/120, an = 0.15*PI + t*7.4*PI, R = 0.17 - 0.06*t; sp.push([cx + R*Math.cos(an), cy + R*Math.sin(an), -0.03 + 0.07*t]); }
  b.add(G.tube(sp, 0.0062, 8), { c: BRASS, g: 0.95 });
  var s0 = sp[0], lead = [[s0[0] + 0.14, s0[1] + 0.34, -0.03], [s0[0] + 0.05, s0[1] + 0.12, -0.03], s0];
  b.add(G.tube(catmull(lead, 6, false), 0.0062, 8), { c: BRASS, g: 0.95 });
  mouthpiece(b, [s0[0] + 0.14, s0[1] + 0.34, -0.03], [0, 0, -28]);
  var e = sp[sp.length-1];
  var bell = catmull([[e[0], e[1], e[2]], [e[0] - 0.05, e[1] + 0.02, 0.1], [e[0] - 0.06, e[1] + 0.0, 0.22], [e[0] - 0.02, e[1] - 0.03, 0.34]], 8, false);
  b.add(G.tube(bell, function (t) { return 0.0065 + 0.16 * Math.pow(t, 3.1); }, 28, { cap: false }), { c: BRASS, g: 0.95 });
  [0, 1, 2].forEach(function (i) {
    b.add(G.cyl(0.026, 0.04, 20), { p: [0.02 + i * 0.001, 0.14 + i*0.05, 0.05], r: [90, 0, 0], c: BRASS, g: 0.95 });
    rod(b, [0.02, 0.14 + i*0.05, 0.07], [-0.09, 0.14 + i*0.05, 0.085], 0.004, { c: K.silver, g: 0.95 });
  });
  return b;
} };
MODELS.tuba = { title: 'Туба', yaw: -25, pitch: 10, build: function () {
  var b = new Builder();
  function loop(x0, x1, y0, y1, rad, z, r) {
    var xm = (x0 + x1)/2; return b.add(G.tube(roundPath([[xm,y0,z],[x1,y0,z],[x1,y1,z],[x0,y1,z],[x0,y0,z],[xm,y0,z]], rad, 8), r, 12), { c: BRASS, g: 0.95 });
  }
  loop(-0.2, 0.3, 0.02, 0.5, 0.14, 0.0, 0.02); loop(-0.14, 0.22, 0.08, 0.44, 0.1, 0.045, 0.018); loop(-0.08, 0.14, 0.14, 0.38, 0.07, -0.045, 0.016);
  var pr = [], k; for (k = 0; k <= 40; k++) { var t = k/40; pr.push([0.035 + 0.17 * Math.pow(t, 2.6), 0.5 + 0.5 * t]); }
  b.add(G.lathe(pr, 40), { p: [-0.1, 0, 0], c: BRASS, g: 0.95 });
  b.add(G.torus(0.205, 0.009, 40, 8), { p: [-0.1, 1.0, 0], c: BRASS, g: 0.95 });
  [0.12, 0.2, 0.28].forEach(function (x) {
    b.add(G.cyl(0.02, 0.15, 18), { p: [x, 0.58, 0.12], c: BRASS, g: 0.95 });
    b.add(G.cyl(0.022, 0.014, 18), { p: [x, 0.66, 0.12], c: K.silver, g: 0.95 });
    b.add(G.cyl(0.018, 0.01, 18), { p: [x, 0.69, 0.12], c: '#f2eee4', g: 0.4 });
  });
  b.add(G.tube(catmull([[0.3, 0.5, 0.0],[0.36, 0.62, 0.06],[0.42, 0.74, 0.12]], 6, false), 0.012, 10), { c: BRASS, g: 0.95 });
  mouthpiece(b, [0.42, 0.74, 0.12], [0, 0, -30]);
  return b;
} };

/* ---------- ударные ---------- */
function tripod(b, x, z, y0, y1, spread, r, c) {
  [90, 210, 330].forEach(function (a) { var an = a * D2R; rod(b, [x, y1, z], [x + spread*Math.cos(an), y0, z + spread*Math.sin(an)], r, { c: c || K.steel, g: 0.95, seg: 8 }); });
}
MODELS.snare = { title: 'Малый барабан', yaw: -30, pitch: 22, build: function () {
  var b = new Builder(), k;
  b.add(G.cyl(0.17, 0.15, 40), { p: [0, 0.76, 0], c: '#c9ced5', g: 1 });
  b.add(G.cyl(0.168, 0.003, 40), { p: [0, 0.8365, 0], c: '#f3f0e6', g: 0.35 });
  [0.835, 0.685].forEach(function (y) { b.add(G.torus(0.173, 0.009, 40, 8), { p: [0, y, 0], c: K.silver, g: 1 }); });
  for (k = 0; k < 10; k++) {
    var a = 2*PI*k/10, x = 0.176*Math.cos(a), z = 0.176*Math.sin(a);
    b.add(G.box(0.02, 0.17, 0.014), { p: [x, 0.76, z], r: [0, -a/D2R, 0], c: K.silver, g: 1 });
  }
  b.add(G.box(0.05, 0.03, 0.03), { p: [0.19, 0.74, 0], c: K.black, g: 0.6 });
  b.add(G.torus(0.13, 0.008, 32, 8), { p: [0, 0.64, 0], c: K.steel, g: 0.95 });
  rod(b, [0, 0.64, 0], [0, 0.2, 0], 0.013, { c: K.steel, g: 0.95 });
  tripod(b, 0, 0, 0.0, 0.42, 0.36, 0.009);
  [90, 210, 330].forEach(function (a) { var an = a*D2R; rod(b, [0.13*Math.cos(an), 0.64, 0.13*Math.sin(an)], [0, 0.45, 0], 0.007, { c: K.steel, g: 0.95, seg: 6 }); });
  return b;
} };
MODELS.bassdrum = { title: 'Большой барабан', yaw: -25, pitch: 10, build: function () {
  var b = new Builder(), k, R = 0.45, y = 0.72;
  b.push({ p: [0, y, 0], r: [90, 0, 0] });
  b.add(G.cyl(R, 0.38, 48), { c: '#4a2214', g: 0.75, w: 0.3 });
  [-1, 1].forEach(function (s) {
    b.add(G.cyl(R - 0.004, 0.004, 48), { p: [0, s * 0.19, 0], c: '#f3f0e6', g: 0.35 });
    b.add(G.torus(R + 0.005, 0.014, 48, 8), { p: [0, s * 0.19, 0], c: K.silver, g: 1 });
  });
  for (k = 0; k < 14; k++) { var a = 2*PI*k/14; b.add(G.cyl(0.006, 0.42, 8), { p: [(R + 0.03)*Math.cos(a), 0, (R + 0.03)*Math.sin(a)], c: K.silver, g: 1 }); b.add(G.box(0.03, 0.04, 0.03), { p: [(R + 0.03)*Math.cos(a), 0.21, (R + 0.03)*Math.sin(a)], c: K.silver, g: 1 }); }
  b.pop();
  [-1, 1].forEach(function (s) {
    b.add(G.box(0.05, 0.9, 0.05), { p: [s * 0.5, 0.45, 0], c: '#2a1710', g: 0.6 });
    b.add(G.box(0.1, 0.04, 0.5), { p: [s * 0.5, 0.02, 0], c: '#2a1710', g: 0.6 });
  });
  return b;
} };
MODELS.timpani = { title: 'Литавры', yaw: -25, pitch: 20, build: function () {
  var b = new Builder(), k, R = 0.33, y0 = 0.3, pr = [];
  for (k = 0; k <= 24; k++) { var ph = (PI/2) * k/24; pr.push([R*Math.sin(ph), y0 + R*(1 - Math.cos(ph))]); }
  b.add(G.lathe(pr, 48), { c: K.copper, g: 0.95 });
  var yr = y0 + R;
  b.add(G.disc(R + 0.005, 48), { p: [0, yr + 0.004, 0], c: '#f1ede2', g: 0.3 });
  b.add(G.torus(R + 0.012, 0.014, 48, 8), { p: [0, yr + 0.006, 0], c: K.silver, g: 1 });
  for (k = 0; k < 8; k++) {
    var a = 2*PI*k/8, x = (R + 0.04)*Math.cos(a), z = (R + 0.04)*Math.sin(a);
    rod(b, [x, yr + 0.03, z], [x*0.8, yr - 0.14, z*0.8], 0.005, { c: K.silver, g: 1, seg: 6 });
    b.add(G.box(0.03, 0.03, 0.03), { p: [x, yr + 0.02, z], r: [0, -a/D2R, 0], c: K.silver, g: 1 });
  }
  b.add(G.torus(0.2, 0.012, 32, 8), { p: [0, 0.42, 0], c: K.steel, g: 0.95 });
  [90, 210, 330].forEach(function (a) { var an = a*D2R; rod(b, [0.2*Math.cos(an), 0.42, 0.2*Math.sin(an)], [0.42*Math.cos(an), 0.0, 0.42*Math.sin(an)], 0.012, { c: K.steel, g: 0.95, seg: 8 }); b.add(G.cyl(0.03, 0.02, 12), { p: [0.42*Math.cos(an), 0.01, 0.42*Math.sin(an)], c: K.black, g: 0.5 }); });
  b.add(G.box(0.05, 0.02, 0.2), { p: [0, 0.03, 0.32], c: K.steel, g: 0.9 });
  return b;
} };
MODELS.cymbals = { title: 'Тарелки', yaw: -28, pitch: 14, build: function () {
  var b = new Builder();
  var pr = [[0,.045],[.03,.043],[.052,.03],[.065,.02],[.22,.0],[.22,-.003],[.065,.017],[.052,.027],[.03,.04],[0,.042]];
  [[1.06, 4], [0.98, 0]].forEach(function (c) {
    b.add(G.lathe(pr, 48), { p: [0, c[0], 0], r: [c[1], 0, c[1] / 2], c: '#d7ae3b', g: 1 });
    b.add(G.cyl(0.012, 0.03, 12), { p: [0, c[0] + 0.03, 0], c: '#7a1f1f', g: 0.4 });
  });
  rod(b, [0, 0.04, 0], [0, 1.02, 0], 0.011, { c: K.steel, g: 0.95 });
  tripod(b, 0, 0, 0.0, 0.5, 0.35, 0.009);
  b.add(G.cyl(0.03, 0.07, 14), { p: [0, 0.5, 0], c: K.black, g: 0.6 });
  return b;
} };
MODELS.xylo = { title: 'Ксилофон / маримба', yaw: -28, pitch: 24, build: function () {
  var b = new Builder(), n = 24, pitch = 0.05, y = 0.8, x0 = -(n - 1) * pitch / 2, i;
  for (i = 0; i < n; i++) {
    var t = i / (n - 1), len = lerp(0.36, 0.2, t), x = x0 + i * pitch;
    b.add(G.box(pitch * 0.86, 0.022, len), { p: [x, y, 0.06], c: '#8a3d20', g: 0.55, w: 0.5 });
    var rl = lerp(0.62, 0.2, t);
    b.add(G.cyl(0.0185, rl, 16), { p: [x, y - 0.03 - rl / 2, 0.06], c: K.silver, g: 1 });
    if ('CDFGA'.indexOf('CDEFGAB'[i % 7]) >= 0 && i < n - 1) b.add(G.box(pitch * 0.7, 0.022, lerp(0.26, 0.16, t)), { p: [x + pitch / 2, y + 0.03, -0.17], c: '#a0502c', g: 0.55, w: 0.5 });
  }
  var W = n * pitch + 0.1;
  [0.2, -0.2].forEach(function (z) { b.add(G.box(W, 0.04, 0.035), { p: [0, y - 0.025, z + 0.06 * (z > 0 ? 1 : 0) - 0.0], c: K.dark, g: 0.5 }); });
  [-1, 1].forEach(function (s) { [0.24, -0.2].forEach(function (z) { b.add(G.box(0.045, 0.8, 0.045), { p: [s * (W / 2 - 0.02), 0.4, z], c: K.dark, g: 0.55 }); }); b.add(G.box(0.05, 0.04, 0.5), { p: [s * (W / 2 - 0.02), 0.03, 0.02], c: K.dark, g: 0.55 }); });
  return b;
} };

/* ---------- аккордеон / баян ---------- */
MODELS.accordion = { title: 'Аккордеон / баян', yaw: -28, pitch: 12, build: function () {
  var b = new Builder(), i, n = 16;
  b.add(G.box(0.2, 0.4, 0.24), { p: [-0.32, 0.4, 0], c: K.gloss, g: 0.8 });
  b.add(G.box(0.27, 0.44, 0.26), { p: [0.22, 0.4, 0], c: K.gloss, g: 0.8 });
  for (i = 0; i < n; i++) {
    var big = i % 2 === 0, x = -0.22 + (i + 0.5) * (0.32 / n);
    b.add(G.box(0.016, big ? 0.37 : 0.34, big ? 0.215 : 0.19), { p: [x, 0.4, 0], c: big ? '#8b1a22' : '#6a1219', g: 0.4 });
  }
  b.add(G.box(0.34, 0.012, 0.23), { p: [-0.06, 0.585, 0], c: K.silver, g: 0.9 });
  b.add(G.box(0.34, 0.012, 0.23), { p: [-0.06, 0.215, 0], c: K.silver, g: 0.9 });
  b.push({ p: [0.12, 0.4, 0.131], r: [90, 0, 90] });
  keyboard(b, { x: 0, y: 0, z: 0.0, n: 22, start: 'C', w: 0.38, len: 0.15, h: 0.014, bh: 0.01 });
  b.pop();
  b.add(G.box(0.1, 0.4, 0.01), { p: [0.3, 0.4, 0.133], c: '#2c2c33', g: 0.5 });
  var r, c;
  for (r = 0; r < 4; r++) for (c = 0; c < 6; c++) b.add(G.cyl(0.0105, 0.012, 14), { p: [-0.375 + c * 0.0325, 0.3 + r * 0.04, 0.123], r: [90, 0, 0], c: (r + c) % 2 ? K.cream : K.silver, g: 0.7 });
  rod(b, [-0.42, 0.25, 0.0], [-0.42, 0.55, 0.0], 0.008, { c: K.silver, g: 0.9 });
  return b;
} };

/* ---------- орган ---------- */
MODELS.organ = { title: 'Орган', yaw: -25, pitch: 10, build: function () {
  var b = new Builder(), i;
  b.add(G.box(2.0, 0.55, 0.8), { p: [0, 0.275, 0], c: '#3d1e12', g: 0.7, w: 0.4 });
  b.add(G.box(2.0, 2.4, 0.08), { p: [0, 1.5, -0.36], c: '#3d1e12', g: 0.7, w: 0.4 });
  [-1, 1].forEach(function (s) { b.add(G.box(0.08, 2.4, 0.8), { p: [s * 0.96, 1.5, 0], c: '#3d1e12', g: 0.7, w: 0.4 }); });
  b.add(G.box(2.06, 0.06, 0.86), { p: [0, 2.73, 0], c: '#3d1e12', g: 0.7, w: 0.4 });
  function pipe(x, h, c, z) {
    var r = 0.02 + h * 0.022, y0 = 0.56;
    b.add(G.lathe([[0,y0],[0.012,y0],[0.012,y0],[r,y0 + 0.14],[r,y0 + 0.14],[r,y0 + h],[r,y0 + h],[r*0.8,y0 + h]], 20), { p: [x, 0, z], c: c, g: 1 });
    b.add(G.box(r * 1.5, 0.03, 0.01), { p: [x, y0 + 0.2, z + r], c: '#0c0806', g: 0.2 });
  }
  for (i = 0; i < 7; i++) pipe(-0.88 + i * 0.1, 1.0 + i * 0.1, i % 2 ? K.silver : '#d6ad40', 0.24);
  for (i = 0; i < 9; i++) pipe(-0.4 + i * 0.1, 2.0 - Math.abs(i - 4) * 0.2, K.silver, 0.28);
  for (i = 0; i < 7; i++) pipe(0.28 + i * 0.1, 1.6 - i * 0.1, i % 2 ? K.silver : '#d6ad40', 0.24);
  for (i = 0; i < 9; i++) pipe(-0.32 + i * 0.08, 1.3, '#d6ad40', -0.1);
  b.add(G.box(1.2, 0.8, 0.5), { p: [0, 0.4, 0.62], c: '#2a1710', g: 0.7, w: 0.4 });
  b.add(G.box(1.25, 0.08, 0.5), { p: [0, 0.84, 0.62], c: '#2a1710', g: 0.7, w: 0.4 });
  keyboard(b, { x: 0, y: 0.88, z: 0.46, n: 29, start: 'C', w: 1.1, len: 0.2 });
  b.add(G.box(0.8, 0.28, 0.02), { p: [0, 1.06, 0.4], r: [-15, 0, 0], c: '#2a1710', g: 0.7 });
  return b;
} };

})(window);
'''

app.jinja_loader = DictLoader(T)


# ---------------------------------------------------------------- запуск
init_db()


if __name__ == "__main__":
    import sys
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "remind":      # python app.py remind  — для планировщика (cron / Планировщик заданий)
        with app.app_context():
            sent, skipped, err = send_reminders()
            con = sqlite3.connect(DB, timeout=15)
            write_audit(con, None, "system", "Отправка напоминаний (по расписанию)", "notifications", "",
                        f"отправлено: {sent}, без e-mail: {skipped}" + (f"; {err}" if err else ""))
            con.commit()
            con.close()
        print(f"Отправлено: {sent}, без e-mail: {skipped}" + (f"; {err}" if err else ""))
    elif cmd == "backup":    # python app.py backup  — разовая копия (для планировщика)
        print("Создана копия:", make_backup("_auto"))
        prune_backups()
    else:
        debug = os.environ.get("APP_DEBUG", "1") == "1"
        if not debug or os.environ.get("WERKZEUG_RUN_MAIN") == "true":
            start_background()
        app.run(debug=debug)
