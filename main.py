"""Dynamic Island для Windows (PyQt6).

Возможности:
- Музыка: обложка, пауза/вперёд/назад, прогресс трека и перемотка (по клику на стрелку)
- Текст песни (LRCLIB), синхронно подсвечивается строка
- Громкость системы колесом мыши над островом (появляется только при прокрутке)
- Буфер обмена: текст и картинки, без лимита, закрепление, поиск
- Загрузки: скорость сети, график скорости, активность по приложениям, уведомление о завершении
- Таймеры: несколько одновременно
- Alt+Space: лаунчер (программы, калькулятор, конвертер валют/единиц, перевод, погода,
  поиск файлов по индексу Windows, таймер, команды, поиск в буфере)

ПКМ по острову -> автозапуск / выход.
"""
import os
import re
import sys
import ast
import math
import time
import html
import ctypes
import hashlib
import sqlite3
import asyncio
import operator
import datetime
import threading
import subprocess
import shutil
import winreg
import traceback
import importlib.util
import webbrowser
import winsound
import urllib.parse
from collections import deque
from ctypes import wintypes
from pathlib import Path

import psutil

try:
    import requests
    HAVE_REQUESTS = True
except Exception:
    HAVE_REQUESTS = False

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL
    from ctypes import POINTER, cast
    HAVE_VOLUME = True
except Exception:
    HAVE_VOLUME = False

try:
    import win32com.client
    HAVE_WIN32 = True
except Exception:
    HAVE_WIN32 = False

from PyQt6.QtCore import (Qt, QRect, QRectF, QSize, QByteArray, QBuffer, QIODevice,
                          QPropertyAnimation, QEasingCurve, QThread, pyqtSignal, QTimer,
                          QObject, QEvent, QAbstractNativeEventFilter)
from PyQt6.QtGui import (QPainter, QColor, QPixmap, QGuiApplication, QPainterPath, QCursor,
                         QFont, QFontMetrics, QLinearGradient, QPen)
from PyQt6.QtWidgets import (QApplication, QWidget, QLabel, QPushButton, QHBoxLayout, QVBoxLayout,
                             QMenu, QFrame, QLineEdit, QListWidget, QListWidgetItem, QScrollArea,
                             QStackedWidget, QProgressBar, QLayout)

from winrt.windows.media.control import (
    GlobalSystemMediaTransportControlsSessionManager as MediaManager,
    GlobalSystemMediaTransportControlsSessionPlaybackStatus as PlaybackStatus,
)
from winrt.windows.storage.streams import DataReader

# ====================== Настройки ======================
APPS = [  # (иконка Segoe Fluent Icons, что запускать)
    ("\uE8A5", "notepad.exe"),
    ("\uE8EF", "calc.exe"),
    ("\uE8B7", "explorer.exe"),
    ("\uE774", "https://www.google.com"),
    ("\uE713", "ms-settings:"),
]
SCREENSHOT_DIR = Path.home() / "Pictures" / "Screenshots"
DOWNLOAD_DIRS = [Path.home() / "Downloads"]
DATA_DIR = Path.home() / ".dynamic_island"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "data.db"
CLIPS_DIR = DATA_DIR / "clips"
RECORD_DIR = Path.home() / "Videos" / "Island"
PLUGINS_DIR = DATA_DIR / "plugins"
PLUGIN_DATA_DIR = DATA_DIR / "plugin_data"

HOTKEY_MOD = 0x0001 | 0x4000   # MOD_ALT | MOD_NOREPEAT
HOTKEY_VK = 0x20               # Space

W0, H0 = 150, 36     # минимальный свёрнутый размер
W1 = 440             # ширина раскрытого острова
TOP_MARGIN = 8

TEXT_FONT = "Segoe UI Variable Display"
ICON_FONTS = ["Segoe Fluent Icons", "Segoe MDL2 Assets"]

TAB_HOME, TAB_CLIP, TAB_DL, TAB_TIMER, TAB_LYRICS, TAB_PLUGINS = range(6)
TAB_NAMES = ["Главная", "Буфер", "Загрузки", "Таймер", "Текст", "Плагины"]
TAB_H = [245, 440, 400, 380, 320, 360]
PROGRESS_EXTRA_H = 30

PARTIAL_EXT = {".crdownload", ".part", ".opdownload", ".download", ".partial"}

EDIT_QSS = ("QLineEdit{background:rgba(255,255,255,18);color:white;"
            "border:1px solid rgba(255,255,255,30);border-radius:12px;padding:6px 12px;"
            "selection-background-color:rgba(255,255,255,80);}"
            "QLineEdit:focus{border:1px solid rgba(255,255,255,90);}")
SCROLL_QSS = ("QScrollArea{background:transparent;border:none;}"
              "QScrollBar:vertical{background:transparent;width:6px;margin:0;}"
              "QScrollBar::handle:vertical{background:rgba(255,255,255,50);border-radius:3px;min-height:24px;}"
              "QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{height:0;}"
              "QScrollBar::add-page:vertical,QScrollBar::sub-page:vertical{background:transparent;}")


# ====================== Утилиты ======================
def fmt_speed(bps):
    if bps < 1024:
        return f"{bps:.0f} Б/с"
    if bps < 1024 ** 2:
        return f"{bps / 1024:.0f} КБ/с"
    return f"{bps / 1024 ** 2:.1f} МБ/с"


def fmt_size(b):
    if b < 1024:
        return f"{b:.0f} Б"
    if b < 1024 ** 2:
        return f"{b / 1024:.0f} КБ"
    if b < 1024 ** 3:
        return f"{b / 1024 ** 2:.1f} МБ"
    return f"{b / 1024 ** 3:.2f} ГБ"


def fmt_clock(sec):
    sec = int(max(0, math.ceil(sec)))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h}:{m:02}:{s:02}" if h else f"{m:02}:{s:02}"


UNITS = {
    "h": 3600, "hr": 3600, "hour": 3600, "hours": 3600, "ч": 3600, "час": 3600, "часа": 3600, "часов": 3600,
    "m": 60, "min": 60, "minute": 60, "minutes": 60, "м": 60, "мин": 60, "минута": 60, "минуты": 60, "минут": 60,
    "s": 1, "sec": 1, "с": 1, "сек": 1, "секунда": 1, "секунды": 1, "секунд": 1,
}
_TOKEN = r"(\d+(?:\.\d+)?)\s*([a-zа-яё]*)"


def parse_duration(text):
    """'5', '5m', '1ч 30м', '90с', '25:00', '1:30:00' -> секунды (или None)."""
    t = text.strip().lower().replace(",", ".")
    if not t:
        return None
    if re.fullmatch(r"\d+(:\d+){1,2}", t):
        parts = [int(x) for x in t.split(":")]
        total = parts[0] * 60 + parts[1] if len(parts) == 2 else parts[0] * 3600 + parts[1] * 60 + parts[2]
        return total or None
    tokens = re.findall(_TOKEN, t)
    if not tokens or re.sub(_TOKEN + r"|\s", "", t):
        return None
    total = 0.0
    for num, unit in tokens:
        mult = 60 if unit == "" else UNITS.get(unit)
        if mult is None:
            return None
        total += float(num) * mult
    return int(total) if total > 0 else None


_OPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv,
        ast.Mod: operator.mod, ast.Pow: operator.pow, ast.FloorDiv: operator.floordiv}


def _ev(n):
    if isinstance(n, ast.Expression):
        return _ev(n.body)
    if isinstance(n, ast.Constant) and isinstance(n.value, (int, float)):
        return n.value
    if isinstance(n, ast.UnaryOp) and isinstance(n.op, (ast.USub, ast.UAdd)):
        v = _ev(n.operand)
        return -v if isinstance(n.op, ast.USub) else v
    if isinstance(n, ast.BinOp) and type(n.op) in _OPS:
        a, b = _ev(n.left), _ev(n.right)
        if isinstance(n.op, ast.Pow) and abs(b) > 1000:
            raise ValueError
        return _OPS[type(n.op)](a, b)
    raise ValueError


def calc(q):
    s = q.strip().replace(",", ".").replace("^", "**").replace("×", "*").replace("÷", "/")
    if not re.fullmatch(r"[\d\s+\-*/().%]+", s) or not re.search(r"\d", s):
        return None
    if not re.search(r"[+\-*/%]", s.lstrip("+- ")):
        return None
    try:
        v = _ev(ast.parse(s, mode="eval"))
        if isinstance(v, float):
            if not math.isfinite(v):
                return None
            v = round(v, 10)
            if v == int(v):
                v = int(v)
        return str(v)
    except Exception:
        return None


# --- конвертер валют и единиц измерения ---
CURRENCY_ALIASES = {
    "руб": "RUB", "рубль": "RUB", "рублей": "RUB", "рубля": "RUB", "р": "RUB", "rub": "RUB",
    "доллар": "USD", "долларов": "USD", "доллара": "USD", "бакс": "USD", "баксов": "USD", "usd": "USD", "$": "USD",
    "евро": "EUR", "eur": "EUR", "€": "EUR",
    "фунт": "GBP", "фунтов": "GBP", "gbp": "GBP", "£": "GBP",
    "юань": "CNY", "юаней": "CNY", "cny": "CNY", "yuan": "CNY",
    "иена": "JPY", "йена": "JPY", "иен": "JPY", "jpy": "JPY",
    "франк": "CHF", "франков": "CHF", "chf": "CHF",
    "злотый": "PLN", "злотых": "PLN", "pln": "PLN",
    "тенге": "KZT", "kzt": "KZT",
    "гривна": "UAH", "гривен": "UAH", "гривны": "UAH", "uah": "UAH",
    "лира": "TRY", "лир": "TRY", "try": "TRY",
}
UNIT_GROUPS = {
    "length": {"m": 1, "м": 1, "метр": 1, "метров": 1, "km": 1000, "км": 1000, "километр": 1000, "километров": 1000,
              "cm": 0.01, "см": 0.01, "mm": 0.001, "мм": 0.001,
              "mi": 1609.344, "миля": 1609.344, "миль": 1609.344, "мили": 1609.344,
              "yd": 0.9144, "ft": 0.3048, "фут": 0.3048, "футов": 0.3048,
              "in": 0.0254, "дюйм": 0.0254, "дюймов": 0.0254},
    "weight": {"kg": 1, "кг": 1, "килограмм": 1, "килограммов": 1, "g": 0.001, "г": 0.001, "грамм": 0.001,
              "граммов": 0.001, "lb": 0.453592, "фунт": 0.453592, "фунтов": 0.453592,
              "oz": 0.0283495, "унция": 0.0283495, "унций": 0.0283495, "t": 1000, "тонна": 1000, "тонн": 1000},
    "volume": {"l": 1, "л": 1, "литр": 1, "литров": 1, "ml": 0.001, "мл": 0.001,
              "gal": 3.78541, "галлон": 3.78541, "галлонов": 3.78541},
}


def _norm_unit(u):
    return u.strip().lower().rstrip(".")


def convert_units(amount, u1, u2):
    a, b = _norm_unit(u1), _norm_unit(u2)
    if a in ("c", "°c", "celsius", "цельсий") and b in ("f", "°f", "fahrenheit", "фаренгейт"):
        return amount * 9 / 5 + 32
    if a in ("f", "°f", "fahrenheit", "фаренгейт") and b in ("c", "°c", "celsius", "цельсий"):
        return (amount - 32) * 5 / 9
    for group in UNIT_GROUPS.values():
        if a in group and b in group:
            return amount * group[a] / group[b]
    return None


def search_windows_index(query, limit=8):
    """Поиск файлов по индексу Windows Search (без обхода диска вручную)."""
    if not HAVE_WIN32:
        return None
    try:
        conn = win32com.client.Dispatch("ADODB.Connection")
        conn.Open("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';")
        safe = query.replace("'", "''")
        sql = ("SELECT TOP %d System.ItemPathDisplay, System.ItemNameDisplay FROM SystemIndex "
              "WHERE System.FileName LIKE '%%%s%%' ORDER BY System.DateModified DESC" % (limit, safe))
        rs = conn.Execute(sql)[0]
        out = []
        while not rs.EOF:
            try:
                path = rs.Fields.Item("System.ItemPathDisplay").Value
                name = rs.Fields.Item("System.ItemNameDisplay").Value
                if path:
                    out.append((name or Path(path).name, path))
            except Exception:
                pass
            rs.MoveNext()
        rs.Close()
        conn.Close()
        return out
    except Exception:
        return None


def midnight_ts():
    return datetime.datetime.combine(datetime.date.today(), datetime.time.min).timestamp()


def draw_surface(p, w, h, r):
    """Чёрная «премиальная» поверхность: основа + блик + тонкая рамка."""
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(QRectF(0.5, 0.5, w - 1, h - 1), r, r)
    p.fillPath(path, QColor(0, 0, 0))
    grad = QLinearGradient(0, 0, 0, h)
    grad.setColorAt(0.0, QColor(255, 255, 255, 20))
    grad.setColorAt(0.5, QColor(255, 255, 255, 0))
    p.fillPath(path, grad)
    border = QLinearGradient(0, 0, 0, h)
    border.setColorAt(0.0, QColor(255, 255, 255, 70))
    border.setColorAt(1.0, QColor(255, 255, 255, 22))
    p.setPen(QPen(border, 1))
    p.drawPath(path)


def icon_font(px):
    f = QFont()
    f.setFamilies(ICON_FONTS)
    f.setPixelSize(px)
    return f


def clear_layout(layout):
    while layout.count():
        item = layout.takeAt(0)
        w = item.widget()
        if w:
            w.deleteLater()


def rounded_pixmap(pm, size, radius):
    pm = pm.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                   Qt.TransformationMode.SmoothTransformation)
    out = QPixmap(size, size)
    out.fill(Qt.GlobalColor.transparent)
    p = QPainter(out)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addRoundedRect(0, 0, size, size, radius, radius)
    p.setClipPath(path)
    p.drawPixmap((size - pm.width()) // 2, (size - pm.height()) // 2, pm)
    p.end()
    return out


def autostart_enabled():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run") as k:
            winreg.QueryValueEx(k, "DynamicIsland")
            return True
    except Exception:
        return False


def set_autostart(on):
    import winreg
    exe = Path(sys.executable)
    pyw = exe.with_name("pythonw.exe")
    runner = pyw if pyw.exists() else exe
    cmd = f'"{runner}" "{Path(__file__).resolve()}"'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run",
                        0, winreg.KEY_SET_VALUE) as k:
        if on:
            winreg.SetValueEx(k, "DynamicIsland", 0, winreg.REG_SZ, cmd)
        else:
            try:
                winreg.DeleteValue(k, "DynamicIsland")
            except FileNotFoundError:
                pass


# --- громкость системы ---
def _volume_iface():
    device = AudioUtilities.GetSpeakers()
    iface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))


def change_volume(delta_pct):
    if not HAVE_VOLUME:
        return None
    try:
        v = _volume_iface()
        cur = v.GetMasterVolumeLevelScalar()
        new = min(1.0, max(0.0, cur + delta_pct / 100))
        v.SetMasterVolumeLevelScalar(new, None)
        return round(new * 100)
    except Exception:
        return None


# ====================== Хранилище ======================
class Store:
    def __init__(self):
        self.db = sqlite3.connect(DB_PATH)
        self.db.create_function("lw", 1, lambda s: (s or "").lower())
        self.db.execute("CREATE TABLE IF NOT EXISTS clips("
                        "id INTEGER PRIMARY KEY, text TEXT UNIQUE NOT NULL, "
                        "ts REAL NOT NULL, pinned INTEGER NOT NULL DEFAULT 0)")
        for col, decl in (("kind", "TEXT NOT NULL DEFAULT 'text'"), ("path", "TEXT"), ("label", "TEXT")):
            try:
                self.db.execute(f"ALTER TABLE clips ADD COLUMN {col} {decl}")
            except sqlite3.OperationalError:
                pass
        self.db.execute("CREATE TABLE IF NOT EXISTS launches(name TEXT PRIMARY KEY, n INTEGER NOT NULL DEFAULT 0)")
        self.db.commit()

    def add_clip(self, text):
        self.db.execute("INSERT INTO clips(text, ts, kind) VALUES(?, ?, 'text') "
                        "ON CONFLICT(text) DO UPDATE SET ts = excluded.ts", (text, time.time()))
        self.db.commit()

    def add_image_clip(self, path, label):
        key = f"\u0000img:{path}"
        self.db.execute("INSERT INTO clips(text, ts, kind, path, label) VALUES(?, ?, 'image', ?, ?) "
                        "ON CONFLICT(text) DO UPDATE SET ts = excluded.ts", (key, time.time(), path, label))
        self.db.commit()

    def clips(self, query="", today=True, limit=100):
        sql = "SELECT id, text, ts, pinned, kind, path, label FROM clips WHERE 1=1"
        args = []
        if today:
            sql += " AND (ts >= ? OR pinned = 1)"
            args.append(midnight_ts())
        if query:
            sql += " AND instr(lw(COALESCE(label, text)), ?) > 0"
            args.append(query.lower())
        sql += " ORDER BY pinned DESC, ts DESC LIMIT ?"
        args.append(limit)
        return self.db.execute(sql, args).fetchall()

    def clip_total(self):
        return self.db.execute("SELECT COUNT(*) FROM clips").fetchone()[0]

    def set_pin(self, cid, pinned):
        self.db.execute("UPDATE clips SET pinned = ? WHERE id = ?", (1 if pinned else 0, cid))
        self.db.commit()

    def delete_clip(self, cid):
        self.db.execute("DELETE FROM clips WHERE id = ?", (cid,))
        self.db.commit()

    def bump_launch(self, name):
        self.db.execute("INSERT INTO launches(name, n) VALUES(?, 1) "
                        "ON CONFLICT(name) DO UPDATE SET n = n + 1", (name,))
        self.db.commit()

    def launch_counts(self):
        return dict(self.db.execute("SELECT name, n FROM launches").fetchall())


# ====================== Медиа ======================
def _sec(v):
    try:
        return v.total_seconds()
    except AttributeError:
        return v / 10_000_000


async def fetch_state():
    mgr = await MediaManager.request_async()
    session = mgr.get_current_session()
    if not session:
        return None
    props = await session.try_get_media_properties_async()
    info = session.get_playback_info()
    thumb = None
    try:
        if props.thumbnail:
            stream = await props.thumbnail.open_read_async()
            size = stream.size
            if size:
                reader = DataReader(stream)
                await reader.load_async(size)
                buf = bytearray(size)
                reader.read_bytes(buf)
                thumb = bytes(buf)
    except Exception:
        thumb = None
    pos = dur = 0.0
    try:
        tl = session.get_timeline_properties()
        pos = _sec(tl.position)
        dur = max(0.0, _sec(tl.end_time) - _sec(tl.start_time))
    except Exception:
        pass
    return {"title": props.title or "", "artist": props.artist or "",
            "playing": info.playback_status == PlaybackStatus.PLAYING, "thumb": thumb,
            "position": pos, "duration": dur}


class MediaThread(QThread):
    state = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.loop = None

    def run(self):
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self._poll())

    async def _poll(self):
        while True:
            try:
                st = await fetch_state()
            except Exception:
                st = None
            self.state.emit(st)
            await asyncio.sleep(1)

    def command(self, name):
        if self.loop:
            asyncio.run_coroutine_threadsafe(self._cmd(name), self.loop)

    async def _cmd(self, name):
        try:
            mgr = await MediaManager.request_async()
            s = mgr.get_current_session()
            if not s:
                return
            if name == "toggle":
                await s.try_toggle_play_pause_async()
            elif name == "next":
                await s.try_skip_next_async()
            elif name == "prev":
                await s.try_skip_previous_async()
        except Exception:
            pass

    def seek(self, seconds):
        if self.loop:
            asyncio.run_coroutine_threadsafe(self._seek(seconds), self.loop)

    async def _seek(self, seconds):
        try:
            mgr = await MediaManager.request_async()
            s = mgr.get_current_session()
            if not s:
                return
            ticks = int(max(0, seconds) * 10_000_000)
            await s.try_change_playback_position_async(ticks)
        except Exception:
            pass


class LyricsThread(QThread):
    result = pyqtSignal(str, object)  # key, [(сек, строка)] | None

    def __init__(self, key, artist, title):
        super().__init__()
        self.key, self.artist, self.title = key, artist, title

    def run(self):
        lines = None
        if HAVE_REQUESTS:
            try:
                r = requests.get("https://lrclib.net/api/get",
                                 params={"track_name": self.title, "artist_name": self.artist}, timeout=6)
                if r.ok:
                    data = r.json()
                    synced = data.get("syncedLyrics")
                    if synced:
                        lines = []
                        for m in re.finditer(r"\[(\d+):(\d+(?:\.\d+)?)\]([^\[]*)", synced):
                            t = int(m.group(1)) * 60 + float(m.group(2))
                            lines.append((t, m.group(3).strip()))
                    elif data.get("plainLyrics"):
                        lines = [(0.0, ln) for ln in data["plainLyrics"].splitlines() if ln.strip()]
            except Exception:
                lines = None
        self.result.emit(self.key, lines)


# ====================== Загрузки ======================
class DownloadMonitor(QObject):
    changed = pyqtSignal()
    file_done = pyqtSignal(str, str)  # имя, папка

    def __init__(self):
        super().__init__()
        self.dirs = [d for d in DOWNLOAD_DIRS if d.exists()]
        self.files = {}       # path -> {"name","size","speed"}
        self.recent = []      # (name, size, ts)
        self.history = deque(maxlen=60)
        self.by_app = []
        self.by_app_error = False
        self.net_down = 0.0
        self.net_up = 0.0
        self._net = psutil.net_io_counters()
        self._t = time.time()
        self._poll_n = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.poll)
        self.timer.start(1000)

    @property
    def speed(self):
        return sum(f["speed"] for f in self.files.values())

    def poll(self):
        now = time.time()
        dt = max(now - self._t, 0.001)
        self._t = now
        net = psutil.net_io_counters()
        self.net_down = max(0.0, (net.bytes_recv - self._net.bytes_recv) / dt)
        self.net_up = max(0.0, (net.bytes_sent - self._net.bytes_sent) / dt)
        self._net = net
        self.history.append(self.net_down)
        self._poll_n += 1
        if self._poll_n % 3 == 0:
            self._update_by_app()

        seen = set()
        for d in self.dirs:
            try:
                for e in os.scandir(d):
                    if not e.is_file() or Path(e.name).suffix.lower() not in PARTIAL_EXT:
                        continue
                    size = e.stat().st_size
                    seen.add(e.path)
                    f = self.files.get(e.path)
                    if f is None:
                        self.files[e.path] = {"name": Path(e.name).stem, "size": size, "speed": 0.0}
                    else:
                        inst = max(0.0, (size - f["size"]) / dt)
                        f["speed"] = inst if f["speed"] == 0 else 0.5 * f["speed"] + 0.5 * inst
                        if inst == 0:
                            f["speed"] *= 0.5
                            if f["speed"] < 1:
                                f["speed"] = 0.0
                        f["size"] = size
            except OSError:
                pass
        for path in list(self.files):
            if path not in seen:
                f = self.files.pop(path)
                if f["size"] > 1024 * 200:
                    self.recent.insert(0, (f["name"], f["size"], now))
                    del self.recent[6:]
                    self.file_done.emit(f["name"], os.path.dirname(path))
        self.changed.emit()

    def _update_by_app(self):
        try:
            counts = {}
            for c in psutil.net_connections(kind="inet"):
                if c.pid and c.status == psutil.CONN_ESTABLISHED:
                    counts[c.pid] = counts.get(c.pid, 0) + 1
            agg = {}
            for pid, n in counts.items():
                try:
                    name = psutil.Process(pid).name()
                except Exception:
                    continue
                agg[name] = agg.get(name, 0) + n
            self.by_app = sorted(agg.items(), key=lambda x: -x[1])[:5]
            self.by_app_error = False
        except Exception:
            self.by_app = []
            self.by_app_error = True


# ====================== Таймеры ======================
class TimerManager(QObject):
    changed = pyqtSignal()
    finished = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self.items = []
        self._id = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(250)

    def add(self, seconds, label=""):
        self._id += 1
        self.items.append({"id": self._id, "total": seconds, "end": time.time() + seconds,
                           "label": label, "done": False})
        self.changed.emit()

    def remove(self, tid):
        self.items = [i for i in self.items if i["id"] != tid]
        self.changed.emit()

    def nearest(self):
        rem = [max(0.0, i["end"] - time.time()) for i in self.items if not i["done"]]
        return min(rem) if rem else None

    def _tick(self):
        now = time.time()
        fired = []
        for i in self.items:
            if not i["done"] and now >= i["end"]:
                i["done"] = True
                fired.append(i)
        for i in fired:
            self.finished.emit(i)
        if fired or any(not i["done"] for i in self.items):
            self.changed.emit()


# ====================== Строки списков ======================
class ClipRow(QFrame):
    def __init__(self, island, cid, text, ts, pinned, kind, path, label):
        super().__init__()
        self.island, self.cid, self.text, self.pinned = island, cid, text, pinned
        self.kind, self.path = kind, path
        self.setObjectName("row")
        self.setFixedHeight(40)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 0, 6, 0)
        lay.setSpacing(8)
        self.setStyleSheet(f"#row{{background:rgba(255,255,255,{22 if pinned else 11});border-radius:12px;}}"
                           "#row:hover{background:rgba(255,255,255,30);}")

        if kind == "image" and path and Path(path).exists():
            thumb = QLabel()
            thumb.setPixmap(rounded_pixmap(QPixmap(path), 30, 8))
            thumb.setFixedSize(30, 30)
            lay.addWidget(thumb)
            shown = label or "Изображение"
            self.setToolTip(shown)
        else:
            shown = " ".join(text.split())[:200]
            self.setToolTip(shown[:600])
        lab = QLabel()
        lab.setFont(QFont(TEXT_FONT, 10))
        lab.setText(QFontMetrics(lab.font()).elidedText(shown, Qt.TextElideMode.ElideRight, 225))
        lab.setFixedWidth(230)
        lay.addWidget(lab)
        lay.addStretch()
        t = QLabel(datetime.datetime.fromtimestamp(ts).strftime("%H:%M"))
        t.setFont(QFont(TEXT_FONT, 8))
        t.setStyleSheet("color: rgba(255,255,255,110); background: transparent;")
        lay.addWidget(t)
        pin = island._btn("\uE842" if pinned else "\uE718", self._pin, 26, 12, soft=pinned)
        pin.setToolTip("Открепить" if pinned else "Закрепить")
        lay.addWidget(pin)
        dele = island._btn("\uE74D", self._delete, 26, 12)
        dele.setToolTip("Удалить")
        lay.addWidget(dele)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            if self.kind == "image":
                self.island.copy_image(self.path)
            else:
                self.island.copy_text(self.text)

    def _pin(self):
        self.island.toggle_pin(self.cid, not self.pinned)

    def _delete(self):
        self.island.delete_clip(self.cid)


class TimerRow(QFrame):
    def __init__(self, island, item):
        super().__init__()
        self.item = item
        self.setObjectName("trow")
        self.setFixedHeight(78)
        self.setStyleSheet("#trow{background:rgba(255,255,255,12);border-radius:14px;}")
        v = QVBoxLayout(self)
        v.setContentsMargins(14, 8, 8, 10)
        v.setSpacing(4)
        top = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(0)
        self.clock = QLabel()
        self.clock.setFont(QFont(TEXT_FONT, 16, QFont.Weight.DemiBold))
        self.sub = QLabel()
        self.sub.setFont(QFont(TEXT_FONT, 8))
        self.sub.setStyleSheet("color: rgba(255,255,255,120); background: transparent;")
        left.addWidget(self.clock)
        left.addWidget(self.sub)
        top.addLayout(left, 1)
        top.addWidget(island._btn("\uE711", lambda: island.timers.remove(item["id"]), 28, 12, soft=True))
        v.addLayout(top)
        self.bar = QProgressBar()
        self.bar.setRange(0, 1000)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(4)
        self.bar.setStyleSheet("QProgressBar{background:rgba(255,255,255,18);border:none;border-radius:2px;}"
                               "QProgressBar::chunk{background:white;border-radius:2px;}")
        v.addWidget(self.bar)
        self.refresh()

    def refresh(self):
        it = self.item
        if it["done"]:
            self.clock.setText("Время вышло")
            self.bar.setValue(0)
        else:
            rem = max(0.0, it["end"] - time.time())
            self.clock.setText(fmt_clock(rem))
            self.bar.setValue(int(1000 * rem / it["total"]))
        self.sub.setText("из " + fmt_clock(it["total"]) + (f" · {it['label']}" if it["label"] else ""))


class SeekBar(QWidget):
    """Полоса прогресса трека: перетащи или кликни, чтобы перемотать."""

    def __init__(self, island):
        super().__init__()
        self.island = island
        self.setFixedHeight(28)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        w, h = self.width(), self.height()
        isl = self.island
        dur = max(isl.media_dur, 1)
        pos = isl.current_position()
        frac = min(1.0, pos / dur)
        y = h / 2 - 2 - 6
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor(255, 255, 255, 40))
        p.drawRoundedRect(QRectF(0, y, w, 4), 2, 2)
        p.setBrush(QColor(255, 255, 255, 230))
        p.drawRoundedRect(QRectF(0, y, w * frac, 4), 2, 2)
        p.setPen(QColor(255, 255, 255, 170))
        p.setFont(QFont(TEXT_FONT, 8))
        p.drawText(QRectF(0, h - 14, 70, 14), Qt.AlignmentFlag.AlignLeft, fmt_clock(pos))
        p.drawText(QRectF(w - 70, h - 14, 70, 14), Qt.AlignmentFlag.AlignRight, fmt_clock(dur))

    def mousePressEvent(self, e):
        self._seek_to(e.position().x())

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.MouseButton.LeftButton:
            self._seek_to(e.position().x())

    def _seek_to(self, x):
        frac = min(1, max(0, x / max(self.width(), 1)))
        seconds = frac * max(self.island.media_dur, 1)
        self.island.media_pos = seconds
        self.island.media_pos_ts = time.time()
        self.update()
        self.island.media.seek(seconds)


class SparkLine(QWidget):
    """Мини-график скорости загрузки за последнюю минуту."""

    def __init__(self, history_getter):
        super().__init__()
        self.get_history = history_getter
        self.setFixedHeight(40)

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        data = list(self.get_history())
        w, h = self.width(), self.height()
        if len(data) < 2:
            return
        mx = max(max(data), 1024)
        step = w / (len(data) - 1)
        path = QPainterPath()
        for i, v in enumerate(data):
            x = i * step
            y = h - (v / mx) * (h - 6) - 3
            path.moveTo(x, y) if i == 0 else path.lineTo(x, y)
        p.setPen(QPen(QColor(255, 255, 255, 210), 1.6))
        p.drawPath(path)


class Toast(QWidget):
    """Всплывающее уведомление о завершённой загрузке."""

    def __init__(self):
        super().__init__(None, Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(300, 76)
        self.folder = None
        self.setStyleSheet("QLabel{background:transparent;}")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 10, 16, 10)
        self.t1 = QLabel("Загрузка завершена")
        self.t1.setFont(QFont(TEXT_FONT, 10, QFont.Weight.DemiBold))
        self.t1.setStyleSheet("color:white;")
        self.t2 = QLabel("")
        self.t2.setFont(QFont(TEXT_FONT, 9))
        self.t2.setStyleSheet("color: rgba(255,255,255,150);")
        self.t3 = QLabel("Нажмите, чтобы открыть папку")
        self.t3.setFont(QFont(TEXT_FONT, 8))
        self.t3.setStyleSheet("color: rgba(255,255,255,110);")
        lay.addWidget(self.t1)
        lay.addWidget(self.t2)
        lay.addWidget(self.t3)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.hide_timer = QTimer(self)
        self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.hide)

    def paintEvent(self, e):
        p = QPainter(self)
        draw_surface(p, self.width(), self.height(), 20)

    def show_toast(self, name, folder):
        self.folder = folder
        self.t2.setText(name if len(name) < 40 else name[:37] + "…")
        sg = QGuiApplication.primaryScreen().availableGeometry()
        self.move(sg.right() - self.width() - 16, sg.bottom() - self.height() - 16)
        self.show()
        self.raise_()
        self.hide_timer.start(7000)

    def mousePressEvent(self, e):
        self.hide()
        if self.folder:
            try:
                os.startfile(self.folder)
            except Exception:
                pass


# ====================== Остров ======================
def weather_emoji(code, is_day=True):
    """WMO weather code (Open-Meteo) -> эмодзи."""
    if code == 0:
        return "☀️" if is_day else "🌙"
    if code in (1, 2):
        return "🌤️" if is_day else "☁️"
    if code == 3:
        return "☁️"
    if code in (45, 48):
        return "🌫️"
    if code in (51, 53, 55, 56, 57):
        return "🌦️"
    if code in (61, 63, 65, 66, 67, 80, 81, 82):
        return "🌧️"
    if code in (71, 73, 75, 77, 85, 86):
        return "❄️"
    if code in (95, 96, 99):
        return "⛈️"
    return "🌡️"


class IslandWeatherThread(QThread):
    """Погода для свёрнутого острова: место по IP, данные из Open-Meteo."""
    done = pyqtSignal(object, object)  # (lat, lon) или None, (температура, код, день) или None

    def __init__(self, loc):
        super().__init__()
        self.loc = loc

    def run(self):
        loc, res = self.loc, None
        try:
            if not loc:
                r = requests.get("https://ipwho.is/", timeout=6).json()
                if "latitude" in r and "longitude" in r:
                    loc = (r["latitude"], r["longitude"])
            if loc:
                r = requests.get("https://api.open-meteo.com/v1/forecast",
                                 params={"latitude": loc[0], "longitude": loc[1],
                                         "current": "temperature_2m,weather_code,is_day"}, timeout=6)
                if r.ok:
                    c = r.json()["current"]
                    res = (round(c["temperature_2m"]), int(c["weather_code"]), bool(c["is_day"]))
        except Exception:
            res = None
        self.done.emit(loc, res)


# ====================== Микрофон / камера (индикатор приватности) ======================
_CONSENT = r"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore"


def _reg_subkeys(path):
    names = []
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
            i = 0
            while True:
                try:
                    names.append(winreg.EnumKey(k, i))
                    i += 1
                except OSError:
                    break
    except OSError:
        pass
    return names


def _reg_qword(k, name):
    try:
        return int(winreg.QueryValueEx(k, name)[0])
    except OSError:
        return 0


def consent_active(kind):
    """Приложения, которые прямо сейчас используют 'microphone' или 'webcam'
    (Windows пишет это в реестр: начало использования позже, чем конец)."""
    base = _CONSENT + "\\" + kind
    found = []

    def check(path, label):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as k:
                start = _reg_qword(k, "LastUsedTimeStart")
                stop = _reg_qword(k, "LastUsedTimeStop")
        except OSError:
            return
        if start and start > stop and label not in found:
            found.append(label)

    for name in _reg_subkeys(base):
        if name == "NonPackaged":
            for sub in _reg_subkeys(base + "\\NonPackaged"):
                label = sub.split("#")[-1]
                if label.lower().endswith(".exe"):
                    label = label[:-4]
                check(base + "\\NonPackaged\\" + sub, label)
        else:
            check(base + "\\" + name, name.split("_")[0].split(".")[-1])
    return found


# ====================== Запись экрана ======================
def find_ffmpeg():
    """ffmpeg.exe рядом с main.py, в ~/.dynamic_island или в PATH."""
    here = Path(__file__).resolve().parent
    for c in (here / "ffmpeg.exe", DATA_DIR / "ffmpeg.exe"):
        if c.exists():
            return str(c)
    return shutil.which("ffmpeg")


# ====================== Плагины ======================
EXAMPLE_PLUGIN = r'''"""Пример плагина для Dynamic Island.

Как включить: переименуйте файл в example.py (без «_» в начале) и нажмите
«Перезагрузить» на вкладке «Плагины» (или ПКМ по острову -> Перезагрузить плагины).

Плагин - это один .py файл (или папка с файлом plugin.py) в папке плагинов.
Обязательна только функция setup(api), всё остальное по желанию.
"""
import datetime

NAME = "Пример: кофе"


def setup(api):
    state = {"n": 0}

    # 1) Короткий текст в свёрнутом острове. None или "" - ничего не показывать.
    #    Функция вызывается раз в секунду, поэтому она должна быть быстрой.
    #    idle_only=True - показывать только когда не играет музыка.
    api.compact(lambda: f"☕ {state['n']}" if state["n"] else None)

    # 2) Команда в лаунчере (Alt+Space): набираете «кофе» и Enter.
    def coffee():
        state["n"] += 1
        api.flash(f"Кофе выпито: {state['n']}")

    api.command(["кофе", "coffee"], "Выпить кофе", coffee, glyph="\uE74C")

    # 3) Карточка на вкладке «Плагины»: строки текста и кнопки.
    def lines():
        return [f"Выпито: {state['n']}", "Сейчас " + datetime.datetime.now().strftime("%H:%M")]

    api.panel(lines, buttons=[("Сбросить", lambda: state.update(n=0))])

    # 4) Повторяющаяся задача (выполняется в главном потоке, не блокируйте её).
    api.every(3600, lambda: api.flash("Пора размяться"), now=False)

    # Ещё доступно:
    #   api.search(fn)           fn(запрос) -> [(глиф, заголовок, подпись, действие), ...] для лаунчера
    #   api.background(f, cb)    f() выполняется в фоне (сеть, диск), cb(результат) - в главном потоке
    #   api.after(сек, fn)       один раз через N секунд
    #   api.copy(текст)          скопировать в буфер обмена
    #   api.open(путь_или_ссылка)
    #   api.data_dir             своя папка для файлов плагина
'''


class Plugin:
    def __init__(self, pid, path):
        self.id, self.path, self.name = pid, path, pid
        self.module = None
        self.error = ""
        self.load_failed = False
        self.compact = []      # [(fn, idle_only)]
        self.cache = []        # [(text, idle_only)]
        self.commands = []     # [(keys, title, glyph, sub, action)]
        self.searches = []
        self.panel_fn = None
        self.buttons = []
        self.input = None      # (подсказка, функция)
        self.markdown = False
        self.timers = []


class PluginAPI:
    """Объект, который плагин получает в setup(api)."""

    def __init__(self, mgr, plugin):
        self._m, self._p = mgr, plugin
        self.name = plugin.name
        self.data_dir = PLUGIN_DATA_DIR / plugin.id
        try:
            self.data_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    def compact(self, fn, idle_only=False):
        self._p.compact.append((fn, bool(idle_only)))

    def command(self, keys, title, action, glyph="\uE74C", sub=None):
        if isinstance(keys, str):
            keys = [keys]
        self._p.commands.append(([str(k).lower() for k in keys], str(title), glyph,
                                 sub or f"Плагин · {self._p.name}", action))

    def search(self, fn):
        self._p.searches.append(fn)

    def panel(self, lines_fn=None, buttons=None, input=None, markdown=False):
        self._p.panel_fn = lines_fn
        self._p.buttons = list(buttons or [])
        self._p.input = input
        self._p.markdown = bool(markdown)
        self._m.changed.emit()

    def show(self, seconds=None):
        self._m.island.show_plugins(seconds)

    def every(self, seconds, fn, now=True):
        t = QTimer(self._m)
        t.timeout.connect(lambda: self._m.safe(self._p, fn))
        t.start(int(max(0.2, float(seconds)) * 1000))
        self._p.timers.append(t)
        if now:
            QTimer.singleShot(0, lambda: self._m.safe(self._p, fn))

    def after(self, seconds, fn):
        QTimer.singleShot(int(max(0.0, float(seconds)) * 1000), lambda: self._m.safe(self._p, fn))

    def background(self, fn, on_done=None):
        m, p = self._m, self._p

        def work():
            try:
                res, err = fn(), None
            except Exception as e:
                res, err = None, e
            m.bg_done.emit(p, on_done, res, err)

        threading.Thread(target=work, daemon=True).start()

    def flash(self, text, ms=1600):
        self._m.island.flash_msg(str(text), ms)

    def copy(self, text):
        self._m.island.copy_text(str(text))

    def open(self, target):
        self._m.island.launch(target)


class PluginManager(QObject):
    changed = pyqtSignal()
    bg_done = pyqtSignal(object, object, object, object)

    def __init__(self, island):
        super().__init__()
        self.island = island
        self.plugins = []
        self.bg_done.connect(self._on_bg_done)

    # ---------- служебное ----------
    def _log(self, pl, text):
        try:
            with open(DATA_DIR / "plugins.log", "a", encoding="utf-8") as f:
                f.write(f"[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}] {pl.id}\n{text}\n")
        except Exception:
            pass

    def safe(self, pl, fn, *args):
        try:
            return fn(*args)
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            if pl.error != msg:
                pl.error = msg
                self._log(pl, traceback.format_exc())
            return None

    def _on_bg_done(self, pl, cb, res, err):
        if err is not None:
            msg = f"{type(err).__name__}: {err}"
            if pl.error != msg:
                pl.error = msg
                self._log(pl, msg)
            return
        if cb:
            self.safe(pl, cb, res)

    # ---------- загрузка ----------
    def _files(self):
        out = []
        try:
            entries = sorted(PLUGINS_DIR.iterdir())
        except Exception:
            return out
        for p in entries:
            if p.name.startswith(("_", ".")):
                continue
            if p.is_file() and p.suffix == ".py":
                out.append(p)
            elif p.is_dir() and (p / "plugin.py").exists():
                out.append(p / "plugin.py")
        return out

    def unload_all(self):
        for pl in self.plugins:
            for t in pl.timers:
                t.stop()
                t.deleteLater()
            pl.timers = []
            td = getattr(pl.module, "teardown", None)
            if callable(td):
                self.safe(pl, td)
        self.plugins = []

    def load_all(self):
        self.unload_all()
        try:
            first = not PLUGINS_DIR.exists()
            PLUGINS_DIR.mkdir(parents=True, exist_ok=True)
            PLUGIN_DATA_DIR.mkdir(parents=True, exist_ok=True)
            if first:
                (PLUGINS_DIR / "_example.py").write_text(EXAMPLE_PLUGIN, encoding="utf-8")
        except Exception:
            self.changed.emit()
            return
        for path in self._files():
            pid = path.parent.name if path.name == "plugin.py" else path.stem
            pl = Plugin(pid, path)
            self.plugins.append(pl)
            try:
                if path.name == "plugin.py" and str(path.parent) not in sys.path:
                    sys.path.insert(0, str(path.parent))
                spec = importlib.util.spec_from_file_location(f"island_plugin_{pid}", path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                pl.module = mod
                pl.name = str(getattr(mod, "NAME", pid))
                setup = getattr(mod, "setup", None)
                if not callable(setup):
                    raise RuntimeError("в плагине нет функции setup(api)")
                setup(PluginAPI(self, pl))
            except Exception as e:
                pl.error = f"{type(e).__name__}: {e}"
                pl.load_failed = True
                pl.compact, pl.commands, pl.searches = [], [], []
                for t in pl.timers:
                    t.stop()
                pl.timers = []
                self._log(pl, traceback.format_exc())
        self.refresh_compact()
        self.changed.emit()

    # ---------- свёрнутый остров ----------
    def refresh_compact(self):
        for pl in self.plugins:
            if pl.load_failed:
                continue
            cache = []
            for fn, idle_only in pl.compact:
                v = self.safe(pl, fn)
                if v not in (None, "", False):
                    cache.append((str(v)[:16], idle_only))
            pl.cache = cache

    def compact_texts(self, has_media):
        out = []
        for pl in self.plugins:
            if pl.load_failed:
                continue
            for text, idle_only in pl.cache:
                if idle_only and has_media:
                    continue
                out.append(text)
        return out[:3]

    # ---------- лаунчер ----------
    def launcher_results(self, q):
        ql = q.strip().lower()
        out = []
        if not ql:
            return out
        for pl in self.plugins:
            if pl.load_failed:
                continue
            for keys, title, glyph, sub, action in pl.commands:
                if any(k.startswith(ql) for k in keys) or ql in title.lower():
                    out.append((glyph, title, sub, lambda a=action, p=pl: self.safe(p, a)))
            for fn in pl.searches:
                items = self.safe(pl, fn, q) or []
                for it in items:
                    try:
                        glyph, title, sub, action = it[0], str(it[1]), str(it[2]), it[3]
                    except Exception:
                        continue
                    out.append((glyph, title, sub, lambda a=action, p=pl: self.safe(p, a)))
        return out[:4]


class Island(QWidget):
    def __init__(self):
        super().__init__(None, Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("QLabel{color:white;background:transparent;}")

        self.store = Store()
        self.dm = DownloadMonitor()
        self.timers = TimerManager()
        self.toast = Toast()
        self.cb = QGuiApplication.clipboard()
        self.last_clip = ""

        self.expanded = False
        self._keep_open_until = 0.0
        self.last_title = None
        self.has_media = False
        self.playing = False
        self.cover_small = None
        self.flash = None
        self.phase = 0.0
        self.timer_rows = {}
        self.cfont = QFont(TEXT_FONT, 9, QFont.Weight.DemiBold)

        self.show_progress = False
        self.media_pos = 0.0
        self.media_dur = 0.0
        self.media_pos_ts = time.time()
        self.media_key = ""
        self.lyrics_cache = {}
        self.lyrics_threads = []
        self.lyrics_line_labels = []

        self.vol_show = False
        self.vol_pct = 0
        self._lyrics_tick_n = 0

        self.weather = None          # (температура, код, день)
        self.weather_loc = None
        self.weather_thread = None
        self.wfont = QFont("Segoe UI Emoji", 10)
        self.mic_apps, self.cam_apps = [], []
        self.rec_proc = None
        self.rec_path = None
        self.rec_start = 0.0
        self.rec_stopping = False
        self._rec_stop_t0 = 0.0
        self.plugins = PluginManager(self)
        self.plug_labels = []

        self._build_ui()
        self.setGeometry(self._rect_for(W0, H0))
        self.full.setVisible(False)

        self.anim = QPropertyAnimation(self, b"geometry")
        self.anim.setDuration(340)
        self.anim.setEasingCurve(QEasingCurve.Type.OutCubic)

        self.collapse_timer = QTimer(self)
        self.collapse_timer.setSingleShot(True)
        self.collapse_timer.timeout.connect(self._maybe_collapse)

        self.eq_timer = QTimer(self)
        self.eq_timer.timeout.connect(self._tick)
        self.eq_timer.start(40)

        self.vol_timer = QTimer(self)
        self.vol_timer.setSingleShot(True)
        self.vol_timer.timeout.connect(self._hide_volume)

        self.clip_refresh_timer = QTimer(self)
        self.clip_refresh_timer.setSingleShot(True)
        self.clip_refresh_timer.timeout.connect(self.refresh_clips)

        self.cb.dataChanged.connect(self._on_clip_changed)
        self.dm.changed.connect(self._on_data_change)
        self.dm.file_done.connect(lambda name, folder: self.toast.show_toast(name, folder))
        self.timers.changed.connect(self._on_timers_changed)
        self.timers.finished.connect(self._on_timer_finished)

        self.media = MediaThread()
        self.media.state.connect(self.on_state)
        self.media.start()

        self.refresh_clips()
        self._render_downloads()

        self.weather_timer = QTimer(self)
        self.weather_timer.timeout.connect(self._refresh_weather)
        self.weather_timer.start(15 * 60 * 1000)
        QTimer.singleShot(1500, self._refresh_weather)

        self.plugins.changed.connect(self._render_plugins)
        self.plugins.load_all()

        self.privacy_timer = QTimer(self)
        self.privacy_timer.timeout.connect(self._poll_privacy)
        self.privacy_timer.start(1500)
        QTimer.singleShot(300, self._poll_privacy)

        self.plugin_timer = QTimer(self)
        self.plugin_timer.timeout.connect(self._poll_plugins)
        self.plugin_timer.start(1000)

        # сторож: гарантированно сворачивает остров, если курсор ушёл
        self.watchdog = QTimer(self)
        self.watchdog.timeout.connect(self._watchdog)
        self.watchdog.start(800)

    def _watchdog(self):
        if not self.expanded or time.time() < self._keep_open_until:
            return
        if self.anim.state() == QPropertyAnimation.State.Running:
            return
        if self.geometry().contains(QCursor.pos()):
            return
        if QApplication.activePopupWidget() is not None:
            return
        fw = QApplication.focusWidget()
        if self.isActiveWindow() and isinstance(fw, QLineEdit) and self.isAncestorOf(fw) and fw.text():
            return
        self.collapse()

    # ---------- построение интерфейса ----------
    def _btn(self, glyph, slot, size=36, px=16, filled=False, soft=False):
        b = QPushButton(glyph)
        b.setFixedSize(size, size)
        b.setFont(icon_font(px))
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.clicked.connect(lambda _=False: slot())
        r = size // 2
        if filled:
            bg, hv, fg = "#ffffff", "#e6e6e6", "#000000"
        elif soft:
            bg, hv, fg = "rgba(255,255,255,20)", "rgba(255,255,255,48)", "#ffffff"
        else:
            bg, hv, fg = "transparent", "rgba(255,255,255,30)", "#ffffff"
        b.setStyleSheet(f"QPushButton{{background:{bg};color:{fg};border:none;border-radius:{r}px;}}"
                        f"QPushButton:hover{{background:{hv};}}"
                        f"QPushButton:pressed{{background:rgba(255,255,255,70);color:{fg};}}")
        return b

    def _text_btn(self, text, slot, filled=False, w=None, h=32, checkable=False):
        b = QPushButton(text)
        b.setFixedHeight(h)
        if w:
            b.setFixedWidth(w)
        b.setFont(QFont(TEXT_FONT, 9, QFont.Weight.DemiBold))
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setCheckable(checkable)
        b.clicked.connect(lambda _=False: slot())
        r = h // 2
        if filled:
            b.setStyleSheet(f"QPushButton{{background:white;color:black;border:none;border-radius:{r}px;padding:0 14px;}}"
                            "QPushButton:hover{background:#e6e6e6;}")
        else:
            b.setStyleSheet(f"QPushButton{{background:rgba(255,255,255,18);color:white;border:none;border-radius:{r}px;padding:0 12px;}}"
                            "QPushButton:hover{background:rgba(255,255,255,42);}"
                            "QPushButton:checked{background:rgba(255,255,255,70);}")
        return b

    def _scroll(self):
        sa = QScrollArea()
        sa.setWidgetResizable(True)
        sa.setFrameShape(QFrame.Shape.NoFrame)
        sa.setStyleSheet(SCROLL_QSS)
        sa.viewport().setAutoFillBackground(False)
        inner = QWidget()
        inner.setAutoFillBackground(False)
        inner.setStyleSheet("background: transparent;")
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(0, 0, 4, 0)
        lay.setSpacing(6)
        sa.setWidget(inner)
        return sa, lay

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSizeConstraint(QLayout.SizeConstraint.SetNoConstraint)
        root.setContentsMargins(22, 0, 22, 0)

        self.full = QWidget()
        fl = QVBoxLayout(self.full)
        fl.setContentsMargins(0, 16, 0, 14)
        fl.setSpacing(10)

        tabs = QHBoxLayout()
        tabs.setSpacing(4)
        self.tab_btns = []
        for i, name in enumerate(TAB_NAMES):
            b = QPushButton(name)
            b.setCheckable(True)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setFont(QFont(TEXT_FONT, 9, QFont.Weight.DemiBold))
            b.setStyleSheet("QPushButton{color:rgba(255,255,255,120);background:transparent;border:none;"
                            "padding:4px 6px;border-bottom:2px solid transparent;}"
                            "QPushButton:hover{color:rgba(255,255,255,200);}"
                            "QPushButton:checked{color:white;border-bottom:2px solid white;}")
            b.clicked.connect(lambda _=False, idx=i: self._set_tab(idx))
            tabs.addWidget(b)
            self.tab_btns.append(b)
        tabs.addStretch()
        fl.addLayout(tabs)

        self.stack = QStackedWidget()
        self.stack.addWidget(self._page_home())
        self.stack.addWidget(self._page_clip())
        self.stack.addWidget(self._page_dl())
        self.stack.addWidget(self._page_timer())
        self.stack.addWidget(self._page_lyrics())
        self.stack.addWidget(self._page_plugins())
        fl.addWidget(self.stack, 1)
        self.tab_btns[0].setChecked(True)
        root.addWidget(self.full)

    def _page_home(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 4, 0, 0)
        pl.setSpacing(8)

        head = QHBoxLayout()
        head.setSpacing(14)
        self.art = QLabel()
        self.art.setFixedSize(72, 72)
        self.art.setStyleSheet("background: rgba(255,255,255,14); border-radius: 16px;")
        head.addWidget(self.art)
        texts = QVBoxLayout()
        texts.setSpacing(2)
        self.title = QLabel("Ничего не играет")
        self.title.setFont(QFont(TEXT_FONT, 12, QFont.Weight.DemiBold))
        self.artist = QLabel("")
        self.artist.setFont(QFont(TEXT_FONT, 10))
        self.artist.setStyleSheet("color: rgba(255,255,255,140); background: transparent;")
        texts.addStretch()
        texts.addWidget(self.title)
        texts.addWidget(self.artist)
        texts.addStretch()
        head.addLayout(texts, 1)
        self.btn_prev = self._btn("\uE892", lambda: self.media.command("prev"), 38, 16)
        self.btn_play = self._btn("\uE768", lambda: self.media.command("toggle"), 46, 18, filled=True)
        self.btn_next = self._btn("\uE893", lambda: self.media.command("next"), 38, 16)
        self.btn_chevron = self._btn("\uE70D", self._toggle_progress, 28, 11, soft=True)
        for b in (self.btn_prev, self.btn_play, self.btn_next, self.btn_chevron):
            head.addWidget(b)
        pl.addLayout(head)

        self.seekbar = SeekBar(self)
        self.seekbar.hide()
        pl.addWidget(self.seekbar)

        line = QFrame()
        line.setFixedHeight(1)
        line.setStyleSheet("background: rgba(255,255,255,22); border: none;")
        pl.addWidget(line)

        apps = QHBoxLayout()
        apps.setSpacing(8)
        for glyph, target in APPS:
            apps.addWidget(self._btn(glyph, lambda t=target: self.launch(t), 34, 15, soft=True))
        apps.addStretch()
        apps.addWidget(self._btn("\uE722", self.screenshot_full, 34, 15, soft=True))
        apps.addWidget(self._btn("\uE7A8", self.screenshot_region, 34, 15, soft=True))
        self.btn_rec = self._btn("\u25CF", self.toggle_recording, 34, 15, soft=True)
        apps.addWidget(self.btn_rec)
        pl.addLayout(apps)
        pl.addStretch()
        return page

    def _page_clip(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 2, 0, 0)
        pl.setSpacing(8)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.clip_search = QLineEdit()
        self.clip_search.setPlaceholderText("Поиск по буферу обмена…")
        self.clip_search.setFont(QFont(TEXT_FONT, 10))
        self.clip_search.setStyleSheet(EDIT_QSS)
        self.clip_search.setFixedHeight(34)
        self.clip_search.textChanged.connect(lambda _: self.refresh_clips())
        row.addWidget(self.clip_search, 1)
        self.btn_today = self._text_btn("Сегодня", lambda: self._clip_scope(True), h=34, checkable=True)
        self.btn_all = self._text_btn("Всё", lambda: self._clip_scope(False), h=34, checkable=True)
        self.btn_today.setChecked(True)
        row.addWidget(self.btn_today)
        row.addWidget(self.btn_all)
        pl.addLayout(row)

        self.clip_info = QLabel("")
        self.clip_info.setFont(QFont(TEXT_FONT, 8))
        self.clip_info.setStyleSheet("color: rgba(255,255,255,110); background: transparent;")
        pl.addWidget(self.clip_info)

        self.clip_scroll, self.clip_list = self._scroll()
        pl.addWidget(self.clip_scroll, 1)
        return page

    def _page_dl(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 4, 0, 0)
        pl.setSpacing(6)
        cap = QLabel("СЕТЬ")
        cap.setFont(QFont(TEXT_FONT, 8, QFont.Weight.DemiBold))
        cap.setStyleSheet("color: rgba(255,255,255,110); background: transparent;")
        pl.addWidget(cap)
        self.net_label = QLabel("")
        self.net_label.setTextFormat(Qt.TextFormat.RichText)
        pl.addWidget(self.net_label)
        self.spark = SparkLine(lambda: self.dm.history)
        pl.addWidget(self.spark)
        line = QFrame()
        line.setFixedHeight(1)
        line.setStyleSheet("background: rgba(255,255,255,22); border: none;")
        pl.addWidget(line)
        self.dl_label = QLabel("")
        self.dl_label.setTextFormat(Qt.TextFormat.RichText)
        self.dl_label.setWordWrap(True)
        pl.addWidget(self.dl_label)
        self.app_label = QLabel("")
        self.app_label.setTextFormat(Qt.TextFormat.RichText)
        pl.addWidget(self.app_label)
        pl.addStretch()
        return page

    def _page_timer(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 2, 0, 0)
        pl.setSpacing(8)

        self.alert_label = QLabel("")
        self.alert_label.setFont(QFont(TEXT_FONT, 11, QFont.Weight.DemiBold))
        self.alert_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.alert_label.setFixedHeight(34)
        self.alert_label.setStyleSheet("background: white; color: black; border-radius: 17px;")
        self.alert_label.hide()
        pl.addWidget(self.alert_label)

        row = QHBoxLayout()
        row.setSpacing(8)
        self.timer_edit = QLineEdit()
        self.timer_edit.setPlaceholderText("5м · 1ч 30м · 25:00 · 90с")
        self.timer_edit.setFont(QFont(TEXT_FONT, 10))
        self.timer_edit.setStyleSheet(EDIT_QSS)
        self.timer_edit.setFixedHeight(34)
        self.timer_edit.returnPressed.connect(self._start_timer_from_edit)
        row.addWidget(self.timer_edit, 1)
        row.addWidget(self._text_btn("Старт", self._start_timer_from_edit, filled=True, h=34))
        pl.addLayout(row)

        presets = QHBoxLayout()
        presets.setSpacing(6)
        for label, sec in [("1м", 60), ("5м", 300), ("10м", 600), ("25м", 1500), ("1ч", 3600)]:
            presets.addWidget(self._text_btn(label, lambda s=sec: self.timers.add(s), h=30))
        presets.addStretch()
        pl.addLayout(presets)

        self.timer_scroll, self.timer_list = self._scroll()
        pl.addWidget(self.timer_scroll, 1)
        self.timer_empty = QLabel("Таймеров нет")
        self.timer_empty.setStyleSheet("color: rgba(255,255,255,100); background: transparent;")
        self.timer_empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        return page

    def _page_lyrics(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 4, 0, 0)
        pl.setSpacing(6)
        self.lyrics_status = QLabel("Нет данных")
        self.lyrics_status.setFont(QFont(TEXT_FONT, 10))
        self.lyrics_status.setStyleSheet("color: rgba(255,255,255,130); background: transparent;")
        pl.addWidget(self.lyrics_status)
        self.lyrics_scroll, self.lyrics_list = self._scroll()
        pl.addWidget(self.lyrics_scroll, 1)
        return page

    # ---------- вкладки ----------
    def _tab_height(self, idx):
        h = TAB_H[idx]
        if idx == TAB_HOME and self.show_progress and self.media_dur > 1:
            h += PROGRESS_EXTRA_H
        return h

    def _set_tab(self, idx, animate=True):
        self.stack.setCurrentIndex(idx)
        for i, b in enumerate(self.tab_btns):
            b.setChecked(i == idx)
        if idx == TAB_CLIP:
            self.refresh_clips()
        if self.expanded and animate:
            self._animate_to(W1, self._tab_height(idx))

    # ---------- геометрия и анимация ----------
    def _rect_for(self, w, h):
        sg = QGuiApplication.primaryScreen().geometry()
        return QRect(sg.x() + (sg.width() - w) // 2, sg.y() + TOP_MARGIN, w, h)

    def _animate_to(self, w, h):
        target = self._rect_for(w, h)
        if self.geometry() == target:
            return
        if self.anim.state() == QPropertyAnimation.State.Running and self.anim.endValue() == target:
            return
        self.anim.stop()
        self.anim.setStartValue(self.geometry())
        self.anim.setEndValue(target)
        self.anim.start()

    def expand(self):
        self.expanded = True
        self._animate_to(W1, self._tab_height(self.stack.currentIndex()))
        QTimer.singleShot(110, lambda: self.full.setVisible(self.expanded))

    def collapse(self):
        if not self.expanded:
            return
        self.expanded = False
        self.full.setVisible(False)
        self.stack.setCurrentIndex(TAB_HOME)
        for i, b in enumerate(self.tab_btns):
            b.setChecked(i == TAB_HOME)
        self._animate_to(self._target_w(), H0)

    def show_plugins(self, seconds=None):
        self.collapse_timer.stop()
        self._keep_open_until = time.time() + (seconds or 0)
        self._set_tab(TAB_PLUGINS, animate=False)
        self.expand()
        if seconds:
            self.collapse_timer.start(int(seconds * 1000) + 100)

    def _maybe_collapse(self):
        if time.time() < self._keep_open_until:
            self.collapse_timer.start(1000)
            return
        if self.geometry().contains(QCursor.pos()):
            return
        fw = QApplication.focusWidget()
        # не сворачиваем, пока в поле ввода есть набранный текст
        if self.isActiveWindow() and isinstance(fw, QLineEdit) and self.isAncestorOf(fw) and fw.text():
            self.collapse_timer.start(1500)
            return
        self.collapse()

    def enterEvent(self, e):
        self.collapse_timer.stop()
        self.expand()

    def leaveEvent(self, e):
        self.collapse_timer.start(450)

    def event(self, e):
        if e.type() == QEvent.Type.WindowDeactivate and self.expanded:
            self.collapse_timer.start(350)
        return super().event(e)

    def contextMenuEvent(self, e):
        menu = QMenu(self)
        act = menu.addAction("Запускать вместе с Windows")
        act.setCheckable(True)
        act.setChecked(autostart_enabled())
        act.triggered.connect(lambda on: self._toggle_autostart(on))
        menu.addSeparator()
        menu.addAction("Остановить запись экрана" if self.rec_proc else "Начать запись экрана",
                       self.toggle_recording)
        menu.addAction("Папка записей", lambda: self._open_dir(RECORD_DIR))
        menu.addSeparator()
        menu.addAction("Папка плагинов", self.open_plugins_dir)
        menu.addAction("Перезагрузить плагины", self.reload_plugins)
        menu.addSeparator()
        menu.addAction("Свернуть остров", lambda: (setattr(self, "_keep_open_until", 0.0), self.collapse()))
        menu.addAction("Выход", QApplication.quit)
        menu.exec(e.globalPos())

    def _toggle_autostart(self, on):
        try:
            set_autostart(on)
            self.flash_msg("Автозапуск включён" if on else "Автозапуск выключен")
        except Exception:
            self.flash_msg("Не удалось изменить автозапуск")

    # ---------- громкость (колесо мыши) ----------
    def wheelEvent(self, e):
        if self.expanded or not HAVE_VOLUME:
            super().wheelEvent(e)
            return
        step = 4 if e.angleDelta().y() > 0 else -4
        pct = change_volume(step)
        if pct is None:
            return
        self.vol_pct = pct
        self.vol_show = True
        self._resize_compact()
        self.update()
        self.vol_timer.start(1100)

    def _hide_volume(self):
        self.vol_show = False
        self._resize_compact()
        self.update()

    # ---------- прогресс/перемотка ----------
    def _toggle_progress(self):
        self.show_progress = not self.show_progress
        visible = self.show_progress and self.media_dur > 1
        self.seekbar.setVisible(visible)
        self.btn_chevron.setText("\uE70E" if self.show_progress else "\uE70D")
        if self.expanded:
            self._animate_to(W1, self._tab_height(self.stack.currentIndex()))

    def current_position(self):
        if not self.playing:
            return min(self.media_pos, max(self.media_dur, 0))
        return min(self.media_dur, self.media_pos + (time.time() - self.media_pos_ts))

    # ---------- свёрнутый вид ----------
    def _compact_items(self):
        fm = QFontMetrics(self.cfont)
        items = []
        if not self.has_media and self.weather:
            items.append(("weather", 20 + 4 + fm.horizontalAdvance(f"{self.weather[0]}°")))
        if self.has_media and self.cover_small:
            items.append(("cover", 22))
        if self.rec_proc is not None:
            wide = (time.time() - self.rec_start) >= 3600
            items.append(("rec", 15 + fm.horizontalAdvance("0:00:00" if wide else "00:00")))
        rem = self.timers.nearest()
        if rem is not None:
            items.append(("timer", fm.horizontalAdvance("0:00:00" if rem >= 3600 else "00:00")))
        if self.dm.speed > 2048:
            items.append(("dl", fm.horizontalAdvance("↓ 99.9 МБ/с")))
        for i, text in enumerate(self.plugins.compact_texts(self.has_media)):
            items.append((f"plug{i}", fm.horizontalAdvance(text) + 4))
        if self.mic_apps:
            items.append(("mic", 16))
        if self.cam_apps:
            items.append(("cam", 16))
        if self.has_media:
            items.append(("eq", 21))
        return items

    def _target_w(self):
        if self.vol_show:
            fm_icon = QFontMetrics(icon_font(16))
            fm_txt = QFontMetrics(self.cfont)
            content = fm_icon.horizontalAdvance("\uE995") + 8 + fm_txt.horizontalAdvance("100%") + 8 + 46
            return max(W0, content + 44)
        if self.flash:
            return max(W0, QFontMetrics(self.cfont).horizontalAdvance(self.flash) + 44)
        items = self._compact_items()
        if not items:
            return W0
        return max(W0, 28 + sum(w for _, w in items) + 12 * (len(items) - 1))

    def _resize_compact(self):
        if not self.expanded:
            self._animate_to(self._target_w(), H0)

    def _tick(self):
        if self.rec_proc is not None and not self.expanded:
            self.update()
            wide = (time.time() - self.rec_start) >= 3600
            if wide != getattr(self, '_rec_wide', False):
                self._rec_wide = wide
                self._resize_compact()
        if self.playing and not self.expanded:
            self.phase += 0.28
            self.update()
        if self.show_progress and self.seekbar.isVisible():
            self.seekbar.update()
        self._lyrics_tick_n += 1
        if self._lyrics_tick_n % 15 == 0:
            self._update_lyrics_highlight()

    def paintEvent(self, e):
        p = QPainter(self)
        w, h = self.width(), self.height()
        draw_surface(p, w, h, min(h / 2, 34))
        if self.expanded or h > 60:
            return

        p.setFont(self.cfont)
        if self.vol_show:
            fm_icon = QFontMetrics(icon_font(16))
            fm_txt = QFontMetrics(self.cfont)
            glyph = "\uE995" if self.vol_pct > 0 else "\uE992"
            text = f"{self.vol_pct}%"
            bar_w = 46
            content = fm_icon.horizontalAdvance(glyph) + 8 + fm_txt.horizontalAdvance(text) + 8 + bar_w
            x = (w - content) / 2
            p.setFont(icon_font(16))
            p.setPen(QColor(255, 255, 255))
            p.drawText(QRectF(x, 0, fm_icon.horizontalAdvance(glyph), h),
                      Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, glyph)
            x += fm_icon.horizontalAdvance(glyph) + 8
            p.setFont(self.cfont)
            p.drawText(QRectF(x, 0, fm_txt.horizontalAdvance(text), h),
                      Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, text)
            x += fm_txt.horizontalAdvance(text) + 8
            p.setPen(Qt.PenStyle.NoPen)
            p.setBrush(QColor(255, 255, 255, 40))
            p.drawRoundedRect(QRectF(x, h / 2 - 3, bar_w, 6), 3, 3)
            p.setBrush(QColor(255, 255, 255, 230))
            p.drawRoundedRect(QRectF(x, h / 2 - 3, bar_w * self.vol_pct / 100, 6), 3, 3)
            return

        if self.flash:
            p.setPen(QColor(255, 255, 255))
            p.drawText(QRectF(0, 0, w, h), Qt.AlignmentFlag.AlignCenter, self.flash)
            return

        items = self._compact_items()
        if not items:
            return
        content = sum(wd for _, wd in items) + 12 * (len(items) - 1)
        x = (w - content) / 2
        for kind, wd in items:
            if kind == "weather" and self.weather:
                temp, code, is_day = self.weather
                p.setFont(self.wfont)
                p.setPen(QColor(255, 255, 255))
                p.drawText(QRectF(x, 0, 20, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                           weather_emoji(code, is_day))
                p.setFont(self.cfont)
                p.drawText(QRectF(x + 24, 0, wd - 24, h),
                           Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, f"{temp}°")
            elif kind == "rec":
                el = time.time() - self.rec_start
                pulse = 0.5 + 0.5 * math.sin(time.time() * 5)
                p.setPen(Qt.PenStyle.NoPen)
                p.setBrush(QColor(255, 59, 48, int(130 + 125 * pulse)))
                p.drawEllipse(QRectF(x, (h - 9) / 2, 9, 9))
                p.setPen(QColor(255, 255, 255))
                p.drawText(QRectF(x + 15, 0, wd - 15, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                           fmt_clock(int(el)))
            elif kind in ("mic", "cam"):
                p.setFont(icon_font(14))
                p.setPen(QColor(255, 159, 10) if kind == "mic" else QColor(48, 209, 88))
                p.drawText(QRectF(x, 0, wd, h), Qt.AlignmentFlag.AlignCenter,
                           "\uE720" if kind == "mic" else "\uE722")
                p.setFont(self.cfont)
            elif kind.startswith("plug"):
                texts = self.plugins.compact_texts(self.has_media)
                idx = int(kind[4:])
                if idx < len(texts):
                    p.setPen(QColor(255, 255, 255, 230))
                    p.drawText(QRectF(x, 0, wd, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                               texts[idx])
            elif kind == "cover":
                p.drawPixmap(int(x), (h - 22) // 2, self.cover_small)
            elif kind == "timer":
                p.setPen(QColor(255, 255, 255))
                rem = self.timers.nearest() or 0
                p.drawText(QRectF(x, 0, wd, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                           fmt_clock(rem))
            elif kind == "dl":
                p.setPen(QColor(255, 255, 255, 215))
                p.drawText(QRectF(x, 0, wd, h), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
                           "↓ " + fmt_speed(self.dm.speed))
            elif kind == "eq":
                p.setPen(Qt.PenStyle.NoPen)
                p.setBrush(QColor(255, 255, 255, 230))
                for i in range(4):
                    bh = 5 + 12 * (0.5 + 0.5 * math.sin(self.phase + i * 1.3)) if self.playing else 4
                    p.drawRoundedRect(QRectF(x + i * 6, (h - bh) / 2, 3, bh), 1.5, 1.5)
            x += wd + 12

    # ---------- медиа ----------
    def on_state(self, st):
        if not st or not st["title"]:
            self.has_media = False
            self.playing = False
            self.title.setText("Ничего не играет")
            self.artist.setText(self._weather_text())
            self.art.clear()
            self.cover_small = None
            self.last_title = None
            self.media_dur = 0.0
            self.seekbar.setVisible(False)
            self._resize_compact()
            self.update()
            return
        self.has_media = True
        self.playing = st["playing"]
        self.title.setText(self.title.fontMetrics().elidedText(
            st["title"], Qt.TextElideMode.ElideRight, 150))
        self.artist.setText(self.artist.fontMetrics().elidedText(
            st["artist"], Qt.TextElideMode.ElideRight, 150))
        self.btn_play.setText("\uE769" if st["playing"] else "\uE768")

        self.media_pos = st.get("position", 0.0)
        self.media_dur = st.get("duration", 0.0)
        self.media_pos_ts = time.time()
        self.seekbar.setVisible(self.show_progress and self.media_dur > 1)

        if st["title"] != self.last_title:
            self.last_title = st["title"]
            pm = QPixmap()
            if st["thumb"] and pm.loadFromData(st["thumb"]):
                self.art.setPixmap(rounded_pixmap(pm, 72, 16))
                self.cover_small = rounded_pixmap(pm, 22, 11)
            else:
                self.art.clear()
                self.cover_small = None

        key = f"{st['artist']}|{st['title']}"
        if key != self.media_key:
            self.media_key = key
            if key in self.lyrics_cache:
                self._render_lyrics()
            elif HAVE_REQUESTS:
                self.lyrics_status.setText("Ищу текст…")
                clear_layout(self.lyrics_list)
                th = LyricsThread(key, st["artist"], st["title"])
                th.result.connect(self._on_lyrics)
                self.lyrics_threads.append(th)
                th.finished.connect(lambda t=th: self.lyrics_threads.remove(t) if t in self.lyrics_threads else None)
                th.start()
            else:
                self.lyrics_status.setText("Текст недоступен (нет библиотеки requests)")

        self._resize_compact()
        self.update()

    # ---------- микрофон / камера ----------
    def _poll_privacy(self):
        try:
            mic = consent_active("microphone")
            cam = consent_active("webcam")
        except Exception:
            return
        if mic == self.mic_apps and cam == self.cam_apps:
            return
        self.mic_apps, self.cam_apps = mic, cam
        parts = []
        if mic:
            parts.append("Микрофон: " + ", ".join(mic))
        if cam:
            parts.append("Камера: " + ", ".join(cam))
        self.setToolTip("\n".join(parts))
        self._resize_compact()
        self.update()

    # ---------- запись экрана ----------
    def _open_dir(self, path):
        try:
            Path(path).mkdir(parents=True, exist_ok=True)
            os.startfile(str(path))
        except Exception:
            self.flash_msg("Не удалось открыть папку")

    def _rec_ui(self, on):
        if not hasattr(self, "_rec_qss"):
            self._rec_qss = self.btn_rec.styleSheet()
        self.btn_rec.setText("\u25A0" if on else "\u25CF")
        self.btn_rec.setStyleSheet(self._rec_qss.replace("color:#ffffff;", "color:#ff453a;") if on else self._rec_qss)

    def toggle_recording(self):
        if self.rec_proc is None:
            self.start_recording()
        elif not self.rec_stopping:
            self.stop_recording()

    def start_recording(self):
        ff = find_ffmpeg()
        if not ff:
            self.flash_msg("Для записи нужен ffmpeg.exe (положите рядом с main.py)", 4500)
            return
        try:
            RECORD_DIR.mkdir(parents=True, exist_ok=True)
            out = RECORD_DIR / datetime.datetime.now().strftime("Запись_%Y-%m-%d_%H-%M-%S.mp4")
            cmd = [ff, "-y", "-f", "gdigrab", "-framerate", "30", "-draw_mouse", "1", "-i", "desktop",
                   "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
                   "-c:v", "libx264", "-preset", "ultrafast", "-crf", "23", "-pix_fmt", "yuv420p",
                   str(out)]
            self.rec_proc = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
                                             stderr=subprocess.DEVNULL, creationflags=0x08000000)
        except Exception:
            self.rec_proc = None
            self.flash_msg("Не удалось начать запись")
            return
        self.rec_path = out
        self.rec_start = time.time()
        self.rec_stopping = False
        self._rec_ui(True)
        QTimer.singleShot(1500, self._check_rec_alive)
        self._resize_compact()
        self.update()

    def _check_rec_alive(self):
        if self.rec_proc is not None and not self.rec_stopping and self.rec_proc.poll() is not None:
            self._rec_cleanup()
            self.flash_msg("Запись не запустилась (ошибка ffmpeg)", 3000)

    def stop_recording(self):
        if self.rec_proc is None or self.rec_stopping:
            return
        self.rec_stopping = True
        try:
            self.rec_proc.stdin.write(b"q")
            self.rec_proc.stdin.flush()
        except Exception:
            try:
                self.rec_proc.terminate()
            except Exception:
                pass
        self._rec_stop_t0 = time.time()
        QTimer.singleShot(250, self._poll_rec_stop)

    def _poll_rec_stop(self):
        p = self.rec_proc
        if p is None:
            return
        if p.poll() is None:
            if time.time() - self._rec_stop_t0 > 10:
                try:
                    p.kill()
                except Exception:
                    pass
            QTimer.singleShot(250, self._poll_rec_stop)
            return
        path = self.rec_path
        self._rec_cleanup()
        try:
            ok = bool(path) and Path(path).exists() and Path(path).stat().st_size > 0
        except Exception:
            ok = False
        self.flash_msg("Запись сохранена" if ok else "Запись не сохранилась", 2400)

    def _rec_cleanup(self):
        p, self.rec_proc = self.rec_proc, None
        self.rec_stopping = False
        try:
            if p and p.stdin:
                p.stdin.close()
        except Exception:
            pass
        self._rec_ui(False)
        self._resize_compact()
        self.update()

    def shutdown(self):
        """Вызывается при выходе: корректно дописать видео."""
        p = self.rec_proc
        if p is not None and p.poll() is None:
            try:
                p.stdin.write(b"q")
                p.stdin.flush()
                p.wait(timeout=8)
            except Exception:
                try:
                    p.kill()
                except Exception:
                    pass

    # ---------- плагины ----------
    def _page_plugins(self):
        page = QWidget()
        pl = QVBoxLayout(page)
        pl.setContentsMargins(0, 2, 0, 0)
        pl.setSpacing(8)

        row = QHBoxLayout()
        row.setSpacing(8)
        cap = QLabel("ПЛАГИНЫ")
        cap.setFont(QFont(TEXT_FONT, 8, QFont.Weight.DemiBold))
        cap.setStyleSheet("color: rgba(255,255,255,110); background: transparent;")
        row.addWidget(cap)
        row.addStretch()
        row.addWidget(self._text_btn("Папка", self.open_plugins_dir, h=28))
        row.addWidget(self._text_btn("Перезагрузить", self.reload_plugins, h=28))
        pl.addLayout(row)

        self.plug_scroll, self.plug_list = self._scroll()
        pl.addWidget(self.plug_scroll, 1)
        self.plug_labels = []
        return page

    def _render_plugins(self):
        clear_layout(self.plug_list)
        self.plug_labels = []
        if not self.plugins.plugins:
            hint = QLabel("Плагинов нет.\n\nПоложите .py файл в папку плагинов (кнопка «Папка»).\n"
                          "Пример: переименуйте _example.py в example.py\nи нажмите «Перезагрузить».")
            hint.setWordWrap(True)
            hint.setFont(QFont(TEXT_FONT, 9))
            hint.setStyleSheet("color: rgba(255,255,255,120); background: transparent;")
            self.plug_list.addWidget(hint)
        for pl in self.plugins.plugins:
            card = QFrame()
            card.setObjectName("plugCard")
            card.setStyleSheet("#plugCard{background:rgba(255,255,255,14);border-radius:14px;}")
            v = QVBoxLayout(card)
            v.setContentsMargins(12, 10, 12, 10)
            v.setSpacing(6)
            name = QLabel(pl.name)
            name.setFont(QFont(TEXT_FONT, 10, QFont.Weight.DemiBold))
            v.addWidget(name)
            status = QLabel("")
            status.setWordWrap(True)
            status.setFont(QFont(TEXT_FONT, 9))
            if pl.markdown:
                status.setTextFormat(Qt.TextFormat.MarkdownText)
            v.addWidget(status)
            if pl.input and not pl.load_failed:
                try:
                    hint, fn = pl.input[0], pl.input[1]
                    edit = QLineEdit()
                    edit.setPlaceholderText(str(hint))
                    edit.setFont(QFont(TEXT_FONT, 10))
                    edit.setStyleSheet(EDIT_QSS)
                    edit.setFixedHeight(32)

                    def submit(e=edit, f=fn, p=pl):
                        t = e.text().strip()
                        if t:
                            e.clear()
                            e.clearFocus()
                            self.plugins.safe(p, f, t)
                    edit.returnPressed.connect(submit)
                    v.insertWidget(1, edit)
                except Exception:
                    pass
            if pl.buttons and not pl.load_failed:
                brow = QHBoxLayout()
                brow.setSpacing(6)
                for item in pl.buttons:
                    try:
                        label, fn = item[0], item[1]
                    except Exception:
                        continue
                    brow.addWidget(self._text_btn(str(label), lambda f=fn, p=pl: self.plugins.safe(p, f), h=28))
                brow.addStretch()
                v.addLayout(brow)
            self.plug_labels.append((pl, status))
            self.plug_list.addWidget(card)
        self.plug_list.addStretch()
        self._update_plugin_labels()

    def _update_plugin_labels(self):
        for pl, lbl in self.plug_labels:
            if pl.error:
                text, err = "⚠ " + pl.error, True
            else:
                lines = self.plugins.safe(pl, pl.panel_fn) if pl.panel_fn else None
                text = "\n".join(str(x) for x in (lines or [])) or "Работает"
                err = bool(pl.error)
            if lbl.text() != text:
                lbl.setText(text)
            if lbl.property("err") != err:
                lbl.setProperty("err", err)
                lbl.setStyleSheet("color: #ff6b6b; background: transparent;" if err
                                  else "color: rgba(255,255,255,150); background: transparent;")

    def _poll_plugins(self):
        before = self.plugins.compact_texts(self.has_media)
        self.plugins.refresh_compact()
        after = self.plugins.compact_texts(self.has_media)
        if before != after:
            self._resize_compact()
            self.update()
        if self.expanded and self.stack.currentIndex() == TAB_PLUGINS:
            self._update_plugin_labels()

    def open_plugins_dir(self):
        self._open_dir(PLUGINS_DIR)

    def reload_plugins(self):
        self.plugins.load_all()
        self.flash_msg("Плагины перезагружены")

    # ---------- погода ----------
    def _weather_text(self):
        if not self.weather:
            return ""
        temp, code, is_day = self.weather
        return f"{weather_emoji(code, is_day)} {temp}°"

    def _refresh_weather(self):
        if not HAVE_REQUESTS or (self.weather_thread and self.weather_thread.isRunning()):
            return
        th = IslandWeatherThread(self.weather_loc)
        th.done.connect(self._on_weather)
        th.finished.connect(lambda: setattr(self, "weather_thread", None))
        self.weather_thread = th
        th.start()

    def _on_weather(self, loc, res):
        if loc:
            self.weather_loc = loc
        if res:
            self.weather = res
            if not self.has_media:
                self.artist.setText(self._weather_text())
            self._resize_compact()
            self.update()

    def _on_lyrics(self, key, lines):
        self.lyrics_cache[key] = lines
        if key == self.media_key:
            self._render_lyrics()

    def _render_lyrics(self):
        lines = self.lyrics_cache.get(self.media_key)
        clear_layout(self.lyrics_list)
        self.lyrics_line_labels = []
        if not lines:
            self.lyrics_status.setText("Текст не найден")
            self.lyrics_list.addStretch()
            return
        self.lyrics_status.setText("")
        for t, text in lines:
            lab = QLabel(text or "♪")
            lab.setWordWrap(True)
            lab.setFont(QFont(TEXT_FONT, 11))
            lab.setStyleSheet("color: rgba(255,255,255,120); background: transparent;")
            self.lyrics_list.addWidget(lab)
            self.lyrics_line_labels.append((t, lab))
        self.lyrics_list.addStretch()

    def _update_lyrics_highlight(self):
        if self.stack.currentIndex() != TAB_LYRICS or not self.lyrics_line_labels:
            return
        pos = self.current_position()
        cur = None
        for t, lab in self.lyrics_line_labels:
            lab.setStyleSheet("color: rgba(255,255,255,120); background: transparent;")
            if t <= pos + 0.3:
                cur = lab
        if cur:
            cur.setStyleSheet("color: white; font-weight: 600; background: transparent;")
            self.lyrics_scroll.ensureWidgetVisible(cur, 0, 60)

    # ---------- буфер обмена ----------
    def _on_clip_changed(self):
        try:
            md = self.cb.mimeData()
            if md is None:
                return
            if md.hasFormat("ExcludeClipboardContentFromMonitorProcessing"):
                return  # менеджеры паролей помечают так секретные данные
            if md.hasImage():
                img = self.cb.image()
                if img.isNull():
                    return
                buf = QByteArray()
                qbuf = QBuffer(buf)
                qbuf.open(QIODevice.OpenModeFlag.WriteOnly)
                img.save(qbuf, "PNG")
                h = hashlib.sha1(bytes(buf)).hexdigest()[:16]
                if h == self.last_clip:
                    return
                self.last_clip = h
                CLIPS_DIR.mkdir(exist_ok=True)
                path = CLIPS_DIR / f"{h}.png"
                if not path.exists():
                    img.save(str(path), "PNG")
                self.store.add_image_clip(str(path), f"Изображение {img.width()}×{img.height()}")
            elif md.hasText():
                text = md.text()
                if not text or not text.strip() or text == self.last_clip:
                    return
                self.last_clip = text
                self.store.add_clip(text)
            else:
                return
        except Exception:
            return
        if self.expanded and self.stack.currentIndex() == TAB_CLIP:
            self.clip_refresh_timer.start(150)

    def _clip_scope(self, today):
        self.btn_today.setChecked(today)
        self.btn_all.setChecked(not today)
        self.refresh_clips()

    def refresh_clips(self):
        q = self.clip_search.text().strip()
        rows = self.store.clips(q, self.btn_today.isChecked(), 100)
        clear_layout(self.clip_list)
        for cid, text, ts, pinned, kind, path, label in rows:
            self.clip_list.addWidget(ClipRow(self, cid, text, ts, bool(pinned), kind, path, label))
        self.clip_list.addStretch()
        scope = "за сегодня" if self.btn_today.isChecked() else "за всё время"
        self.clip_info.setText(f"Показано {len(rows)} {scope} · всего в памяти {self.store.clip_total()}"
                               + (" · ищу: " + q if q else ""))

    def copy_text(self, text):
        self.last_clip = text
        self.cb.setText(text)
        self.flash_msg("Скопировано")

    def copy_image(self, path):
        try:
            self.last_clip = hashlib.sha1(Path(path).read_bytes()).hexdigest()[:16]
            self.cb.setPixmap(QPixmap(path))
            self.flash_msg("Скопировано")
        except Exception:
            self.flash_msg("Не удалось скопировать")

    def toggle_pin(self, cid, pinned):
        self.store.set_pin(cid, pinned)
        self.refresh_clips()

    def delete_clip(self, cid):
        self.store.delete_clip(cid)
        self.refresh_clips()

    # ---------- загрузки ----------
    def _on_data_change(self):
        self._render_downloads()
        self._resize_compact()
        if not self.expanded:
            self.update()
        elif self.stack.currentIndex() == TAB_DL:
            self.spark.update()

    def _render_downloads(self):
        dm = self.dm
        self.net_label.setText(
            f"<span style='font-size:24px;font-weight:600;color:white'>↓ {fmt_speed(dm.net_down)}</span>"
            f"&nbsp;&nbsp;<span style='font-size:15px;color:#8a8a8a'>↑ {fmt_speed(dm.net_up)}</span>")
        parts = []
        if dm.files:
            parts.append("<div style='color:#8a8a8a;font-size:11px'>СКАЧИВАЕТСЯ</div>")
            for f in dm.files.values():
                name = html.escape(f["name"] if len(f["name"]) < 46 else f["name"][:43] + "…")
                state = fmt_speed(f["speed"]) if f["speed"] > 0 else "ожидание / пауза"
                parts.append(f"<div style='margin-top:6px'><span style='color:white;font-size:13px'>{name}</span><br>"
                             f"<span style='color:#8a8a8a;font-size:12px'>{fmt_size(f['size'])} · {state}</span></div>")
        else:
            parts.append("<div style='color:#6e6e6e;font-size:12px'>Файлы, которые качает браузер "
                         "(Chrome, Edge, Firefox…), появятся здесь.</div>")
        if dm.recent:
            parts.append("<div style='color:#8a8a8a;font-size:11px;margin-top:10px'>ЗАВЕРШЕНО</div>")
            for name, size, _ in dm.recent[:4]:
                name = html.escape(name if len(name) < 46 else name[:43] + "…")
                parts.append(f"<div style='margin-top:3px;color:#cfcfcf;font-size:12px'>{name} "
                             f"<span style='color:#6e6e6e'>· {fmt_size(size)}</span></div>")
        self.dl_label.setText("".join(parts))

        if dm.by_app:
            rows = "".join(f"<div style='color:#cfcfcf;font-size:12px;margin-top:2px'>{html.escape(n)} "
                           f"<span style='color:#6e6e6e'>· {c} соед.</span></div>" for n, c in dm.by_app)
            self.app_label.setText("<div style='color:#8a8a8a;font-size:11px;margin-top:10px'>"
                                   "АКТИВНОСТЬ ПО ПРИЛОЖЕНИЯМ (приблизительно)</div>" + rows)
        elif dm.by_app_error:
            self.app_label.setText("<div style='color:#6e6e6e;font-size:11px;margin-top:10px'>"
                                   "Активность по приложениям недоступна на этом ПК</div>")
        else:
            self.app_label.setText("")

    # ---------- таймеры ----------
    def _start_timer_from_edit(self):
        d = parse_duration(self.timer_edit.text())
        if d:
            self.timers.add(d)
            self.timer_edit.clear()
        else:
            self.flash_msg("Не понял время")

    def _on_timers_changed(self):
        keys = [(i["id"], i["done"]) for i in self.timers.items]
        if keys != list(self.timer_rows.keys()) or len(keys) != len(self.timer_rows):
            clear_layout(self.timer_list)
            self.timer_rows = {}
            for it in self.timers.items:
                row = TimerRow(self, it)
                self.timer_rows[(it["id"], it["done"])] = row
                self.timer_list.addWidget(row)
            if not self.timers.items:
                self.timer_list.addWidget(self.timer_empty)
                self.timer_empty.show()
            self.timer_list.addStretch()
        else:
            for row in self.timer_rows.values():
                row.refresh()
        self._resize_compact()
        if not self.expanded:
            self.update()

    def _on_timer_finished(self, it):
        label = it["label"] or fmt_clock(it["total"])
        self.alert_label.setText(f"Время вышло · {label}")
        self.alert_label.show()
        for i in range(3):
            QTimer.singleShot(i * 450, lambda: winsound.MessageBeep(winsound.MB_ICONEXCLAMATION))
        self.collapse_timer.stop()
        self._keep_open_until = time.time() + 12
        self._set_tab(TAB_TIMER, animate=False)
        self.expand()
        QTimer.singleShot(12000, self._dismiss_alert)

    def _dismiss_alert(self):
        self.alert_label.hide()
        self._maybe_collapse()

    # ---------- действия ----------
    def flash_msg(self, text, ms=1600):
        self.flash = text
        self._resize_compact()
        self.update()
        QTimer.singleShot(ms, self._clear_flash)

    def _clear_flash(self):
        self.flash = None
        self._resize_compact()
        self.update()

    def launch(self, target):
        try:
            os.startfile(target)
        except Exception:
            self.flash_msg("Не удалось открыть")

    def screenshot_full(self):
        self.collapse_timer.stop()
        self.hide()
        QTimer.singleShot(350, self._do_screenshot)

    def _do_screenshot(self):
        try:
            SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
            pm = QGuiApplication.primaryScreen().grabWindow(0)
            name = datetime.datetime.now().strftime("island_%Y-%m-%d_%H-%M-%S.png")
            pm.save(str(SCREENSHOT_DIR / name))
            self.cb.setPixmap(pm)
            msg = "Скриншот сохранён"
        except Exception:
            msg = "Ошибка скриншота"
        self.show()
        self.collapse()
        self.flash_msg(msg)

    def screenshot_region(self):
        self.collapse()
        try:
            os.startfile("ms-screenclip:")
        except Exception:
            self.flash_msg("Не удалось открыть")


# ====================== Лаунчер (Alt+Space) ======================
class Result:
    def __init__(self, glyph, title, sub, action):
        self.glyph, self.title, self.sub, self.action = glyph, title, sub, action


def match_score(name, q):
    n = name.lower()
    if n.startswith(q):
        return 100
    if any(w.startswith(q) for w in re.split(r"[\s\-_.]+", n)):
        return 80
    if q in n:
        return 60
    it = iter(n)
    if all(ch in it for ch in q):
        return 30
    return None


def scan_start_menu():
    apps = []
    roots = [Path(os.environ.get("ProgramData", "C:/ProgramData")) / "Microsoft/Windows/Start Menu/Programs",
             Path(os.environ.get("APPDATA", "")) / "Microsoft/Windows/Start Menu/Programs"]
    bad = ("uninstall", "удал", "readme", "справк", "help")
    for root in roots:
        if not root.exists():
            continue
        for lnk in root.rglob("*.lnk"):
            name = lnk.stem
            if any(b in name.lower() for b in bad):
                continue
            apps.append({"name": name, "kind": "path", "target": str(lnk)})
    return apps


def scan_start_apps():
    import json
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-StartApps | ConvertTo-Json -Compress"],
            capture_output=True, timeout=30, creationflags=0x08000000)
        data = json.loads(out.stdout.decode("utf-8", "ignore") or "[]")
        if isinstance(data, dict):
            data = [data]
        return [{"name": d["Name"], "kind": "appid", "target": d["AppID"]} for d in data if d.get("Name")]
    except Exception:
        return []


BUILTIN_APPS = [
    ("Командная строка", "cmd.exe"), ("PowerShell", "powershell.exe"), ("Диспетчер задач", "taskmgr.exe"),
    ("Проводник", "explorer.exe"), ("Панель управления", "control.exe"), ("Редактор реестра", "regedit.exe"),
    ("Параметры", "ms-settings:"), ("Блокнот", "notepad.exe"), ("Калькулятор", "calc.exe"),
    ("Paint", "mspaint.exe"), ("Ножницы", "ms-screenclip:"),
]


class RateThread(QThread):
    done = pyqtSignal(str, object)

    def __init__(self, base):
        super().__init__()
        self.base = base

    def run(self):
        data = None
        try:
            r = requests.get("https://api.frankfurter.app/latest", params={"from": self.base}, timeout=6)
            if r.ok:
                data = r.json().get("rates")
        except Exception:
            data = None
        self.done.emit(self.base, data)


class TranslateThread(QThread):
    done = pyqtSignal(str, str, str)  # текст, целевой язык, перевод

    def __init__(self, text, target):
        super().__init__()
        self.text, self.target = text, target

    def run(self):
        translated = ""
        try:
            r = requests.get("https://translate.googleapis.com/translate_a/single",
                             params={"client": "gtx", "sl": "auto", "tl": self.target, "dt": "t", "q": self.text},
                             timeout=6)
            data = r.json()
            translated = "".join(seg[0] for seg in data[0])
        except Exception:
            translated = ""
        self.done.emit(self.text, self.target, translated)


class WeatherThread(QThread):
    done = pyqtSignal(str, str)  # город, результат

    def __init__(self, city):
        super().__init__()
        self.city = city

    def run(self):
        text = ""
        try:
            loc = urllib.parse.quote(self.city) if self.city else ""
            r = requests.get(f"https://wttr.in/{loc}", params={"format": "3"}, timeout=6)
            if r.ok:
                text = r.text.strip()
        except Exception:
            text = ""
        self.done.emit(self.city, text)


class Launcher(QWidget):
    PW = 640
    ROW_H = 54

    def __init__(self, island):
        super().__init__(None, Qt.WindowType.FramelessWindowHint
                         | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("QLabel{color:white;background:transparent;}")
        self.island = island
        self.store = island.store
        self.apps = [{"name": n, "kind": "path", "target": t} for n, t in BUILTIN_APPS]
        self.results = []

        self.rate_cache = {}
        self.rate_threads = []
        self._fetching_rate = set()
        self.translate_cache = {}
        self.translate_threads = []
        self._fetching_translate = set()
        self.weather_cache = {}
        self.weather_threads = []
        self._fetching_weather = set()

        threading.Thread(target=self._index_apps, daemon=True).start()

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 14)
        lay.setSpacing(8)
        self.edit = QLineEdit()
        self.edit.setPlaceholderText("Программа, 2+2, 100 usd в eur, t привет, погода, f имя файла, таймер 5м")
        self.edit.setFont(QFont(TEXT_FONT, 15))
        self.edit.setFixedHeight(48)
        self.edit.setStyleSheet("QLineEdit{background:transparent;color:white;border:none;padding:0 6px;"
                                "selection-background-color:rgba(255,255,255,80);}")
        self.edit.textChanged.connect(self.refresh)
        self.edit.returnPressed.connect(self.run_current)
        self.edit.installEventFilter(self)
        lay.addWidget(self.edit)

        self.sep = QFrame()
        self.sep.setFixedHeight(1)
        self.sep.setStyleSheet("background: rgba(255,255,255,22); border: none;")
        lay.addWidget(self.sep)

        self.list = QListWidget()
        self.list.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.list.setFrameShape(QFrame.Shape.NoFrame)
        self.list.setStyleSheet("QListWidget{background:transparent;border:none;outline:0;}"
                                "QListWidget::item{border-radius:12px;}"
                                "QListWidget::item:selected{background:rgba(255,255,255,26);}")
        self.list.itemClicked.connect(lambda _: self.run_current())
        lay.addWidget(self.list)

        self.setFixedWidth(self.PW)

    def _index_apps(self):
        found = scan_start_menu() + scan_start_apps()
        seen = {a["name"].lower() for a in self.apps}
        extra = []
        for a in found:
            if a["name"].lower() not in seen:
                seen.add(a["name"].lower())
                extra.append(a)
        self.apps = self.apps + extra

    def paintEvent(self, e):
        p = QPainter(self)
        draw_surface(p, self.width(), self.height(), 26)

    def event(self, e):
        if e.type() == QEvent.Type.WindowDeactivate and self.isVisible():
            self.hide()
        return super().event(e)

    def eventFilter(self, obj, ev):
        if obj is self.edit and ev.type() == QEvent.Type.KeyPress:
            k = ev.key()
            if k == Qt.Key.Key_Down:
                self._move(1)
                return True
            if k == Qt.Key.Key_Up:
                self._move(-1)
                return True
            if k == Qt.Key.Key_Escape:
                self.hide()
                return True
        return super().eventFilter(obj, ev)

    def _move(self, d):
        n = self.list.count()
        if n:
            self.list.setCurrentRow((self.list.currentRow() + d) % n)

    def toggle(self):
        if self.isVisible():
            self.hide()
            return
        self.edit.clear()
        self.refresh("")
        self.show()
        self.raise_()
        self.activateWindow()
        try:
            ctypes.windll.user32.SetForegroundWindow(int(self.winId()))
        except Exception:
            pass
        self.edit.setFocus()

    # ----- асинхронные источники -----
    def _fetch_rate(self, base):
        if base in self._fetching_rate:
            return
        self._fetching_rate.add(base)
        th = RateThread(base)
        th.done.connect(self._on_rate)
        self.rate_threads.append(th)
        th.finished.connect(lambda t=th: self.rate_threads.remove(t) if t in self.rate_threads else None)
        th.start()

    def _on_rate(self, base, data):
        self._fetching_rate.discard(base)
        if data:
            self.rate_cache[base] = (data, time.time())
            self.refresh(self.edit.text())

    def _fetch_translate(self, text, target):
        key = (text, target)
        if key in self._fetching_translate:
            return
        self._fetching_translate.add(key)
        th = TranslateThread(text, target)
        th.done.connect(self._on_translate)
        self.translate_threads.append(th)
        th.finished.connect(lambda t=th: self.translate_threads.remove(t) if t in self.translate_threads else None)
        th.start()

    def _on_translate(self, text, target, translated):
        self._fetching_translate.discard((text, target))
        if translated:
            self.translate_cache[(text, target)] = translated
            self.refresh(self.edit.text())

    def _fetch_weather(self, city):
        key = city.lower()
        if key in self._fetching_weather:
            return
        self._fetching_weather.add(key)
        th = WeatherThread(city)
        th.done.connect(self._on_weather)
        self.weather_threads.append(th)
        th.finished.connect(lambda t=th: self.weather_threads.remove(t) if t in self.weather_threads else None)
        th.start()

    def _on_weather(self, city, text):
        self._fetching_weather.discard(city.lower())
        if text:
            self.weather_cache[city.lower()] = (text, time.time())
            self.refresh(self.edit.text())

    # ----- результаты -----
    def _launch_app(self, a):
        self.store.bump_launch(a["name"])
        try:
            if a["kind"] == "appid":
                subprocess.Popen(["explorer.exe", "shell:AppsFolder\\" + a["target"]])
            else:
                os.startfile(a["target"])
        except Exception:
            self.island.flash_msg("Не удалось запустить")

    def _build(self, q):
        res = []
        ql = q.lower()
        if not q:
            counts = self.store.launch_counts()
            top = sorted(self.apps, key=lambda a: (-counts.get(a["name"], 0), a["name"].lower()))[:6]
            for a in top:
                res.append(Result("\uE71D", a["name"], "Приложение", lambda a=a: self._launch_app(a)))
            return res

        # таймер
        m = re.match(r"^(t|timer|таймер|т)\s+(.+)$", q, re.I)
        dur = parse_duration(m.group(2)) if m else None
        if not dur and re.fullmatch(r"\d+\s*[a-zа-яё]+(\s*\d+\s*[a-zа-яё]+)*", q, re.I):
            dur = parse_duration(q)
        if dur:
            res.append(Result("\uE121", f"Таймер на {fmt_clock(dur)}", "Enter — запустить",
                              lambda d=dur: self.island.timers.add(d)))

        # калькулятор
        v = calc(q)
        if v is not None:
            res.append(Result("\uE8EF", f"= {v}", "Enter — скопировать результат",
                              lambda v=v: self.island.copy_text(v)))

        # валюты / единицы измерения: "100 usd в eur", "5 км in mile"
        m = re.match(r"^(\d+(?:[.,]\d+)?)\s*([a-zа-яё°$€£]+)\s*(?:в|to|->|=|in)?\s*([a-zа-яё°$€£]+)$", ql, re.I)
        if m:
            amount = float(m.group(1).replace(",", "."))
            u1raw, u2raw = m.group(2), m.group(3)
            c1 = CURRENCY_ALIASES.get(u1raw.lower())
            c2 = CURRENCY_ALIASES.get(u2raw.lower())
            if c1 and c2 and HAVE_REQUESTS:
                cached = self.rate_cache.get(c1)
                if cached and time.time() - cached[1] < 3600 and c2 in cached[0]:
                    val = amount * cached[0][c2]
                    res.append(Result("\uE8EF", f"{amount:g} {c1} = {val:.2f} {c2}", "Enter — скопировать",
                                      lambda v=f"{val:.2f}": self.island.copy_text(v)))
                else:
                    res.append(Result("\uE8EF", f"{amount:g} {c1} → {c2}", "Загружаю курс…", lambda: None))
                    self._fetch_rate(c1)
            else:
                conv = convert_units(amount, u1raw, u2raw)
                if conv is not None:
                    val = f"{conv:.4g}"
                    res.append(Result("\uE9E9", f"{amount:g} {u1raw} = {val} {u2raw}", "Enter — скопировать",
                                      lambda v=val: self.island.copy_text(v)))

        # перевод: "t текст" / "т текст" / "пер текст"
        m = re.match(r"^(t|т|пер|translate)\s+(.+)$", q, re.I)
        if m and HAVE_REQUESTS:
            text = m.group(2)
            target = "en" if not re.search("[а-яё]", text, re.I) else "ru"
            cached = self.translate_cache.get((text, target))
            if cached:
                res.append(Result("\uE8C1", cached, text, lambda v=cached: self.island.copy_text(v)))
            else:
                res.append(Result("\uE8C1", "Перевожу…", text, lambda: None))
                self._fetch_translate(text, target)

        # погода: "погода" / "погода берлин"
        m = re.match(r"^(погода|weather)(\s+(.+))?$", q, re.I)
        if m and HAVE_REQUESTS:
            city = (m.group(3) or "").strip()
            cached = self.weather_cache.get(city.lower())
            if cached and time.time() - cached[1] < 600:
                res.append(Result("\uE9C4", cached[0], "Погода", lambda: None))
            else:
                res.append(Result("\uE9C4", "Узнаю погоду…", city or "Текущее местоположение", lambda: None))
                self._fetch_weather(city)

        # поиск файлов по индексу Windows: "f имя" / "файл имя"
        m = re.match(r"^(f|файл|find)\s+(.+)$", q, re.I)
        if m:
            found = search_windows_index(m.group(2))
            if found is None:
                res.append(Result("\uE7C3", "Поиск файлов недоступен", "Установите pywin32: pip install pywin32",
                                  lambda: None))
            elif not found:
                res.append(Result("\uE7C3", "Ничего не найдено", m.group(2), lambda: None))
            else:
                for name, path in found:
                    res.append(Result("\uE8A5", name, path, lambda p=path: os.startfile(p)))

        # выполнить в cmd
        if q.startswith(">") and q[1:].strip():
            cmd = q[1:].strip()
            res.append(Result("\uE756", f"Выполнить: {cmd}", "в командной строке",
                              lambda c=cmd: subprocess.Popen(f'start "" cmd /k {c}', shell=True)))

        # поиск по буферу обмена
        m = re.match(r"^(c|к|clip|буфер)\s+(.+)$", q, re.I)
        if m:
            for _, text, ts, _p, kind, path, label in self.store.clips(m.group(2), today=False, limit=6):
                if kind == "image":
                    one, sub_extra = (label or "Изображение"), "картинка"
                else:
                    one, sub_extra = " ".join(text.split())[:90], None
                when = datetime.datetime.fromtimestamp(ts).strftime("%d.%m %H:%M")
                sub = f"Из буфера · {when}" + (f" · {sub_extra}" if sub_extra else "") + " · Enter — скопировать"
                if kind == "image":
                    res.append(Result("\uE8B9", one, sub, lambda p=path: self.island.copy_image(p)))
                else:
                    res.append(Result("\uE8C8", one, sub, lambda t=text: self.island.copy_text(t)))

        # путь / ссылка
        path = os.path.expandvars(q.strip('"'))
        if os.path.exists(path):
            res.append(Result("\uE8B7", path, "Открыть", lambda p=path: os.startfile(p)))
        elif re.match(r"^(https?://|www\.)\S+$", q, re.I) or re.match(r"^\S+\.(com|ru|org|net|io|dev|de)(/\S*)?$", q, re.I):
            url = q if q.lower().startswith("http") else "https://" + q
            res.append(Result("\uE774", url, "Открыть в браузере", lambda u=url: webbrowser.open(u)))

        # приложения
        counts = self.store.launch_counts()
        scored = []
        for a in self.apps:
            s = match_score(a["name"], ql)
            if s is not None:
                scored.append((s + min(counts.get(a["name"], 0), 20), a))
        scored.sort(key=lambda x: -x[0])
        for _, a in scored[:6]:
            res.append(Result("\uE71D", a["name"], "Приложение", lambda a=a: self._launch_app(a)))

        # плагины
        for glyph, title, sub, action in self.island.plugins.launcher_results(q):
            res.append(Result(glyph, title, sub, action))

        # системные команды
        rec_title = "Остановить запись экрана" if self.island.rec_proc else "Начать запись экрана"
        system = [
            (rec_title, ["rec", "record", "запись", "запис"], "\uE714", self.island.toggle_recording),
            ("Заблокировать компьютер", ["lock", "блок", "заблок"], "\uE72E",
             lambda: ctypes.windll.user32.LockWorkStation()),
            ("Сделать скриншот", ["screenshot", "скрин", "снимок"], "\uE722", self.island.screenshot_full),
            ("Открыть папку загрузок", ["downloads", "загруз"], "\uE896",
             lambda: os.startfile(str(Path.home() / "Downloads"))),
            ("Закрыть Dynamic Island", ["exit", "quit", "выход", "закрыть"], "\uE711", QApplication.quit),
        ]
        for title, keys, glyph, action in system:
            if any(k.startswith(ql) or ql in title.lower() for k in keys):
                res.append(Result(glyph, title, "Команда", action))

        url = "https://www.google.com/search?q=" + urllib.parse.quote_plus(q)
        res.append(Result("\uE721", f"Искать в интернете: {q}", "Google", lambda u=url: webbrowser.open(u)))
        return res[:9]

    def refresh(self, text):
        self.results = self._build(text.strip())
        self.list.clear()
        for r in self.results:
            item = QListWidgetItem()
            item.setSizeHint(QSize(0, self.ROW_H))
            self.list.addItem(item)
            self.list.setItemWidget(item, self._row(r))
        if self.results:
            self.list.setCurrentRow(0)
        self.list.setFixedHeight(len(self.results) * self.ROW_H)
        self.sep.setVisible(bool(self.results))
        self.list.setVisible(bool(self.results))
        self.adjustSize()
        self.setFixedWidth(self.PW)
        sg = QGuiApplication.primaryScreen().geometry()
        self.move(sg.x() + (sg.width() - self.PW) // 2, sg.y() + int(sg.height() * 0.18))

    def _row(self, r):
        w = QWidget()
        w.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(12, 0, 12, 0)
        lay.setSpacing(14)
        ic = QLabel(r.glyph)
        ic.setFont(icon_font(20))
        ic.setFixedWidth(28)
        ic.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(ic)
        col = QVBoxLayout()
        col.setSpacing(0)
        t = QLabel(QFontMetrics(QFont(TEXT_FONT, 12)).elidedText(r.title, Qt.TextElideMode.ElideRight, 520))
        t.setFont(QFont(TEXT_FONT, 12, QFont.Weight.DemiBold))
        s = QLabel(r.sub)
        s.setFont(QFont(TEXT_FONT, 9))
        s.setStyleSheet("color: rgba(255,255,255,120); background: transparent;")
        col.addStretch()
        col.addWidget(t)
        col.addWidget(s)
        col.addStretch()
        lay.addLayout(col, 1)
        return w

    def run_current(self):
        row = self.list.currentRow()
        if 0 <= row < len(self.results):
            action = self.results[row].action
            self.hide()
            QTimer.singleShot(60, action)


# ====================== Глобальный хоткей ======================
class HotkeyFilter(QAbstractNativeEventFilter):
    WM_HOTKEY = 0x0312

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def nativeEventFilter(self, event_type, message):
        try:
            msg = wintypes.MSG.from_address(int(message))
            if msg.message == self.WM_HOTKEY and msg.wParam == 1:
                QTimer.singleShot(0, self.callback)
                return True, 0
        except Exception:
            pass
        return False, 0


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    island = Island()
    island.show()
    launcher = Launcher(island)

    hotkey_filter = HotkeyFilter(launcher.toggle)
    app.installNativeEventFilter(hotkey_filter)
    if not ctypes.windll.user32.RegisterHotKey(None, 1, HOTKEY_MOD, HOTKEY_VK):
        QTimer.singleShot(800, lambda: island.flash_msg("Alt+Space занят другой программой", 3500))
    app.aboutToQuit.connect(lambda: ctypes.windll.user32.UnregisterHotKey(None, 1))
    app.aboutToQuit.connect(island.shutdown)

    sys.exit(app.exec())