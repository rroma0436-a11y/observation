"""Локальне зберігання даних: акаунти, профіль педагога, діти, записи, щоденні Excel-файли.

Усе зберігається у приватній папці застосунку (працює повністю офлайн).
"""
import os
import re
import json
import uuid
import hashlib
import secrets
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

ROLES = ["Вихователь", "Асистент вихователя", "Помічник вихователя"]
CRITERIA = ["Мовлення", "Соціалізація", "Дрібна моторика", "Фізична активність"]

MONTHS_NOM = ["січень", "лютий", "березень", "квітень", "травень", "червень",
              "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"]
MONTHS_GEN = ["січня", "лютого", "березня", "квітня", "травня", "червня",
              "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]

IMAGE_EXT = (".jpg", ".jpeg", ".png", ".webp")
VIDEO_EXT = (".mp4", ".3gp", ".mkv", ".webm", ".mov")

ALL_GROUP_ID = ""  # запис для всієї групи


# ----------------------------------------------------------------------
# Допоміжні функції
# ----------------------------------------------------------------------
def plural(n, one, few, many):
    """Українські форми множини: 1 запис, 2 записи, 5 записів."""
    n = abs(int(n))
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


def format_day(day):
    """'2026-11-13' -> '13 листопада 2026'"""
    d = datetime.strptime(day, "%Y-%m-%d")
    return f"{d.day} {MONTHS_GEN[d.month - 1]} {d.year}"


def format_day_short(day):
    """'2026-11-13' -> '13 листопада'"""
    d = datetime.strptime(day, "%Y-%m-%d")
    return f"{d.day} {MONTHS_GEN[d.month - 1]}"


def format_month(month):
    """'2026-11' -> 'листопад 2026'"""
    y, m = month.split("-")
    return f"{MONTHS_NOM[int(m) - 1]} {y}"


def today():
    return datetime.now().strftime("%Y-%m-%d")


def safe_filename(text):
    text = re.sub(r'[\\/:*?"<>|\n\r\t]+', " ", text).strip()
    return re.sub(r"\s+", "_", text)[:60] or "file"


def _read_json(path, default):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return default


def _write_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    os.replace(tmp, path)  # атомарна заміна — файл не пошкодиться при збої


# ----------------------------------------------------------------------
# Акаунти
# ----------------------------------------------------------------------
class UserStore:
    ITERATIONS = 100_000

    def __init__(self, root):
        self.root = root
        self.path = os.path.join(root, "users.json")
        os.makedirs(root, exist_ok=True)

    def _load(self):
        return _read_json(self.path, {"users": {}, "last_login": ""})

    @staticmethod
    def normalize(login):
        return login.strip().lower()

    def _hash(self, password, salt_hex):
        return hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                                   bytes.fromhex(salt_hex), self.ITERATIONS).hex()

    def register(self, login, password, password2):
        login_n = self.normalize(login)
        if not re.fullmatch(r"[\w.\-]{3,32}", login_n):
            raise ValueError("Логін: від 3 до 32 символів — літери, цифри, крапка, дефіс або _.")
        if len(password) < 6:
            raise ValueError("Пароль має містити щонайменше 6 символів.")
        if password != password2:
            raise ValueError("Паролі не збігаються.")
        data = self._load()
        if login_n in data["users"]:
            raise ValueError("Такий логін уже зареєстровано на цьому пристрої.")
        salt = secrets.token_hex(16)
        data["users"][login_n] = {
            "salt": salt,
            "hash": self._hash(password, salt),
            "folder": uuid.uuid4().hex[:12],
            "created": datetime.now().isoformat(timespec="seconds"),
        }
        data["last_login"] = login_n
        _write_json(self.path, data)
        return login_n

    def verify(self, login, password):
        login_n = self.normalize(login)
        user = self._load()["users"].get(login_n)
        if not user:
            return False
        return secrets.compare_digest(user["hash"], self._hash(password, user["salt"]))

    def set_last_login(self, login):
        data = self._load()
        data["last_login"] = self.normalize(login)
        _write_json(self.path, data)

    def last_login(self):
        return self._load().get("last_login", "")

    def user_dir(self, login):
        user = self._load()["users"][self.normalize(login)]
        path = os.path.join(self.root, "accounts", user["folder"])
        os.makedirs(path, exist_ok=True)
        return path


# ----------------------------------------------------------------------
# Щоденник одного педагога
# ----------------------------------------------------------------------
class Diary:
    def __init__(self, user_dir):
        self.dir = user_dir
        self.media_dir = os.path.join(user_dir, "media")
        self.excel_dir = os.path.join(user_dir, "excel")
        self.reports_dir = os.path.join(user_dir, "reports")
        self.records_dir = os.path.join(user_dir, "records")
        for d in (self.media_dir, self.excel_dir, self.reports_dir, self.records_dir):
            os.makedirs(d, exist_ok=True)
        self.profile_path = os.path.join(user_dir, "profile.json")
        self.profile = _read_json(self.profile_path, {})
        self.profile.setdefault("full_name", "")
        self.profile.setdefault("role", "")
        self.profile.setdefault("group_number", "")
        self.profile.setdefault("group_spec", "")
        self.profile.setdefault("children", [])
        self.profile.setdefault("onboarded", False)

    # --- профіль ---
    def save_profile(self):
        _write_json(self.profile_path, self.profile)

    @property
    def group_title(self):
        parts = [self.profile["group_number"].strip(), self.profile["group_spec"].strip()]
        return ", ".join(p for p in parts if p)

    # --- діти ---
    @property
    def children(self):
        return sorted(self.profile["children"], key=lambda c: c["name"].lower())

    def get_child(self, child_id):
        for c in self.profile["children"]:
            if c["id"] == child_id:
                return c
        return None

    def add_child(self, name, age="", notes=""):
        name = name.strip()
        if not name:
            raise ValueError("Вкажіть прізвище та ім'я дитини.")
        child = {"id": uuid.uuid4().hex[:10], "name": name,
                 "age": age.strip(), "notes": notes.strip()}
        self.profile["children"].append(child)
        self.save_profile()
        return child

    def update_child(self, child_id, name, age, notes):
        child = self.get_child(child_id)
        if child is None:
            raise ValueError("Дитину не знайдено.")
        if not name.strip():
            raise ValueError("Вкажіть прізвище та ім'я дитини.")
        child.update(name=name.strip(), age=age.strip(), notes=notes.strip())
        self.save_profile()

    def remove_child(self, child_id):
        self.profile["children"] = [c for c in self.profile["children"] if c["id"] != child_id]
        self.save_profile()

    # --- записи ---
    def _day_path(self, day):
        return os.path.join(self.records_dir, f"{day}.json")

    def records(self, day):
        return _read_json(self._day_path(day), [])

    def add_record(self, child_id, criteria, note, media, day=None):
        day = day or today()
        if not note.strip() and not criteria and not media:
            raise ValueError("Додайте текст спостереження, критерій або медіафайл.")
        child = self.get_child(child_id) if child_id else None
        rec = {
            "id": uuid.uuid4().hex[:10],
            "time": datetime.now().strftime("%H:%M"),
            "child_id": child["id"] if child else ALL_GROUP_ID,
            "child_name": child["name"] if child else "Уся група",
            "criteria": [c for c in CRITERIA if c in criteria],
            "note": note.strip(),
            "media": list(media),
        }
        recs = self.records(day)
        recs.append(rec)
        _write_json(self._day_path(day), recs)
        self.write_day_excel(day)
        return rec

    def delete_record(self, day, rec_id):
        recs = self.records(day)
        removed = [r for r in recs if r["id"] == rec_id]
        recs = [r for r in recs if r["id"] != rec_id]
        for r in removed:
            for name in r.get("media", []):
                try:
                    os.remove(os.path.join(self.media_dir, name))
                except OSError:
                    pass
        if recs:
            _write_json(self._day_path(day), recs)
        else:
            try:
                os.remove(self._day_path(day))
            except OSError:
                pass
        self.write_day_excel(day)

    def days(self):
        """Список (день, кількість записів), найновіші першими."""
        out = []
        for name in os.listdir(self.records_dir):
            if name.endswith(".json"):
                day = name[:-5]
                n = len(self.records(day))
                if n:
                    out.append((day, n))
        return sorted(out, reverse=True)

    def months(self):
        months = {day[:7] for day, _ in self.days()}
        months.add(today()[:7])
        return sorted(months, reverse=True)

    def child_month_records(self, child_id, month):
        """Записи дитини за місяць: список (day, record) у хронологічному порядку."""
        out = []
        for day, _ in sorted(self.days()):
            if day.startswith(month):
                for r in self.records(day):
                    if r.get("child_id") == child_id:
                        out.append((day, r))
        return out

    # --- Excel ---
    def excel_path(self, day):
        return os.path.join(self.excel_dir, f"schodennik_{day}.xlsx")

    def write_day_excel(self, day):
        path = self.excel_path(day)
        recs = self.records(day)
        if not recs:
            if os.path.exists(path):
                os.remove(path)
            return None

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = day

        headers = ["№", "Час", "Дитина"] + CRITERIA + ["Спостереження", "Медіафайли"]
        last_col = openpyxl.utils.get_column_letter(len(headers))

        ws["A1"] = f"Щоденник спостережень — {format_day(day)}"
        ws["A1"].font = Font(bold=True, size=14)
        ws.merge_cells(f"A1:{last_col}1")
        ws["A2"] = (f"Педагог: {self.profile['full_name']} ({self.profile['role']})"
                    f"    Група: {self.group_title}")
        ws.merge_cells(f"A2:{last_col}2")

        ws.append([])
        ws.append(headers)
        head_fill = PatternFill("solid", fgColor="2E7D60")
        for cell in ws[4]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for i, r in enumerate(recs, start=1):
            ws.append([i, r["time"], r["child_name"]]
                      + ["✓" if c in r["criteria"] else "" for c in CRITERIA]
                      + [r["note"], ", ".join(r.get("media", []))])
            for cell in ws[ws.max_row]:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            for col in range(4, 4 + len(CRITERIA)):
                ws.cell(row=ws.max_row, column=col).alignment = Alignment(horizontal="center", vertical="top")

        widths = [5, 8, 26, 12, 14, 16, 18, 60, 30]
        for idx, w in enumerate(widths, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = w
        ws.freeze_panes = "A5"
        wb.save(path)
        return path
