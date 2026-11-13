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
ROLE_GENITIVE = {"Вихователь": "вихователя", "Асистент вихователя": "асистента вихователя",
                 "Помічник вихователя": "помічника вихователя"}

# Режимні моменти — як у паперовій формі щоденника.
# (ключ, назва, підказка, швидкі варіанти діяльності)
PERIODS = [
    ("morning", "Ранок", "Ранкова гімнастика, спостереження, бесіди, доручення",
     ["Ранкова гімнастика", "Спостереження", "Бесіди", "Доручення"]),
    ("education", "Освітня діяльність", "Заняття, освітні ситуації, індивідуальна робота",
     ["Заняття", "Індивідуальна робота", "Творча діяльність", "Музичне заняття"]),
    ("walk", "Прогулянка", "Спостереження в природі, рухливі ігри, трудові доручення",
     ["Спостереження в природі", "Рухливі ігри", "Трудові доручення", "Вільна гра"]),
    ("free", "Самостійна діяльність", "Ігри, малювання, конструювання за власним вибором",
     ["Ігри", "Малювання", "Конструювання", "Книжки"]),
]
PERIOD_TITLES = {key: title for key, title, _, _ in PERIODS}
TASK_MODES = ["Самостійно", "З допомогою", "За зразком"]
NEEDS = ["Допомога в навідних запитаннях", "Багаторазове повторення",
         "Більше наочності", "Більше часу на виконання"]
GAMES = ["Рухлива", "Дидактична", "Сюжетно-рольова"]

MONTHS_NOM = ["січень", "лютий", "березень", "квітень", "травень", "червень",
              "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"]
MONTHS_GEN = ["січня", "лютого", "березня", "квітня", "травня", "червня",
              "липня", "серпня", "вересня", "жовтня", "листопада", "грудня"]

IMAGE_EXT = (".jpg", ".jpeg", ".png", ".webp")
VIDEO_EXT = (".mp4", ".3gp", ".mkv", ".webm", ".mov")

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
# Дані одного режимного моменту
# ----------------------------------------------------------------------
def empty_period():
    return {"activities": [], "description": "", "tasks": {m: 0 for m in TASK_MODES},
            "difficulties": "", "needs": [], "needs_note": "", "interaction": "",
            "games": [], "games_note": "", "media": []}


def period_filled(p):
    if not p:
        return False
    return bool(p.get("activities") or p.get("description", "").strip()
                or any(p.get("tasks", {}).values()) or p.get("difficulties", "").strip()
                or p.get("needs") or p.get("needs_note", "").strip()
                or p.get("interaction", "").strip() or p.get("games")
                or p.get("games_note", "").strip() or p.get("media"))


def join_text(items, note=""):
    parts = list(items) + ([note.strip()] if note and note.strip() else [])
    return "; ".join(parts)


def tasks_text(tasks):
    return ", ".join(f"{m.lower()} — {tasks.get(m, 0)}" for m in TASK_MODES if tasks.get(m, 0))


def period_lines(p):
    """Короткий текстовий опис етапу: список (назва поля, значення)."""
    out = []
    desc = join_text(p.get("activities", []), p.get("description", ""))
    if desc:
        out.append(("Діяльність", desc))
    t = tasks_text(p.get("tasks", {}))
    if t:
        out.append(("Виконані завдання", t))
    if p.get("difficulties", "").strip():
        out.append(("Труднощі", p["difficulties"].strip()))
    needs = join_text(p.get("needs", []), p.get("needs_note", ""))
    if needs:
        out.append(("Потреби", needs))
    if p.get("interaction", "").strip():
        out.append(("Взаємодія з ровесниками", p["interaction"].strip()))
    games = join_text(p.get("games", []), p.get("games_note", ""))
    if games:
        out.append(("Ігрова діяльність", games))
    return out


# ----------------------------------------------------------------------
# Щоденник одного педагога
# ----------------------------------------------------------------------
class Diary:
    def __init__(self, user_dir):
        self.dir = user_dir
        self.media_dir = os.path.join(user_dir, "media")
        self.excel_dir = os.path.join(user_dir, "excel")
        self.reports_dir = os.path.join(user_dir, "reports")
        self.entries_dir = os.path.join(user_dir, "entries")
        for d in (self.media_dir, self.excel_dir, self.reports_dir, self.entries_dir):
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

    # --- записи: один запис = одна дитина за один день, 4 режимні моменти ---
    def _day_path(self, day):
        return os.path.join(self.entries_dir, f"{day}.json")

    def entries(self, day):
        return _read_json(self._day_path(day), [])

    def _save_day(self, day, entries):
        entries = [e for e in entries if any(period_filled(p) for p in e["periods"].values())]
        if entries:
            _write_json(self._day_path(day), entries)
        elif os.path.exists(self._day_path(day)):
            os.remove(self._day_path(day))
        self.write_day_excel(day)

    def get_entry(self, day, child_id):
        for e in self.entries(day):
            if e["child_id"] == child_id:
                return e
        return None

    def get_period(self, day, child_id, key):
        e = self.get_entry(day, child_id)
        data = empty_period()
        if e and key in e["periods"]:
            data.update(e["periods"][key])
            data["tasks"] = {**empty_period()["tasks"], **e["periods"][key].get("tasks", {})}
        return data

    def save_period(self, day, child_id, key, data):
        child = self.get_child(child_id)
        if child is None:
            raise ValueError("Оберіть дитину.")
        entries = self.entries(day)
        entry = next((e for e in entries if e["child_id"] == child_id), None)
        if entry is None:
            entry = {"id": uuid.uuid4().hex[:10], "child_id": child_id, "periods": {}}
            entries.append(entry)
        entry["child_name"] = child["name"]
        entry["updated"] = datetime.now().strftime("%H:%M")
        if period_filled(data):
            entry["periods"][key] = data
        else:
            entry["periods"].pop(key, None)
        self._save_day(day, entries)

    def delete_entry(self, day, entry_id):
        entries = self.entries(day)
        for e in entries:
            if e["id"] == entry_id:
                for p in e["periods"].values():
                    for name in p.get("media", []):
                        try:
                            os.remove(os.path.join(self.media_dir, name))
                        except OSError:
                            pass
        self._save_day(day, [e for e in entries if e["id"] != entry_id])

    def days(self):
        """Список (день, кількість дітей із записами), найновіші першими."""
        out = []
        for name in os.listdir(self.entries_dir):
            if name.endswith(".json"):
                day = name[:-5]
                n = len(self.entries(day))
                if n:
                    out.append((day, n))
        return sorted(out, reverse=True)

    def months(self):
        months = {day[:7] for day, _ in self.days()}
        months.add(today()[:7])
        return sorted(months, reverse=True)

    def child_month_entries(self, child_id, month):
        """Записи дитини за місяць: список (day, entry) у хронологічному порядку."""
        out = []
        for day, _ in sorted(self.days()):
            if day.startswith(month):
                e = self.get_entry(day, child_id)
                if e:
                    out.append((day, e))
        return out

    # --- Excel у формі паперового щоденника ---
    def excel_path(self, day):
        return os.path.join(self.excel_dir, f"schodennik_{day}.xlsx")

    @property
    def diary_title(self):
        who = ROLE_GENITIVE.get(self.profile.get("role", ""), "педагога")
        return f"Щоденник спостережень {who} ЗДО за дитиною з ООП"

    def write_day_excel(self, day):
        from openpyxl.styles import Border, Side
        from openpyxl.utils import get_column_letter
        path = self.excel_path(day)
        entries = sorted(self.entries(day), key=lambda e: e["child_name"].lower())
        if not entries:
            if os.path.exists(path):
                os.remove(path)
            return None

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = day
        ncols = 12
        last = get_column_letter(ncols)
        thin = Side(style="thin", color="7F8C8D")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        head_fill = PatternFill("solid", fgColor="2E7D60")
        sub_fill = PatternFill("solid", fgColor="D9EAD3")
        wrap_top = Alignment(vertical="top", wrap_text=True)
        center = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws["A1"] = self.diary_title
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws.merge_cells(f"A1:{last}1")
        ws["A2"] = f"Група: {self.group_title}"
        ws["A2"].alignment = Alignment(horizontal="center")
        ws.merge_cells(f"A2:{last}2")
        ws["A3"] = f"{self.profile['role']}: {self.profile['full_name']}    Дата: {format_day(day)}"
        ws["A3"].alignment = Alignment(horizontal="center")
        ws.merge_cells(f"A3:{last}3")

        # Двоярусна шапка, як у паперовій формі
        top = {1: "Дата", 2: "Прізвище, ім'я дитини", 3: "Режимні моменти",
               4: "Короткий опис діяльності", 5: "Характеристика діяльності дитини",
               10: "Поведінка", 12: "Фото / відео"}
        sub = {5: "Завдань самостійно", 6: "Завдань з допомогою", 7: "Завдань за зразком",
               8: "Труднощі",
               9: "Потреби (навідні запитання, повторення, наочність, час)",
               10: "Взаємодія з ровесниками",
               11: "Ігрова діяльність (рухлива, дидактична, сюжетно-рольова)"}
        for col in range(1, ncols + 1):
            for row in (5, 6):
                c = ws.cell(row=row, column=col)
                c.fill = head_fill if row == 5 else sub_fill
                c.font = Font(bold=True, color="FFFFFF" if row == 5 else "1E3A2F")
                c.alignment = center
                c.border = border
        for col, text in top.items():
            ws.cell(row=5, column=col, value=text)
        for col, text in sub.items():
            ws.cell(row=6, column=col, value=text)
        for col in (1, 2, 3, 4, 12):
            ws.merge_cells(start_row=5, start_column=col, end_row=6, end_column=col)
            ws.cell(row=5, column=col).fill = head_fill
        ws.merge_cells("E5:I5")
        ws.merge_cells("J5:K5")

        row = 7
        for e in entries:
            first = row
            for key, title, _, _ in PERIODS:
                p = {**empty_period(), **e["periods"].get(key, {})}
                tasks = p.get("tasks", {})
                values = [
                    format_day(day), e["child_name"], title,
                    join_text(p["activities"], p["description"]),
                    tasks.get("Самостійно") or None, tasks.get("З допомогою") or None,
                    tasks.get("За зразком") or None,
                    p["difficulties"].strip(), join_text(p["needs"], p["needs_note"]),
                    p["interaction"].strip(), join_text(p["games"], p["games_note"]),
                    ", ".join(p["media"]),
                ]
                for col, v in enumerate(values, start=1):
                    c = ws.cell(row=row, column=col, value=v if v != "" else None)
                    c.alignment = center if col in (5, 6, 7) else wrap_top
                    c.border = border
                ws.cell(row=row, column=3).font = Font(bold=True)
                row += 1
            for col in (1, 2):
                ws.merge_cells(start_row=first, start_column=col, end_row=row - 1, end_column=col)
                ws.cell(row=first, column=col).alignment = Alignment(vertical="center", wrap_text=True)
            ws.cell(row=first, column=2).font = Font(bold=True)

        widths = [13, 22, 16, 40, 12, 12, 12, 28, 32, 28, 30, 22]
        for idx, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = w
        ws.row_dimensions[6].height = 48
        ws.freeze_panes = "D7"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        wb.save(path)
        return path
