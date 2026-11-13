"""Щоденник спостережень — офлайн-застосунок для педагогів ЗДО."""
import os
import threading
from datetime import datetime, timedelta

from kivy.app import App
from kivy.clock import Clock, mainthread
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle, RoundedRectangle
from kivy.metrics import dp, sp
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.modalview import ModalView
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import NoTransition, Screen, ScreenManager
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner, SpinnerOption
from kivy.uix.textinput import TextInput
from kivy.uix.togglebutton import ToggleButton
from kivy.uix.widget import Widget
from kivy.utils import escape_markup

import android_bridge as ab
import pdf_report
from storage import (GAMES, IMAGE_EXT, NEEDS, PERIODS, ROLES, TASK_MODES, Diary, UserStore, format_day,
                     format_day_short, format_month, period_filled, period_lines, plural,
                     safe_filename, today)

# ----------------------------------------------------------------------
# Розміри (dp/sp масштабуються під щільність екрана телефону)
# ----------------------------------------------------------------------
BTN_H = dp(60)
BTN_H_BIG = dp(68)
BTN_H_SMALL = dp(50)
INPUT_H = dp(58)
FS = sp(19)
FS_SMALL = sp(16)
FS_TITLE = sp(25)

BG = (0.11, 0.13, 0.15, 1)
CARD = (0.18, 0.21, 0.24, 1)
CARD_LIGHT = (0.24, 0.28, 0.32, 1)
PRIMARY = (0.18, 0.49, 0.38, 1)
ACCENT = (0.85, 0.45, 0.22, 1)
SECOND = (0.29, 0.35, 0.42, 1)
DANGER = (0.70, 0.25, 0.22, 1)
TEXT = (1, 1, 1, 1)
MUTED = (0.72, 0.76, 0.80, 1)
GREEN_TEXT = (0.55, 0.86, 0.70, 1)

EXPORT_ROOT = "Shodennyk"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def darker(c, k=0.72):
    return (c[0] * k, c[1] * k, c[2] * k, c[3])


# ----------------------------------------------------------------------
# Базові віджети з великими розмірами
# ----------------------------------------------------------------------
class Btn(Button):
    def __init__(self, bg=PRIMARY, **kw):
        opts = dict(size_hint_y=None, height=BTN_H, font_size=FS, background_normal="",
                    background_down="", halign="center", valign="middle")
        if "size_hint" in kw:
            opts.pop("size_hint_y")
        opts.update(kw)
        super().__init__(**opts)
        self._bg = bg
        self.bind(state=self._paint, size=self._wrap)
        self._paint()

    def _paint(self, *args):
        self.background_color = darker(self._bg) if self.state == "down" else self._bg

    def _wrap(self, *args):
        self.text_size = (max(self.width - dp(14), 1), None)


class ToggleBtn(ToggleButton):
    def __init__(self, bg_off=CARD_LIGHT, bg_on=PRIMARY, **kw):
        opts = dict(size_hint_y=None, height=BTN_H, font_size=FS, background_normal="",
                    background_down="", halign="center", valign="middle")
        if "size_hint" in kw:
            opts.pop("size_hint_y")
        opts.update(kw)
        super().__init__(**opts)
        self._off, self._on = bg_off, bg_on
        self.bind(state=self._paint, size=self._wrap)
        self._paint()

    def _paint(self, *args):
        self.background_color = self._on if self.state == "down" else self._off

    def _wrap(self, *args):
        self.text_size = (max(self.width - dp(14), 1), None)


class Lbl(Label):
    """Текст із перенесенням рядків; висота підлаштовується під текст."""
    def __init__(self, **kw):
        opts = dict(size_hint_y=None, font_size=FS, halign="left", valign="middle", color=TEXT)
        opts.update(kw)
        super().__init__(**opts)
        self.bind(width=self._wrap, texture_size=self._fit)
        self._wrap()

    def _wrap(self, *args):
        self.text_size = (self.width, None)

    def _fit(self, *args):
        self.height = self.texture_size[1] + dp(6)


class Panel(BoxLayout):
    """BoxLayout із кольоровим тлом."""
    def __init__(self, bg=CARD, radius=0, **kw):
        super().__init__(**kw)
        with self.canvas.before:
            Color(*bg)
            if radius:
                self._bg = RoundedRectangle(pos=self.pos, size=self.size, radius=[radius])
            else:
                self._bg = Rectangle(pos=self.pos, size=self.size)
        self.bind(pos=self._upd, size=self._upd)

    def _upd(self, *args):
        self._bg.pos = self.pos
        self._bg.size = self.size


def card():
    c = Panel(bg=CARD, radius=dp(12), orientation="vertical", size_hint_y=None,
              padding=dp(14), spacing=dp(8))
    c.bind(minimum_height=c.setter("height"))
    return c


def make_input(hint="", multiline=False, height=None, password=False):
    h = height or INPUT_H
    pad_v = dp(12) if multiline else max(dp(6), (h - FS * 1.3) / 2)
    return TextInput(hint_text=hint, multiline=multiline, password=password,
                     size_hint_y=None, height=h, font_size=FS, write_tab=False,
                     padding=[dp(12), pad_v, dp(12), pad_v])


class VoiceField(BoxLayout):
    """Поле вводу з кнопкою «Голос» праворуч."""
    def __init__(self, hint="", multiline=False, height=None, **kw):
        h = height or INPUT_H
        super().__init__(orientation="horizontal", size_hint_y=None, height=h, spacing=dp(8), **kw)
        self.input = make_input(hint, multiline, h)
        self.add_widget(self.input)
        mic = Btn(text="Голос", bg=ACCENT, size_hint_x=None, width=dp(92), height=h)
        mic.bind(on_release=lambda *_: App.get_running_app().voice_into(self.input, append=multiline))
        self.add_widget(mic)

    @property
    def text(self):
        return self.input.text.strip()

    @text.setter
    def text(self, value):
        self.input.text = value


class BigOption(SpinnerOption):
    def __init__(self, **kw):
        super().__init__(**kw)
        self.height = dp(58)
        self.font_size = FS
        self.background_normal = ""
        self.background_color = CARD_LIGHT
        self.halign = "center"
        self.bind(size=lambda *_: setattr(self, "text_size", (self.width - dp(14), None)))


def make_spinner(values=(), text=""):
    s = Spinner(text=text, values=list(values), option_cls=BigOption, size_hint_y=None,
                height=BTN_H, font_size=FS, background_normal="", background_color=CARD_LIGHT,
                halign="center", valign="middle")
    s.bind(size=lambda *_: setattr(s, "text_size", (s.width - dp(14), None)))
    return s


def scroll_column(padding=None, spacing=None):
    sv = ScrollView(do_scroll_x=False, bar_width=dp(6))
    col = BoxLayout(orientation="vertical", size_hint_y=None,
                    padding=dp(16) if padding is None else padding,
                    spacing=dp(12) if spacing is None else spacing)
    col.bind(minimum_height=col.setter("height"))
    sv.add_widget(col)
    return sv, col


def section(text):
    return Lbl(text=text, bold=True, font_size=sp(20))


def hint(text):
    return Lbl(text=text, color=MUTED, font_size=FS_SMALL)


def spacer(h=dp(8)):
    return Widget(size_hint_y=None, height=h)


# ----------------------------------------------------------------------
# Спливаючі вікна
# ----------------------------------------------------------------------
@mainthread
def show_message(text, title="Щоденник спостережень"):
    box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(12))
    sv, col = scroll_column(padding=dp(4))
    col.add_widget(Lbl(text=text, halign="center"))
    box.add_widget(sv)
    popup = Popup(title=title, title_size=FS, title_align="center", content=box,
                  size_hint=(0.92, None), height=dp(340), separator_color=PRIMARY)
    box.add_widget(Btn(text="Гаразд", on_release=popup.dismiss))
    popup.open()


def confirm(text, on_yes, yes_text="Так", danger=True):
    box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(12))
    box.add_widget(Lbl(text=text, halign="center"))
    box.add_widget(Widget())
    popup = Popup(title="Підтвердження", title_size=FS, title_align="center", content=box,
                  size_hint=(0.92, None), height=dp(300), separator_color=PRIMARY)
    row = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=BTN_H)
    row.add_widget(Btn(text="Скасувати", bg=SECOND, on_release=popup.dismiss))

    def yes(*_):
        popup.dismiss()
        on_yes()
    row.add_widget(Btn(text=yes_text, bg=DANGER if danger else PRIMARY, on_release=yes))
    box.add_widget(row)
    popup.open()


def open_day_preview(day):
    """Попередній перегляд усіх записів за день."""
    app = App.get_running_app()
    entries = sorted(app.diary.entries(day), key=lambda e: e["child_name"].lower())
    box = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(4))
    sv, col = scroll_column(padding=dp(4), spacing=dp(10))
    popup = Popup(title=f"Записи за {format_day(day)}", title_size=FS, title_align="center",
                  content=box, size_hint=(0.96, 0.92), separator_color=PRIMARY)

    if not entries:
        col.add_widget(Lbl(text="За цей день записів ще немає.", halign="center"))
    for e in entries:
        c = card()
        c.add_widget(Lbl(text=e["child_name"], bold=True, font_size=sp(21)))
        for key, title, _, _ in PERIODS:
            p = e["periods"].get(key)
            if not period_filled(p):
                continue
            c.add_widget(Lbl(text=title, bold=True, color=GREEN_TEXT))
            for label, value in period_lines(p):
                c.add_widget(Lbl(text=f"{label}: {value}", font_size=FS_SMALL))
            if p.get("media"):
                n = len(p["media"])
                c.add_widget(hint(f"{n} {plural(n, 'медіафайл', 'медіафайли', 'медіафайлів')}"))

        def edit(*_, cid=e["child_id"]):
            popup.dismiss()
            app.open_period(day, cid, 0)

        def ask_delete(*_, eid=e["id"], name=e["child_name"]):
            def do_delete():
                app.diary.delete_entry(day, eid)
                popup.dismiss()
                app.refresh_tabs()
                open_day_preview(day)
            confirm(f"Видалити весь запис про «{name}» за цей день разом із фото та відео?",
                    do_delete, "Видалити")
        row = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=BTN_H_SMALL)
        b_edit = Btn(text="Редагувати", bg=SECOND, height=BTN_H_SMALL, font_size=FS_SMALL, on_release=edit)
        b_edit.disabled = app.diary.get_child(e["child_id"]) is None
        row.add_widget(b_edit)
        row.add_widget(Btn(text="Видалити", bg=DANGER, height=BTN_H_SMALL, font_size=FS_SMALL,
                           on_release=ask_delete))
        c.add_widget(row)
        col.add_widget(c)
    box.add_widget(sv)

    row = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=BTN_H)
    send = Btn(text="Надіслати Excel", bg=ACCENT, on_release=lambda *_: app.share_day(day))
    send.disabled = not entries
    row.add_widget(send)
    row.add_widget(Btn(text="Закрити", bg=SECOND, on_release=popup.dismiss))
    box.add_widget(row)
    popup.open()


def recent_days(n=14):
    """{'Сьогодні, 13 листопада': '2026-11-13', ...} за останні n днів."""
    out = {}
    now = datetime.now()
    for i in range(n):
        d = (now - timedelta(days=i)).strftime("%Y-%m-%d")
        prefix = "Сьогодні, " if i == 0 else ("Вчора, " if i == 1 else "")
        out[prefix + format_day_short(d)] = d
    return out


# ----------------------------------------------------------------------
# Вхід і реєстрація
# ----------------------------------------------------------------------
class LoginScreen(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="login", **kw)
        self.app = app
        sv, col = scroll_column(padding=dp(24), spacing=dp(14))
        col.add_widget(spacer(dp(30)))
        col.add_widget(Lbl(text="Щоденник спостережень", font_size=FS_TITLE, bold=True, halign="center"))
        col.add_widget(hint("Вхід до облікового запису"))
        col.add_widget(spacer())
        col.add_widget(Lbl(text="Логін"))
        self.login = make_input("Ваш логін")
        col.add_widget(self.login)
        col.add_widget(Lbl(text="Пароль"))
        self.password = make_input("Пароль", password=True)
        col.add_widget(self.password)
        col.add_widget(spacer())
        col.add_widget(Btn(text="Увійти", height=BTN_H_BIG, bold=True, on_release=self.do_login))
        col.add_widget(Btn(text="Створити обліковий запис", bg=SECOND,
                           on_release=lambda *_: setattr(app.sm, "current", "register")))
        self.add_widget(sv)

    def on_pre_enter(self, *args):
        self.login.text = self.app.users.last_login()
        self.password.text = ""

    def do_login(self, *_):
        if not self.login.text.strip() or not self.password.text:
            show_message("Введіть логін і пароль.")
            return
        if not self.app.users.verify(self.login.text, self.password.text):
            show_message("Неправильний логін або пароль.")
            return
        self.app.open_account(self.login.text)


class RegisterScreen(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="register", **kw)
        self.app = app
        sv, col = scroll_column(padding=dp(24), spacing=dp(14))
        col.add_widget(spacer(dp(20)))
        col.add_widget(Lbl(text="Новий обліковий запис", font_size=FS_TITLE, bold=True, halign="center"))
        col.add_widget(hint("Дані зберігаються лише на цьому телефоні. Пароль захищає їх від сторонніх."))
        col.add_widget(Lbl(text="Логін"))
        self.login = make_input("Латиницею або кирилицею, від 3 символів")
        col.add_widget(self.login)
        col.add_widget(Lbl(text="Пароль"))
        self.password = make_input("Щонайменше 6 символів", password=True)
        col.add_widget(self.password)
        col.add_widget(Lbl(text="Повторіть пароль"))
        self.password2 = make_input("Ще раз пароль", password=True)
        col.add_widget(self.password2)
        col.add_widget(spacer())
        col.add_widget(Btn(text="Зареєструватися", height=BTN_H_BIG, bold=True, on_release=self.do_register))
        col.add_widget(Btn(text="Назад до входу", bg=SECOND,
                           on_release=lambda *_: setattr(app.sm, "current", "login")))
        self.add_widget(sv)

    def on_pre_enter(self, *args):
        self.login.text = self.password.text = self.password2.text = ""

    def do_register(self, *_):
        try:
            login = self.app.users.register(self.login.text, self.password.text, self.password2.text)
        except ValueError as e:
            show_message(str(e))
            return
        self.app.open_account(login)


# ----------------------------------------------------------------------
# Редактор списку дітей (у майстрі та на вкладці «Діти»)
# ----------------------------------------------------------------------
class ChildrenEditor(BoxLayout):
    def __init__(self, app, **kw):
        super().__init__(orientation="vertical", **kw)
        self.app = app
        self.editing_id = None
        sv, col = scroll_column()
        self.add_widget(sv)

        self.form_title = section("Додати дитину")
        col.add_widget(self.form_title)
        col.add_widget(hint("Прізвище та ім'я"))
        self.f_name = VoiceField("Наприклад: Шевченко Степан")
        col.add_widget(self.f_name)
        col.add_widget(hint("Вік"))
        self.f_age = VoiceField("Наприклад: 4 роки")
        col.add_widget(self.f_age)
        col.add_widget(hint("Примітки / особливості (бачите лише ви)"))
        self.f_notes = VoiceField("Алергії, особливості розвитку тощо", multiline=True, height=dp(110))
        col.add_widget(self.f_notes)
        self.btn_save = Btn(text="Додати дитину", on_release=self.save)
        col.add_widget(self.btn_save)
        self.cancel_slot = BoxLayout(size_hint_y=None, height=0)
        self.btn_cancel = Btn(text="Скасувати редагування", bg=SECOND, on_release=lambda *_: self.reset_form())
        col.add_widget(self.cancel_slot)

        col.add_widget(spacer())
        self.list_title = section("Список групи")
        col.add_widget(self.list_title)
        self.list_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10))
        self.list_box.bind(minimum_height=self.list_box.setter("height"))
        col.add_widget(self.list_box)
        self.refresh()

    def reset_form(self):
        self.editing_id = None
        self.f_name.text = self.f_age.text = self.f_notes.text = ""
        self.form_title.text = "Додати дитину"
        self.btn_save.text = "Додати дитину"
        self.cancel_slot.clear_widgets()
        self.cancel_slot.height = 0

    def save(self, *_):
        diary = self.app.diary
        try:
            if self.editing_id:
                diary.update_child(self.editing_id, self.f_name.text, self.f_age.text, self.f_notes.text)
            else:
                diary.add_child(self.f_name.text, self.f_age.text, self.f_notes.text)
        except ValueError as e:
            show_message(str(e))
            return
        self.reset_form()
        self.refresh()

    def edit(self, child):
        self.editing_id = child["id"]
        self.f_name.text = child["name"]
        self.f_age.text = child.get("age", "")
        self.f_notes.text = child.get("notes", "")
        self.form_title.text = "Редагування"
        self.btn_save.text = "Зберегти зміни"
        if not self.cancel_slot.children:
            self.cancel_slot.add_widget(self.btn_cancel)
            self.cancel_slot.height = BTN_H
        self.parent_scroll_to_top()

    def parent_scroll_to_top(self):
        sv = self.children[-1] if self.children else None
        if isinstance(sv, ScrollView):
            sv.scroll_y = 1

    def remove(self, child):
        def do_remove():
            self.app.diary.remove_child(child["id"])
            if self.editing_id == child["id"]:
                self.reset_form()
            self.refresh()
        confirm(f"Видалити «{child['name']}» зі списку групи?\nЗаписи про дитину залишаться в архіві.",
                do_remove, "Видалити")

    def refresh(self):
        self.list_box.clear_widgets()
        kids = self.app.diary.children if self.app.diary else []
        self.list_title.text = f"Список групи ({len(kids)})"
        if not kids:
            self.list_box.add_widget(hint("Поки що порожньо. Додайте дітей через форму вище."))
        for child in kids:
            c = card()
            c.add_widget(Lbl(text=child["name"], bold=True))
            details = " · ".join(x for x in (child.get("age", ""), child.get("notes", "")) if x)
            if details:
                c.add_widget(hint(details))
            row = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=BTN_H_SMALL)
            row.add_widget(Btn(text="Змінити", bg=SECOND, height=BTN_H_SMALL, font_size=FS_SMALL,
                               on_release=lambda *_, ch=child: self.edit(ch)))
            row.add_widget(Btn(text="Видалити", bg=DANGER, height=BTN_H_SMALL, font_size=FS_SMALL,
                               on_release=lambda *_, ch=child: self.remove(ch)))
            c.add_widget(row)
            self.list_box.add_widget(c)


# ----------------------------------------------------------------------
# Майстер налаштування
# ----------------------------------------------------------------------
class WizardScreen(Screen):
    TITLES = ["Про вас", "Ваша група", "Діти групи"]

    def __init__(self, app, **kw):
        super().__init__(name="wizard", **kw)
        self.app = app
        self.editing = False
        self.step = 0
        root = BoxLayout(orientation="vertical")
        header = Panel(bg=PRIMARY, orientation="vertical", size_hint_y=None, height=dp(78),
                       padding=[dp(16), dp(8)])
        self.title = Label(font_size=sp(21), bold=True, halign="left", valign="middle", markup=True)
        self.title.bind(size=lambda *_: setattr(self.title, "text_size", self.title.size))
        header.add_widget(self.title)
        root.add_widget(header)
        self.body = BoxLayout()
        root.add_widget(self.body)
        nav = Panel(bg=CARD, orientation="horizontal", size_hint_y=None, height=BTN_H + dp(20),
                    padding=dp(10), spacing=dp(10))
        self.btn_back = Btn(text="Назад", bg=SECOND, on_release=lambda *_: self.go_back())
        self.btn_next = Btn(text="Далі", on_release=lambda *_: self.go_next())
        nav.add_widget(self.btn_back)
        nav.add_widget(self.btn_next)
        root.add_widget(nav)
        self.add_widget(root)

    def start(self, editing=False):
        self.editing = editing
        self.step = 0
        p = self.app.diary.profile

        # Крок 1 — про педагога
        sv1, c1 = scroll_column()
        c1.add_widget(section("Прізвище, ім'я та по батькові"))
        self.f_name = VoiceField("Наприклад: Коваленко Олена Петрівна")
        self.f_name.text = p["full_name"]
        c1.add_widget(self.f_name)
        c1.add_widget(spacer())
        c1.add_widget(section("Посада"))
        self.role_btns = []
        for role in ROLES:
            t = ToggleBtn(text=role, group="role", allow_no_selection=False)
            if role == p["role"]:
                t.state = "down"
            self.role_btns.append(t)
            c1.add_widget(t)

        # Крок 2 — група
        sv2, c2 = scroll_column()
        c2.add_widget(section("Номер групи"))
        self.f_group = VoiceField("Наприклад: Група №11")
        self.f_group.text = p["group_number"]
        c2.add_widget(self.f_group)
        c2.add_widget(spacer())
        c2.add_widget(section("Специфікація групи"))
        c2.add_widget(hint("Вікова категорія, профіль, особливості"))
        self.f_spec = VoiceField("Наприклад: молодша група, 3–4 роки")
        self.f_spec.text = p["group_spec"]
        c2.add_widget(self.f_spec)

        # Крок 3 — діти
        self.children_editor = ChildrenEditor(self.app)

        self.steps = [sv1, sv2, self.children_editor]
        self.show()

    def show(self):
        self.body.clear_widgets()
        self.body.add_widget(self.steps[self.step])
        self.title.text = (f"{self.TITLES[self.step]}\n[size={int(FS_SMALL)}]"
                           f"Крок {self.step + 1} з {len(self.steps)}[/size]")
        last = self.step == len(self.steps) - 1
        self.btn_next.text = "Завершити" if last else "Далі"
        if self.step == 0:
            self.btn_back.text = "Скасувати" if self.editing else "Вийти"
        else:
            self.btn_back.text = "Назад"

    def _save_step(self):
        p = self.app.diary.profile
        if self.step == 0:
            role = next((b.text for b in self.role_btns if b.state == "down"), "")
            if not self.f_name.text:
                show_message("Вкажіть ваше прізвище, ім'я та по батькові.")
                return False
            if not role:
                show_message("Оберіть вашу посаду.")
                return False
            p["full_name"], p["role"] = self.f_name.text, role
        elif self.step == 1:
            if not self.f_group.text:
                show_message("Вкажіть номер групи.")
                return False
            p["group_number"], p["group_spec"] = self.f_group.text, self.f_spec.text
        self.app.diary.save_profile()
        return True

    def go_next(self):
        if not self._save_step():
            return
        if self.step < len(self.steps) - 1:
            self.step += 1
            self.show()
            return
        self.app.diary.profile["onboarded"] = True
        self.app.diary.save_profile()
        self.app.show_main()

    def go_back(self):
        if self.step > 0:
            self.step -= 1
            self.show()
        elif self.editing:
            self.app.show_main()
        else:
            self.app.logout()


# ----------------------------------------------------------------------
# Вкладка «Запис»: дата, дитина, чотири режимні моменти
# ----------------------------------------------------------------------
class RecordTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="record", **kw)
        self.app = app
        self.child_map = {}
        self.day_map = {}
        sv, col = scroll_column()

        col.add_widget(section("Дата"))
        self.day_spinner = make_spinner()
        self.day_spinner.bind(text=lambda *_: self.refresh_periods())
        col.add_widget(self.day_spinner)

        col.add_widget(section("Дитина"))
        self.child_spinner = make_spinner(text="Оберіть дитину")
        self.child_spinner.bind(text=lambda *_: self.refresh_periods())
        col.add_widget(self.child_spinner)

        col.add_widget(section("Режимні моменти"))
        col.add_widget(hint("Натисніть на етап, щоб заповнити характеристику діяльності дитини. "
                            "Фото та відео знімаються всередині кожного етапу — кнопки вгорі екрана."))
        self.period_btns = []
        for i, (key, title, _, _) in enumerate(PERIODS):
            b = Btn(text=title, bg=CARD_LIGHT, height=dp(80),
                    on_release=lambda *_, idx=i: self.open_period(idx))
            self.period_btns.append(b)
            col.add_widget(b)

        col.add_widget(spacer())
        self.btn_next = Btn(text="Далі", height=BTN_H_BIG, bold=True, on_release=self.go_next)
        col.add_widget(self.btn_next)
        col.add_widget(Btn(text="Попередній перегляд за день", bg=SECOND,
                           on_release=lambda *_: open_day_preview(self.day)))
        col.add_widget(Btn(text="Надіслати таблицю за день", bg=ACCENT,
                           on_release=lambda *_: app.share_day(self.day)))
        self.add_widget(sv)

    @property
    def day(self):
        return self.day_map.get(self.day_spinner.text, today())

    @property
    def child_id(self):
        return self.child_map.get(self.child_spinner.text)

    def on_pre_enter(self, *args):
        if not self.app.diary:
            return
        self.day_map = recent_days()
        self.day_spinner.values = list(self.day_map)
        if self.day_spinner.text not in self.day_map:
            self.day_spinner.text = next(iter(self.day_map))
        self.child_map = {}
        for child in self.app.diary.children:
            label = child["name"]
            while label in self.child_map:
                label += " "
            self.child_map[label] = child["id"]
        self.child_spinner.values = list(self.child_map)
        if self.child_spinner.text not in self.child_map:
            self.child_spinner.text = ("Оберіть дитину" if self.child_map
                                       else "Спершу додайте дітей на вкладці «Діти»")
        self.refresh_periods()

    def refresh_periods(self):
        if not self.app.diary:
            return
        cid = self.child_id
        next_title = None
        for btn, (key, title, _, _) in zip(self.period_btns, PERIODS):
            filled = bool(cid) and period_filled(self.app.diary.get_period(self.day, cid, key))
            btn.text = f"{title}\n[size={int(FS_SMALL)}]{'заповнено' if filled else 'не заповнено'}[/size]"
            btn.markup = True
            btn._bg = PRIMARY if filled else CARD_LIGHT
            btn._paint()
            if not filled and next_title is None:
                next_title = title
        self.btn_next.text = f"Далі: {next_title}" if next_title else "Далі: переглянути день"

    def open_period(self, index):
        if not self.child_id:
            show_message("Спершу оберіть дитину.")
            return
        self.app.open_period(self.day, self.child_id, index)

    def go_next(self, *_):
        if not self.child_id:
            show_message("Спершу оберіть дитину.")
            return
        for i, (key, _, _, _) in enumerate(PERIODS):
            if not period_filled(self.app.diary.get_period(self.day, self.child_id, key)):
                self.open_period(i)
                return
        open_day_preview(self.day)


# ----------------------------------------------------------------------
# Екран одного режимного моменту з підказками, як у паперовій таблиці
# ----------------------------------------------------------------------
class PeriodScreen(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="period", **kw)
        self.app = app
        self.day = today()
        self.child_id = None
        self.index = 0
        self.undo = []
        root = BoxLayout(orientation="vertical")
        header = Panel(bg=PRIMARY, orientation="vertical", size_hint_y=None, height=dp(86),
                       padding=[dp(16), dp(8)])
        self.title = Label(font_size=sp(22), bold=True, halign="left", valign="middle",
                           markup=True, shorten=True, shorten_from="right")
        self.title.bind(size=lambda *_: setattr(self.title, "text_size", self.title.size))
        header.add_widget(self.title)
        root.add_widget(header)
        self.body = BoxLayout()
        root.add_widget(self.body)
        nav = Panel(bg=CARD, orientation="horizontal", size_hint_y=None, height=BTN_H + dp(20),
                    padding=dp(10), spacing=dp(8))
        self.btn_back = Btn(text="Назад", bg=SECOND, on_release=lambda *_: self.go_back())
        self.btn_save = Btn(text="Зберегти", bg=CARD_LIGHT, on_release=lambda *_: self.save_close())
        self.btn_next = Btn(text="Далі", on_release=lambda *_: self.go_next())
        for b in (self.btn_back, self.btn_save, self.btn_next):
            nav.add_widget(b)
        root.add_widget(nav)
        self.add_widget(root)

    def open(self, day, child_id, index):
        self.day, self.child_id, self.index = day, child_id, index
        self.build_form()

    # --- побудова форми ---
    def _toggles(self, col, options, selected, cols=2, height=None):
        grid = GridLayout(cols=cols, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))
        btns = []
        for opt in options:
            t = ToggleBtn(text=opt, font_size=FS_SMALL, height=height or BTN_H)
            if opt in selected:
                t.state = "down"
            btns.append(t)
            grid.add_widget(t)
        col.add_widget(grid)
        return btns

    def _stepper(self, col, mode, value):
        row = BoxLayout(orientation="horizontal", size_hint_y=None, height=BTN_H, spacing=dp(8))
        lbl = Label(text=mode, font_size=FS, halign="left", valign="middle")
        lbl.bind(size=lambda *_: setattr(lbl, "text_size", lbl.size))
        row.add_widget(lbl)
        minus = Btn(text="−", bg=SECOND, size_hint_x=None, width=dp(64), font_size=sp(26))
        val = Label(text=str(value), font_size=sp(24), bold=True, size_hint_x=None, width=dp(56))
        plus = Btn(text="+", bg=PRIMARY, size_hint_x=None, width=dp(64), font_size=sp(26))
        minus.bind(on_release=lambda *_: setattr(val, "text", str(max(0, int(val.text) - 1))))
        plus.bind(on_release=lambda *_: setattr(val, "text", str(min(99, int(val.text) + 1))))
        row.add_widget(minus)
        row.add_widget(val)
        row.add_widget(plus)
        col.add_widget(row)
        return val

    def build_form(self):
        diary = self.app.diary
        key, title, period_hint, options = PERIODS[self.index]
        child = diary.get_child(self.child_id)
        data = diary.get_period(self.day, self.child_id, key)
        self.media = list(data["media"])
        self.undo = []
        sv, col = scroll_column()

        # --- Фото та відео: на самому верху, щоб було одразу видно ---
        media_card = card()
        media_card.add_widget(section("Фото та відео"))
        mrow = GridLayout(cols=3, spacing=dp(8), size_hint_y=None, height=dp(70))
        mrow.add_widget(Btn(text="Зняти фото", bg=ACCENT, height=dp(70), font_size=FS_SMALL,
                            on_release=lambda *_: self.app.take_photo(self._media_added)))
        mrow.add_widget(Btn(text="Зняти відео", bg=ACCENT, height=dp(70), font_size=FS_SMALL,
                            on_release=lambda *_: self.app.take_video(self._media_added)))
        mrow.add_widget(Btn(text="З галереї", bg=SECOND, height=dp(70), font_size=FS_SMALL,
                            on_release=lambda *_: self.app.pick_media(self._media_added)))
        media_card.add_widget(mrow)
        self.media_list = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(6))
        self.media_list.bind(minimum_height=self.media_list.setter("height"))
        media_card.add_widget(self.media_list)
        col.add_widget(media_card)
        self._update_media()

        col.add_widget(spacer())
        col.add_widget(section("Короткий опис діяльності"))
        col.add_widget(hint(period_hint))
        self.activity_btns = self._toggles(col, options, data["activities"])
        self.desc = make_input("Що робила дитина? Можна надиктувати...", multiline=True, height=dp(150))
        self.desc.text = data["description"]
        col.add_widget(self.desc)
        row = GridLayout(cols=3, spacing=dp(8), size_hint_y=None, height=BTN_H)
        row.add_widget(Btn(text="Голос", bg=ACCENT, on_release=self.voice))
        row.add_widget(Btn(text="Скасувати", bg=SECOND, font_size=FS_SMALL, on_release=self.undo_last))
        row.add_widget(Btn(text="Стерти", bg=DANGER, on_release=self.clear_desc))
        col.add_widget(row)

        col.add_widget(spacer())
        col.add_widget(section("Кількість виконаних завдань"))
        self.task_vals = {m: self._stepper(col, m, data["tasks"].get(m, 0)) for m in TASK_MODES}

        col.add_widget(spacer())
        col.add_widget(section("Труднощі"))
        self.diff = VoiceField("Що давалося важко?", multiline=True, height=dp(110))
        self.diff.text = data["difficulties"]
        col.add_widget(self.diff)

        col.add_widget(spacer())
        col.add_widget(section("Потреби"))
        col.add_widget(hint("У допомозі асистента: навідні запитання, багаторазове повторення, "
                            "більше наочності чи часу на виконання завдання"))
        self.need_btns = self._toggles(col, NEEDS, data["needs"], height=dp(66))
        self.needs_note = VoiceField("Інше (за потреби)")
        self.needs_note.text = data["needs_note"]
        col.add_widget(self.needs_note)

        col.add_widget(spacer())
        col.add_widget(section("Поведінка"))
        col.add_widget(Lbl(text="Взаємодія з ровесниками"))
        self.inter = VoiceField("Як дитина взаємодіяла з дітьми?", multiline=True, height=dp(110))
        self.inter.text = data["interaction"]
        col.add_widget(self.inter)
        col.add_widget(Lbl(text="Ігрова діяльність"))
        self.game_btns = self._toggles(col, GAMES, data["games"], cols=3)
        self.games_note = VoiceField("Деталі гри (за потреби)")
        self.games_note.text = data["games_note"]
        col.add_widget(self.games_note)

        col.add_widget(spacer(dp(20)))

        self.body.clear_widgets()
        self.body.add_widget(sv)
        name = escape_markup(child["name"]) if child else ""
        self.title.text = (f"{title}\n[size={int(FS_SMALL)}]{name} · {format_day_short(self.day)} · "
                           f"етап {self.index + 1} з {len(PERIODS)}[/size]")
        self.btn_back.text = "До списку" if self.index == 0 else "Назад"
        self.btn_next.text = "Завершити" if self.index == len(PERIODS) - 1 else "Далі"

    # --- текст ---
    def _remember(self):
        self.undo.append(self.desc.text)
        del self.undo[:-30]

    def voice(self, *_):
        self.app.voice_into(self.desc, append=True, before=self._remember)

    def undo_last(self, *_):
        if self.undo:
            self.desc.text = self.undo.pop()
        else:
            self.desc.do_undo()

    def clear_desc(self, *_):
        if self.desc.text:
            self._remember()
            self.desc.text = ""

    # --- медіа ---
    def _media_added(self, names):
        self.media.extend(names)
        self._update_media()
        self.save()  # одразу зберігаємо, щоб файли не загубились

    def remove_media(self, name):
        def do_remove():
            try:
                os.remove(os.path.join(self.app.diary.media_dir, name))
            except OSError:
                pass
            if name in self.media:
                self.media.remove(name)
            self._update_media()
            self.save()
        confirm("Прибрати цей файл із запису?", do_remove, "Прибрати")

    def _update_media(self):
        self.media_list.clear_widgets()
        if not self.media:
            self.media_list.add_widget(hint("Файли ще не додано"))
            return
        photos = videos = 0
        for name in self.media:
            is_photo = name.lower().endswith(IMAGE_EXT)
            if is_photo:
                photos += 1
                label = f"Фото {photos}"
            else:
                videos += 1
                label = f"Відео {videos}"
            try:
                size = os.path.getsize(os.path.join(self.app.diary.media_dir, name)) / 1048576
                label += f"  ({size:.1f} МБ)"
            except OSError:
                pass
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=BTN_H_SMALL, spacing=dp(8))
            lbl = Label(text=label, font_size=FS_SMALL, halign="left", valign="middle")
            lbl.bind(size=lambda inst, *_: setattr(inst, "text_size", inst.size))
            row.add_widget(lbl)
            row.add_widget(Btn(text="Прибрати", bg=DANGER, size_hint_x=None, width=dp(120),
                               height=BTN_H_SMALL, font_size=FS_SMALL,
                               on_release=lambda *_, n=name: self.remove_media(n)))
            self.media_list.add_widget(row)

    # --- збереження та навігація ---
    def collect(self):
        return {
            "activities": [b.text for b in self.activity_btns if b.state == "down"],
            "description": self.desc.text.strip(),
            "tasks": {m: int(v.text) for m, v in self.task_vals.items()},
            "difficulties": self.diff.text,
            "needs": [b.text for b in self.need_btns if b.state == "down"],
            "needs_note": self.needs_note.text,
            "interaction": self.inter.text,
            "games": [b.text for b in self.game_btns if b.state == "down"],
            "games_note": self.games_note.text,
            "media": list(self.media),
        }

    def save(self):
        key = PERIODS[self.index][0]
        try:
            self.app.diary.save_period(self.day, self.child_id, key, self.collect())
        except Exception as e:
            show_message(f"Не вдалося зберегти:\n{e}")
            return False
        return True

    def go_next(self):
        if not self.save():
            return
        if self.index < len(PERIODS) - 1:
            self.index += 1
            self.build_form()
        else:
            self.app.show_main()
            show_message("Запис за день збережено.\nЙого можна переглянути або надіслати на вкладці «Запис».")

    def go_back(self):
        if not self.save():
            return
        if self.index > 0:
            self.index -= 1
            self.build_form()
        else:
            self.app.show_main()

    def save_close(self):
        if self.save():
            self.app.show_main()
            show_message(f"Збережено: {PERIODS[self.index][1]}.")


class ChildrenTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="children", **kw)
        self.editor = ChildrenEditor(app)
        self.add_widget(self.editor)

    def on_pre_enter(self, *args):
        if self.editor.app.diary:
            self.editor.reset_form()
            self.editor.refresh()


class ArchiveTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="archive", **kw)
        self.app = app
        sv, col = scroll_column()
        col.add_widget(Btn(text="Експортувати все в «Завантаження»", on_release=lambda *_: app.export_all()))
        col.add_widget(hint("Таблиці Excel, фото, відео та звіти буде скопійовано в папку "
                            f"«Завантаження/{EXPORT_ROOT}» — їх видно у файловому менеджері."))
        col.add_widget(spacer())
        col.add_widget(section("Дні із записами"))
        self.list_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10))
        self.list_box.bind(minimum_height=self.list_box.setter("height"))
        col.add_widget(self.list_box)
        self.add_widget(sv)

    def on_pre_enter(self, *args):
        if self.app.diary:
            self.refresh()

    def refresh(self):
        self.list_box.clear_widgets()
        days = self.app.diary.days()
        if not days:
            self.list_box.add_widget(hint("Записів ще немає. Збережіть перший запис на вкладці «Запис»."))
        for day, n in days:
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(76), spacing=dp(8))
            row.add_widget(Btn(
                text=f"{format_day(day)}\n[size={int(FS_SMALL)}]{n} {plural(n, 'дитина', 'дитини', 'дітей')}[/size]",
                markup=True, bg=CARD_LIGHT, height=dp(76), on_release=lambda *_, d=day: open_day_preview(d)))
            row.add_widget(Btn(text="Надіслати", bg=ACCENT, size_hint_x=None, width=dp(124), height=dp(76),
                               font_size=FS_SMALL, on_release=lambda *_, d=day: self.app.share_day(d)))
            self.list_box.add_widget(row)


class ReportsTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="reports", **kw)
        self.app = app
        self.child_map = {}
        self.month_map = {}
        self.last_pdf = None
        sv, col = scroll_column()
        col.add_widget(Lbl(text="Звіт для батьків", font_size=sp(23), bold=True))
        col.add_widget(hint("Усі спостереження за режимними моментами та фото дитини за місяць "
                            "в одному PDF — готовому для надсилання у Viber чи Telegram. "
                            "Примітки про особливості дитини до звіту не потрапляють."))
        col.add_widget(section("Дитина"))
        self.child_spinner = make_spinner()
        col.add_widget(self.child_spinner)
        col.add_widget(section("Місяць"))
        self.month_spinner = make_spinner()
        col.add_widget(self.month_spinner)
        col.add_widget(spacer())
        self.btn_make = Btn(text="Сформувати звіт (PDF)", height=BTN_H_BIG, bold=True, on_release=self.make)
        col.add_widget(self.btn_make)
        self.result_box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(10))
        self.result_box.bind(minimum_height=self.result_box.setter("height"))
        col.add_widget(self.result_box)
        self.add_widget(sv)

    def on_pre_enter(self, *args):
        if not self.app.diary:
            return
        kids = self.app.diary.children
        self.child_map = {}
        for child in kids:
            label = child["name"]
            while label in self.child_map:
                label += " "
            self.child_map[label] = child
        self.child_spinner.values = list(self.child_map)
        if self.child_spinner.text not in self.child_map:
            self.child_spinner.text = next(iter(self.child_map), "Спершу додайте дітей")
        self.month_map = {format_month(m): m for m in self.app.diary.months()}
        self.month_spinner.values = list(self.month_map)
        if self.month_spinner.text not in self.month_map:
            self.month_spinner.text = format_month(today()[:7])
        self.btn_make.disabled = not kids

    def make(self, *_):
        child = self.child_map.get(self.child_spinner.text)
        month = self.month_map.get(self.month_spinner.text)
        if not child or not month:
            show_message("Оберіть дитину та місяць.")
            return
        diary = self.app.diary
        records = diary.child_month_entries(child["id"], month)
        if not records:
            show_message(f"За {self.month_spinner.text} записів про «{child['name']}» немає.")
            return
        name = f"Звіт_{safe_filename(child['name'])}_{month}.pdf"
        path = os.path.join(diary.reports_dir, name)
        self.btn_make.disabled = True
        self.btn_make.text = "Формуємо звіт..."
        self.result_box.clear_widgets()

        def work():
            return pdf_report.generate(path, child, month, dict(diary.profile), records,
                                       diary.media_dir, android=ab.IS_ANDROID)

        def done(result):
            self._reset_btn()
            self.last_pdf = result
            ab.notify("Звіт для батьків готовий", name)
            self.result_box.add_widget(Lbl(text=f"Готово: {name}", color=GREEN_TEXT, bold=True))
            self.result_box.add_widget(Btn(text="Надіслати (Viber, Telegram…)", bg=ACCENT,
                                           on_release=lambda *_: self.app.share_file(
                                               self.last_pdf, f"{EXPORT_ROOT}/Звіти", "application/pdf")))
            self.result_box.add_widget(Btn(text="Зберегти в «Завантаження»", bg=SECOND,
                                           on_release=lambda *_: self.app.export_one(
                                               self.last_pdf, f"{EXPORT_ROOT}/Звіти")))

        def failed(error):
            self._reset_btn()
            show_message(f"Не вдалося сформувати звіт:\n{error}")

        self.app.run_bg(work, done, failed)

    def _reset_btn(self):
        self.btn_make.disabled = False
        self.btn_make.text = "Сформувати звіт (PDF)"


class MainScreen(Screen):
    TABS = [("record", "Запис"), ("children", "Діти"), ("archive", "Архів"), ("reports", "Звіти")]

    def __init__(self, app, **kw):
        super().__init__(name="main", **kw)
        self.app = app
        root = BoxLayout(orientation="vertical")

        header = Panel(bg=PRIMARY, orientation="horizontal", size_hint_y=None, height=dp(72),
                       padding=[dp(14), dp(10)], spacing=dp(10))
        texts = BoxLayout(orientation="vertical")
        self.header_name = Label(font_size=FS, bold=True, halign="left", valign="middle",
                                 shorten=True, shorten_from="right")
        self.header_sub = Label(font_size=FS_SMALL, halign="left", valign="middle", color=(1, 1, 1, 0.85),
                                shorten=True, shorten_from="right")
        for lbl in (self.header_name, self.header_sub):
            lbl.bind(size=lambda inst, *_: setattr(inst, "text_size", inst.size))
            texts.add_widget(lbl)
        header.add_widget(texts)
        header.add_widget(Btn(text="Меню", bg=darker(PRIMARY, 0.8), size_hint=(None, 1),
                              width=dp(96), on_release=lambda *_: self.open_menu()))
        root.add_widget(header)

        self.tabs = ScreenManager(transition=NoTransition())
        self.tabs.add_widget(RecordTab(app))
        self.tabs.add_widget(ChildrenTab(app))
        self.tabs.add_widget(ArchiveTab(app))
        self.tabs.add_widget(ReportsTab(app))
        root.add_widget(self.tabs)

        nav = GridLayout(cols=len(self.TABS), size_hint_y=None, height=dp(68), spacing=dp(2))
        self.nav_btns = {}
        for key, text in self.TABS:
            b = ToggleBtn(text=text, group="nav", allow_no_selection=False, height=dp(68),
                          bg_off=CARD, bg_on=PRIMARY)
            b.bind(on_release=lambda btn, k=key: self.select(k))
            self.nav_btns[key] = b
            nav.add_widget(b)
        root.add_widget(nav)
        self.add_widget(root)

    def refresh_header(self):
        p = self.app.diary.profile
        sub = " · ".join(x for x in (p["role"], self.app.diary.group_title) if x)
        self.header_name.text = p["full_name"] or "Щоденник спостережень"
        self.header_sub.text = sub

    def select(self, key):
        for k, b in self.nav_btns.items():
            b.state = "down" if k == key else "normal"
        if self.tabs.current == key:
            self.tabs.current_screen.on_pre_enter()
        self.tabs.current = key

    @property
    def current_tab(self):
        return self.tabs.current

    def open_menu(self):
        box = BoxLayout(orientation="vertical", padding=dp(8), spacing=dp(12))
        popup = Popup(title="Меню", title_size=FS, title_align="center", content=box,
                      size_hint=(0.9, None), height=BTN_H * 3 + dp(130), separator_color=PRIMARY)

        def edit_profile(*_):
            popup.dismiss()
            self.app.edit_profile()

        def logout(*_):
            popup.dismiss()
            self.app.logout()
        box.add_widget(Btn(text="Профіль, група та діти", bg=SECOND, on_release=edit_profile))
        box.add_widget(Btn(text="Вийти з облікового запису", bg=DANGER, on_release=logout))
        box.add_widget(Btn(text="Закрити", bg=CARD_LIGHT, on_release=popup.dismiss))
        popup.open()


# ----------------------------------------------------------------------
# Застосунок
# ----------------------------------------------------------------------
class DiaryApp(App):
    title = "Щоденник спостережень"

    def build(self):
        Window.clearcolor = BG
        Window.softinput_mode = "below_target"
        self.users = UserStore(os.path.join(self.user_data_dir, "diary"))
        self.diary = None
        self._voice_target = None
        self._voice_append = False
        self._voice_before = None
        self._media_callback = None
        ab.bind_activity_result(self.on_activity_result)
        Window.bind(on_keyboard=self.on_key)

        self.sm = ScreenManager(transition=NoTransition())
        self.sm.add_widget(LoginScreen(self))
        self.sm.add_widget(RegisterScreen(self))
        self.wizard = WizardScreen(self)
        self.sm.add_widget(self.wizard)
        self.main = MainScreen(self)
        self.sm.add_widget(self.main)
        self.period = PeriodScreen(self)
        self.sm.add_widget(self.period)
        self.sm.current = "login"
        # Запит усіх дозволів одразу після встановлення (при першому запуску)
        Clock.schedule_once(lambda dt: ab.request_startup_permissions(), 1.0)
        return self.sm

    # --- обліковий запис ---
    def open_account(self, login):
        self.users.set_last_login(login)
        self.diary = Diary(self.users.user_dir(login))
        if self.diary.profile.get("onboarded"):
            self.show_main()
        else:
            self.wizard.start(editing=False)
            self.sm.current = "wizard"

    def edit_profile(self):
        self.wizard.start(editing=True)
        self.sm.current = "wizard"

    def show_main(self):
        self.main.refresh_header()
        self.sm.current = "main"
        self.main.select("record")

    def open_period(self, day, child_id, index):
        if not self.diary.get_child(child_id):
            show_message("Цієї дитини вже немає у списку групи.")
            return
        self.period.open(day, child_id, index)
        self.sm.current = "period"

    def logout(self):
        self.diary = None
        self.sm.current = "login"

    def refresh_tabs(self):
        if self.sm.current == "main":
            self.main.tabs.current_screen.on_pre_enter()

    # --- фонові задачі ---
    def run_bg(self, fn, on_done=None, on_error=None):
        def worker():
            result, error = None, None
            try:
                result = fn()
            except Exception as e:
                error = e
            finally:
                ab.detach()
            Clock.schedule_once(lambda dt: self._bg_finish(result, error, on_done, on_error))
        threading.Thread(target=worker, daemon=True).start()

    @staticmethod
    def _bg_finish(result, error, on_done, on_error):
        if error is not None:
            (on_error or (lambda e: show_message(f"Помилка:\n{e}")))(error)
        elif on_done:
            on_done(result)

    # --- голос ---
    def voice_into(self, target, append=False, before=None):
        self._voice_target, self._voice_append, self._voice_before = target, append, before
        if not ab.IS_ANDROID:
            self._voice_result("Тестовий голосовий текст")
            return
        try:
            ab.start_speech()
        except Exception as e:
            show_message("Розпізнавання мовлення недоступне.\n"
                         f"Встановіть або оновіть застосунок Google.\n\n{e}")

    @mainthread
    def _voice_result(self, text):
        target = self._voice_target
        if target is None:
            return
        if self._voice_before:
            self._voice_before()
        if self._voice_append and target.text.strip():
            target.text = target.text.rstrip() + " " + text
        else:
            target.text = text

    # --- медіа ---
    def pick_media(self, callback):
        if not ab.IS_ANDROID:
            show_message("Вибір фото та відео доступний на телефоні.")
            return
        self._media_callback = callback
        try:
            ab.pick_media()
        except Exception as e:
            show_message(f"Не вдалося відкрити галерею:\n{e}")

    def take_photo(self, callback):
        self._camera(callback, ab.take_photo, "камеру")

    def take_video(self, callback):
        self._camera(callback, ab.take_video, "відеозйомку")

    def _camera(self, callback, action, what):
        if not ab.IS_ANDROID:
            show_message("Зйомка доступна на телефоні.")
            return
        if not ab.has_camera_permission():
            def on_result(granted):
                if granted:
                    Clock.schedule_once(lambda dt: self._camera(callback, action, what))
                else:
                    show_message("Щоб знімати фото й відео, дозвольте доступ до камери "
                                 "(Налаштування → Застосунки → Щоденник спостережень → Дозволи).")
            ab.request_camera_permission(on_result)
            return
        self._media_callback = callback
        try:
            action()
        except Exception as e:
            show_message(f"Не вдалося відкрити {what}:\n{e}")

    def _copy_to_media(self, uris):
        if uris and self.diary:
            media_dir = self.diary.media_dir
            self.run_bg(lambda: ab.copy_uris(uris, media_dir), self._media_copied)

    def on_activity_result(self, request_code, result_code, intent):
        ok = result_code == ab.RESULT_OK
        try:
            if request_code == ab.REQUEST_PHOTO:
                # камера з EXTRA_OUTPUT часто повертає порожній intent — це нормально
                uri = ab.photo_result(ok)
                if uri is not None:
                    self._copy_to_media([uri])
                return
        except Exception as e:
            show_message(f"Помилка обробки фото:\n{e}")
            return
        if not ok or intent is None:
            return
        try:
            if request_code == ab.REQUEST_SPEECH:
                text = ab.speech_text(intent)
                if text:
                    self._voice_result(text)
            elif request_code in (ab.REQUEST_MEDIA, ab.REQUEST_VIDEO):
                self._copy_to_media(ab.uris_from_result(intent))
        except Exception as e:
            show_message(f"Помилка обробки результату:\n{e}")

    def _media_copied(self, result):
        saved, errors = result
        if self._media_callback and saved:
            self._media_callback(saved)
        if errors:
            show_message(f"Додано файлів: {len(saved)}\nНе вдалося додати: {len(errors)}\n{errors[0]}")

    # --- експорт і надсилання ---
    def _with_storage(self, action):
        if ab.needs_storage_permission():
            def on_result(granted):
                if granted:
                    Clock.schedule_once(lambda dt: action())
                else:
                    show_message("Без дозволу на доступ до пам'яті файл зберегти не вийде.")
            ab.request_storage_permission(on_result)
        else:
            action()

    def export_one(self, path, folder):
        def work():
            ab.export_file(path, folder)
        self._with_storage(lambda: self.run_bg(
            work, lambda _: show_message(f"Збережено в «Завантаження/{folder}».")))

    def share_file(self, path, folder, mime):
        if not path or not os.path.exists(path):
            show_message("Файл не знайдено.")
            return
        if not ab.IS_ANDROID:
            show_message(f"Файл:\n{path}")
            return

        def work():
            target = ab.export_file(path, folder)
            ab.share(target, mime)
        self._with_storage(lambda: self.run_bg(work))

    def share_day(self, day):
        path = self.diary.excel_path(day)
        if not os.path.exists(path):
            show_message(f"За {format_day(day)} записів ще немає.")
            return
        self.share_file(path, f"{EXPORT_ROOT}/Таблиці", XLSX_MIME)

    def export_all(self):
        diary = self.diary
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        base = f"{EXPORT_ROOT}/{stamp}"

        def work():
            counts = {"Таблиці": 0, "Медіа": 0, "Звіти": 0}
            for sub, src_dir in (("Таблиці", diary.excel_dir), ("Медіа", diary.media_dir),
                                 ("Звіти", diary.reports_dir)):
                for name in sorted(os.listdir(src_dir)):
                    ab.export_file(os.path.join(src_dir, name), f"{base}/{sub}")
                    counts[sub] += 1
            return counts

        def done(c):
            if not any(c.values()):
                show_message("Поки що немає файлів для експорту.")
                return
            ab.notify("Експорт завершено", f"Файли в «Завантаження/{base}»")
            show_message(f"Скопійовано в «Завантаження/{base}»:\n"
                         f"таблиць — {c['Таблиці']}, медіафайлів — {c['Медіа']}, звітів — {c['Звіти']}")
        self._with_storage(lambda: self.run_bg(work, done))

    # --- кнопка «Назад» на телефоні ---
    def on_key(self, window, key, *args):
        if key != 27:
            return False
        if window.children and isinstance(window.children[0], ModalView):
            window.children[0].dismiss()
            return True
        current = self.sm.current
        if current == "register":
            self.sm.current = "login"
            return True
        if current == "wizard":
            self.wizard.go_back()
            return True
        if current == "period":
            self.period.go_back()
            return True
        if current == "main" and self.main.current_tab != "record":
            self.main.select("record")
            return True
        return False


if __name__ == "__main__":
    DiaryApp().run()
