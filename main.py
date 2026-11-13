"""Щоденник спостережень — офлайн-застосунок для педагогів ЗДО."""
import os
import threading
from datetime import datetime

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
from storage import (CRITERIA, ROLES, Diary, UserStore, format_day, format_month,
                     plural, safe_filename, today)

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
    recs = app.diary.records(day)
    box = BoxLayout(orientation="vertical", spacing=dp(10), padding=dp(4))
    sv, col = scroll_column(padding=dp(4), spacing=dp(10))
    popup = Popup(title=f"Записи за {format_day(day)}", title_size=FS, title_align="center",
                  content=box, size_hint=(0.96, 0.92), separator_color=PRIMARY)

    if not recs:
        col.add_widget(Lbl(text="За цей день записів ще немає.", halign="center"))
    for r in recs:
        c = card()
        c.add_widget(Lbl(text=f"{r['time']}  ·  {r['child_name']}", bold=True))
        if r.get("criteria"):
            c.add_widget(Lbl(text=", ".join(r["criteria"]), color=GREEN_TEXT, font_size=FS_SMALL))
        if r.get("note"):
            c.add_widget(Lbl(text=r["note"]))
        if r.get("media"):
            n = len(r["media"])
            c.add_widget(hint(f"{n} {plural(n, 'медіафайл', 'медіафайли', 'медіафайлів')}"))

        def ask_delete(*_, rec_id=r["id"]):
            def do_delete():
                app.diary.delete_record(day, rec_id)
                popup.dismiss()
                app.refresh_tabs()
                open_day_preview(day)
            confirm("Видалити цей запис разом із його фото та відео?", do_delete, "Видалити")
        c.add_widget(Btn(text="Видалити запис", bg=DANGER, height=BTN_H_SMALL,
                         font_size=FS_SMALL, on_release=ask_delete))
        col.add_widget(c)
    box.add_widget(sv)

    row = GridLayout(cols=2, spacing=dp(10), size_hint_y=None, height=BTN_H)
    send = Btn(text="Надіслати Excel", on_release=lambda *_: app.share_file(
        app.diary.excel_path(day), f"{EXPORT_ROOT}/Таблиці", XLSX_MIME))
    send.disabled = not recs
    row.add_widget(send)
    row.add_widget(Btn(text="Закрити", bg=SECOND, on_release=popup.dismiss))
    box.add_widget(row)
    popup.open()


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
# Вкладки головного екрана
# ----------------------------------------------------------------------
class RecordTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="record", **kw)
        self.app = app
        self.pending = []   # медіафайли до наступного запису
        self.undo = []      # історія тексту для кнопки «Назад»
        self.child_map = {}
        sv, col = scroll_column()

        self.date_lbl = hint("")
        col.add_widget(self.date_lbl)
        col.add_widget(section("Дитина"))
        self.child_spinner = make_spinner()
        col.add_widget(self.child_spinner)

        col.add_widget(section("Критерії розвитку"))
        grid = GridLayout(cols=2, spacing=dp(8), size_hint_y=None)
        grid.bind(minimum_height=grid.setter("height"))
        self.crit = {}
        for c in CRITERIA:
            t = ToggleBtn(text=c, font_size=FS_SMALL)
            self.crit[c] = t
            grid.add_widget(t)
        col.add_widget(grid)

        col.add_widget(section("Спостереження"))
        self.note = make_input("Введіть або надиктуйте спостереження...", multiline=True, height=dp(190))
        col.add_widget(self.note)
        row = GridLayout(cols=3, spacing=dp(8), size_hint_y=None, height=BTN_H)
        row.add_widget(Btn(text="Голос", bg=ACCENT, on_release=self.voice))
        row.add_widget(Btn(text="Назад", bg=SECOND, on_release=self.undo_last))
        row.add_widget(Btn(text="Стерти", bg=DANGER, on_release=self.clear_note))
        col.add_widget(row)

        col.add_widget(section("Фото та відео"))
        mrow = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=BTN_H)
        mrow.add_widget(Btn(text="Додати фото / відео", bg=SECOND, font_size=FS_SMALL, on_release=self.add_media))
        mrow.add_widget(Btn(text="Прибрати додані", bg=SECOND, font_size=FS_SMALL, on_release=self.drop_media))
        col.add_widget(mrow)
        self.media_lbl = hint("")
        col.add_widget(self.media_lbl)

        col.add_widget(spacer())
        col.add_widget(Btn(text="Зберегти запис", height=BTN_H_BIG, bold=True, on_release=self.save))
        col.add_widget(Btn(text="Попередній перегляд за сьогодні", bg=SECOND,
                           on_release=lambda *_: open_day_preview(today())))
        self.add_widget(sv)
        self._update_media()

    def on_pre_enter(self, *args):
        if not self.app.diary:
            return
        self.date_lbl.text = f"Сьогодні: {format_day(today())}"
        self.child_map = {"Уся група": ""}
        for child in self.app.diary.children:
            label = child["name"]
            while label in self.child_map:
                label += " "
            self.child_map[label] = child["id"]
        self.child_spinner.values = list(self.child_map)
        if self.child_spinner.text not in self.child_map:
            self.child_spinner.text = "Уся група"

    # --- текст ---
    def _remember(self):
        self.undo.append(self.note.text)
        del self.undo[:-30]

    def voice(self, *_):
        self.app.voice_into(self.note, append=True, before=self._remember)

    def undo_last(self, *_):
        if self.undo:
            self.note.text = self.undo.pop()
        else:
            self.note.do_undo()

    def clear_note(self, *_):
        if self.note.text:
            self._remember()
            self.note.text = ""

    # --- медіа ---
    def add_media(self, *_):
        self.app.pick_media(self._media_added)

    def _media_added(self, names):
        self.pending.extend(names)
        self._update_media()

    def drop_media(self, *_):
        for name in self.pending:
            try:
                os.remove(os.path.join(self.app.diary.media_dir, name))
            except OSError:
                pass
        self.pending = []
        self._update_media()

    def _update_media(self):
        n = len(self.pending)
        self.media_lbl.text = (f"Додано до запису: {n} {plural(n, 'файл', 'файли', 'файлів')}"
                               if n else "Файли не додано")

    # --- збереження ---
    def save(self, *_):
        criteria = [c for c, t in self.crit.items() if t.state == "down"]
        child_id = self.child_map.get(self.child_spinner.text, "")
        try:
            rec = self.app.diary.add_record(child_id, criteria, self.note.text, self.pending)
        except ValueError as e:
            show_message(str(e))
            return
        except Exception as e:
            show_message(f"Не вдалося зберегти запис:\n{e}")
            return
        self.note.text = ""
        self.undo = []
        self.pending = []
        for t in self.crit.values():
            t.state = "normal"
        self._update_media()
        show_message(f"Запис збережено о {rec['time']}.\nДитина: {rec['child_name']}")


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
            self.list_box.add_widget(Btn(
                text=f"{format_day(day)}\n{n} {plural(n, 'запис', 'записи', 'записів')}",
                bg=CARD_LIGHT, height=dp(72), on_release=lambda *_, d=day: open_day_preview(d)))


class ReportsTab(Screen):
    def __init__(self, app, **kw):
        super().__init__(name="reports", **kw)
        self.app = app
        self.child_map = {}
        self.month_map = {}
        self.last_pdf = None
        sv, col = scroll_column()
        col.add_widget(Lbl(text="Звіт для батьків", font_size=sp(23), bold=True))
        col.add_widget(hint("Усі спостереження, критерії розвитку та фото дитини за місяць "
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
        records = diary.child_month_records(child["id"], month)
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
        self.header_lbl = Label(markup=True, font_size=FS, halign="left", valign="middle")
        self.header_lbl.bind(size=lambda *_: setattr(self.header_lbl, "text_size", self.header_lbl.size))
        header.add_widget(self.header_lbl)
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
        self.header_lbl.text = (f"[b]{escape_markup(p['full_name'] or 'Щоденник')}[/b]\n"
                                f"[size={int(FS_SMALL)}]{escape_markup(sub)}[/size]")

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
        self.sm.current = "login"
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

    def on_activity_result(self, request_code, result_code, intent):
        if result_code != ab.RESULT_OK or intent is None:
            return
        try:
            if request_code == ab.REQUEST_SPEECH:
                text = ab.speech_text(intent)
                if text:
                    self._voice_result(text)
            elif request_code == ab.REQUEST_MEDIA:
                uris = ab.uris_from_result(intent)
                if uris and self.diary:
                    media_dir = self.diary.media_dir
                    self.run_bg(lambda: ab.copy_uris(uris, media_dir), self._media_copied)
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
        if current == "main" and self.main.current_tab != "record":
            self.main.select("record")
            return True
        return False


if __name__ == "__main__":
    DiaryApp().run()
