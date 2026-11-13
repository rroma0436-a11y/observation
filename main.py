import os
import shutil
import mimetypes
import threading
from datetime import datetime

import openpyxl
from kivy.app import App
from kivy.clock import Clock, mainthread
from kivy.core.window import Window
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.textinput import TextInput
from kivy.utils import platform

EXCEL_NAME = "Shodennyk_Sposterejen_ZDO.xlsx"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
EXPORT_FOLDER = "Shodennyk"

REQUEST_SPEECH = 101
REQUEST_MEDIA = 102
RESULT_OK = -1  # Activity.RESULT_OK

HEADERS = [
    "Дата", "Група", "Прізвище, ім'я дитини", "Режимні моменти",
    "Кількість завдань", "Труднощі", "Потреби", "Поведінка", "Медіафайли",
]

# Підключення Android API через pyjnius
if platform == 'android':
    from jnius import autoclass, detach
    from android import activity as android_activity
    from android.permissions import request_permissions, check_permission, Permission

    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    Intent = autoclass('android.content.Intent')
    RecognizerIntent = autoclass('android.speech.RecognizerIntent')
    MimeTypeMap = autoclass('android.webkit.MimeTypeMap')
    ContentValues = autoclass('android.content.ContentValues')
    Environment = autoclass('android.os.Environment')
    SDK_INT = autoclass('android.os.Build$VERSION').SDK_INT
else:
    SDK_INT = 0


class ObservationApp(App):
    def build(self):
        self.title = "Щоденник спостережень"
        self.active_text_input = None   # поле, у яке вставляється надиктований текст
        self.pending_media = []         # медіафайли, прикріплені до наступного запису
        self.excel_path = os.path.join(self.user_data_dir, EXCEL_NAME)
        self.media_dir = os.path.join(self.user_data_dir, "media")
        os.makedirs(self.media_dir, exist_ok=True)

        # Клавіатура не перекриватиме поле вводу
        Window.softinput_mode = 'below_target'

        if platform == 'android':
            # Результати голосового вводу та вибору файлів приходять сюди
            android_activity.bind(on_activity_result=self.on_activity_result)

        root = TabbedPanel(do_default_tab=False)

        # --- Вкладка 1: Група ---
        tab_group = TabbedPanelItem(text='Група')
        layout_group = BoxLayout(orientation='vertical', padding=15, spacing=15)

        layout_group.add_widget(Label(text="Номер та специфікація групи:", size_hint_y=None, height=30))
        self.group_input = TextInput(
            hint_text='Наприклад: Група №11, Молодшого віку',
            size_hint_y=None, height=50, multiline=False
        )
        layout_group.add_widget(self.group_input)

        btn_voice_group = Button(text='Надиктувати назву групи голосом', size_hint_y=None, height=50)
        btn_voice_group.bind(on_press=lambda x: self.start_speech_to_text(self.group_input))
        layout_group.add_widget(btn_voice_group)

        layout_group.add_widget(Label(text="Прізвище, ім'я дитини:", size_hint_y=None, height=30))
        self.student_input = TextInput(
            hint_text='Введіть або надиктуйте ПІБ дитини...',
            size_hint_y=None, height=50, multiline=False
        )
        layout_group.add_widget(self.student_input)

        btn_voice_student = Button(text='Надиктувати ПІБ голосом', size_hint_y=None, height=50)
        btn_voice_student.bind(on_press=lambda x: self.start_speech_to_text(self.student_input))
        layout_group.add_widget(btn_voice_student)

        layout_group.add_widget(Label())  # заповнювач, щоб елементи були зверху
        tab_group.add_widget(layout_group)
        root.add_widget(tab_group)

        # --- Вкладка 2: Запис ---
        tab_work = TabbedPanelItem(text='Запис')
        layout_work = BoxLayout(orientation='vertical', padding=15, spacing=15)

        layout_work.add_widget(Label(text="Режимні моменти / Короткий опис діяльності:", size_hint_y=None, height=30))
        self.note_input = TextInput(
            hint_text='Введіть нотатку або скористайтеся мікрофоном...',
            size_hint=(1, 1)
        )
        layout_work.add_widget(self.note_input)

        btn_voice_note = Button(text='Надиктувати нотатку голосом', size_hint_y=None, height=50,
                                background_color=(0.3, 0.2, 0.1, 1))
        btn_voice_note.bind(on_press=lambda x: self.start_speech_to_text(self.note_input))
        layout_work.add_widget(btn_voice_note)

        btn_media = Button(text='Додати фото / відео / аудіо', size_hint_y=None, height=50,
                           background_color=(0.2, 0.1, 0.3, 1))
        btn_media.bind(on_press=self.open_media_picker)
        layout_work.add_widget(btn_media)

        self.media_label = Label(text="Медіафайли не додано", size_hint_y=None, height=30)
        layout_work.add_widget(self.media_label)

        btn_save = Button(text='Зберегти у таблицю Excel', size_hint_y=None, height=50,
                          background_color=(0.1, 0.3, 0.2, 1))
        btn_save.bind(on_press=self.save_observation_to_excel)
        layout_work.add_widget(btn_save)

        tab_work.add_widget(layout_work)
        root.add_widget(tab_work)

        # --- Вкладка 3: Архів ---
        tab_archive = TabbedPanelItem(text='Архів')
        layout_archive = BoxLayout(orientation='vertical', padding=15, spacing=10)

        btn_export = Button(text='Експортувати в «Завантаження»', size_hint_y=None, height=50,
                            background_color=(0.1, 0.2, 0.35, 1))
        btn_export.bind(on_press=self.export_files)
        layout_archive.add_widget(btn_export)

        self.archive_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.archive_layout.bind(minimum_height=self.archive_layout.setter('height'))
        scroll_archive = ScrollView(size_hint=(1, 1))
        scroll_archive.add_widget(self.archive_layout)
        layout_archive.add_widget(scroll_archive)

        tab_archive.add_widget(layout_archive)
        root.add_widget(tab_archive)

        self.refresh_archive()
        return root

    # ------------------------------------------------------------------
    # Повідомлення користувачу (print на телефоні не видно)
    # ------------------------------------------------------------------
    @mainthread
    def show_message(self, text, title="Щоденник спостережень"):
        content = BoxLayout(orientation='vertical', padding=10, spacing=10)
        lbl = Label(text=text, halign='center', valign='middle')
        lbl.bind(size=lambda inst, val: setattr(inst, 'text_size', val))
        content.add_widget(lbl)
        btn = Button(text='OK', size_hint_y=None, height=50)
        content.add_widget(btn)
        popup = Popup(title=title, content=content, size_hint=(0.9, 0.45))
        btn.bind(on_press=popup.dismiss)
        popup.open()

    # ------------------------------------------------------------------
    # Голосовий ввід
    # ------------------------------------------------------------------
    def start_speech_to_text(self, target_input):
        self.active_text_input = target_input
        if platform != 'android':
            self._insert_spoken_text("Тестовий голосовий запис")
            return
        try:
            intent = Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH)
            intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE_MODEL, RecognizerIntent.LANGUAGE_MODEL_FREE_FORM)
            intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE, "uk-UA")
            intent.putExtra(RecognizerIntent.EXTRA_PROMPT, "Говоріть...")
            PythonActivity.mActivity.startActivityForResult(intent, REQUEST_SPEECH)
        except Exception as e:
            self.show_message(
                "Розпізнавання мовлення недоступне.\n"
                "Встановіть або оновіть застосунок Google.\n\n" + str(e)
            )

    @mainthread
    def _insert_spoken_text(self, spoken_text):
        target = self.active_text_input
        if target is None:
            return
        if target.multiline and target.text.strip():
            # У нотатку дописуємо, а не затираємо вже введене
            target.text = target.text.rstrip() + " " + spoken_text
        else:
            target.text = spoken_text

    # ------------------------------------------------------------------
    # Результати з Android (викликається НЕ в потоці Kivy)
    # ------------------------------------------------------------------
    def on_activity_result(self, request_code, result_code, intent):
        if result_code != RESULT_OK or intent is None:
            return
        try:
            if request_code == REQUEST_SPEECH:
                results = intent.getStringArrayListExtra(RecognizerIntent.EXTRA_RESULTS)
                if results is not None and results.size() > 0:
                    self._insert_spoken_text(results.get(0))

            elif request_code == REQUEST_MEDIA:
                uris = []
                clip = intent.getClipData()
                if clip is not None:
                    for i in range(clip.getItemCount()):
                        uris.append(clip.getItemAt(i).getUri())
                elif intent.getData() is not None:
                    uris.append(intent.getData())
                if uris:
                    threading.Thread(target=self._copy_media, args=(uris,), daemon=True).start()
        except Exception as e:
            self.show_message(f"Помилка обробки результату:\n{e}")

    # ------------------------------------------------------------------
    # Медіафайли
    # ------------------------------------------------------------------
    def open_media_picker(self, instance):
        if platform != 'android':
            self.show_message("Вибір медіафайлів доступний лише на Android.")
            return
        try:
            intent = Intent(Intent.ACTION_GET_CONTENT)
            intent.setType("*/*")
            intent.addCategory(Intent.CATEGORY_OPENABLE)
            try:
                intent.putExtra(Intent.EXTRA_MIME_TYPES, ["image/*", "video/*", "audio/*"])
            except Exception:
                pass  # без фільтра просто покаже всі файли
            intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, True)
            PythonActivity.mActivity.startActivityForResult(intent, REQUEST_MEDIA)
        except Exception as e:
            self.show_message(f"Не вдалося відкрити галерею:\n{e}")

    def _copy_media(self, uris):
        """Копіює вибрані файли у папку застосунку (працює у фоновому потоці)."""
        saved, errors = [], []
        try:
            resolver = PythonActivity.mActivity.getContentResolver()
            mime_map = MimeTypeMap.getSingleton()
            stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            for idx, uri in enumerate(uris, start=1):
                try:
                    mime = resolver.getType(uri)
                    ext = mime_map.getExtensionFromMimeType(mime) if mime else None
                    name = f"media_{stamp}_{idx}.{ext or 'bin'}"
                    dest = os.path.join(self.media_dir, name)
                    pfd = resolver.openFileDescriptor(uri, "r")
                    fd = pfd.detachFd()
                    with os.fdopen(fd, 'rb') as src, open(dest, 'wb') as dst:
                        shutil.copyfileobj(src, dst, 1024 * 1024)
                    saved.append(name)
                except Exception as e:
                    errors.append(str(e))
        finally:
            detach()  # обов'язково для потоків, що викликали Java
        self._on_media_copied(saved, errors)

    @mainthread
    def _on_media_copied(self, saved, errors):
        self.pending_media.extend(saved)
        self._update_media_label()
        self.refresh_archive()
        if errors:
            self.show_message(f"Додано файлів: {len(saved)}\nПомилок: {len(errors)}\n{errors[0]}")

    def _update_media_label(self):
        if self.pending_media:
            self.media_label.text = f"До запису прикріплено файлів: {len(self.pending_media)}"
        else:
            self.media_label.text = "Медіафайли не додано"

    # ------------------------------------------------------------------
    # Excel
    # ------------------------------------------------------------------
    def save_observation_to_excel(self, instance):
        group_name = self.group_input.text.strip()
        student_name = self.student_input.text.strip()
        note_text = self.note_input.text.strip()

        if not note_text and not self.pending_media:
            self.show_message("Введіть нотатку або додайте медіафайл перед збереженням.")
            return

        try:
            if os.path.exists(self.excel_path):
                wb = openpyxl.load_workbook(self.excel_path)
                ws = wb.active
                # Дописуємо заголовки нових колонок у старих файлах
                for col, header in enumerate(HEADERS, start=1):
                    if ws.cell(row=1, column=col).value is None:
                        ws.cell(row=1, column=col, value=header)
            else:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Щоденник спостережень"
                ws.append(HEADERS)
                for letter, width in zip("ABCDEFGHI", (17, 20, 28, 50, 22, 28, 18, 14, 30)):
                    ws.column_dimensions[letter].width = width

            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                group_name or "Не вказано",
                student_name or "Не вказано",
                note_text,
                "Самостійно / з допомогою",
                "Спостереження за діяльністю",
                "Згідно з планом",
                "Спокійна",
                ", ".join(self.pending_media),
            ])
            wb.save(self.excel_path)
        except Exception as e:
            self.show_message(f"Помилка запису в Excel:\n{e}")
            return

        self.note_input.text = ""
        self.pending_media = []
        self._update_media_label()
        self.refresh_archive()
        self.show_message("Запис збережено у таблицю.")

    # ------------------------------------------------------------------
    # Експорт у загальну папку «Завантаження»
    # (внутрішня папка застосунку недоступна з файлового менеджера)
    # ------------------------------------------------------------------
    def export_files(self, *args):
        if not os.path.exists(self.excel_path):
            self.show_message("Ще немає жодного запису для експорту.")
            return
        if platform != 'android':
            self.show_message(f"Файл таблиці:\n{self.excel_path}")
            return

        # Android 9 і старіші: потрібен дозвіл на запис
        if SDK_INT < 29 and not check_permission(Permission.WRITE_EXTERNAL_STORAGE):
            request_permissions([Permission.WRITE_EXTERNAL_STORAGE], self._on_storage_permission)
            return

        threading.Thread(target=self._export_worker, daemon=True).start()

    def _on_storage_permission(self, permissions, grants):
        if grants and all(grants):
            Clock.schedule_once(lambda dt: self.export_files())
        else:
            self.show_message("Без дозволу на сховище експорт неможливий.")

    def _export_worker(self):
        folder = f"{EXPORT_FOLDER}/{datetime.now().strftime('%Y-%m-%d_%H-%M')}"
        try:
            self._export_one(self.excel_path, EXCEL_NAME, XLSX_MIME, folder)
            media_count = 0
            for name in sorted(os.listdir(self.media_dir)):
                mime = mimetypes.guess_type(name)[0] or "application/octet-stream"
                self._export_one(os.path.join(self.media_dir, name), name, mime, folder)
                media_count += 1
            self.show_message(
                f"Експортовано в:\nЗавантаження/{folder}\n\n"
                f"Таблиця Excel + медіафайлів: {media_count}"
            )
        except Exception as e:
            self.show_message(f"Помилка експорту:\n{e}")
        finally:
            detach()

    def _export_one(self, src_path, display_name, mime, folder):
        downloads = Environment.DIRECTORY_DOWNLOADS
        if SDK_INT >= 29:
            MediaStoreDownloads = autoclass('android.provider.MediaStore$Downloads')
            resolver = PythonActivity.mActivity.getContentResolver()
            values = ContentValues()
            values.put("_display_name", display_name)
            values.put("mime_type", mime)
            values.put("relative_path", f"{downloads}/{folder}")
            uri = resolver.insert(MediaStoreDownloads.EXTERNAL_CONTENT_URI, values)
            if uri is None:
                raise RuntimeError(f"Не вдалося створити файл {display_name}")
            pfd = resolver.openFileDescriptor(uri, "w")
            fd = pfd.detachFd()
            with open(src_path, 'rb') as src, os.fdopen(fd, 'wb') as dst:
                shutil.copyfileobj(src, dst, 1024 * 1024)
        else:
            base = Environment.getExternalStoragePublicDirectory(downloads).getAbsolutePath()
            dest_dir = os.path.join(base, folder)
            os.makedirs(dest_dir, exist_ok=True)
            shutil.copy2(src_path, os.path.join(dest_dir, display_name))

    # ------------------------------------------------------------------
    # Архів
    # ------------------------------------------------------------------
    def refresh_archive(self):
        self.archive_layout.clear_widgets()
        entries = []

        if os.path.exists(self.excel_path):
            try:
                wb = openpyxl.load_workbook(self.excel_path, read_only=True)
                rows = max(wb.active.max_row - 1, 0)
                wb.close()
                entries.append(f"Таблиця: {EXCEL_NAME} (записів: {rows})")
            except Exception:
                entries.append(f"Таблиця: {EXCEL_NAME}")

        if os.path.isdir(self.media_dir):
            for name in sorted(os.listdir(self.media_dir), reverse=True):
                size_kb = os.path.getsize(os.path.join(self.media_dir, name)) // 1024
                entries.append(f"Медіа: {name} ({size_kb} КБ)")

        if not entries:
            entries.append("Архів порожній. Збережіть записи.")

        for text in entries:
            lbl = Label(text=text, size_hint_y=None, height=40, halign='left', valign='middle')
            lbl.bind(size=lambda inst, val: setattr(inst, 'text_size', val))
            self.archive_layout.add_widget(lbl)


if __name__ == '__main__':
    ObservationApp().run()
