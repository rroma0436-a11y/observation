import os
import openpyxl
from datetime import datetime
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.utils import platform

# Підключення Android API через pyjnius для голосового вводу та дозволів
if platform == 'android':
    from jnius import autoclass
    from android.permissions import request_permissions, Permission
    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    Intent = autoclass('android.content.Intent')
    RecognizerIntent = autoclass('speech.RecognizerIntent')

class ObservationApp(App):
    def build(self):
        self.title = "Щоденник спостережень"
        self.active_text_input = None  # Запам'ятовуємо активне поле для вставки тексту
        
        # Запит необхідних дозволів Android при запуску
        if platform == 'android':
            try:
                request_permissions([
                    Permission.RECORD_AUDIO,
                    Permission.CAMERA,
                    Permission.WRITE_EXTERNAL_STORAGE,
                    Permission.READ_EXTERNAL_STORAGE
                ])
            except Exception as e:
                print(f"Помилка запиту дозволів: {e}")
            
            try:
                PythonActivity.mActivity.bind(on_activity_result=self.on_activity_result)
            except Exception as e:
                print(f"Помилка прив'язки on_activity_result: {e}")

        root = TabbedPanel()
        root.do_default_tab = False
        
        # --- Вкладка 1: Налаштування групи ---
        tab_group = TabbedPanelItem(text='Група')
        layout_group = BoxLayout(orientation='vertical', padding=15, spacing=15)
        
        layout_group.add_widget(Label(text="Номер та специфікація групи:", size_hint_y=None, height=30))
        
        self.group_input = TextInput(
            hint_text='Наприклад: Група №11, Молодшого віку',
            size_hint_y=None, height=50, multiline=False
        )
        self.group_input.bind(focus=self.on_text_focus)
        layout_group.add_widget(self.group_input)
        
        btn_voice_group = Button(text='🎤 Надиктувати назву групи голосом', size_hint_y=None, height=50)
        btn_voice_group.bind(on_press=lambda x: self.start_speech_to_text(self.group_input))
        layout_group.add_widget(btn_voice_group)
        
        layout_group.add_widget(Label(text="Прізвище, ім'я дитини / Список вихованців:", size_hint_y=None, height=30))
        
        self.student_input = TextInput(
            hint_text='Введіть або надиктуйте ПІБ дитини...',
            size_hint_y=None, height=50, multiline=False
        )
        self.student_input.bind(focus=self.on_text_focus)
        layout_group.add_widget(self.student_input)
        
        btn_voice_student = Button(text='🎤 Надиктувати ПІБ голосом', size_hint_y=None, height=50)
        btn_voice_student.bind(on_press=lambda x: self.start_speech_to_text(self.student_input))
        layout_group.add_widget(btn_voice_student)
        
        tab_group.add_widget(layout_group)
        root.add_widget(tab_group)
        
        # --- Вкладка 2: Робочий кабінет (Запис) ---
        tab_work = TabbedPanelItem(text='Запис')
        layout_work = BoxLayout(orientation='vertical', padding=15, spacing=15)
        
        layout_work.add_widget(Label(text="Режимні моменти / Короткий опис діяльності:", size_hint_y=None, height=30))
        
        self.note_input = TextInput(
            hint_text='Введіть нотатку або скористайтеся мікрофоном...',
            size_hint=(1, 1)
        )
        self.note_input.bind(focus=self.on_text_focus)
        layout_work.add_widget(self.note_input)
        
        btn_voice_note = Button(text='🎤 Надиктувати нотатку голосом (UA)', size_hint_y=None, height=50, background_color=(0.3, 0.2, 0.1, 1))
        btn_voice_note.bind(on_press=lambda x: self.start_speech_to_text(self.note_input))
        layout_work.add_widget(btn_voice_note)
        
        btn_media = Button(text='📷 Додати медіафайл / фото з галереї', size_hint_y=None, height=50, background_color=(0.2, 0.1, 0.3, 1))
        btn_media.bind(on_press=self.open_media_picker)
        layout_work.add_widget(btn_media)
        
        btn_save = Button(text='💾 Зберегти у таблицю Excel', size_hint_y=None, height=50, background_color=(0.1, 0.3, 0.2, 1))
        btn_save.bind(on_press=self.save_observation_to_excel)
        layout_work.add_widget(btn_save)
        
        tab_work.add_widget(layout_work)
        root.add_widget(tab_work)
        
        # --- Вкладка 3: Архів ---
        self.tab_archive = TabbedPanelItem(text='Архів')
        self.archive_layout = GridLayout(cols=1, spacing=10, size_hint_y=None)
        self.archive_layout.bind(minimum_height=self.archive_layout.setter('height'))
        
        scroll_archive = ScrollView(size_hint=(1, 1))
        scroll_archive.add_widget(self.archive_layout)
        
        self.tab_archive.add_widget(scroll_archive)
        root.add_widget(self.tab_archive)
        
        self.refresh_archive()
        return root

    def on_text_focus(self, instance, value):
        if value:
            self.active_text_input = instance

    def start_speech_to_text(self, target_input):
        self.active_text_input = target_input
        if platform == 'android':
            try:
                activity = PythonActivity.mActivity
                intent = Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH)
                intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE_MODEL, RecognizerIntent.LANGUAGE_MODEL_FREE_FORM)
                intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE, "uk-UA")
                intent.putExtra(RecognizerIntent.EXTRA_PROMPT, "Говоріть...")
                activity.startActivityForResult(intent, 101)
            except Exception as e:
                target_input.text = "[Помилка мікрофона]"
        else:
            # Тестовий режим для ПК
            target_input.text = "Тестовий голосовий запис"

    def on_activity_result(self, requestCode, resultCode, intent):
        # Обробка результату розпізнавання голосу з Android Intent
        if requestCode == 101 and resultCode == -1:  # -1 це RESULT_OK
            if intent is not None:
                try:
                    results = intent.getStringArrayListExtra(RecognizerIntent.EXTRA_RESULTS)
                    if results is not None and results.size() > 0:
                        spoken_text = results.get(0)
                        if self.active_text_input:
                            self.active_text_input.text = spoken_text
                except Exception as e:
                    print(f"Помилка отримання голосу: {e}")

    def open_media_picker(self, instance):
        # Логіка додавання медіафайлу до папки сховища
        try:
            demo_file = os.path.join(self.user_data_dir, f"media_{datetime.now().strftime('%H%M%S')}.jpg")
            with open(demo_file, 'w') as f:
                f.write("media placeholder")
        except Exception as e:
            print(f"Помилка додавання медіа: {e}")
        self.refresh_archive()

    def save_observation_to_excel(self, instance):
        current_date = datetime.now().strftime("%Y-%m-%d %H:%M")
        group_name = self.group_input.text if hasattr(self, 'group_input') else "Група №11"
        student_name = self.student_input.text if hasattr(self, 'student_input') else "Не вказано"
        note_text = self.note_input.text if hasattr(self, 'note_input') else ""
        
        file_path = os.path.join(self.user_data_dir, "Shodennyk_Sposterejen_ZDO.xlsx")
        
        try:
            if not os.path.exists(file_path):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Щоденник спостережень"
                ws.append(["Дата", "Група", "Прізвище, ім'я дитини", "Режимні моменти", "Кількість завдань", "Труднощі", "Потреби", "Поведінка"])
                wb.save(file_path)
            
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            
            new_row = [
                current_date,
                group_name,
                student_name,
                note_text,
                "Самостійно / з допомогою",
                "Спостереження за діяльністю",
                "Згідно з планом",
                "Спокійна"
            ]
            ws.append(new_row)
            wb.save(file_path)
            
            if hasattr(self, 'note_input'):
                self.note_input.text = ""
                
            self.refresh_archive()
            print("Дані успішно збережено у Excel!")
        except Exception as e:
            print(f"Помилка запису в Excel: {e}")

    def refresh_archive(self):
        self.archive_layout.clear_widgets()
        user_dir = self.user_data_dir
        if os.path.exists(user_dir):
            files = [f for f in os.listdir(user_dir) if f.endswith(('.xlsx', '.png', '.jpg', '.txt'))]
            if not files:
                self.archive_layout.add_widget(Label(text="Архів порожній. Збережіть записи.", size_hint_y=None, height=40))
            for file in files:
                self.archive_layout.add_widget(Label(text=f"Файл: {file}", size_hint_y=None, height=40))

if __name__ == '__main__':
    ObservationApp().run()