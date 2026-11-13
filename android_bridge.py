"""Робота з Android: голосовий ввід, вибір фото/відео, експорт у «Завантаження», надсилання файлів.

На комп'ютері всі функції працюють у тестовому режимі, щоб програму можна було запустити й там.
"""
import os
import shutil
import mimetypes
from datetime import datetime

from kivy.utils import platform

IS_ANDROID = platform == "android"

REQUEST_SPEECH = 101
REQUEST_MEDIA = 102
RESULT_OK = -1

if IS_ANDROID:
    from jnius import autoclass, cast, detach as _detach
    from android import activity as _activity
    from android.permissions import request_permissions, check_permission, Permission

    PythonActivity = autoclass("org.kivy.android.PythonActivity")
    Intent = autoclass("android.content.Intent")
    RecognizerIntent = autoclass("android.speech.RecognizerIntent")
    MimeTypeMap = autoclass("android.webkit.MimeTypeMap")
    ContentValues = autoclass("android.content.ContentValues")
    Environment = autoclass("android.os.Environment")
    JFile = autoclass("java.io.File")
    Uri = autoclass("android.net.Uri")
    JString = autoclass("java.lang.String")
    SDK_INT = autoclass("android.os.Build$VERSION").SDK_INT
else:
    SDK_INT = 0


def cs(text):
    """Python-рядок -> Java CharSequence (потрібно для заголовків вікон і сповіщень)."""
    return cast("java.lang.CharSequence", JString(text))


def detach():
    """Викликати в кінці кожного фонового потоку, що звертався до Java."""
    if IS_ANDROID:
        try:
            _detach()
        except Exception:
            pass


def bind_activity_result(callback):
    if IS_ANDROID:
        _activity.bind(on_activity_result=callback)


def _activity_obj():
    return PythonActivity.mActivity


# ----------------------------------------------------------------------
# Голос
# ----------------------------------------------------------------------
def start_speech(prompt="Говоріть..."):
    intent = Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH)
    intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE_MODEL, RecognizerIntent.LANGUAGE_MODEL_FREE_FORM)
    intent.putExtra(RecognizerIntent.EXTRA_LANGUAGE, "uk-UA")
    intent.putExtra(RecognizerIntent.EXTRA_PROMPT, prompt)
    _activity_obj().startActivityForResult(intent, REQUEST_SPEECH)


def speech_text(intent):
    results = intent.getStringArrayListExtra(RecognizerIntent.EXTRA_RESULTS)
    if results is not None and results.size() > 0:
        return results.get(0)
    return None


# ----------------------------------------------------------------------
# Медіафайли
# ----------------------------------------------------------------------
def pick_media():
    intent = Intent(Intent.ACTION_GET_CONTENT)
    intent.setType("*/*")
    intent.addCategory(Intent.CATEGORY_OPENABLE)
    try:
        intent.putExtra(Intent.EXTRA_MIME_TYPES, ["image/*", "video/*"])
    except Exception:
        pass
    intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, True)
    _activity_obj().startActivityForResult(intent, REQUEST_MEDIA)


def uris_from_result(intent):
    uris = []
    clip = intent.getClipData()
    if clip is not None:
        for i in range(clip.getItemCount()):
            uris.append(clip.getItemAt(i).getUri())
    elif intent.getData() is not None:
        uris.append(intent.getData())
    return uris


def copy_uris(uris, dest_dir):
    """Копіює вибрані файли в папку застосунку. Виконувати у фоновому потоці."""
    saved, errors = [], []
    resolver = _activity_obj().getContentResolver()
    mime_map = MimeTypeMap.getSingleton()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for idx, uri in enumerate(uris, start=1):
        try:
            mime = resolver.getType(uri)
            ext = mime_map.getExtensionFromMimeType(mime) if mime else None
            name = f"media_{stamp}_{idx}.{ext or 'bin'}"
            pfd = resolver.openFileDescriptor(uri, "r")
            fd = pfd.detachFd()
            with os.fdopen(fd, "rb") as src, open(os.path.join(dest_dir, name), "wb") as dst:
                shutil.copyfileobj(src, dst, 1024 * 1024)
            saved.append(name)
        except Exception as e:
            errors.append(str(e))
    return saved, errors


# ----------------------------------------------------------------------
# Експорт у загальну папку «Завантаження»
# ----------------------------------------------------------------------
def guess_mime(name):
    if name.lower().endswith(".xlsx"):
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return mimetypes.guess_type(name)[0] or "application/octet-stream"


def needs_storage_permission():
    return IS_ANDROID and SDK_INT < 29 and not check_permission(Permission.WRITE_EXTERNAL_STORAGE)


def request_storage_permission(callback):
    """callback(granted: bool) викликається з іншого потоку."""
    def _cb(permissions, grants):
        callback(bool(grants) and all(grants))
    request_permissions([Permission.WRITE_EXTERNAL_STORAGE], _cb)


def export_file(src_path, folder, display_name=None):
    """Копіює файл у «Завантаження/<folder>». Повертає Uri (Android) або шлях (ПК)."""
    display_name = display_name or os.path.basename(src_path)
    if not IS_ANDROID:
        dest_dir = os.path.join(os.path.expanduser("~"), "Downloads", folder)
        os.makedirs(dest_dir, exist_ok=True)
        dest = os.path.join(dest_dir, display_name)
        shutil.copy2(src_path, dest)
        return dest

    downloads = Environment.DIRECTORY_DOWNLOADS
    if SDK_INT >= 29:
        MediaStoreDownloads = autoclass("android.provider.MediaStore$Downloads")
        resolver = _activity_obj().getContentResolver()
        values = ContentValues()
        values.put("_display_name", display_name)
        values.put("mime_type", guess_mime(display_name))
        values.put("relative_path", f"{downloads}/{folder}")
        uri = resolver.insert(MediaStoreDownloads.EXTERNAL_CONTENT_URI, values)
        if uri is None:
            raise RuntimeError(f"Не вдалося створити файл {display_name}")
        pfd = resolver.openFileDescriptor(uri, "w")
        fd = pfd.detachFd()
        with open(src_path, "rb") as src, os.fdopen(fd, "wb") as dst:
            shutil.copyfileobj(src, dst, 1024 * 1024)
        return uri

    base = Environment.getExternalStoragePublicDirectory(downloads).getAbsolutePath()
    dest_dir = os.path.join(base, folder)
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, display_name)
    shutil.copy2(src_path, dest)
    return dest


# ----------------------------------------------------------------------
# Надсилання (Viber, Telegram, пошта…)
# ----------------------------------------------------------------------
def share(target, mime, title="Надіслати"):
    """target — результат export_file()."""
    if not IS_ANDROID:
        return
    if isinstance(target, str):
        # Android 8–9: дозволяємо передати file:// посилання
        StrictMode = autoclass("android.os.StrictMode")
        VmPolicyBuilder = autoclass("android.os.StrictMode$VmPolicy$Builder")
        StrictMode.setVmPolicy(VmPolicyBuilder().build())
        uri = Uri.fromFile(JFile(target))
    else:
        uri = target
    intent = Intent(Intent.ACTION_SEND)
    intent.setType(mime)
    intent.putExtra(Intent.EXTRA_STREAM, cast("android.os.Parcelable", uri))
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
    chooser = Intent.createChooser(intent, cs(title))
    _activity_obj().startActivity(chooser)


# ----------------------------------------------------------------------
# Дозволи при першому запуску
# ----------------------------------------------------------------------
def startup_permissions():
    """Список дозволів, які застосунок запитує одразу після встановлення."""
    perms = ["android.permission.RECORD_AUDIO", "android.permission.CAMERA"]
    if SDK_INT >= 33:
        perms += ["android.permission.POST_NOTIFICATIONS",
                  "android.permission.READ_MEDIA_IMAGES",
                  "android.permission.READ_MEDIA_VIDEO"]
    else:
        perms.append("android.permission.READ_EXTERNAL_STORAGE")
        if SDK_INT < 29:
            perms.append("android.permission.WRITE_EXTERNAL_STORAGE")
    return perms


def request_startup_permissions():
    if not IS_ANDROID:
        return
    missing = [p for p in startup_permissions() if not check_permission(p)]
    if missing:
        request_permissions(missing)


# ----------------------------------------------------------------------
# Сповіщення
# ----------------------------------------------------------------------
_CHANNEL = "diary_main"
_notify_id = [1000]


def notify(title, text):
    """Показує сповіщення в шторці телефону. Помилки ігноруються."""
    if not IS_ANDROID:
        return
    try:
        Context = autoclass("android.content.Context")
        NotificationManager = autoclass("android.app.NotificationManager")
        Builder = autoclass("android.app.Notification$Builder")
        act = _activity_obj()
        nm = cast("android.app.NotificationManager",
                  act.getSystemService(Context.NOTIFICATION_SERVICE))
        if SDK_INT >= 26:
            Channel = autoclass("android.app.NotificationChannel")
            nm.createNotificationChannel(
                Channel(_CHANNEL, cs("Щоденник спостережень"), NotificationManager.IMPORTANCE_DEFAULT))
            builder = Builder(act, _CHANNEL)
        else:
            builder = Builder(act)
        builder.setSmallIcon(act.getApplicationInfo().icon)
        builder.setContentTitle(cs(title))
        builder.setContentText(cs(text))
        builder.setAutoCancel(True)
        _notify_id[0] += 1
        nm.notify(_notify_id[0], builder.build())
    except Exception:
        pass
