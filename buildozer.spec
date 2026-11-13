[app]

# (str) Title of your application
title = Щоденник спостережень

# (str) Package name
package.name = observationdiary

# (str) Package domain (needed for android packaging)
package.domain = org.observation

# (str) Source files where the let's go (relative to directory of spec)
source.include_exts = py,png,jpg,kv,atlas

# (str) Directory where the source files are located
source.dir = .

# (string) Version of the application
version = 1.0

# (list) Application requirements (включно з openpyxl та et_xmlfile для Excel)
requirements = python3,kivy,setuptools,plyer,openpyxl,et_xmlfile

# (str) Icon of the application
icon.filename = %(source.dir)s/icon.png

# (list) Permissions (камера, мікрофон, зовнішня пам'ять та медіафайли)
android.permissions = INTERNET, CAMERA, RECORD_AUDIO, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, READ_MEDIA_IMAGES, READ_MEDIA_AUDIO, READ_MEDIA_VIDEO

# (list) Supported orientations
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

# (bool) Enable AndroidX support
android.androidx = True

# (list) The android api to target that you want for the application
android.api = 34

# (list) Minimum API your APK / AAB will support (Піднято до 26 для сумісності з Android 8.0+)
android.minapi = 26

# (str) Explicit Android SDK version to use
android.sdk = 34

# (str) Android NDK version to use
android.ndk = 25b

# (str) The Android arch to build for (Додано arm64-v8a для сучасних 64-бітних пристроїв)
android.archs = arm64-v8a

# (bool) Use --private data storage (True) or public storage (False)
android.private_storage = True

# --- Налаштування релізного цифрового підпису ---
android.release_keystore = observation-release-key.keystore
android.release_keyalias = observation_alias

# --- Змушуємо збирати саме APK замість AAB ---
android.release_artifact = apk

[buildozer]
