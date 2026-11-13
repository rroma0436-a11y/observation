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

# (list) Application requirements
requirements = python3,kivy,setuptools,plyer,openpyxl,et_xmlfile

# (str) Icon of the application
icon.filename = %(source.dir)s/icon.png

# (list) Permissions
android.permissions = INTERNET, CAMERA, RECORD_AUDIO, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, READ_MEDIA_IMAGES, READ_MEDIA_AUDIO, READ_MEDIA_VIDEO

# (list) Supported orientations
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

# (bool) Enable AndroidX support
android.androidx = True

# (list) The android api to target that you want for the application
android.api = 34

# (list) Minimum API your APK / AAB will support
android.minapi = 26

# (str) Explicit Android SDK version to use (можна залишити закоментованим або вказати 34)
#android.sdk = 34

# (str) Android NDK version to use (закоментовуємо, щоб дія сама підібрала сумісну версію r25b або новішу)
#android.ndk = 25b

# (str) The Android arch to build for
android.archs = arm64-v8a

# (bool) Use --private data storage (True) or public storage (False)
android.private_storage = True

# --- Налаштування релізного цифрового підпису ---
android.release_keystore = observation-release-key.keystore
android.release_keyalias = observation_alias

# --- Змушуємо збирати саме APK замість AAB ---
android.release_artifact = apk

[buildozer]

# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root (0 = False, 1 = True)
warn_on_root = 1
