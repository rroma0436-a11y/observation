[app]

# (str) Title of your application
title = Щоденник спостережень

# (str) Package name
package.name = observationdiary

# (str) Package domain (needed for android packaging)
package.domain = org.observation

# (str) Source code where the main.py live
source.dir = .

# (list) Source files to include (додайте сюди інші розширення, якщо є: ttf, json, db, wav тощо)
source.include_exts = py,png,jpg,kv,atlas,ttf,json

# (str) Application versioning
version = 2.0

# (list) Application requirements
requirements = python3,kivy,setuptools,plyer,openpyxl,et_xmlfile

# (str) Icon of the application
icon.filename = %(source.dir)s/icon.png

# (list) Supported orientations
orientation = portrait

# (bool) Indicate if the application should be fullscreen or not
fullscreen = 0

#
# Android specific
#

# (list) Permissions
# CAMERA і RECORD_AUDIO також треба запитати в коді через android.permissions.request_permissions
# WRITE/READ_EXTERNAL_STORAGE діють лише на старих Android (до 10-12), READ_MEDIA_* — на Android 13+
android.permissions = INTERNET, CAMERA, RECORD_AUDIO, WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE, READ_MEDIA_IMAGES, READ_MEDIA_AUDIO, READ_MEDIA_VIDEO

# (int) Target Android API (для Google Play потрібно 35)
android.api = 34

# (int) Minimum API your APK / AAB will support
android.minapi = 26

# (str) Android NDK version to use (закоментовано — buildozer підбере сумісну версію)
#android.ndk = 25b

# (bool) Automatically accept SDK license (потрібно для збірки в GitHub Actions / CI)
android.accept_sdk_license = True

# (bool) Enable AndroidX support
android.androidx = True

# (list) The Android archs to build for
# Лише одна архітектура: поточний python-for-android ламає pip, коли збирає дві архітектури поспіль
android.archs = arm64-v8a

# (bool) Use --private data storage (True) or --dir public storage (False)
android.private_storage = True

# (str) Формат релізної збірки: apk замість aab
android.release_artifact = apk

# --- Релізний підпис ---
# Buildozer НЕ читає ключ зі spec-файлу. Перед "buildozer android release" задайте змінні оточення:
#   export P4A_RELEASE_KEYSTORE=$PWD/observation-release-key.keystore
#   export P4A_RELEASE_KEYSTORE_PASSWD=<пароль сховища>
#   export P4A_RELEASE_KEYALIAS=observation_alias
#   export P4A_RELEASE_KEYALIAS_PASSWD=<пароль ключа>
# У GitHub Actions передайте їх через env: з secrets.


[buildozer]

# (int) Log level (0 = error only, 1 = info, 2 = debug (with command output))
log_level = 2

# (int) Display warning if buildozer is run as root (0 = False, 1 = True)
warn_on_root = 1
