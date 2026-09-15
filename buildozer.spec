[app]
title = Avancement Horizontal
package.name = avancementhorizontal
package.domain = dz.avancement
source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,xlsx,xlsm,csv
version = 16.1
requirements = python3,kivy,openpyxl
orientation = portrait
fullscreen = 0

[buildozer]
log_level = 2
warn_on_root = 1

[app:android]
android.api = 35
android.permissions = INTERNET
android.minapi = 23
android.archs = arm64-v8a,armeabi-v7a
