[app]
title = Transactions
package.name = transactions
package.domain = org.perso

source.dir = .
source.include_exts = py,png,jpg,kv,atlas
source.exclude_dirs = .github,bin,.buildozer
source.exclude_patterns = test_logic.py,README.md,buildozer_b.spec
version = 1.1.0

# VARIANTE B (secours) : on fige python-for-android à la version de janvier 2024,
# celle qui accompagnait Kivy 2.3.0 (Python 3.11, NDK 25b, API 33).
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile,pyjnius,android
p4a.branch = v2024.01.21

orientation = all
fullscreen = 0
android.permissions =

android.api = 33
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a
android.accept_sdk_license = True
android.allow_backup = False

[buildozer]
log_level = 2
warn_on_root = 1
