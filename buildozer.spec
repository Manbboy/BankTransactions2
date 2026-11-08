[app]
title = Transactions
package.name = transactions
package.domain = org.perso

source.dir = .
source.include_exts = py,png,jpg,kv,atlas
source.exclude_dirs = .github,bin,.buildozer
source.exclude_patterns = test_logic.py,README.md,buildozer_b.spec
version = 1.1.0

# VARIANTE A : python-for-android actuel, mais avec Python 3.11 imposé.
# (Par défaut il prend Python 3.14, avec lequel Kivy 2.3 ne compile pas.)
# openpyxl est du Python pur ; et_xmlfile est sa seule dépendance obligatoire.
requirements = python3==3.11.5,hostpython3==3.11.5,kivy==2.3.0,openpyxl,et_xmlfile,pyjnius,android

orientation = all
fullscreen = 0

# Aucune permission : le fichier est choisi via le sélecteur de documents Android (SAF)
# puis copié dans le stockage privé de l'app. Pas de INTERNET -> 100 % hors ligne.
android.permissions =

android.api = 34
android.minapi = 24
android.archs = arm64-v8a
android.accept_sdk_license = True
android.allow_backup = False

[buildozer]
log_level = 2
warn_on_root = 1
