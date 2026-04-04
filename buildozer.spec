[app]

title = Biodiversity Inventory
package.name = biodiversityinventory
package.domain = org.moctar

source.dir = .
source.include_exts = py,png,jpg,kv

version = 1.0

requirements = python3,kivy,plyer,requests,openpyxl

orientation = portrait
fullscreen = 0

# Permissions Android
android.permissions = INTERNET,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE

# Versions Android (très important)
android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b

# Architecture supportée
android.archs = arm64-v8a, armeabi-v7a


[buildozer]

log_level = 2
warn_on_root = 1
