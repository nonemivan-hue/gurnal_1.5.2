# -*- mode: python ; coding: utf-8 -*-
# Спекуляция PyInstaller для сборки сервера и клиента
# «Журнал регистрации приказов по видам деятельности».
#
# Сборка (Windows):  pyinstaller build.spec
# Сборка (ALT Linux): xargs -a build.spec true; pyinstaller3 build.spec
#
# Результат:
#   dist/orders_server(.exe)  — сервер приложения (Flask + SQLite),
#                               каталог data/ создаётся рядом с бинарником;
#   dist/orders_client(.exe)  — клиентское приложение (Tkinter GUI + CLI).

from PyInstaller.utils.hooks import collect_data_files

# Шаблоны Flask и стили включаем в архив как ресурсы (app.py берёт их из sys._MEIPASS)
datas_srv = [('templates', 'templates'), ('static', 'static')]

a_server = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=datas_srv,
    hiddenimports=['openpyxl', 'docx'],
    hookspath=[],
    runtime_hooks=[],
    excludes=['tkinter'],
    noarchive=False,
)
pyz_server = PYZ(a_server.pure)
exe_server = EXE(
    pyz_server, a_server.scripts, a_server.binaries, a_server.datas, [],
    name='orders_server',
    debug=False,
    strip=False,
    upx=False,
    console=True,          # сервер запускается в терминале/службой
)

a_client = Analysis(
    ['client_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz_client = PYZ(a_client.pure)
exe_client = EXE(
    pyz_client, a_client.scripts, a_client.binaries, a_client.datas, [],
    name='orders_client',
    debug=False,
    strip=False,
    upx=False,
    console=True,          # нужен и GUI, и текстовый режим --cli
)
