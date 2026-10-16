# Сборка приложения в исполняемые файлы (Windows и ALT Linux)

Настоящая инструкция описывает, как собрать сервер и клиент программы
«Журнал регистрации приказов по видам деятельности» в самостоятельные
исполняемые файлы (без установленного Python на целевом компьютере) с помощью
PyInstaller. Готовые сборки различаются для Windows и ALT Linux — **собирать
нужно на той ОС, под которую планируется запуск** (кросс-сборки PyInstaller не
поддерживает).

Состав сборки:

| Бинарник | Что это | Окно консоли |
|---|---|---|
| `orders_server(.exe)` | Сервер приложения (Flask + SQLite, веб-интерфейс, REST API) | есть (можно запускать службой) |
| `orders_client(.exe)` | Клиентское приложение: графический интерфейс Tkinter + текстовый режим `--cli` | есть |

Результат сборки помещается в каталог `dist/`. Каталог с данными `data/`
(база, сканы, логи, архивы) создаётся автоматически рядом с бинарником сервера
при первом запуске; при переносе между платформами достаточно скопировать его
целиком (см. раздел 5).

---

## 0. Краткая инструкция (для опытных)

**Windows** (из каталога проекта, в `cmd` или PowerShell):

```bat
py -m pip install -r requirements.txt pyinstaller
py -m PyInstaller --noconfirm build.spec
dist\orders_server.exe
```

**ALT Linux** (из каталога проекта):

```sh
pip3 install -r requirements.txt pyinstaller   # либо apt-get, см. п. 1.2
python3 -m PyInstaller --noconfirm build.spec
./dist/orders_server
```

Результат — в каталоге `dist/`: `orders_server(.exe)` и `orders_client(.exe)`.
Подробности подготовки, запуска, автозапуска и переноса данных — далее.

---

## 1. Подготовка: установка инструментов сборки

### 1.1. Windows

1. Установите Python 3.10–3.13 с https://python.org (при установке отметьте
   **«Add Python to PATH»**). Tkinter входит в стандартный установщик.
2. Откройте «Командную строку» (`cmd`) или PowerShell и выполните:

   ```bat
   cd C:\orders_journal
   py -m pip install -r requirements.txt
   py -m pip install pyinstaller
   ```

### 1.2. ALT Linux (Альт, Alt Server и т.п.)

Установка из репозитория APT (от root):

```sh
apt-get install python3 python3-module-pip python3-module-flask \
    python3-module-openpyxl python3-module-docx python3-module-pyinstaller \
    python3-tkinter
```

Если пакета `python3-module-pyinstaller` нет в вашем источнике:

```sh
pip3 install --user pyinstaller flask openpyxl python-docx
```

Проверка готовности: `python3 -c "import flask, openpyxl, docx; print('ok')"`
и `pyinstaller3 --version` (или `python3 -m PyInstaller --version`).

---

## 2. Сборка сервера и клиента

В каталоге проекта лежит готовая спекуляция **`build.spec`**, которая собирает
оба приложения одной командой (шаблоны `templates/` и стили `static/`
включаются внутрь бинарника как ресурсы).

### 2.1. Windows

```bat
cd C:\orders_journal
py -m PyInstaller --noconfirm build.spec
```

Готово: `dist\orders_server.exe`, `dist\orders_client.exe`.

### 2.2. ALT Linux

```sh
cd /opt/orders_journal
pyinstaller3 --noconfirm build.spec        # либо:
python3 -m PyInstaller --noconfirm build.spec
```

Готово: `dist/orders_server`, `dist/orders_client` (ELF-исполняемые файлы).

### 2.3. Сборка без spec-файла (одной командой)

Если нужно собрать только сервер или только клиент:

```sh
# сервер (ресурсы включаются параметрами --add-data; ; — Windows, : — Linux)
pyinstaller --noconfirm --onefile --name orders_server ^
    --add-data "templates;templates" --add-data "static;static" ^
    --hidden-import openpyxl --hidden-import docx app.py          # Windows

pyinstaller --noconfirm --onefile --name orders_server \
    --add-data "templates:templates" --add-data "static:static" \
    --hidden-import openpyxl --hidden-import docx app.py          # ALT Linux

# клиент (зависимостей из stdlib нет)
pyinstaller --noconfirm --onefile --name orders_client client_app.py
```

Типичное время сборки: 1–3 минуты. Служебные каталоги `build/` и файлы `*.spec`
можно удалить после получения результата — в `dist/` остаётся сам бинарник.

---

## 3. Запуск собранного приложения

### 3.1. Сервер

```bat
:: Windows
cd C:\orders_journal
dist\orders_server.exe                  :: порт по умолчанию 8080
dist\orders_server.exe --port 8090      :: другой порт
set HOST=0.0.0.0 & set PORT=8090 & dist\orders_server.exe
```

```sh
# ALT Linux (из терминала; адрес и порт — так же, как у run.sh)
./dist/orders_server --port 8090
HOST=0.0.0.0 PORT=8090 ./dist/orders_server
```

Порт по умолчанию — **8080** (порт 5000 занят другой программой); если порт
занят, сервер автоматически подбирает ближайший свободный и сообщает об этом.
После первого запуска рядом с бинарником создаётся каталог `data/` с базой;
вход: `admin / admin` (немедленно смените пароль!). Веб-интерфейс —
`http://адрес_сервера:порт`.

Автозапуск сервера как службы в ALT Linux — см. INSTALL.md, раздел 7
(systemd unit), только вместо `ExecStart=/usr/bin/python3 .../app.py` укажите
`ExecStart=/opt/orders_journal/dist/orders_server`.

**Автозапуск сервера на Windows (подробное описание — INSTALL.md, раздел 7.4).**

Способ 1 — служба Windows через NSSM (рекомендуется; запуск при загрузке ОС,
автоматический перезапуск при сбое). PowerShell от имени администратора:

```powershell
nssm install JournalPrikazov "C:\orders_journal\dist\orders_server.exe" "--port 8080"
nssm set JournalPrikazov AppDirectory "C:\orders_journal"
nssm set JournalPrikazov Start SERVICE_AUTO_START
nssm set JournalPrikazov AppStdout "C:\orders_journal\data\logs\service.out.log"
nssm set JournalPrikazov AppStderr "C:\orders_journal\data\logs\service.err.log"
nssm set JournalPrikazov AppRestartDelay 5000
nssm start JournalPrikazov
```

Способ 2 — готовый скрипт `scripts\windows\install_autostart.ps1`: сам выбирает NSSM,
а если его нет — регистрирует задачу Планировщика заданий с запуском при входе
пользователя в систему. Управление службой: `Get-Service JournalPrikazov`,
`Restart-Service JournalPrikazov` или оснастка services.msc. Удаление:
`scripts\windows\uninstall_autostart.ps1`.

Способ 3 — задача Планировщика заданий без NSSM (запуск при старте системы):

```bat
schtasks /create /tn JournalPrikazov /tr "C:\orders_journal\dist\orders_server.exe --port 8080" ^
    /sc onstart /ru SYSTEM /rl HIGHEST
```

После установки откройте порт в брандмауэре для доступа сотрудников по сети:

```bat
netsh advfirewall firewall add rule name="JournalPrikazov" dir=in action=allow protocol=TCP localport=8080
```

### 3.2. Клиент

Графический интерфейс (нужен рабочий стол; на сервере без X/Tkinter клиент
автоматически переключится в текстовый режим):

```bat
dist\orders_client.exe
dist\orders_client.exe --server http://192.168.0.10:8080
```

```sh
./dist/orders_client --server http://192.168.0.10:8080
```

Текстовый режим по SSH / без GUI:

```sh
./dist/orders_client --server http://192.168.0.10:8080 --cli --login ivanov
```

На ALT Linux для GUI-режима должны быть установлены `python3-tkinter` (для
сборки) и системные библиотеки Tk (`libtk`, `libtcl`) — обычно тянутся
зависимостями. На минимальных серверных образах ALT без X-окружения
используйте `--cli`.

---

## 4. Особенности бинарной сборки (что нужно знать)

1. **Поддержка frozen-режима вшита в `app.py`**: при запуске из PyInstaller
   ресурсы (шаблоны, стили) берутся из внутриархивного каталога `_MEIPASS`,
   а рабочие данные `data/` — из каталога рядом с `.exe`/бинарником. Поэтому
   обновление программы заменой файла НЕ трогает базу, сканы и логи.
2. **Однофайловая сборка (`--onefile`)** при старте распаковывается во временный
   каталог — первый запуск может занимать 2–5 секунд. Это нормально.
3. **Антивирусы Windows** иногда реагируют на PyInstaller-EXE (упаковка
   загрузчика). Если корпоративный антивирус блокирует запуск — добавьте файл
   в исключения или соберите вариант `--onedir` (папка с exe + dll):

   ```bat
   pyinstaller --noconfirm --onedir --name orders_server ^
       --add-data "templates;templates" --add-data "static;static" app.py
   ```

   В этом случае распространяйте всю папку `dist\orders_server\`.
4. **Мягкий перезапуск (обновление без выключения сервера)** в бинарном режиме
   работает иначе: watchdog-цикл `run.sh` недоступен внутри EXE. После загрузки
   ZIP-обновления через «Сервис» программа создаст флаг `data/restart_pending`
   и завершится; перезапустите её службой/systemd/schtasks (Restart=always или
   `/sc onstart`) — новая версия поднимется автоматически. Подробности —
   UPDATE_HOT.md.
5. **Клиент** не требует установки зависимостей: GUI (Tkinter) и REST-API-клиент
   собираются из стандартной библиотеки.

---

## 5. Перенос данных между платформами (Windows ↔ ALT Linux)

Формат базы (SQLite) и файлов одинаков на всех платформах:

1. Остановьте сервер на старой платформе.
2. Скопируйте каталог `data/` целиком (или сделайте резервную копию через
   «Сервис → Архивирование» и скачайте ZIP).
3. Разместите его рядом с бинарником новой платформы и запустите — база,
   сканы, пользователи и права сохранятся.

Пересобранная под другую ОС программа читает ту же базу без миграции.

---

## 6. Чистка после сборки

```sh
rm -rf build/ *.egg-info __pycache__     # служебные артефакты PyInstaller
```

`dist/` и `build.spec` оставьте — там результат и конфигурация сборки.
В репозитории артефакты сборки не версионируются (см. .gitignore).

---

## 7. Краткая шпаргалка

| Действие | Windows | ALT Linux |
|---|---|---|
| Установка инструментов | `py -m pip install pyinstaller -r requirements.txt` | `apt-get install python3-module-pyinstaller python3-module-flask python3-module-openpyxl python3-module-docx python3-tkinter` |
| Сборка всего | `py -m PyInstaller --noconfirm build.spec` | `pyinstaller3 --noconfirm build.spec` |
| Сервер | `dist\orders_server.exe` | `./dist/orders_server` |
| Клиент (GUI) | `dist\orders_client.exe` | `./dist/orders_client` |
| Клиент (текст) | `dist\orders_client.exe --server http://ИП:8080 --cli` | `./dist/orders_client --server http://ИП:8080 --cli` |
| Порт | `--port 8090` / `set PORT=8090` | `--port 8090` / `PORT=8090` |
