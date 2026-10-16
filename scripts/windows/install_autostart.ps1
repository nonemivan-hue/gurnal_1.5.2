# ============================================================================
# Установка автоматического запуска сервера «Журнал приказов» на Windows
# ============================================================================
# Требуется: PowerShell от имени администратора.
# Скрипт регистрирует службу Windows (NSSM) или, как альтернативу, задачу
# Планировщика заданий, которая запускает сервер при входе пользователя.
#
# Запуск:  powershell -ExecutionPolicy Bypass -File install_autostart.ps1
# Параметры (необязательные):
#   -AppDir   "C:\journal"          каталог с app.py (по умолчанию — родительний каталог скрипта)
#   -PythonExe "python"             путь к python.exe (по умолчанию из PATH)
#   -Port     8080                  порт сервера (по умолчанию 8080)
#   -Method   service|task          способ автозапуска (по умолчанию service, если есть NSSM, иначе task)
# ============================================================================

param(
    [string]$AppDir = "",
    [string]$PythonExe = "python",
    [int]$Port = 8080,
    [ValidateSet("service", "task")]
    [string]$Method = ""
)

$ErrorActionPreference = "Stop"

# Каталог приложения: если не задан — ищем рядом со скриптом (scripts/windows -> корень)
if (-not $AppDir) {
    $candidate = Resolve-Path (Join-Path $PSScriptRoot "..\..")
    if (Test-Path (Join-Path $Candidate "app.py")) { $AppDir = $Candidate.Path }
    else { $AppDir = (Get-Location).Path }
}
if (-not (Test-Path (Join-Path $AppDir "app.py"))) {
    Write-Host "ОШИБКА: в каталоге $AppDir не найден app.py. Укажите -AppDir." -ForegroundColor Red
    exit 1
}

$ServiceName = "JournalPrikazov"
Write-Host "Каталог приложения : $AppDir"
Write-Host "Порт сервера       : $Port"

# ---------- Способ 1: служба Windows через NSSM ----------
$nssm = Get-Command nssm -ErrorAction SilentlyContinue
if (-not $nssm) {
    foreach ($p in @("$env:ProgramFiles\nssm\nssm.exe", "$AppDir\nssm.exe")) {
        if (Test-Path $p) { $nssm = $p; break }
    }
}
if ($nssm -and ($Method -eq "" -or $Method -eq "service")) {
    Write-Host "Способ: служба Windows (NSSM)" -ForegroundColor Green
    & $nssm install $ServiceName $PythonExe "-m flask --version >nul" | Out-Null  # проверка запускаемости
    & $nssm remove $ServiceName confirm | Out-Null
    & $nssm install $ServiceName $PythonExe "`"$AppDir\app.py`" --port $Port"
    & $nssm set $ServiceName AppDirectory $AppDir
    & $nssm set $ServiceName DisplayName "Сервер журнала приказов"
    & $nssm set $ServiceName Description "Веб-приложение регистрации приказов по видам деятельности"
    & $nssm set $ServiceName Start SERVICE_AUTO_START
    & $nssm set $ServiceName AppStdout "$AppDir\data\logs\service.out.log"
    & $nssm set $ServiceName AppStderr "$AppDir\data\logs\service.err.log"
    & $nssm set $ServiceName AppRotateFiles 1
    & $nssm set $ServiceName AppRestartDelay 5000   # перезапуск через 5 сек после сбоя
    New-Item -ItemType Directory -Force -Path "$AppDir\data\logs" | Out-Null
    & $nssm start $ServiceName
    Write-Host "Служба '$ServiceName' установлена и запущена (автозапуск при загрузке ОС)." -ForegroundColor Green
    exit 0
}

# ---------- Способ 2: задача Планировщика заданий ----------
Write-Host "Способ: задача Планировщика заданий (запуск при входе в систему)" -ForegroundColor Yellow

$batContent = @"
@echo off
rem Автостарт сервера «Журнал приказов». Не закрывать окно при работе программы.
cd /d "$AppDir"
set PORT=$Port
"$PythonExe" app.py --port $Port
"@
$batPath = Join-Path $AppDir "start_server.bat"
Set-Content -Path $batPath -Value $batContent -Encoding OEMCP866

$action  = New-ScheduledTaskAction -Execute $batPath -WorkingDirectory $AppDir
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) `
            -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
Register-ScheduledTask -TaskName "JournalPrikazov" -Action $action -Trigger $trigger `
    -Settings $settings -Description "Автозапуск веб-сервера журнала приказов" -Force | Out-Null

Start-ScheduledTask -TaskName "JournalPrikazov"
Write-Host "Задача 'JournalPrikazov' зарегистрирована и запущена." -ForegroundColor Green
Write-Host "Проверка: http://localhost:$Port  (лог: data\logs\app.log)"
