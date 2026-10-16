# Удаление автозапуска сервера «Журнал приказов» с Windows
# Запуск от имени администратора:
#   powershell -ExecutionPolicy Bypass -File uninstall_autostart.ps1
$ErrorActionPreference = "SilentlyContinue"
$nssm = Get-Command nssm
if ($nssm) { & $nssm stop JournalPrikazov; & $nssm remove JournalPrikazov confirm }
sc.exe stop JournalPrikazov 2>$null
sc.exe delete JournalPrikazov 2>$null
Unregister-ScheduledTask -TaskName "JournalPrikazov" -Confirm:$false
Write-Host "Автозапуск удален (служба и/или задача Планировщика)."
