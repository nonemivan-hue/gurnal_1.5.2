#!/bin/sh
# Запуск «Журнала приказов».
# Порт по умолчанию — 8080 (порт 5000 на сервере занят другой программой).
# Способ задания порта (приоритет слева направо):
#   ./run.sh --port 8090   |   PORT=8090 ./run.sh   |   по умолчанию 8080
# Если указанный порт занят — автоматически подбирается ближайший свободный.
#
# Режимы:
#   ./run.sh                 — запуск с автоперезапуском (watchdog-цикл):
#                              при мягком перезапуске из веб-интерфейса
#                              («Обновление программы») процесс стартует заново
#                              БЕЗ выключения сервера администратором;
#   ./run.sh restart         — немедленный мягкий перезапуск работающего
#                              сервера (создаёт флаг data/restart_pending);
#   ./run.sh stop            — остановить сервер;
#   ./run.sh run [--port N]  — запуск без цикла автоперезапуска (однократно).
cd "$(dirname "$0")"

PIDFILE="data/app.pid"
FLAG="data/restart_pending"

case "$1" in
  restart)
    mkdir -p data
    date '+%Y-%m-%d %H:%M:%S' > "$FLAG"
    if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
      kill "$(cat "$PIDFILE")"
      echo "Мягкий перезапуск запрошен (флаг $FLAG, процесс остановлен)."
    else
      echo "Флаг перезапуска создан. Если сервер запущен через ./run.sh, он перезапустится сам."
    fi
    exit 0
    ;;
  stop)
    if [ -f "$PIDFILE" ] && kill -0 "$(cat "$PIDFILE")" 2>/dev/null; then
      rm -f "$FLAG"
      kill "$(cat "$PIDFILE")"
      echo "Сервер остановлен."
    else
      echo "Рабочий процесс не найден (PID-файл отсутствует или процесс завершён)."
    fi
    exit 0
    ;;
  run)
    shift
    exec python3 app.py "$@"
    ;;
esac

if [ "$1" = "--port" ] || [ -n "$PORT" ]; then
  : # параметры порта пробрасываются ниже как есть
fi

mkdir -p data
# watchdog-цикл: пока существует флаг data/restart_pending (создаётся при
# обновлении программы из интерфейса), сервер перезапускается автоматически.
while true; do
  python3 app.py "$@" &
  APP_PID=$!
  echo "$APP_PID" > "$PIDFILE"
  wait "$APP_PID"
  code=$?
  if [ -f "$FLAG" ]; then
    rm -f "$FLAG"
    echo "[run.sh] Мягкий перезапуск: старт новой версии кода..."
    sleep 1
    continue
  fi
  echo "[run.sh] Приложение завершилось (код $code). Перезапуск прекращён."
  rm -f "$PIDFILE"
  break
done
