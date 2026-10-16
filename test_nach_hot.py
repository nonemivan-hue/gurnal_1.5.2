# -*- coding: utf-8 -*-
"""Сквозной тест: роль «Начальник отдела» + мягкий перезапуск без выключения сервера."""
import io, os, re, sqlite3, subprocess, time, zipfile
import requests

BASE = 'http://localhost:8080'
DB = os.path.join(os.path.dirname(__file__), 'data', 'orders.db')
ok = fail = 0


def check(name, cond):
    global ok, fail
    if cond:
        ok += 1; print(f'PASS  {name}')
    else:
        fail += 1; print(f'FAIL  {name}')


def login(s, u, p):
    r = s.post(f'{BASE}/login', data={'login': u, 'password': p}, allow_redirects=True)
    return r.status_code == 200 and 'Выйти' in r.text


# ---- подготовка: создать сотрудника-начальника отдела через админа
sa = requests.Session()
check('вход администратора', login(sa, 'admin', 'admin'))
conn = sqlite3.connect(DB)
role_id = conn.execute("SELECT id FROM roles WHERE name='Начальник отдела'").fetchone()
conn.close()
check('роль «Начальник отдела» существует в БД (миграция seed)', role_id is not None)

r = sa.post(f'{BASE}/employees/save', data={
    'full_name': 'Тестов Начальник Тестович', 'position': 'Начальник отдела',
    'role': 'Начальник отдела', 'login': 'nach_test', 'password': 'test1234'})
check('админ создал сотрудника с ролью «Начальник отдела»', r.status_code in (200, 302))

# ---- куратор для сравнения (создан ранее? создадим если нет)
if not login(requests.Session(), 'kur_test', 'test1234'):
    sa.post(f'{BASE}/employees/save', data={
        'full_name': 'Тестов Куратор Тестович', 'position': 'Куратор',
        'role': 'Куратор', 'login': 'kur_test', 'password': 'test1234'})

sn = requests.Session()
check('вход начальника отдела', login(sn, 'nach_test', 'test1234'))

# ---- доступ ко всем журналам
for j, t in [('opeka', 'Опека'), ('nuzhdaemost', 'Нуждаемость'), ('profilaktika', 'Профилактика')]:
    r = sn.get(f'{BASE}/journal/{j}')
    check(f'начальник отдела открывает журнал «{t}»', r.status_code == 200 and t in r.text)

# ---- создание приказа во всех журналах
conn = sqlite3.connect(DB); conn.row_factory = sqlite3.Row
types = {row['kind']: row['id'] for row in conn.execute('SELECT kind, MIN(id) id FROM order_types GROUP BY kind')}
conn.close()
for j in ('opeka', 'nuzhdaemost', 'profilaktika'):
    r = sn.post(f'{BASE}/journal/{j}/create', data={
        'order_date': '2026-10-15', 'order_type_id': types[j],
        'person_fio': 'Гражданин Т. Т. Т.', 'specialist_fio': 'Тестов Н. Т.',
        'note': 'тест'} | ({'order_number': ''}), allow_redirects=True)
    check(f'начальник отдела регистрирует приказ в «{j}»', r.status_code == 200 and 'Гражданин Т. Т. Т.' in r.text)

# ---- все отчёты
for path, key in [('/reports/by_period', 'Количество приказов'),
                  ('/reports/by_period_detail', 'расшифровкой'),
                  ('/reports/journal_print?journal=opeka', 'Журнал')]:
    r = sn.get(f'{BASE}{path}')
    check(f'отчёт {path.split("?")[0]} доступен начальнику отдела', r.status_code == 200 and key in r.text)
    r = sn.get(f'{BASE}{path}&fmt=xlsx' if '?' in path else f'{BASE}{path}?fmt=xlsx')
    check(f'выгрузка Excel {path.split("?")[0]}', r.status_code == 200 and len(r.content) > 1000)

# ---- справочники
r = sn.get(f'{BASE}/employees'); check('начальник отдела видит справочник «Сотрудники»', r.status_code == 200)
r = sn.get(f'{BASE}/order_types'); check('видит «Виды приказов»', r.status_code == 200)
r = sn.post(f'{BASE}/order_types/save', data={'kind': 'opeka', 'name': 'Тест вид приказа НО'}, allow_redirects=True)
check('редактирует справочник «Виды приказов»', r.status_code == 200 and 'Тест вид приказа НО' in r.text)
conn = sqlite3.connect(DB); tid = conn.execute("SELECT id FROM order_types WHERE name='Тест вид приказа НО'").fetchone()
conn.close()
if tid:
    r = sn.post(f'{BASE}/order_types/delete/{tid[0]}', allow_redirects=True)
    check('удаляет вид приказа', r.status_code == 200 and 'Тест вид приказа НО' not in r.text)
r = sn.post(f'{BASE}/employees/save', data={
    'full_name': 'Тест Подчинённый П. П.', 'position': 'Специалист',
    'role': 'Опека', 'login': 'podt_test', 'password': 'test1234'}, allow_redirects=True)
check('редактирует справочник «Сотрудники»', r.status_code == 200 and 'Тест Подчинённый' in r.text)
r = sn.get(f'{BASE}/employees/template'); check('шаблон Excel сотрудников доступен', r.status_code == 200)

# ---- сервис — недоступен начальнику отдела
for path in ['/service', '/logs']:
    r = sn.get(f'{BASE}{path}')
    check(f'{path} запрещён начальнику отдела (403)', r.status_code == 403)

# ---- куратор не получил доступ к сотрудникам
sk = requests.Session()
if login(sk, 'kur_test', 'test1234'):
    r = sk.post(f'{BASE}/employees/save', data={'full_name': 'X', 'role': 'Опека', 'login': 'x_kur', 'password': '1'})
    check('куратору сотрудники по-прежнему запрещены (403)', r.status_code == 403)

# ---- мягкий перезапуск через ./run.sh restart (без выключения сервера)
before_pid = open('data/app.pid').read().strip() if os.path.exists('data/app.pid') else ''
subprocess.run(['./run.sh', 'restart'], capture_output=True, text=True)
up = False
new_pid = before_pid
for _ in range(30):
    time.sleep(0.7)
    try:
        if requests.get(f'{BASE}/login', timeout=2).status_code == 200:
            new_pid = open('data/app.pid').read().strip() if os.path.exists('data/app.pid') else 'n/a'
            up = True
            break
    except Exception:
        pass
check('сервер поднялся после restart без ручного запуска', up)
check('процесс реально перезапущен (новый PID)', up and new_pid != before_pid)
check('флаг restart_pending удалён', not os.path.exists('data/restart_pending'))

# ---- данные на месте после перезапуска
sn2 = requests.Session()
check('повторный вход после перезапуска', login(sn2, 'nach_test', 'test1234'))
r = sn2.get(f'{BASE}/journal/opeka')
check('приказы сохранились после перезапуска', 'Гражданин Т. Т. Т.' in r.text)

# ---- проверка записей аудита
log = open('data/logs/app.log', encoding='utf-8').read()
check('в логе есть запись о мягком перезапуске', 'МЯГКИЙ ПЕРЕЗАПУСК' in log or 'Мягкий перезапуск' in (open('server.out', encoding='utf-8', errors='ignore').read()))

# ---- уборка тестовых данных
conn = sqlite3.connect(DB)
conn.execute("DELETE FROM employees WHERE login IN ('nach_test','kur_test','podt_test','x_kur')")
conn.execute("DELETE FROM orders WHERE person_fio='Гражданин Т. Т. Т.'")
conn.commit(); conn.close()
print(f'\nИТОГО: PASS={ok} FAIL={fail}')
raise SystemExit(1 if fail else 0)
