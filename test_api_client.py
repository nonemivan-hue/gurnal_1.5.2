# -*- coding: utf-8 -*-
"""Сквозной тест REST API v1 + логики клиентского приложения (без GUI)."""
import sys, json, io
sys.path.insert(0, '.')
from client_app import ServerClient, ApiError

BASE = 'http://localhost:8080'
ok = fail = 0
def check(name, cond, extra=''):
    global ok, fail
    if cond: ok += 1; print(f'  PASS  {name}')
    else: fail += 1; print(f'  FAIL  {name} {extra}')

c = ServerClient(BASE, 'admin', 'admin')
print('== Сессия и права ==')
s = c.get('/api/v1/session')
check('session роль', s['role'] == 'Администратор', s)
check('session 3 журнала', len(s['journals']) == 3, s['journals'])
check('session reports', s['reports'] is True)
emps = c.get('/api/v1/employees')
check('employees список', isinstance(emps, list) and len(emps) > 0)
check('employees login виден админу', 'login' in emps[0])

print('== CRUD приказа в журнале opeka ==')
types = c.get('/api/v1/order_types?journal=opeka')
check('order_types opeka', isinstance(types, list))
made = c.post('/api/v1/journal/opeka/orders', {
    'person_fio': 'Тестовый Т.Т.', 'specialist_fio': 'Специалист С.С.',
    'note': 'тест API', 'order_type_id': types[0]['id'] if types else None})
check('создание приказа', made.get('ok') and made.get('order_number'), made)
oid = made['id']
lst = c.get('/api/v1/journal/opeka/orders?q=Тестовый')
check('поиск созданного', any(r['id'] == oid for r in lst), len(lst))
one = c.get(f'/api/v1/orders/{oid}')
check('чтение приказа', one['person_fio'] == 'Тестовый Т.Т.')
r = c.put(f'/api/v1/orders/{oid}', {'note': 'обновлено тест'})
check('изменение приказа', r.get('ok'))
check('изменение применилось', c.get(f'/api/v1/orders/{oid}')['note'] == 'обновлено тест')

print('== Скан PDF ==')
pdf = b'%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF'
r = c.post(f'/api/v1/orders/{oid}/scan', files=[('file', 'test.pdf', pdf)])
check('загрузка скана', r.get('ok'), r)
data, hdr = c.get_raw(f'/api/v1/orders/{oid}/scan')
check('скачивание скана', data == pdf)
check('флаг has_scan', c.get(f'/api/v1/orders/{oid}')['has_scan'] is True)
r = c.delete(f'/api/v1/orders/{oid}/scan')
check('удаление скана', r.get('ok'))
try:
    c.get(f'/api/v1/orders/{oid}/scan'); check('скан отсутствует -> 404', False)
except ApiError as e:
    check('скан отсутствует -> 404', e.status == 404)
try:
    c.post(f'/api/v1/orders/{oid}/scan', files=[('file', 'x.txt', b'not pdf')])
    check('не-PDF отклонён', False)
except ApiError as e:
    check('не-PDF отклонён', e.status == 400)

print('== Отчёты ==')
rp = c.get('/api/v1/reports/by_period?d1=2020-01-01&d2=2030-12-31')
check('отчёт по периоду', rp['total'] >= 1, rp)
rd = c.get('/api/v1/reports/by_period_detail?d1=2020-01-01&d2=2030-12-31')
check('отчёт с расшифровкой', 'cols' in rd and 'table' in rd)
rj = c.get('/api/v1/reports/journal_print?d1=2020-01-01&d2=2030-12-31&journal=opeka')
check('распечатка журнала', rj['total'] >= 1)

print('== Справочник видов приказов ==')
newt = c.post('/api/v1/order_types', {'kind': 'profilaktika', 'name': 'Тест вид API'})
check('добавление вида', newt.get('ok'), newt)
allt = c.get('/api/v1/order_types')
tid = [t['id'] for t in allt if t['name'] == 'Тест вид API'][0]
r = c.post('/api/v1/order_types', {'id': tid, 'kind': 'profilaktika', 'name': 'Тест вид API 2'})
check('изменение вида', r.get('ok'))
dup = None
try:
    c.post('/api/v1/order_types', {'kind': 'profilaktika', 'name': 'Тест вид API 2'})
except ApiError as e: dup = e
check('дубликат отклонён', dup is not None and dup.status == 400)
r = c.delete(f'/api/v1/order_types/{tid}')
check('удаление неиспользуемого вида', r.get('ok'))

print('== Права: куратор не видит справочник сотрудников с логами ==')
# найдём куратора в БД
import sqlite3
conn = sqlite3.connect('data/orders.db'); conn.row_factory = sqlite3.Row
cur_row = conn.execute("SELECT e.login, e.password_hash FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name='Куратор' LIMIT 1").fetchone()
adm_row = conn.execute("SELECT e.login FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name='Начальник отдела' LIMIT 1").fetchone()
conn.close()
if cur_row:
    from werkzeug.security import check_password_hash
    cc = None
    for pwd in ('curator','1234','test','password','12345','kurator'):
        try:
            cc = ServerClient(BASE, cur_row['login'], pwd)
            cc.get('/api/v1/session'); break
        except ApiError: cc = None
    if cc:
        e2 = cc.get('/api/v1/employees')
        check('куратор: login скрыт', 'login' not in e2[0], e2[0])
        try:
            cc.get('/api/v1/journal/profilaktika/orders'); check('куратор профиль недоступен->403?', False)
        except ApiError as ex:
            check('куратор: чужой журнал 403', ex.status == 403)
    else:
        print('  SKIP  пароль куратора неизвестен')
else:
    print('  SKIP  куратор не заведён')

print('== Ошибки входа ==')
try:
    ServerClient(BASE, 'admin', 'bad').get('/api/v1/session')
    check('неверный пароль 401', False)
except ApiError as e:
    check('неверный пароль 401', e.status == 401)

print('== Удаление приказа ==')
r = c.delete(f'/api/v1/orders/{oid}')
check('удаление приказа', r.get('ok'))
try:
    c.get(f'/api/v1/orders/{oid}'); check('после удаления 404', False)
except ApiError as e:
    check('после удаления 404', e.status == 404)

print(f'\nИТОГ: PASS={ok} FAIL={fail}')
sys.exit(1 if fail else 0)
