# -*- coding: utf-8 -*-
"""Сквозная проверка панели «👥 Онлайн» для ролей Администратор, Куратор, Начальник отдела."""
import requests, re, sys

BASE = 'http://localhost:8080'
ok = fail = 0

def check(name, cond):
    global ok, fail
    if cond:
        ok += 1; print(f'PASS  {name}')
    else:
        fail += 1; print(f'FAIL  {name}')

def login(login_name, password):
    s = requests.Session()
    s.post(BASE + '/login', data={'login': login_name, 'password': password})
    assert s.get(BASE + '/api/active_users').status_code in (200, 403), \
        f'сессия {login_name} не установлена'
    return s

# --- создаём тестовых сотрудников под админом
adm = login('admin', 'admin')
check('вход администратора', adm.get(BASE + '/', timeout=5).status_code == 200)

emps = [('ТестОнлайн Куратор', 'cur_test', 'Куратор'),
        ('ТестОнлайн Начотдел', 'head_test', 'Начальник отдела')]
def delete_test_emps():
    """Удалить всех тестовых сотрудников «ТестОнлайн*» (по data-fio в кнопках ✏️)."""
    page = adm.get(BASE + '/employees').text
    for m in re.finditer(r'data-id="(\d+)" data-fio="ТестОнлайн[^"]*"', page):
        adm.post(BASE + f"/employees/delete/{m.group(1)}", allow_redirects=True)

delete_test_emps()
for full, lg, role in emps:
    r = adm.post(BASE + '/employees/save', data={
        'full_name': full, 'position': 'Тест', 'login': lg, 'password': 'test1234',
        'role': role}, allow_redirects=True)
    check(f'создан сотрудник {lg} ({role})', r.status_code == 200 and full in r.text)

# роль Опека — отдельный пользователь для сравнения прав
r = adm.post(BASE + '/employees/save', data={
    'full_name': 'ТестОнлайн Опекун', 'position': 'Тест', 'login': 'opeka_test',
    'password': 'test1234', 'role': 'Опека'}, allow_redirects=True)
check('создан сотрудник opeka_test (Опека)', r.status_code == 200 and 'ТестОнлайн Опекун' in r.text)
cur = login('cur_test', 'test1234')
head = login('head_test', 'test1234')
opeka = login('opeka_test', 'test1234')

# --- панель у админа / куратора / начальника
for who, sess in [('Администратор', adm), ('Куратор', cur), ('Начальник отдела', head)]:
    page = sess.get(BASE + '/').text
    check(f'{who}: кнопка «👥 Онлайн» в меню', '👥 Онлайн:' in page and 'onlineBox' in page)
    api = sess.get(BASE + '/api/active_users')
    check(f'{who}: /api/active_users -> 200', api.status_code == 200)
    users = api.json().get('users', [])
    names = [u['full_name'] for u in users]
    # ожидаем всех, кроме уже вышедших (Опекун выйдет позже в этом же тесте)
    expected = ['ТестОнлайн Куратор', 'ТестОнлайн Начотдел', 'ТестОнлайн Опекун']
    check(f'{who}: все вошедшие тестовые видны в списке',
          all(n in names for n in expected))
    check(f'{who}: админ виден в списке', any('администратор' in n.lower() for n in names))

# --- у роли Опека панели нет
page = opeka.get(BASE + '/').text
check('Опека: панель «Онлайн» НЕ показывается', 'onlineBox' not in page)
check('Опека: /api/active_users -> 403', opeka.get(BASE + '/api/active_users').status_code == 403)

# --- счётчик уменьшается после выхода
before = len(adm.get(BASE + '/api/active_users').json()['users'])
opeka.get(BASE + '/logout')
after = len(adm.get(BASE + '/api/active_users').json()['users'])
check(f'после выхода пользователя счётчик уменьшился ({before}->{after})', after < before)

# --- активность обновляется: запрос куратора отражается в seconds_ago
cur.get(BASE + '/')
users = adm.get(BASE + '/api/active_users').json()['users']
cu = [u for u in users if u['full_name'] == 'ТестОнлайн Куратор'][0]
check('у куратора свежая отметка активности (<60 с)', cu['seconds_ago'] < 60)

# --- запись в БД active_sessions существует
import sqlite3
conn = sqlite3.connect('data/orders.db')
rows = conn.execute('SELECT user_id FROM active_sessions').fetchall()
conn.close()
check('таблица active_sessions заполняется', len(rows) >= 1)

# --- убираем тестовых сотрудников из справочника
for m in re.finditer(r'data-id="(\d+)" data-fio="ТестОнлайн[^"]*"',
                     adm.get(BASE + '/employees').text):
    adm.post(BASE + f"/employees/delete/{m.group(1)}", allow_redirects=True)
left = len(re.findall(r'ТестОнлайн', adm.get(BASE + '/employees').text))
check('тестовые сотрудники удалены из справочника', left == 0)

print(f'\nИТОГ: PASS={ok} FAIL={fail}')
sys.exit(1 if fail else 0)
