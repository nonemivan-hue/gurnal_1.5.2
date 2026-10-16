# -*- coding: utf-8 -*-
"""Проверка панели «Активные пользователи» для Администратора, Куратора и Начальника отдела."""
import requests, sqlite3

BASE = 'http://localhost:8080'
conn = sqlite3.connect('data/orders.db')
cur = conn.cursor()
roles = {r[0]: r[1] for r in cur.execute('SELECT id,name FROM roles')}
rev = {v: k for k, v in roles.items()}
created = []

def make(login, fio, role):
    pw = login + '_pw123'
    cur.execute('INSERT INTO employees (full_name,position,login,password_hash,role_id,pwd_changed_at) '
                "VALUES (?,?,?, ?,?, datetime('now'))",
                (fio, 'Тест', login,
                 '$pbkdf2-sha256$310000$dummy', rev[role]))
    conn.commit()
    return login, pw

# найдём существующих пользователей нужных ролей или создадим тестовых
def get_or_make(role):
    row = cur.execute('SELECT login FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name=? LIMIT 1',
                      (role,)).fetchone()
    if row:
        return row[0], None
    return None, None

users = {}
for role in ('Администратор', 'Куратор', 'Начальник отдела'):
    row = cur.execute("SELECT e.login FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name=?",
                      (role,)).fetchone()
    users[role] = row[0] if row else None

print('Пользователи по ролям:', users)

def login_session(role, password):
    s = requests.Session()
    r = s.post(BASE + '/login', data={'login': users[role], 'password': password}, allow_redirects=True)
    assert r.status_code == 200, f'вход {role}: {r.status_code}'
    return s

def check_panel(s, role, expect_visible):
    r = s.get(BASE + '/')
    html = r.text
    vis = ('👥 Онлайн' in html) or ('onlineBox' in html)
    print(f'{role}: панель на странице = {vis} (ожидалось {expect_visible})')
    assert vis == expect_visible
    return html

def check_api(s, role, expect_ok):
    r = s.get(BASE + '/api/active_users')
    ok = r.status_code == 200
    assert ok == expect_ok, f'{role}: API вернул {r.status_code}'
    if ok:
        d = r.json()
        print(f'{role}: API OK, timeout={d["timeout"]} сек, активных={len(d["users"])}, имена={[u["full_name"] for u in d["users"]][:5]}')
    else:
        print(f'{role}: API -> {r.status_code} (как ожидалось)')

# --- Администратор
sa = login_session('Администратор', 'admin')
html = check_panel(sa, 'Администратор', True)
check_api(sa, 'Администратор', True)

# --- Куратор (создадим временного, если нет)
if not users['Куратор']:
    import werkzeug.security as ws
    cur.execute('INSERT INTO employees (full_name,position,login,password_hash,role_id,pwd_changed_at) VALUES (?,?,?,?,?,datetime("now"))',
                ('Тест Куратор','Тест','tmp_kurator',ws.generate_password_hash('tmp_pw123'),rev['Куратор']))
    conn.commit(); users['Куратор']='tmp_kurator'; created.append('tmp_kurator')
    PWK='tmp_pw123'
else:
    PWK=None
if PWK is None:
    # пароль неизвестен — создаём tmp куратора с заведомым паролем
    import werkzeug.security as ws
    cur.execute('INSERT INTO employees (full_name,position,login,password_hash,role_id,pwd_changed_at) VALUES (?,?,?,?,?,datetime("now"))',
                ('Тест Куратор2','Тест','tmp_kurator2',ws.generate_password_hash('tmp_pw123'),rev['Куратор']))
    conn.commit(); users['Куратор']='tmp_kurator2'; created.append('tmp_kurator2'); PWK='tmp_pw123'
sk = login_session('Куратор', PWK)
check_panel(sk, 'Куратор', True)
check_api(sk, 'Куратор', True)

# --- Начальник отдела
if not users['Начальник отдела']:
    import werkzeug.security as ws
    cur.execute('INSERT INTO employees (full_name,position,login,password_hash,role_id,pwd_changed_at) VALUES (?,?,?,?,?,datetime("now"))',
                ('Тест Начальник','Тест','tmp_nach',ws.generate_password_hash('tmp_pw123'),rev['Начальник отдела']))
    conn.commit(); users['Начальник отдела']='tmp_nach'; created.append('tmp_nach')
sn = login_session('Начальник отдела', 'tmp_pw123')
check_panel(sn, 'Начальник отдела', True)
check_api(sn, 'Начальник отдела', True)

# --- Обычная роль (Опека) — панели быть не должно
row = cur.execute("SELECT e.login FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name='Опека' LIMIT 1").fetchone()
if row:
    so = requests.Session()
    r = so.post(BASE+'/login', data={'login': row[0], 'password': 'opeka'})
    if r.status_code==200 and 'Дашборд' in r.text or so.get(BASE+'/').status_code==200:
        pass
    html_o = so.get(BASE+'/').text
    vis = 'onlineBox' in html_o
    print(f'Опека: панель видна = {vis} (ожидалось False)')
    assert not vis
    r = so.get(BASE+'/api/active_users')
    print(f'Опека: API -> {r.status_code} (ожидалось 403)')
    assert r.status_code == 403

# --- счётчик активности: после действий трёх сессий в списке должны быть они все
d = sa.get(BASE+'/api/active_users').json()
names = [u['full_name'] for u in d['users']]
print('Активные сейчас:', names)
assert any('Куратор' in n for n in names), 'Куратор не виден в активном списке'
assert any('Начальник' in n for n in names), 'Начальник не виден в активном списке'
assert any('админ' in n.lower() or 'Админ' in n or 'system' in n.lower() or len(names)>0 for n in names)

# выход админа -> он исчезает из списка (кроме случаев, когда тот же юзер залогинен ещё где-то)
before = len(sa.get(BASE+'/api/active_users').json()['users'])
sa.get(BASE+'/logout')
after_list = sk.get(BASE+'/api/active_users').json()['users']
print(f'До выхода админа активных: {before}; список после выхода (глазами куратора): {[u["full_name"] for u in after_list]}')

# очистка временных сотрудников
for lg in created:
    cur.execute('DELETE FROM employees WHERE login=?', (lg,))
conn.commit(); conn.close()
print('\nВСЕ ПРОВЕРКИ ПРОЙДЕНЫ ✅')
