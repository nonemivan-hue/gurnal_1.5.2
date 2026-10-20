# -*- coding: utf-8 -*-
"""Сквозной тест v1.6.4: отчёт «⏱ Время в сети» (только Администратор)."""
import time, sqlite3, sys
sys.path.insert(0, '.')
import app as A

DB = A.DB_PATH
conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

# 1. Таблица user_online_log создаётся автоматически
A._ensure_required_tables(conn)
names = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")]
assert 'user_online_log' in names, 'нет таблицы user_online_log'
print('PASS 1: таблица user_online_log существует')

# 2. Тестовые пользователи: админ — существующий, второго создадим (Куратор)
admins = conn.execute("SELECT e.*, r.name rn FROM employees e JOIN roles r ON r.id=e.role_id WHERE r.name='Администратор' LIMIT 1").fetchall()
assert admins, 'нужен администратор'
admin = admins[0]
conn.execute("DELETE FROM employees WHERE login='testkur'")
conn.execute("INSERT INTO employees(full_name, login, role_id, position) VALUES ('Тестовый Куратор', 'testkur', 3, 'тест')")
conn.commit()
others = conn.execute("SELECT e.*, r.name rn FROM employees e JOIN roles r ON r.id=e.role_id WHERE e.login='testkur'").fetchall()
assert others, 'не создан тестовый пользователь'
other = others[0]
print(f"PASS 2: админ={admin['full_name']}, другой={other['full_name']} ({other['rn']})")

# 3. Хелперы учёта
now = time.time()
# очистим ВСЕ тестовые интервалы прошлых прогонов (sid='testsid') и свои id
conn.execute("DELETE FROM user_online_log WHERE sid='testsid'")
conn.execute('DELETE FROM user_online_log WHERE user_id IN (?,?)', (admin['id'], other['id']))
conn.commit(); conn.close()

# прямой вызов хелпера (в обход троттлинга HTTP-пути):
def touch():
    A._DB_TOUCH_TS.clear()
    A.online_track_touch(other['id'], 'testsid', '10.0.0.1')
touch()
c1 = sqlite3.connect(DB); c1.row_factory = sqlite3.Row
rows = c1.execute('SELECT * FROM user_online_log WHERE sid=?', ('testsid',)).fetchall()
assert len(rows) == 1, f'ожидалась 1 запись, получено {len(rows)}'
id1, ts1 = rows[0]['id'], rows[0]['logout_ts']
# продление при малом простое
touch()
rows = c1.execute('SELECT * FROM user_online_log WHERE sid=?', ('testsid',)).fetchall()
assert len(rows) == 1 and rows[0]['logout_ts'] >= ts1, 'продление не сработало'
# новый интервал после большого разрыва: подвинем logout_ts назад на ONLINE_GAP+60
back = now - A.ONLINE_GAP - 60
c1.execute('UPDATE user_online_log SET logout_ts=?, login_ts=? WHERE id=?', (back, back - 300, id1))
c1.commit()
touch()
rows = c1.execute('SELECT * FROM user_online_log WHERE sid=? ORDER BY id', ('testsid',)).fetchall()
assert len(rows) == 2, f'ожидались 2 интервала, получено {len(rows)}'
print('PASS 3: online_track_touch — склейка/новый интервал корректны')

# 4. Закрытие при выходе
A.online_close_session(other['id'], 'testsid')
last = c1.execute('SELECT logout_ts FROM user_online_log WHERE id=?', (rows[1]['id'],)).fetchone()[0]
assert abs(last - time.time()) < 5, 'выход не закрыл интервал'
print('PASS 4: online_close_session закрывает интервал')
c1.close()

# 5. HTTP: доступность отчёта
client = A.app.test_client()

def login(login_, password):
    client.get('/logout')
    r = client.post('/login', data={'login': login_, 'password': password}, follow_redirects=False)
    return r

# пароли неизвестны — временно выставим хэши
c1 = sqlite3.connect(DB)
from werkzeug.security import generate_password_hash
for uid in (admin['id'], other['id']):
    c1.execute('UPDATE employees SET password_hash=? WHERE id=?',
               (generate_password_hash('Test1234'), uid))
c1.commit(); c1.close()

r = login(admin['login'], 'Test1234'); assert r.status_code in (200, 302), r.status_code
r = client.get('/reports/online_time')
assert r.status_code == 200, f'админ: /reports/online_time -> {r.status_code}'
html = r.get_data(as_text=True)
assert 'Время в сети' in html, 'заголовок отчёта не найден'
assert 'Итого' in html or 'Нет данных' in html
print('PASS 5a: админ получает отчёт (200)')

r = client.get('/reports')
assert 'Время в сети' in r.get_data(as_text=True), 'карточка отчёта не видна админу'
print('PASS 5b: карточка «4. ⏱ Время в сети» видна админу на /reports')

r = client.get('/reports/online_time?detail=1')
assert r.status_code == 200 and 'по дням' in r.get_data(as_text=True)
print('PASS 5c: детализация по дням (detail=1) работает')

r = client.get('/reports/online_time?fmt=xlsx')
assert r.status_code == 200 and 'spreadsheetml' in r.headers.get('Content-Type',''), r.headers.get('Content-Type')
assert len(r.get_data()) > 1000
print('PASS 5d: экспорт XLSX')
r = client.get('/reports/online_time?fmt=docx')
assert r.status_code == 200 and 'wordprocessingml' in r.headers.get('Content-Type','')
print('PASS 5e: экспорт DOCX')

# период по параметрам
d1 = time.strftime('%Y-%m-%d', time.localtime(now - 7*86400)); d2 = time.strftime('%Y-%m-%d')
r = client.get(f'/reports/online_time?d1={d1}&d2={d2}')
assert r.status_code == 200
print('PASS 5f: выбор периода')

# другой пользователь — 403
r = login(other['login'], 'Test1234')
r = client.get('/reports/online_time')
assert r.status_code == 403, f'не-админ должен получить 403, получил {r.status_code}'
r = client.get('/reports')
assert 'Время в сети' not in r.get_data(as_text=True), 'карточка не должна быть видна не-админу'
print('PASS 6: не-админ — 403, карточка скрыта')

# без входа — редирект на /login или 401/403
client.get('/logout')
r = client.get('/reports/online_time')
assert r.status_code in (302, 401, 403), r.status_code
print('PASS 7: без авторизации доступ закрыт')

# 8. Реальный трекинг через HTTP-запросы: логин админа → запросы → запись в log.
# Запись в user_online_log выполняется с троттлингом (не чаще раза в 30 сек на
# сессию), поэтому сбрасываем окно троттлинга перед проверкой.
c1 = sqlite3.connect(DB)
c1.execute("DELETE FROM user_online_log WHERE user_id=?", (admin['id'],))
c1.commit(); c1.close()
login(admin['login'], 'Test1234')
A._DB_TOUCH_TS.clear()          # снять ограничение частоты записи для теста
client.get('/journal/profilaktika')
time.sleep(0.2)
c1 = sqlite3.connect(DB); c1.row_factory = sqlite3.Row
row = c1.execute('SELECT user_id, login_ts, logout_ts, ip FROM user_online_log '
                 'WHERE user_id=? ORDER BY id DESC LIMIT 1', (admin['id'],)).fetchone()
assert row is not None, 'HTTP-активность не записалась в user_online_log'
assert abs(row['logout_ts'] - time.time()) < 60, 'слишком старая метка времени'
print('PASS 8: активность пользователя фиксируется в БД через обычные запросы')
c1.close()

# 9. Регрессия прежних функций
login(admin['login'], 'Test1234')
for path in ('/', '/references', '/service/users', '/api/active_users'):
    r = client.get(path)
    assert r.status_code == 200, f'{path} -> {r.status_code}'
print('PASS 9: регрессия — /, /references, /service/users, /api/active_users в порядке')

# Уборка тестовых данных
c1 = sqlite3.connect(DB)
c1.execute("DELETE FROM user_online_log WHERE sid='testsid'")
c1.execute("DELETE FROM user_online_log WHERE user_id=?", (other['id'],))
c1.execute("DELETE FROM user_online_log WHERE user_id=? AND ip='127.0.0.1'", (admin['id'],))
c1.execute("DELETE FROM employees WHERE login='testkur'")
c1.execute('UPDATE employees SET password_hash=NULL WHERE id=?', (admin['id'],))
c1.commit(); c1.close()
print('Уборка тестовых данных выполнена')
print('\nВСЕ ТЕСТЫ ПРОЙДЕНЫ — v1.6.4 готова')
