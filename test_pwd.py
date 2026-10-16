# -*- coding: utf-8 -*-
"""Сквозной тест самостоятельной смены пароля сотрудником."""
import requests, sqlite3

BASE = 'http://localhost:8080'
DB = 'data/orders.db'

def q(sql, args=()):
    c = sqlite3.connect(DB); c.row_factory = sqlite3.Row
    r = c.execute(sql, args).fetchall(); c.close(); return r

ok = []
def check(name, cond):
    ok.append((name, bool(cond)))
    print(('PASS' if cond else 'FAIL'), '-', name)

# 1. Админ создаёт сотрудника с логином/паролем
s = requests.Session()
r = s.post(f'{BASE}/login', data={'login': 'admin', 'password': 'admin'}, allow_redirects=True)
check('вход админа', '/login' not in r.url or 'Моя' not in r.text)
q_data = {'full_name': 'Тестовый Тест Тестович', 'position': 'Юрист',
          'login': 'testemp', 'password': 'temp123', 'role': 'Опека'}
# id существующего?
existing = q("SELECT id FROM employees WHERE login='testemp'")
if existing:
    q_data['id'] = existing[0]['id']
r = s.post(f'{BASE}/employees/save', data=q_data, allow_redirects=True)
check('админ создал сотрудника testemp', 'Сотрудник сохранён' in r.text)
emp = q("SELECT * FROM employees WHERE login='testemp'")[0]
check('pwd_changed_at пустой после установки админом', emp['pwd_changed_at'] is None)

# 2. Сотрудник входит -> видит рекомендацию сменить пароль
s2 = requests.Session()
r = s2.post(f'{BASE}/login', data={'login': 'testemp', 'password': 'temp123'})
check('вход сотрудника', 'testemp' in r.url or r.status_code == 200)
check('показана рекомендация сменить пароль', 'Рекомендуем сменить пароль' in r.text)

# 3. Страница профиля доступна
r = s2.get(f'{BASE}/profile')
check('страница профиля открылась', 'Смена пароля' in r.text and 'Текущий пароль' in r.text)
check('предупреждение о несменённом пароле', 'ещё не меняли пароль' in r.text)

# 4. Неверный текущий пароль -> отказ
r = s2.post(f'{BASE}/profile', data={'old_password': 'WRONG', 'new_password': 'new456', 'confirm_password': 'new456'})
check('отказ при неверном текущем пароле', 'Текущий пароль указан неверно' in r.text)

# 5. Несовпадение подтверждения -> отказ
r = s2.post(f'{BASE}/profile', data={'old_password': 'temp123', 'new_password': 'new456', 'confirm_password': 'other456'})
check('отказ при несовпадении подтверждения', 'не совпадают' in r.text)

# 6. Совпадение со старым -> отказ
r = s2.post(f'{BASE}/profile', data={'old_password': 'temp123', 'new_password': 'temp123', 'confirm_password': 'temp123'})
check('отказ при совпадении с текущим', 'не должен совпадать' in r.text)

# 7. Успешная смена
r = s2.post(f'{BASE}/profile', data={'old_password': 'temp123', 'new_password': 'new456', 'confirm_password': 'new456'})
check('успешная смена пароля', 'Пароль успешно изменён' in r.text)
emp = q("SELECT * FROM employees WHERE login='testemp'")[0]
check('pwd_changed_at заполнен', emp['pwd_changed_at'] is not None)
print('   дата смены:', emp['pwd_changed_at'])

# 8. Старый пароль не работает, новый работает
r = requests.post(f'{BASE}/login', data={'login': 'testemp', 'password': 'temp123'})
check('старый пароль отклонён', 'Неверный логин или пароль' in r.text)
s3 = requests.Session()
r = s3.post(f'{BASE}/login', data={'login': 'testemp', 'password': 'new456'})
check('новый пароль принят', r.status_code == 200)
r = s3.get(f'{BASE}/profile')
check('рекомендация больше не показывается', 'ещё не меняли пароль' not in r.text)
check('показана дата последней смены', 'Пароль последний раз изменялся' in r.text)

# 9. Аудит в логе
log = open('data/logs/app.log', encoding='utf-8').read()
check('в логе запись ПАРОЛЬ ИЗМЕНЁН', 'ПАРОЛЬ ИЗМЕНЁН' in log)
check('в логе запись об ошибке текущего пароля', 'ПАРОЛЬ НЕ ИЗМЕНЁН' in log)

# 10. Админ сбрасывает пароль -> снова рекомендуется смена
s.post(f'{BASE}/employees/save', data={'id': emp['id'], 'full_name': 'Тестовый Тест Тестович',
     'position': 'Юрист', 'login': 'testemp', 'password': 'reset77', 'role': 'Опека'})
emp = q("SELECT * FROM employees WHERE login='testemp'")[0]
check('после сброса админом pwd_changed_at сброшен', emp['pwd_changed_at'] is None)
s4 = requests.Session()
r = s4.post(f'{BASE}/login', data={'login': 'testemp', 'password': 'reset77'})
check('сбросенный пароль работает + рекомендация', 'Рекомендуем сменить пароль' in r.text)

fails = [n for n, v in ok if not v]
print('\nИТОГ:', len(ok) - len(fails), 'OK,', len(fails), 'FAILED')
exit(1 if fails else 0)
