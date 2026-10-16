# -*- coding: utf-8 -*-
"""
Клиентское приложение «Журнал регистрации приказов»
====================================================
Десктопный клиент (Tkinter, стандартная библиотека Python) к серверу
веб-приложения. Подключается по REST API v1 (http://адрес_сервера:порт/api/v1/...).

Работает на Windows и ALT Linux без установки сторонних библиотек.

Запуск:
    python3 client_app.py                     # окно с настройкой подключения
    python3 client_app.py --server http://192.168.0.10:8080

Возможности клиента (в пределах прав роли пользователя):
  - просмотр журналов Опека / Нуждаемость / Профилактика, поиск;
  - создание, редактирование и удаление приказов (+ примечание);
  - прикрепление, скачивание и удаление PDF-сканов;
  - справочник «Виды приказов» (просмотр; редактирование — куратору и выше);
  - справочник «Сотрудники» (просмотр; логин виден админу/начальнику отдела);
  - отчёты за период (все три), выгрузка таблицы в CSV/Excel;
  - самостоятельная смена пароля.

Пароль хранится только в памяти процесса (в реестр/файлы не пишется).

Если графическая библиотека Tkinter недоступна, запуск из командной строки
позволяет работать с сервером в текстовом режиме:
    python3 client_app.py --server http://хост:порт --cli --login admin
"""
import sys
import json
import argparse
import urllib.request
import urllib.error
import urllib.parse
import base64
import uuid
import csv
import os
from datetime import date

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog
    HAS_TK = True
except Exception:            # Tkinter отсутствует — доступен режим --cli
    tk = None
    HAS_TK = False

CLIENT_VERSION = '1.4.0'


# ================================================================ HTTP-клиент
class ApiError(Exception):
    def __init__(self, msg, status=None):
        super().__init__(msg)
        self.status = status


class ServerClient:
    """Минимальный HTTP-клиент к REST API сервера (stdlib urllib)."""

    def __init__(self, base_url, login, password):
        self.base = base_url.rstrip('/')
        self.login = login
        token = f'{login}:{password}'.encode('utf-8')
        self.auth = 'Basic ' + base64.b64encode(token).decode('ascii')

    def _request(self, path, method='GET', data=None, files=None, raw=False):
        url = self.base + path
        headers = {'Authorization': self.auth}
        body = None
        if files is not None:                      # multipart/form-data
            boundary = '----pyclient' + uuid.uuid4().hex
            parts = []
            for name, value in (data or {}).items():
                parts.append(f'--{boundary}\r\nContent-Disposition: form-data; '
                             f'name="{name}"\r\n\r\n{value}\r\n'.encode('utf-8'))
            for name, fname, content in files:
                parts.append(('--%s\r\nContent-Disposition: form-data; name="%s"; '
                              'filename="%s"\r\nContent-Type: application/octet-stream'
                              '\r\n\r\n' % (boundary, name, fname)).encode('utf-8'))
                parts.append(content)
                parts.append(b'\r\n')
            parts.append(('--%s--\r\n' % boundary).encode('utf-8'))
            body = b''.join(parts)
            headers['Content-Type'] = 'multipart/form-data; boundary=' + boundary
        elif data is not None:                     # JSON
            body = json.dumps(data, ensure_ascii=False).encode('utf-8')
            headers['Content-Type'] = 'application/json; charset=utf-8'
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                payload = resp.read()
                if raw:
                    return payload, resp.headers
                ct = resp.headers.get('Content-Type', '')
                if 'json' in ct:
                    return json.loads(payload.decode('utf-8'))
                return payload
        except urllib.error.HTTPError as e:
            payload = e.read()
            if e.code == 401:
                raise ApiError('Неверный логин или пароль', 401)
            if e.code == 403:
                raise ApiError('Нет прав для этого действия', 403)
            try:
                msg = json.loads(payload.decode('utf-8')).get('error')
            except Exception:
                msg = None
            raise ApiError(msg or f'Ошибка сервера (HTTP {e.code})', e.code)
        except urllib.error.URLError as e:
            raise ApiError(f'Сервер недоступен: {e.reason}', None)

    # --- convenience wrappers -------------------------------------------
    def get(self, path):
        return self._request(path)

    def post(self, path, data=None, files=None):
        return self._request(path, 'POST', data, files)

    def put(self, path, data):
        return self._request(path, 'PUT', data)

    def delete(self, path):
        return self._request(path, 'DELETE')

    def get_raw(self, path):
        return self._request(path, raw=True)


def check_server(base_url):
    """Пинг публичного /api/v1/info — для проверки адреса при входе."""
    c = ServerClient(base_url, '__probe__', '__probe__')
    try:
        info = c.get('/api/v1/info')
        return info
    except ApiError as e:
        if e.status == 401:                 # сервер живой, но требует авторизацию
            return {}
        raise


# ================================================== графический интерфейс (Tkinter)
def _build_gui_classes():
    """Возвращает классы GUI, определённые только при наличии Tkinter."""
    if not HAS_TK:
        raise RuntimeError('Tkinter недоступен — используйте режим --cli')

    # ------------------------------------------------------------- окно входа
    class LoginWindow(tk.Tk):
        """Диалог подключения: адрес сервера, логин, пароль + проверка сервера."""

        def __init__(self, server=''):
            super().__init__()
            self.title('Вход — Журнал приказов (клиент)')
            self.result = None                      # (client, session) после успеха
            self.resizable(False, False)
            frm = ttk.Frame(self, padding=14)
            frm.grid(sticky='nsew')
            ttk.Label(frm, text='Журнал регистрации приказов',
                      font=('Segoe UI', 13, 'bold')).grid(
                column=0, columnspan=3, row=0, pady=(0, 10))
            ttk.Label(frm, text='Сервер:').grid(row=1, column=0, sticky='w')
            self.e_srv = ttk.Entry(frm, width=38)
            self.e_srv.grid(row=1, column=1, columnspan=2, sticky='we', pady=2)
            self.e_srv.insert(0, server or 'http://localhost:8080')
            ttk.Button(frm, text='Проверить', command=self.ping).grid(
                row=1, column=3, padx=(6, 0))
            ttk.Label(frm, text='Логин:').grid(row=2, column=0, sticky='w')
            self.e_login = ttk.Entry(frm, width=38)
            self.e_login.grid(row=2, column=1, columnspan=2, sticky='we', pady=2)
            ttk.Label(frm, text='Пароль:').grid(row=3, column=0, sticky='w')
            self.e_pwd = ttk.Entry(frm, width=38, show='*')
            self.e_pwd.grid(row=3, column=1, columnspan=2, sticky='we', pady=2)
            self.lbl_info = ttk.Label(frm, text='', foreground='#555', wraplength=380)
            self.lbl_info.grid(row=4, column=0, columnspan=4, sticky='w', pady=(4, 0))
            btns = ttk.Frame(frm)
            btns.grid(row=5, column=0, columnspan=4, pady=(12, 0))
            ttk.Button(btns, text='Войти', command=self.do_login).pack(side='left', padx=4)
            ttk.Button(btns, text='Отмена', command=self.destroy).pack(side='left', padx=4)
            self.e_pwd.bind('<Return>', lambda e: self.do_login())
            self.e_login.bind('<Return>', lambda e: self.do_login())
            self.bind('<Escape>', lambda e: self.destroy())

        def ping(self):
            try:
                info = check_server(self.e_srv.get().strip())
                ver = info.get('version', '?')
                self.lbl_info.config(text=f'Сервер доступен (v{ver})', foreground='#0a7d28')
            except ApiError as e:
                self.lbl_info.config(text=str(e), foreground='#b00020')

        def do_login(self):
            base = self.e_srv.get().strip()
            login = self.e_login.get().strip()
            pwd = self.e_pwd.get()
            if not base or not login:
                messagebox.showwarning('Вход', 'Укажите адрес сервера и логин')
                return
            try:
                client = ServerClient(base, login, pwd)
                sess = client.get('/api/v1/session')
                self.result = (client, sess)
                self.destroy()
            except ApiError as e:
                messagebox.showerror('Ошибка входа', str(e))

    # ------------------------------------------------------- диалог приказа
    class OrderDialog(tk.Toplevel):
        """Форма создания/редактирования приказа (данные — из REST API)."""

        FIELDS = (('order_number', '№ приказа (пусто = авто)'),
                  ('order_date', 'Дата приказа (ГГГГ-ММ-ДД)'),
                  ('person_fio', 'Ф.И.О. в отношении кого'),
                  ('specialist_fio', 'ФИО специалиста'),
                  ('note', 'Примечание'))

        def __init__(self, master, client, journal_key, order=None):
            super().__init__(master)
            self.client = client
            self.journal_key = journal_key
            self.saved = None                       # dict приказа после сохранения
            self.title('Приказ — изменение' if order else 'Приказ — создание')
            self.resizable(False, False)
            frm = ttk.Frame(self, padding=12)
            frm.grid(sticky='nsew')
            self.vars = {}
            row = 0
            for f, label in self.FIELDS:
                ttk.Label(frm, text=label + ':').grid(row=row, column=0, sticky='w')
                v = tk.StringVar(value=(order or {}).get(f, '') or
                                 (date.today().isoformat() if f == 'order_date' and not order else ''))
                ttk.Entry(frm, textvariable=v, width=46).grid(
                    row=row, column=1, sticky='we', pady=2)
                self.vars[f] = v
                row += 1
            # выбор вида приказа из справочника (по журналу)
            ttk.Label(frm, text='Наименование приказа:').grid(row=row, column=0, sticky='w')
            self.cbo_type = ttk.Combobox(frm, width=44, state='readonly')
            self.cbo_type.grid(row=row, column=1, sticky='we', pady=2)
            row += 1
            self.types = []
            try:
                self.types = client.get(f'/api/v1/order_types?journal={journal_key}')
                names = ['— без вида —'] + [t['name'] for t in self.types]
                self.cbo_type['values'] = names
                sel = 0
                if order and order.get('order_type_id'):
                    for i, t in enumerate(self.types, 1):
                        if t['id'] == order['order_type_id']:
                            sel = i
                self.cbo_type.current(sel)
            except ApiError as e:
                messagebox.showerror('Справочник', str(e))
            # выбор причины из справочника «Причины» (поле журнала «Опека»)
            self.cbo_reason = None
            self.reasons = []
            if journal_key == 'opeka':
                ttk.Label(frm, text='Причина:').grid(row=row, column=0, sticky='w')
                self.cbo_reason = ttk.Combobox(frm, width=44, state='readonly')
                self.cbo_reason.grid(row=row, column=1, sticky='we', pady=2)
                row += 1
                try:
                    self.reasons = client.get('/api/v1/order_types?reasons=1')
                    rnames = ['— не выбрана —'] + [r['name'] for r in self.reasons]
                    self.cbo_reason['values'] = rnames
                    rsel = 0
                    if order and order.get('reason_id'):
                        for i, r in enumerate(self.reasons, 1):
                            if r['id'] == order['reason_id']:
                                rsel = i
                    self.cbo_reason.current(rsel)
                except ApiError as e:
                    messagebox.showerror('Справочник «Причины»', str(e))
            btns = ttk.Frame(frm)
            btns.grid(row=row, column=0, columnspan=2, pady=(12, 0))
            ttk.Button(btns, text='Сохранить', command=self.save).pack(side='left', padx=4)
            ttk.Button(btns, text='Отмена', command=self.destroy).pack(side='left', padx=4)

        def save(self):
            data = {f: v.get().strip() for f, v in self.vars.items()}
            idx = self.cbo_type.current()
            data['order_type_id'] = self.types[idx - 1]['id'] if idx > 0 else None
            if getattr(self, 'cbo_reason', None) is not None:
                ridx = self.cbo_reason.current()
                data['reason_id'] = self.reasons[ridx - 1]['id'] if ridx > 0 else None
            try:
                if self.master_order_id:
                    self.saved = self.client.put(f"/api/v1/orders/{self.master_order_id}", data)
                else:
                    self.saved = self.client.post(
                        f'/api/v1/journal/{self.journal_key}/orders', data)
                self.destroy()
            except ApiError as e:
                messagebox.showerror('Сохранение', str(e))

        @property
        def master_order_id(self):
            return getattr(self, '_oid', None)

    # ------------------------------------------------------------ главное окно
    class AppWindow(tk.Tk):
        """Основное окно: журнал слева, карточка приказа справа, панель действий."""

        def __init__(self, client, sess):
            super().__init__()
            self.client = client
            self.sess = sess
            # 'journals' — список {key, title, can_edit} от сервера (права по роли)
            self.journals = {j['title']: j for j in sess['journals']}
            self.rows = []                          # текущий список приказов
            self.can_edit = True                    # роль с доступом к журналу
            self.title(f"Журнал приказов — {sess['full_name']} ({sess['role']})")
            self.geometry('980x600')
            self._build_ui()
            if self.journals:
                self.cbo_journal.current(0)
                self.reload()

        # ---------- построение интерфейса ----------
        def _build_ui(self):
            top = ttk.Frame(self, padding=6)
            top.pack(fill='x')
            ttk.Label(top, text='Журнал:').pack(side='left')
            self.cbo_journal = ttk.Combobox(top, width=22, state='readonly',
                                            values=list(self.journals))
            self.cbo_journal.pack(side='left', padx=4)
            self.cbo_journal.bind('<<ComboboxSelected>>', lambda e: self.reload())
            self.e_search = ttk.Entry(top, width=30)
            self.e_search.pack(side='left', padx=4)
            self.e_search.bind('<Return>', lambda e: self.reload())
            ttk.Button(top, text='Обновить', command=self.reload).pack(side='left', padx=2)
            ttk.Button(top, text='Поиск', command=self.reload).pack(side='left', padx=2)

            body = ttk.Frame(self)
            body.pack(fill='both', expand=True, padx=6, pady=2)
            cols = ('order_number', 'order_date', 'type', 'person', 'spec', 'scan', 'note')
            heads = {'order_number': '№', 'order_date': 'Дата', 'type': 'Вид приказа',
                     'person': 'Ф.И.О.', 'spec': 'Специалист', 'scan': 'Скан', 'note': 'Примечание'}
            widths = {'order_number': 70, 'order_date': 90, 'type': 220, 'person': 240,
                      'spec': 200, 'scan': 50, 'note': 200}
            # журнал «Опека»: столбец «Причина» (справочник «Причины»)
            if self.current_journal() == 'opeka':
                cols = cols[:6] + ('reason',) + cols[6:]
                heads['reason'] = 'Причина'
                widths['reason'] = 180
            self.tree = ttk.Treeview(body, columns=cols, show='tree headings', height=16)
            for c in cols:
                self.tree.heading(c, text=heads[c])
                self.tree.column(c, width=widths[c], anchor='w')
            vsb = ttk.Scrollbar(body, orient='vertical', command=self.tree.yview)
            self.tree.configure(yscrollcommand=vsb.set)
            self.tree.pack(side='left', fill='both', expand=True)
            vsb.pack(side='left', fill='y')
            self.tree.bind('<Double-1>', lambda e: self.edit_selected())

            bar = ttk.Frame(self, padding=6)
            bar.pack(fill='x')
            if self.can_edit:
                ttk.Button(bar, text='➕ Создать', command=self.create).pack(side='left', padx=2)
                ttk.Button(bar, text='✏️ Изменить', command=self.edit_selected).pack(side='left', padx=2)
                ttk.Button(bar, text='🗑 Удалить', command=self.delete_selected).pack(side='left', padx=2)
            ttk.Button(bar, text='📎 Скан…', command=self.scan_menu).pack(side='left', padx=2)
            ttk.Button(bar, text='📊 Отчёты', command=self.reports).pack(side='left', padx=2)
            ttk.Button(bar, text='🔑 Сменить пароль', command=self.change_password).pack(side='left', padx=2)
            ttk.Button(bar, text='Выход', command=self.destroy).pack(side='right', padx=2)
            self.lbl_status = ttk.Label(self, text='', padding=(8, 4), foreground='#555')
            self.lbl_status.pack(fill='x')

        # ---------- данные ----------
        def current_journal(self):
            """Ключ (key) выбранного журнала из ответа сервера."""
            j = self.journals.get(self.cbo_journal.get())
            return j['key'] if j else None

        def journal_can_edit(self):
            """Право редактирования текущего журнала — по данным сервера."""
            j = self.journals.get(self.cbo_journal.get())
            return bool(j and j.get('can_edit'))

        def reload(self):
            jk = self.current_journal()
            if not jk:
                return
            q = self.e_search.get().strip()
            url = f'/api/v1/journal/{jk}/orders'
            if q:
                url += '?q=' + urllib.parse.quote(q)
            try:
                self.rows = self.client.get(url)
            except ApiError as e:
                messagebox.showerror('Журнал', str(e))
                return
            self.tree.delete(*self.tree.get_children())
            has_reason = 'reason' in self.tree['columns']
            for r in self.rows:
                vals = [r['order_number'], r['order_date'], r['order_type_name'] or '',
                        r['person_fio'], r['specialist_fio'],
                        'есть' if r['has_scan'] else '']
                if has_reason:
                    vals.append(r.get('reason_name') or '')
                vals.append(r.get('note') or '')
                self.tree.insert('', 'end', iid=str(r['id']), values=vals)
            self.lbl_status.config(text=f'{self.cbo_journal.get()}: {len(self.rows)} записей')

        def selected_id(self):
            sel = self.tree.selection()
            if not sel:
                messagebox.showinfo('Выбор', 'Выделите приказ в списке')
                return None
            return int(sel[0])

        def _find_row(self, oid):
            return next((r for r in self.rows if r['id'] == oid), None)

        # ---------- действия ----------
        def create(self):
            dlg = OrderDialog(self, self.client, self.current_journal())
            self.wait_window(dlg)
            if dlg.saved:
                self.reload()

        def edit_selected(self):
            oid = self.selected_id()
            if not oid:
                return
            try:
                order = self.client.get(f'/api/v1/orders/{oid}')
            except ApiError as e:
                messagebox.showerror('Приказ', str(e))
                return
            dlg = OrderDialog(self, self.client, order['journal'], order)
            dlg._oid = oid
            self.wait_window(dlg)
            if dlg.saved:
                self.reload()

        def delete_selected(self):
            oid = self.selected_id()
            if not oid:
                return
            r = self._find_row(oid)
            num = r['order_number'] if r else str(oid)
            if not messagebox.askyesno('Удаление', f'Удалить приказ № {num}?'):
                return
            try:
                self.client.delete(f'/api/v1/orders/{oid}')
                self.reload()
            except ApiError as e:
                messagebox.showerror('Удаление', str(e))

        def scan_menu(self):
            oid = self.selected_id()
            if not oid:
                return
            choice = messagebox.askyesno(
                'Скан PDF', 'Да — скачать существующий скан\nНет — выбрать и загрузить PDF')
            if choice:
                out = filedialog.asksaveasfilename(defaultextension='.pdf',
                                                   filetypes=[('PDF', '*.pdf')])
                if not out:
                    return
                try:
                    payload, _ = self.client.get_raw(f'/api/v1/orders/{oid}/scan')
                    with open(out, 'wb') as fh:
                        fh.write(payload)
                    messagebox.showinfo('Скан', f'Сохранён: {out}')
                except ApiError as e:
                    messagebox.showerror('Скан', str(e))
            else:
                path = filedialog.askopenfilename(filetypes=[('PDF', '*.pdf')])
                if not path:
                    return
                try:
                    with open(path, 'rb') as fh:
                        content = fh.read()
                    self.client.post(f'/api/v1/orders/{oid}/scan', files=[
                        ('file', os.path.basename(path), content)])
                    self.reload()
                    messagebox.showinfo('Скан', 'Загружен (только PDF)')
                except ApiError as e:
                    messagebox.showerror('Скан', str(e))

        def change_password(self):
            old = simpledialog.askstring('Пароль', 'Текущий пароль:', show='*', parent=self)
            new = simpledialog.askstring('Пароль', 'Новый пароль:', show='*', parent=self)
            if not old or not new:
                return
            try:
                self.client.post('/api/v1/password', {
                    'old_password': old, 'new_password': new, 'confirm_password': new})
                messagebox.showinfo('Пароль', 'Пароль изменён')
            except ApiError as e:
                messagebox.showerror('Пароль', str(e))

        # ---------- отчёты ----------
        def reports(self):
            win = ReportsWindow(self, self.client, self.sess)
            win.grab_set()

    # ---------------------------------------------------------- окно отчётов
    def _normalize_report(t, data):
        """Нормализация ответа сервера к виду (заголовок, шапка, строки).

        Сервер возвращает разные форматы для трёх отчётов — приводим их
        к единому табличному виду для вывода и экспорта в CSV.
        Возвращает: (title, headers, rows, total)
        """
        title = data.get('title', t)
        if t == 'by_period':                       # rows: [{'kind','count'}]
            headers = ['Вид приказа', 'Кол-во приказов']
            rows = [[r['kind'], str(r['count'])] for r in data.get('rows', [])]
            return title, headers, rows, data.get('total')
        if t == 'by_period_matrix':                # cols + table + totals
            headers = ['Вид приказа'] + list(data.get('cols', []))
            rows = [[str(c) for c in row] for row in data.get('table', [])]
            if data.get('totals'):
                rows.append(['ИТОГО'] + [str(x) for x in data['totals']])
            return title, headers, rows, None
        # journal_print: rows dict
        headers = ['№ п/п', '№ приказа', 'Дата', 'Ф.И.О.',
                   'Наименование приказа', 'ФИО специалиста']
        rows = [[str(r['n']), r['number'], r['date'], r['fio'],
                 r['name'], r['specialist']] for r in data.get('rows', [])]
        return title, headers, rows, data.get('total')

    class ReportsWindow(tk.Toplevel):
        """Выбор периода и типа отчёта; вывод таблицей; сохранение в CSV."""

        TYPES = (('by_period', 'Приказы за период'),
                 ('by_period_matrix', 'Приказы за период с расшифровкой'),
                 ('journal_print', 'Распечатка журнала за период'))

        def __init__(self, master, client, sess):
            super().__init__(master)
            self.client = client
            self.title('Отчёты')
            self.geometry('860x520')
            top = ttk.Frame(self, padding=8)
            top.pack(fill='x')
            ttk.Label(top, text='С:').pack(side='left')
            self.d1 = ttk.Entry(top, width=12)
            self.d1.pack(side='left', padx=2)
            self.d1.insert(0, date.today().replace(day=1).isoformat())
            ttk.Label(top, text='По:').pack(side='left')
            self.d2 = ttk.Entry(top, width=12)
            self.d2.pack(side='left', padx=2)
            self.d2.insert(0, date.today().isoformat())
            self.var_type = tk.StringVar(value=self.TYPES[0][0])
            for key, title in self.TYPES:
                ttk.Radiobutton(top, text=title, variable=self.var_type,
                                value=key).pack(side='left', padx=6)
            self.cbo_j = ttk.Combobox(top, width=16, state='readonly',
                                      values=list(sess['journals'] and
                                                 [j['title'] for j in sess['journals']] or []))
            self.cbo_j.pack(side='left', padx=4)
            if sess['journals']:
                self.cbo_j.current(0)
            ttk.Button(top, text='Построить', command=self.build).pack(side='left', padx=6)
            ttk.Button(top, text='Сохранить CSV', command=self.export_csv).pack(side='left')
            self.txt = tk.Text(self, font=('Consolas', 10))
            self.txt.pack(fill='both', expand=True, padx=8, pady=6)
            self.headers, self.rows_data = [], []

        def build(self):
            """Запрос отчёта к REST API сервера и вывод таблицей."""
            qs = f'?d1={self.d1.get().strip()}&d2={self.d2.get().strip()}'
            t = self.var_type.get()
            try:
                if t == 'journal_print':
                    jk = {j['title']: j['key'] for j in self.master.sess['journals']}
                    data = self.client.get(
                        '/api/v1/reports/journal_print' + qs +
                        '&journal=' + urllib.parse.quote(jk[self.cbo_j.get()]))
                elif t == 'by_period_matrix':
                    data = self.client.get('/api/v1/reports/by_period_detail' + qs)
                else:
                    data = self.client.get('/api/v1/reports/by_period' + qs)
            except ApiError as e:
                messagebox.showerror('Отчёт', str(e))
                return
            title, self.headers, self.rows_data, total = _normalize_report(t, data)
            self.txt.delete('1.0', 'end')
            self.txt.insert('end', f"Отчёт «{title}»\n")
            self.txt.insert('end', f"Период: {data.get('d1')} — {data.get('d2')}\n\n")
            self.txt.insert('end', _table(self.rows_data, self.headers) + '\n')
            if total is not None:
                self.txt.insert('end', f"\nИтого приказов: {total}\n")

        def export_csv(self):
            if not self.rows_data:
                messagebox.showinfo('CSV', 'Сначала постройте отчёт')
                return
            path = filedialog.asksaveasfilename(defaultextension='.csv',
                                                filetypes=[('CSV', '*.csv')])
            if not path:
                return
            with open(path, 'w', newline='', encoding='utf-8-sig') as fh:
                w = csv.writer(fh, delimiter=';')
                if self.headers:
                    w.writerow(self.headers)
                for r in self.rows_data:
                    w.writerow(r if isinstance(r, list) else [r])
            messagebox.showinfo('CSV', f'Сохранено: {path}')

    return LoginWindow, OrderDialog, AppWindow


# ==================================================================== CLI-режим
def _pad(s, w):
    s = str(s)
    return s if len(s) >= w else s + ' ' * (w - len(s))


def _table(rows, headers):
    widths = [max(len(str(h)), *(len(str(r[i])) for r in rows)) if rows else len(str(h))
              for i, h in enumerate(headers)]
    line = '+-' + '-+-'.join('-' * w for w in widths) + '-+'
    out = [line, '| ' + ' | '.join(_pad(h, w) for h, w in zip(headers, widths)) + ' |',
           line]
    for r in rows:
        out.append('| ' + ' | '.join(_pad(c, w) for c, w in zip(r, widths)) + ' |')
    out.append(line)
    return '\n'.join(out)


def cli_main(base, login, password):
    """Интерактивный текстовый клиент (для систем без Tkinter / по SSH)."""
    c = ServerClient(base, login, password)
    sess = c.get('/api/v1/session')
    journals = {j['title']: j['key'] for j in sess['journals']}
    print(f"Вход выполнен: {sess['full_name']} ({sess['role']}), "
          f"сервер v{sess['version']}")
    help_txt = ('l <№ журнала>            — список приказов (с поиском)\n'
                'a <№ журнала>            — создать приказ (по шагам)\n'
                'e <id приказа>           — изменить приказ\n'
                'd <id приказа>           — удалить приказ\n'
                'u <id приказа> <файл.pdf> — прикрепить скан\n'
                'p <id приказа> [файл]    — скачать скан\n'
                't                        — справочник видов приказов\n'
                'r                        — отчёт «Приказы за период»\n'
                'w                        — сотрудники\n'
                'x                        — сменить пароль\n'
                'h                        — эта справка, q — выход')
    print(help_txt)
    while True:
        try:
            cmd = input('\nклиент> ').strip()
        except EOFError:
            break
        if not cmd:
            continue
        parts = cmd.split()
        op, arg = parts[0].lower(), parts[1] if len(parts) > 1 else ''
        try:
            if op in ('q', 'exit'):
                break
            elif op == 'h':
                print(help_txt)
            elif op in ('l', 'list'):
                # журнал можно задать названием («Опека») или порядковым номером («1»)
                titles = list(journals)
                if arg in journals:
                    title = arg
                else:
                    try:
                        title = titles[int(arg) - 1]
                    except (ValueError, IndexError):
                        print('Укажите номер или название журнала:',
                              ', '.join(f'{i+1}:{t}' for i, t in enumerate(titles)))
                        continue
                q = input('Строка поиска (Enter — все): ').strip()
                url = f'/api/v1/journal/{journals[title]}/orders'
                if q:
                    url += '?q=' + urllib.parse.quote(q)
                rows = c.get(url)
                print(_table([[r['id'], r['order_number'], r['order_date'],
                               r['order_type_name'] or '', r['person_fio'],
                               'да' if r['has_scan'] else ''] for r in rows],
                             ['ID', '№', 'Дата', 'Вид приказа', 'Ф.И.О.', 'Скан']))
            elif op == 'a':
                titles = list(journals)
                try:
                    title = titles[int(arg) - 1]
                except (ValueError, IndexError):
                    print('Номера журналов:',
                          ', '.join(f'{i+1}={t}' for i, t in enumerate(titles)))
                    continue
                data = {}
                data['order_number'] = input('№ приказа (пусто = авто): ').strip()
                data['order_date'] = input(
                    f'Дата ГГГГ-ММ-ДД [{date.today()}]: ').strip() or date.today().isoformat()
                data['person_fio'] = input('Ф.И.О. в отношении кого: ').strip()
                data['specialist_fio'] = input('ФИО специалиста: ').strip()
                types = c.get(f"/api/v1/order_types?journal={journals[title]}")
                for i, t in enumerate(types, 1):
                    print(f'  {i}. {t["name"]}')
                tn = input('Номер вида приказа (0 — без): ').strip()
                data['order_type_id'] = types[int(tn) - 1]['id'] \
                    if tn.isdigit() and 0 < int(tn) <= len(types) else None
                data['note'] = input('Примечание: ').strip()
                res = c.post(f"/api/v1/journal/{journals[title]}/orders", data)
                print(f"Создан приказ № {res['order_number']} (id={res['id']})")
            elif op == 'e':
                order = c.get(f'/api/v1/orders/{arg}')
                print(f"Редактирование № {order['order_number']}; Enter = оставить как есть")
                upd = {}
                for f, label in (('order_number', '№'), ('order_date', 'Дата'),
                                 ('person_fio', 'Ф.И.О.'),
                                 ('specialist_fio', 'Специалист'), ('note', 'Примечание')):
                    v = input(f'{label} [{order[f]}]: ').strip()
                    if v:
                        upd[f] = v
                if upd:
                    c.put(f'/api/v1/orders/{arg}', upd)
                    print('Сохранено')
            elif op == 'd':
                if input(f'Удалить приказ id={arg}? (д/н) ').strip().lower() == 'д':
                    c.delete(f'/api/v1/orders/{arg}')
                    print('Удалён')
            elif op == 'u':
                with open(arg and parts[2], 'rb') as fh:
                    content = fh.read()
                c.post(f'/api/v1/orders/{arg}/scan', files=[
                    ('file', os.path.basename(parts[2]), content)])
                print('Скан прикреплён')
            elif op == 'p':
                payload, _ = c.get_raw(f'/api/v1/orders/{arg}/scan')
                out = parts[2] if len(parts) > 2 else f'scan_{arg}.pdf'
                with open(out, 'wb') as fh:
                    fh.write(payload)
                print(f'Скан сохранён: {out} ({len(payload)} байт)')
            elif op == 't':
                rows = c.get('/api/v1/order_types')
                print(_table([[r['id'], r['kind_title'], r['name']] for r in rows],
                             ['ID', 'Журнал', 'Наименование']))
            elif op == 'r':
                d1 = input('С [первый день месяца]: ').strip()
                d2 = input('По [сегодня]: ').strip()
                qs = f'?d1={d1 or date.today().replace(day=1)}&d2={d2 or date.today()}'
                data = c.get('/api/v1/reports/by_period' + qs)
                print(_table([[r['kind'], r['count']] for r in data['rows']]
                             + [['ИТОГО', data['total']]],
                             ['Вид приказа', 'Кол-во']))
            elif op == 'm':
                d1 = input('С [первый день месяца]: ').strip()
                d2 = input('По [сегодня]: ').strip()
                qs = f'?d1={d1 or date.today().replace(day=1)}&d2={d2 or date.today()}'
                data = c.get('/api/v1/reports/by_period_detail' + qs)
                heads = ['Вид приказа'] + data['cols']
                rows = [[str(x) for x in row] for row in data['table']]
                if data.get('totals'):
                    rows.append(['ИТОГО'] + [str(x) for x in data['totals']])
                print(_table(rows, heads))
            elif op == 'w':
                rows = c.get('/api/v1/employees')
                print(_table([[r['full_name'], r['position'], r['role']]
                              for r in rows], ['ФИО', 'Должность', 'Роль']))
            elif op == 'x':
                import getpass
                old = getpass.getpass('Текущий пароль: ')
                new = getpass.getpass('Новый пароль: ')
                c.post('/api/v1/password', {'old_password': old, 'new_password': new,
                                            'confirm_password': new})
                print('Пароль изменён')
            else:
                print('Неизвестная команда, h — справка')
        except ApiError as e:
            print('Ошибка:', e)
        except (ValueError, IndexError):
            print('Некорректные параметры команды (h — справка)')
        except FileNotFoundError as e:
            print('Файл не найден:', e)


# ====================================================================== main
def main():
    ap = argparse.ArgumentParser(description='Клиент журнала приказов')
    ap.add_argument('--server', default='',
                    help='адрес сервера, напр. http://192.168.0.10:8080')
    ap.add_argument('--cli', action='store_true',
                    help='текстовый режим (без графического окна)')
    ap.add_argument('--login', default='', help='логин для режима --cli')
    args = ap.parse_args()

    if args.cli or not HAS_TK:
        if not HAS_TK and not args.cli:
            print('Tkinter недоступен — переключаюсь в текстовый режим '
                  '(--cli). Для графического интерфейса установите python3-tk.')
        base = args.server or input('Адрес сервера [http://localhost:8080]: ').strip() \
            or 'http://localhost:8080'
        login = args.login or input('Логин: ').strip()
        import getpass
        password = getpass.getpass('Пароль: ')
        try:
            cli_main(base, login, password)
        except KeyboardInterrupt:
            print('\nВыход')
        return

    LoginWindow, _, AppWindow = _build_gui_classes()
    win = LoginWindow(args.server)
    win.mainloop()
    if not win.result:
        return
    client, sess = win.result
    app = AppWindow(client, sess)
    app.mainloop()


if __name__ == '__main__':
    main()
