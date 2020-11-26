import json
import logging
import os
import threading
import time
import tkinter
import tkinter as tk
from logging.handlers import TimedRotatingFileHandler
from tkinter.scrolledtext import ScrolledText

from cryptography.fernet import Fernet

from crawl import eastmoney, licai
from export import excel


log_fmt = (
    '[%(asctime)s] [%(levelname)s] [%(name)s:%(lineno)d:%(funcName)s] )=> %(message)s'
)
formatter = logging.Formatter(log_fmt)

file_handler = TimedRotatingFileHandler(
    filename="./.fund.log", when="midnight", backupCount=7
)

file_handler.setFormatter(formatter)
file_handler.setLevel(logging.DEBUG)
logging.getLogger().addHandler(file_handler)

stream_handler = logging.StreamHandler()
stream_handler.setLevel(logging.DEBUG)
stream_handler.setFormatter(formatter)
logging.getLogger().addHandler(stream_handler)
logging.getLogger().setLevel(logging.DEBUG)


class Application(tk.Frame):

    key = 'Zjc5b3NeaGdEUSJhMjZsZVU0U3B8PmlOWjg8dl0sS1g='

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.running = False
        self.interval_val = tk.StringVar()
        self.username_val = tk.StringVar()
        self.password_val = tk.StringVar()
        self.public_fids = []  # ['策略 1:']
        self.private_fids = []
        self.config = {}
        if os.path.exists('.config'):
            with open('.config', 'r') as f:
                c = f.read()
                if c:
                    self.config = json.loads(c)
                    self.public_fids = self.config.get('public_fids', [])
                    self.private_fids = self.config.get('private_fids', [])
        self.public_fids_st = None
        self.private_urls_st = None
        self.log_st = None
        self.interval_text = None
        self.start_btn = None
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        public_label = tk.Label(self, text='eastmoney ID 列表（每个 ID 一行）：')
        public_label.grid(row=0, column=0)
        self.public_fids_st = ScrolledText(self, wrap=tk.WORD)
        self.public_fids_st.grid(row=1, column=0, columnspan=2)
        if self.config:
            fids = self.config.get('public_fids')
            if fids:
                self.public_fids_st.insert(tk.INSERT, '\n'.join(fids))

        private_label = tk.Label(self, text='licai ID 列表（每个 ID 一行）：')
        private_label.grid(row=0, column=2)
        self.private_urls_st = ScrolledText(self)
        self.private_urls_st.grid(row=1, column=2, columnspan=2)
        if self.config:
            fids = self.config.get('private_fids')
            if fids:
                self.private_urls_st.insert(tk.INSERT, '\n'.join(fids))

        interval_label = tk.Label(self, text='抓取时间间隔（秒）：')
        interval_label.grid(row=2, column=0)
        interval = '3'
        if self.config:
            interval = self.config.get('interval', '3')
        self.interval_text = tk.Entry(self, textvariable=self.interval_val)
        self.interval_text.grid(row=2, column=1)
        self.interval_val.set(interval)

        login_view = tk.Label(self)
        login_view.grid(row=2, column=2, columnspan=2)

        username_label = tk.Label(login_view, text='（licai）用户名：')
        username_label.grid(row=1, column=1)
        username_text = tk.Entry(login_view, textvariable=self.username_val)
        username_text.grid(row=1, column=2)
        self.username_val.set(self.config.get('username', ''))

        passwd_label = tk.Label(login_view, text='（licai）密码：')
        passwd_label.grid(row=1, column=3)
        passwd_text = tk.Entry(login_view, textvariable=self.password_val, show='*')
        passwd_text.grid(row=1, column=4)
        passwd = self.config.get('password', '')
        if passwd:
            f = Fernet(key=self.key)
            passwd = f.decrypt(passwd.encode()).decode()
        self.password_val.set(passwd)

        log_view = tk.Label(self)
        log_view.grid(row=3, rowspan=2, columnspan=2)

        log_label = tk.Label(log_view, text='日志：')
        log_label.grid(row=1, columnspan=1, column=0)
        self.log_st = ScrolledText(log_view)
        self.log_st.grid(row=2, rowspan=9)

        self.start_btn = tk.Button(self, text='开始', fg='green', command=self.start)
        self.start_btn.grid(row=3, column=2)

        quit_btn = tk.Button(self, text="停止", fg="red", command=self.quit)
        quit_btn.grid(row=4, column=2)

    def log(self, msg):
        self.log_st.insert(tkinter.END, msg + '\n')
        self.log_st.see(tkinter.END)
        logging.getLogger(__name__).debug(msg)

    def get_fids(self, id_type):
        if id_type not in [1, 2]:
            raise Exception('fid type error')
        if id_type == 1:
            content = self.public_fids_st.get("1.0", tkinter.END)
        else:
            content = self.private_urls_st.get("1.0", tkinter.END)
        if content.find('\n\t') > 0:
            _fids = content.split('\n\t')
        else:
            _fids = content.split('\n')
        if id_type == 1:
            self.public_fids = []
        else:
            self.private_fids = []
        for fid in _fids:
            fid = fid.strip()
            if not fid:
                continue
            if id_type == 1 and fid in self.public_fids:
                continue
            elif id_type == 2 and fid in self.private_fids:
                continue
            if id_type == 1:
                self.public_fids.append(fid)
            else:
                self.private_fids.append(fid)
        if id_type == 1:
            return self.public_fids
        else:
            return self.private_fids

    def work(self):
        self.log('Start...')
        self.log('Init config...\n')
        self.get_fids(1)
        self.get_fids(2)

        for name in self.public_fids:
            if not self.running:
                self.running = False
                return
            self.log('Crawl({}) data...'.format(name))
            tactics, fid, title = name.split(':')
            net_worth = eastmoney.get_net_worth(fid)
            phase_increases = eastmoney.get_phase_increases(fid)
            if not net_worth or not phase_increases:
                self.log('Crawl({}) failed...'.format(name))
                self.log('Crawl({}) stopped...\n'.format(name))
                self.running = False
                return
            increase = ['阶段涨幅', '同类平均', '沪深 300', '同类排名']
            cycle = ['近 1 周', '近 1 月', '近 3 月', '近 6 月', '近 1 年']
            increase_cycle = ['{}（{}）'.format(i, c) for i in increase for c in cycle]
            headers = {
                'items': increase_cycle,
                'title': '{}（{}）'.format(title, fid),
                'subs': [
                    {'title': '单位净值'},
                    {'title': '累计净值'},
                ]
            }
            self.log('Storage({}) data...'.format(name))
            excel.make(tactics, headers, (phase_increases, net_worth))
            try:
                self.log('Sleep({}s)...\n'.format(self.interval_val.get()))
                time.sleep(int(self.interval_val.get()))
            except Exception as e:
                self.log('Interval error...'.format(name))
                self.log('Sleep(3s)...\n')
                time.sleep(3)

        username = self.username_val.get()
        password = self.password_val.get()
        session = licai.login(username, password)
        if not session:
            self.log('Login failed({} : {})...'.format(username, password))
            self.log('Crawl stopped...\n')
            self.running = False
            return
        for name in self.private_fids:
            if not self.running:
                self.running = False
                return
            self.log('Crawl({}) data...'.format(name))
            tactics, fid, title = name.split(':')
            earnings, sharpe_ratio, result = licai.get_net_worth(session, fid)
            if not tactics:
                self.log('Crawl({}) failed...'.format(name))
                self.log('Crawl({}) stopped...\n'.format(name))
                continue
            hd = earnings
            hd.extend(sharpe_ratio)
            headers = {
                'items': ['年化收益', '近一年收益', '夏普比率'],
                'title': '{}（{}）'.format(title, fid),
                'subs': [
                    {'title': '单位净值'},
                    {'title': '复权净值'},
                    {'title': '累计净值'},
                ]
            }
            self.log('Storage({}) data...'.format(name))
            excel.make(tactics, headers, (hd, result))
            try:
                self.log('Sleep({}s)...\n'.format(self.interval_val.get()))
                time.sleep(int(self.interval_val.get()))
            except Exception as e:
                self.log('Interval error...'.format(name))
                self.log('Sleep(3s)...\n')
                time.sleep(3)
        self.log('Done...')
        self.running = False
        print('Done')

    def start(self):
        if self.running:
            self.log('\nWaiting...\n')
        else:
            self.running = True
            threading.Thread(target=self.work).start()

    def quit(self):
        self.running = False
        self.config['public_fids'] = self.public_fids
        self.config['private_fids'] = self.private_fids
        self.config['interval'] = self.interval_val.get()
        self.config['username'] = self.username_val.get()
        passwd = self.password_val.get()
        if passwd:
            f = Fernet(self.key)
            self.config['password'] = f.encrypt(passwd.encode()).decode()
        with open('.config', 'w') as f:
            f.write(json.dumps(self.config))
        super(Application, self).quit()
        print('quit')
