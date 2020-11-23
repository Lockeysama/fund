import json
import os
import tkinter
import tkinter as tk
from tkinter.messagebox import showinfo
from tkinter.scrolledtext import ScrolledText

from crawl import eastmoney
from export import excel


class Application(tk.Frame):

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.public_fids = []
        self.private_fids = []
        self.config = {}
        if os.path.exists('.config'):
            with open('.config', 'r') as f:
                self.config = json.loads(f.read())
                self.public_fids = self.config.get('public_fids', [])
                self.private_fids = self.config.get('private_fids', [])
        self.public_fids_st = None
        self.private_urls_st = None
        self.interval_text = None
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        public_label = tk.Label(self, text='公募 ID 列表（每个 ID 一行）：')
        public_label.grid(row=0, column=0)
        self.public_fids_st = ScrolledText(self, wrap=tk.WORD)
        self.public_fids_st.grid(row=1, column=0, columnspan=2)
        if self.config:
            fids = self.config.get('public_fids')
            if fids:
                self.public_fids_st.insert(tk.INSERT, '\n'.join(fids))

        private_label = tk.Label(self, text='私募 ID 列表（每个 ID 一行）：')
        private_label.grid(row=0, column=2)
        self.private_urls_st = ScrolledText(self)
        self.private_urls_st.grid(row=1, column=2, columnspan=2)
        if self.config:
            fids = self.config.get('private_fids')
            if fids:
                self.private_urls_st.insert(tk.INSERT, '\n'.join(fids))

        interval_label = tk.Label(self, text='抓取时间间隔：')
        interval_label.grid(row=2, column=0)
        self.interval_text = tk.Entry(self)
        self.interval_text.grid(row=2, column=1)

        start_btn = tk.Button(self, text='开始', fg='green', command=self.start)
        start_btn.grid(row=7, column=0)

        quit_btn = tk.Button(self, text="停止", fg="red", command=self.quit)
        quit_btn.grid(row=7, column=2)

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

    def start(self):
        self.running = True
        self.get_fids(1)
        self.get_fids(2)

        for fid in self.public_fids:
            net_worth = eastmoney.get_net_worth(fid)
            # phase_increases = eastmoney.get_phase_increases(fid)
            headers = {
                'title': '易方达（{}）'.format(fid),
                'subs': [
                    {'title': '单位净值'},
                    {'title': '累计净值'},
                ]
            }
            excel.make('公募', headers, net_worth)


        # r2 = eastmoney.get_phase_increases(self.get_public_fids()[0])
        # showinfo('test', r1)
        # showinfo('test2', r2)

    def quit(self):
        self.running = False
        self.config['public_fids'] = self.public_fids
        self.config['private_fids'] = self.private_fids
        with open('.config', 'w') as f:
            f.write(json.dumps(self.config))
        super(Application, self).quit()
        print('quit')
