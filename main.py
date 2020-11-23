import tkinter as tk

import app


if __name__ == '__main__':
    root = tk.Tk()
    app = app.Application(master=root)
    app.mainloop()
