'''
Transparent Mini Timer (grid layout, Hide button separated)
'''
import tkinter as tk
from tkinter import filedialog, ttk
import winsound
import os
import json
CHROMA = 'magenta'
BG_PANEL = CHROMA
FG_TIME_DEFAULT = '#ff4444'
CONFIG_FILE = 'timer_config.json'

class TransparentTimer(tk.Toplevel):
    pass
# WARNING: Decompyle incomplete


def clockmain():
    app = TransparentTimer()
    app.mainloop()

if __name__ == '__main__':
    clockmain()
    return None
