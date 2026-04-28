import os
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
import base64
from datetime import datetime
GITHUB_TOKEN = os.environ.get('MIXEASY_GITHUB_TOKEN', 'YOUR_TOKEN_HERE')
GIST_ID = '22b1a55c4103343ce287737f2e29171c'
GIST_FILENAME = 'shortlinks.json'
SHORT_BASE = 'https://mixeasy.github.io/Quiz/link.html?id='
HEADERS = {
    'Authorization': f'''token {GITHUB_TOKEN}''',
    'Accept': 'application/vnd.github+json',
    'User-Agent': 'gist_manager' }
API_GIST = f'''https://api.github.com/gists/{GIST_ID}'''

def load_links():
    r = requests.get(API_GIST, headers = HEADERS)
    if r.status_code != 200:
        return { }
    files = None.json()['files']
    if GIST_FILENAME not in files:
        return { }
    content = None[GIST_FILENAME]['content']
    return json.loads(content)


def save_links(obj):
    payload = {
        'files': {
            GIST_FILENAME: {
                'content': json.dumps(obj, indent = 2, ensure_ascii = False) } } }
    r = requests.patch(API_GIST, headers = HEADERS, json = payload)
    return r.status_code == 200


def open_shortlink_manager(win):
    pass
# WARNING: Decompyle incomplete

if __name__ == '__main__':
    win = tk.Tk()
    win.title('Test ShortLink Manager')
    win.geometry('300x120')
    btn = tk.Button(win, text = 'Open ShortLink Manager', command = (lambda : open_shortlink_manager(win)))
    btn.pack(expand = True, pady = 30)
    win.mainloop()
    return None
