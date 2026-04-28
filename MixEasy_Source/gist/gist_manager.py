import tkinter as tk

from tkinter import ttk, messagebox, filedialog, simpledialog

import requests

import os

import base64

import json

import re

from gist import shortlink_huy as huy_gist

from gist import shortlink_mixeasy as mixeasy_gist

GITHUB_TOKEN = os.environ.get('HUY_GITHUB_TOKEN', 'YOUR_TOKEN_HERE')

BASE_URL = 'https://huyquochoc.github.io/Quiz/OnlineV2.html'

SECRET_KEY = 'MixEasy_2026'

SHORT_BASE = 'https://huyquochoc.github.io/Quiz/link.html?id='

GIST_ID = '206546f01ff4ad63c8b8755d28696fb5'

GIST_FILENAME = 'shortlinks.json'

HEADERS = {

    'Authorization': f'''token {GITHUB_TOKEN}''',

    'Accept': 'application/vnd.github+json',

    'User-Agent': 'gist_manager' }



def get_all_gists():

    r = requests.get('https://api.github.com/gists', headers = HEADERS)

    if r.status_code == 200:

        return r.json()





def create_folder_gist(folder_name):

    payload = {

        'description': f'''FOLDER::{folder_name}''',

        'public': True,

        'files': {

            'README.txt': {

                'content': f'''Folder: {folder_name}''' } } }

    r = requests.post('https://api.github.com/gists', headers = HEADERS, json = payload)

    if r.status_code == 201:

        return r.json()





def upload_file_to_folder(gist_id, file_path):

    pass

# WARNING: Decompyle incomplete





def create_home_gist_from_file(file_path):

    pass

# WARNING: Decompyle incomplete





def upload_file():

    sel = tree.selection()

    if not sel:

        messagebox.showerror('Lỗi', 'Chưa chọn Home hoặc Folder')

        return None

    target = tree.item(sel[0], 'text')

    path = filedialog.askopenfilename(filetypes = [

        ('JSON', '*.json')])

    if not path:

        return None

    if target == '🏠 Home':

        raw = create_home_gist_from_file(path)

        messagebox.showinfo('OK', 'Upload lên HOME thành công')

    else:

        gist = FOLDERS.get(target)

        if not gist:

            messagebox.showerror('Lỗi', 'Folder không hợp lệ')

            return None

        raw = upload_file_to_folder(gist['id'], path)

        messagebox.showinfo('OK', 'Upload vào folder thành công')

    load_folders()

    entry_raw.delete(0, tk.END)

    entry_raw.insert(0, raw)

    return None

# WARNING: Decompyle incomplete



FOLDERS = { }

FILES = { }

HOME_GISTS = []



def load_folders():

    pass

# WARNING: Decompyle incomplete





def on_folder_select(event):

    pass

# WARNING: Decompyle incomplete





def on_file_select(event):

    sel = files_view.selection()

    if not sel:

        return None

    fname = files_view.item(sel[0], 'text')

    raw = FILES.get(fname, '')

    entry_raw.delete(0, tk.END)

    entry_raw.insert(0, raw)

    entry_link.delete(0, tk.END)





def create_folder():

    name = simpledialog.askstring('Tạo folder', 'Tên folder:')

    if not name:

        return None

    create_folder_gist(name)

    load_folders()





def rename_folder_gist(gist_id, new_name):

    payload = {

        'description': f'''FOLDER::{new_name}''' }

    r = requests.patch(f'''https://api.github.com/gists/{gist_id}''', headers = HEADERS, json = payload)

    if r.status_code != 200:

        raise Exception('Đổi tên folder thất bại')





def rename_selected_folder():

    sel = tree.selection()

    if not sel:

        messagebox.showwarning('Thiếu chọn', 'Hãy chọn 1 folder')

        return None

    item = sel[0]

    name = tree.item(item, 'text')

    if name == '🏠 Home':

        messagebox.showinfo('Không hợp lệ', 'Không thể đổi tên Home')

        return None

    gist = FOLDERS.get(name)

    if not gist:

        messagebox.showerror('Lỗi', 'Folder không hợp lệ')

        return None

    new_name = simpledialog.askstring('Đổi tên folder', f'''Tên mới cho \'{name}\':''', initialvalue = name)

    if new_name or new_name.strip() == name:

        return None

    rename_folder_gist(gist['id'], new_name.strip())

    load_folders()

    messagebox.showinfo('OK', 'Đã đổi tên folder')

    return None

# WARNING: Decompyle incomplete





def delete_gist(gist_id):

    r = requests.delete(f'''https://api.github.com/gists/{gist_id}''', headers = HEADERS)

    if r.status_code != 204:

        raise Exception('Xoá gist thất bại')





def delete_selected_folder():

    sel = tree.selection()

    if not sel:

        messagebox.showwarning('Thiếu chọn', 'Chọn folder cần xoá')

        return None

    item = sel[0]

    name = tree.item(item, 'text')

    if name == '🏠 Home':

        messagebox.showinfo('Không hợp lệ', 'Không thể xoá Home')

        return None

    gist = FOLDERS.get(name)

    if not gist:

        messagebox.showerror('Lỗi', 'Folder không hợp lệ')

        return None

    ok = messagebox.askyesno('Xác nhận xoá', f'''Xoá folder \'{name}\'?\n\nToàn bộ file bên trong sẽ bị xoá!''')

    if not ok:

        return None

    delete_gist(gist['id'])

    load_folders()

# WARNING: Decompyle incomplete





def delete_file_from_gist(gist_id, filename):

    payload = {

        'files': {

            filename: None } }

    r = requests.patch(f'''https://api.github.com/gists/{gist_id}''', headers = HEADERS, json = payload)

    if r.status_code != 200:

        raise Exception('Xoá file gốc thất bại')





def choose_target_folder():

    pass

# WARNING: Decompyle incomplete





def move_selected_file():

    sel_file = files_view.selection()

    if not sel_file:

        messagebox.showwarning('Thiếu chọn', 'Chọn file cần move')

        return None

    filename = files_view.item(sel_file[0], 'text')

    raw_url = FILES.get(filename)

    if not raw_url:

        messagebox.showerror('Lỗi', 'Không xác định được file')

        return None

    target_folder = choose_target_folder()

    if not target_folder:

        return None

    target_gist = FOLDERS.get(target_folder)

    if not target_gist:

        return None

    content = requests.get(raw_url).text

    payload = {

        'files': {

            filename: {

                'content': content } } }

    r = requests.patch(f'''https://api.github.com/gists/{target_gist['id']}''', headers = HEADERS, json = payload)

    if r.status_code != 200:

        raise Exception('Không thêm được file')

    src_gist_id = raw_url.split('/')[4]

    delete_file_from_gist(src_gist_id, filename)

    load_folders()

    messagebox.showinfo('OK', f'''Đã move \'{filename}\' → {target_folder}''')

    return None

# WARNING: Decompyle incomplete





def delete_selected_file():

    sel = files_view.selection()

    if not sel:

        messagebox.showwarning('Thiếu chọn', 'Chọn file cần xoá')

        return None

    filename = files_view.item(sel[0], 'text')

    raw_url = FILES.get(filename)

    if not raw_url:

        messagebox.showerror('Lỗi', 'Không xác định được file')

        return None

    ok = messagebox.askyesno('Xác nhận xoá', f'''Xoá file \'{filename}\'?\n\nHành động không thể hoàn tác!''')

    if not ok:

        return None

    gist_id = raw_url.split('/')[4]

    delete_file_from_gist(gist_id, filename)

    files_view.delete(sel[0])

    FILES.pop(filename, None)

    entry_raw.delete(0, tk.END)

    entry_link.delete(0, tk.END)

    messagebox.showinfo('OK', 'Đã xoá file')

    return None

# WARNING: Decompyle incomplete





def rename_selected_file():

    sel = files_view.selection()

    if not sel:

        messagebox.showwarning('Thiếu chọn', 'Chọn file cần đổi tên')

        return None

    old_name = files_view.item(sel[0], 'text')

    raw_url = FILES.get(old_name)

    if not raw_url:

        messagebox.showerror('Lỗi', 'Không xác định được file')

        return None

    new_name = simpledialog.askstring('Đổi tên file', 'Tên file mới:', initialvalue = old_name)

    if new_name or new_name.strip() == old_name:

        return None

    new_name = new_name.strip()

    gist_id = raw_url.split('/')[4]

    content = requests.get(raw_url).text

    payload = {

        'files': {

            new_name: {

                'content': content } } }

    r = requests.patch(f'''https://api.github.com/gists/{gist_id}''', headers = HEADERS, json = payload)

    if r.status_code != 200:

        raise Exception('Không tạo được file mới')

    delete_file_from_gist(gist_id, old_name)

    files_view.delete(sel[0])

    FILES.pop(old_name, None)

    username = raw_url.split('/')[3]

    FILES[new_name] = f'''https://gist.githubusercontent.com/{username}/{gist_id}/raw/{new_name}'''

    files_view.insert('', 'end', text = new_name)

    entry_raw.delete(0, tk.END)

    entry_link.delete(0, tk.END)

    messagebox.showinfo('OK', 'Đã đổi tên file')

    return None

# WARNING: Decompyle incomplete





def xor_encrypt(text, key):

    pass

# WARNING: Decompyle incomplete





def create_exam_link(data_url):

    xor_text = xor_encrypt(data_url, SECRET_KEY)

    encoded_url = base64.b64encode(xor_text.encode()).decode()

    return f'''{BASE_URL}?data={encoded_url}'''





def create_link():

    raw = entry_raw.get().strip()

    if not raw:

        messagebox.showerror('Lỗi', 'Chưa có RAW')

        return None

    entry_link.delete(0, tk.END)

    entry_link.insert(0, create_exam_link(raw))





def copy_link(root, entry_link):

    link = entry_link.get().strip()

    if not link:

        messagebox.showwarning('Chưa có link', 'Chưa tạo link đề thi để copy')

        return None

    root.clipboard_clear()

    root.clipboard_append(link)

    root.update()

    messagebox.showinfo('Đã copy', 'Link đề thi đã được copy vào clipboard')





def load_links():

    url = f'''https://api.github.com/gists/{GIST_ID}'''

    r = requests.get(url, headers = HEADERS)

    if r.status_code != 200:

        messagebox.showerror('Lỗi', 'Không load được Gist shortlinks')

        return ({ }, None)

    gist = None.json()

    files = gist.get('files', { })

    if GIST_FILENAME not in files:

        return ({ }, None)

    content = None[GIST_FILENAME]['content']

    return (json.loads(content), None)

# WARNING: Decompyle incomplete





def save_links(obj, _sha = (None,)):

    content = json.dumps(obj, indent = 2, ensure_ascii = False)

    payload = {

        'files': {

            GIST_FILENAME: {

                'content': content } } }

    url = f'''https://api.github.com/gists/{GIST_ID}'''

    r = requests.patch(url, headers = HEADERS, json = payload)

    if r.status_code != 200:

        print('Gist error:', r.status_code, r.text)

        return False

    return True





def generate_key(name, existing_keys):

    base = re.sub('\\W+', '', name)

    if not base:

        base = 'link'

    key = base

    i = 1

    if key in existing_keys:

        key = f'''{base}{i}'''

        i += 1

        if key in existing_keys:

            continue

    return key





def create_short_link(entry_link, entry_short_name, entry_short_link, entry_expire = (None,)):

    name = entry_short_name.get().strip()

    if not name:

        messagebox.showerror('Lỗi', 'Chưa đặt tên đề thi')

        return None

    long_url = entry_link.get().strip()

    (data, _) = load_links()

# WARNING: Decompyle incomplete





def decode_exam_link(entry_link, entry_base64_decode, entry_xor_raw):

    encoded = entry_link.get().strip()

    if '?data=' in encoded:

        encoded = encoded.split('?data=')[1]

    if not encoded:

        messagebox.showerror('Error', 'Chưa có chuỗi mã hóa')

        return None

    decoded = base64.b64decode(encoded).decode()

    entry_base64_decode.delete(0, tk.END)

    entry_base64_decode.insert(0, decoded)

    original = xor_encrypt(decoded, SECRET_KEY)

    entry_xor_raw.delete(0, tk.END)

    entry_xor_raw.insert(0, original)

    return None

# WARNING: Decompyle incomplete





def gui_gist(root):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    gui_gist(root)

    root.mainloop()

    return None

