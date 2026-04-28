import os

import tkinter as tk

from tkinter import ttk, filedialog, messagebox

from docx import Document

from docx.text.paragraph import Paragraph

from docx.table import Table

from docx.oxml.ns import qn

from docx.shared import RGBColor

from lxml import etree

from PIL import Image

from officemath2latex import process_math_string

import re

import zipfile

import tempfile

import base64

import io

import json

import csv

import re

import requests

from datetime import datetime

import unicodedata

import webbrowser

from html_generators import trac_nghiem_online_html as html

from formatting import chuan_hoa as chd

from core import functions as vbf

from mixing import mixeasy as vme

from conversion import word_to_html as cv2html

from conversion import omml_to_latex as OMML_latex



def open_link_online():

    url = 'https://docs.google.com/spreadsheets/d/1t-Z3oZyuFtn7hgzmRG9ZwDr168CSnJUfpqY7D_I7rCs/edit?usp=drive_link'

    webbrowser.open_new(url)





def open_link_url(url):

    webbrowser.open_new(url)





def safe_filename(text, max_length = (80,)):

    text = unicodedata.normalize('NFD', text)

    text = (lambda .0: pass# WARNING: Decompyle incomplete

)(text())

    text = text.lower()

    text = re.sub('[^a-z0-9]+', '_', text)

    text = text.strip('_')

    return text[:max_length]





def omml_to_latex(omml_xml = None):

    latex = OMML_latex.omml_to_latex(omml_xml)

    return latex





def run_to_html(run):

    pass

# WARNING: Decompyle incomplete





def para_to_html(p):

    parts = []

# WARNING: Decompyle incomplete





def table_to_html(tbl):

    pass

# WARNING: Decompyle incomplete





def remove_prefix(paragraph, regex):

    '''Xóa tiền tố khỏi một đoạn văn bản (paragraph) mà không làm mất định dạng.'''

    full_text = paragraph.text

    match = re.match(regex, full_text)

    if not match:

        return None

    prefix_len = len(match.group(0))

    for run in paragraph.runs:

        if prefix_len <= 0:

            paragraph.runs

            return None

        run_text = run.text

        if len(run_text) <= prefix_len:

            prefix_len -= len(run_text)

            run.text = ''

            continue

        run.text = run_text[prefix_len:]

        prefix_len = 0





def process_block_type1(block_content, start_q_num, original_index_counter):

    pass

# WARNING: Decompyle incomplete





def process_block_type2(block_content, start_q_num, original_index_counter):

    pass

# WARNING: Decompyle incomplete





def process_block_type3(block_content, start_q_num, original_index_counter):

    pass

# WARNING: Decompyle incomplete





def process_block_general_content(block_content, start_q_num, original_index_counter):

    pass

# WARNING: Decompyle incomplete





def _is_group_start(text = None):

    if not '<SCHOICE@' in text:

        '<SCHOICE@' in text

    return '<SCHUM@' in text





def _is_group_end(text = None):

    if not '<ECHOICE@' in text:

        '<ECHOICE@' in text

    return '<ECHUM@' in text





def _is_change_start(text = None):

    return text.startswith('<SCHANGE@')





def _is_change_end(text = None):

    return text.startswith('<ECHANGE@')





def _try_parse_var_decl(line = None):

    '''

    Parse 1 dòng khai báo biến dạng:   x in {1;2;3}

    Trả về (name, [values]) hoặc None

    '''

    m = re.match('^\\s*([A-Za-z_]\\w*)\\s+in\\s+\\{([^}]*)\\}\\s*$', line)

    if not m:

        return None

    (name, values_str) = m.groups()

# WARNING: Decompyle incomplete





def parse_docx_to_quiz_data(docx_path):

    '''

    - Đọc docx và tạo grouped_data theo chuẩn mới (list các group; mỗi group là list các câu).

    - Hỗ trợ:

        * S1@/E1@, S2@/E2@, S3@/E3@

        * S4@/E4@: phần nội dung hiển thị nguyên khối (không tách từng câu)

        * <SCHOICE@>/<ECHOICE@>

        * <SCHANGE@>/<ECHANGE@>

    '''

    pass

# WARNING: Decompyle incomplete



answer_key_cache = None

quiz_title_cache = None



def generate_files(entry_name_tieu_de, combo_script_url, entry_time_limit, entry_sheet_name, shuffle_vars, xuatketqua_var, entry_diem_loai1, entry_diem_loai21, entry_diem_loai22, entry_diem_loai23, entry_diem_loai24, entry_diem_loai3, default_local_var, csv_var, View_all_var, cal_time_var, create_vars, entry_pass):

    if not entry_name_tieu_de.get().strip():

        entry_name_tieu_de.get().strip()

    quiz_title = 'LUYỆN TẬP'

    file_name = quiz_title

    script_url = combo_script_url.get().strip()

    if not entry_time_limit.get().strip():

        entry_time_limit.get().strip()

    time_limit = '20'

    if not entry_sheet_name.get().strip():

        entry_sheet_name.get().strip()

    sheet_name = 'KetQua'

# WARNING: Decompyle incomplete





def add_ki_hieu_nhan_dien(word):

    doc = word.ActiveDocument

    vbf.STT_2025_new(word)

    vbf.thay_the_replace_1(word, '(^13PHẦN II.)', '^13E1@\\1')

    vbf.thay_the_replace_1(word, '(^13Phần II.)', '^13E1@\\1')

    vbf.thay_the_replace_1(word, '(^13PHẦN 2.)', '^13E1@\\1')

    vbf.thay_the_replace_1(word, '(^13Phần 2.)', '^13E1@\\1')

    vbf.thay_the_replace_1(word, '(^13PHẦN III.)', '^13E2@\\1')

    vbf.thay_the_replace_1(word, '(^13Phần III.)', '^13E2@\\1')

    vbf.thay_the_replace_1(word, '(^13PHẦN 3.)', '^13E2@\\1')

    vbf.thay_the_replace_1(word, '(^13Phần 3.)', '^13E2@\\1')

    myrange = doc.Range()

    for i in range(1, 4):

        find = myrange.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = '(^13Câu 1[.:])'

        find.MatchWildcards = True

        find.Replacement.Text = f'''^13S{i}@\\1'''

        find.MatchWildcards = True

        find.MatchCase = True

        find.Forward = True

        find.Format = True

        if not find.Execute(Replace = 1):

            continue

        myrange = doc.Range(myrange.End, doc.Range().End)

    selection = word.Selection

    selection.EndKey(Unit = 6)

    word.Selection.TypeParagraph()

    word.Selection.TypeText('E3@')

    word.Selection.TypeParagraph()

    return None

# WARNING: Decompyle incomplete





def dem_cong_thuc_mathtype_1(word):

    '''

    Đếm số công thức MathType có trong vùng Range (hoặc toàn bộ Document)

    '''

    doc = word.ActiveDocument

    myrange = doc.Range()

    count = 0

    for shape in myrange.InlineShapes:

        if shape.Type == 1:

            class_type = shape.OLEFormat.ClassType

            if class_type and 'Equation.DSMT' in class_type:

                count += 1

    continue

    return count

# WARNING: Decompyle incomplete





def dem_cong_thuc_mathtype(word):

    '''

    Đếm và highlight công thức MathType trong tài liệu Word.

    '''

    doc = word.ActiveDocument

    myrange = doc.Range()

    count = 0

    for shape in myrange.InlineShapes:

        if shape.Type == 1:

            class_type = shape.OLEFormat.ClassType

            if class_type:

                if 'Equation.DSMT' in class_type or 'Equation.3' in class_type:

                    count += 1

                    shape.Range.HighlightColorIndex = 6

    continue

    return count

# WARNING: Decompyle incomplete





def check_du_lieu_online(word):

    doc = word.ActiveDocument

    so_mathtype = dem_cong_thuc_mathtype(word)

    if so_mathtype > 0:

        messagebox.showinfo('Thông báo', f'''có {so_mathtype} công thức còn ở dạng Mathtype chưa chuyển qua Latex\nHãy chuyển nó qua Latex và kiểm tra lại các lỗi khác\n(Xem những chỗ Highlight màu đỏ)''')

        return None

    vme.check_data_nhanh_combine_docx(word)

    return None

# WARNING: Decompyle incomplete



GITHUB_TOKEN = os.environ.get('MIXEASY_GITHUB_TOKEN', 'YOUR_TOKEN_HERE')

GITHUB_USER = 'mixeasy'

REPO_NAME = 'Quiz'

BRANCH = 'main'

BASE_URL = 'https://mixeasy.github.io/Quiz/OnlineV2.html'

SHORT_BASE = 'https://mixeasy.github.io/Quiz/link.html?id='

GIST_ID = '22b1a55c4103343ce287737f2e29171c'

GIST_FILENAME = 'shortlinks.json'

HEADERS = {

    'Authorization': f'''token {GITHUB_TOKEN}''',

    'Accept': 'application/vnd.github+json',

    'User-Agent': 'MixEasy-Upjson-Linkonline' }



def choose_file(entry_json):

    path = filedialog.askopenfilename(filetypes = [

        ('JSON files', '*.json')])

    if path:

        entry_json.delete(0, tk.END)

        entry_json.insert(0, path)

        return None





def upload_json_to_gist(json_path):

    pass

# WARNING: Decompyle incomplete





def upload_action2(entry_json):

    path = entry_json.get().strip()

    if not path or os.path.isfile(path):

        messagebox.showerror('Lỗi', 'File JSON không hợp lệ, có thể chưa chọn file')

        return None

    gist = upload_json_to_gist(path)

    username = gist['owner']['login']

    gist_id = gist['id']

    filename = list(gist['files'].keys())[0]

    raw_link = f'''https://gist.githubusercontent.com/{username}/{gist_id}/raw/{filename}'''

    return raw_link

# WARNING: Decompyle incomplete





def upload_action(path):

    gist = upload_json_to_gist(path)

    username = gist['owner']['login']

    gist_id = gist['id']

    filename = list(gist['files'].keys())[0]

    return f'''https://gist.githubusercontent.com/{username}/{gist_id}/raw/{filename}'''





def xor_encrypt(text, key):

    pass

# WARNING: Decompyle incomplete





def create_exam_link(data_url):

    BASE_URL = 'https://mixeasy.github.io/Quiz/OnlineV2.html'

    SECRET_KEY = 'MixEasy_2026'

    xor_text = xor_encrypt(data_url, SECRET_KEY)

    encoded_url = base64.b64encode(xor_text.encode()).decode()

    return f'''{BASE_URL}?data={encoded_url}'''





def create_link_action2(entry_json):

    raw = upload_action(entry_json)

    if not raw:

        messagebox.showerror('Lỗi', 'Thiếu RAW link do chưa upload Json được')

        return None

    link = create_exam_link(raw)

    return link

# WARNING: Decompyle incomplete





def create_link_action(path):

    raw = upload_action(path)

    return create_exam_link(raw)





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





def create_short_link2(entry_json, entry_short_name, entry_short_link, entry_expire = (None,)):

    long_url = create_link_action(entry_json)

    name = entry_short_name.get().strip()

    expire = None

    if not long_url or name:

        messagebox.showerror('Lỗi', 'Thiếu link dài hoặc tên đề')

        return None

    (data, _) = load_links()

# WARNING: Decompyle incomplete





def create_short_link(entry_json, entry_short_name, entry_short_link, entry_expire = (None,)):

    path = entry_json.get().strip()

    name = entry_short_name.get().strip()

    if not path:

        messagebox.showerror('Lỗi', 'không lấy được đường dẫn file json để up')

        return None

    if not name:

        messagebox.showerror('Lỗi', 'Chưa đặt tên đề thi')

        return None

    long_url = create_link_action(path)

    (data, _) = load_links()

# WARNING: Decompyle incomplete





def copy_link_action(root, entry_link):

    link = entry_link.get().strip()

    if not link:

        messagebox.showwarning('Chưa có link', 'Chưa tạo link đề thi để copy')

        return None

    root.clipboard_clear()

    root.clipboard_append(link)

    root.update()

    messagebox.showinfo('Đã copy', 'Link đề thi đã được copy vào clipboard')



FILE_PATH = 'C:/html/mixonline/urlonline/url_list.txt'



def load_url_key():

    url_dict = { }

# WARNING: Decompyle incomplete





def save_url_key(combo_script_url, combo_keys_on):

    url = combo_script_url.get().strip()

    key_name = combo_keys_on.get().strip()

    if not url or key_name:

        messagebox.showwarning('Thiếu thông tin', 'Bạn phải nhập URL và key!')

        return None

    os.makedirs(os.path.dirname(FILE_PATH), exist_ok = True)

# WARNING: Decompyle incomplete





def on_key_select(event, combo_script_url, combo_keys_on):

    selected_key = combo_keys_on.get()

    url_dict = load_url_key()

    if selected_key in url_dict:

        combo_script_url.delete(0, tk.END)

        combo_script_url.insert(0, url_dict[selected_key])

        return None





def on_key_select_ctk(selected_key, combo_script_url):

    url_dict = load_url_key()

    if selected_key in url_dict:

        combo_script_url.delete(0, 'end')

        combo_script_url.insert(0, url_dict[selected_key])

        return None





def on_key_select_url(event, combo_script_url):

    selected_key = combo_script_url.get()

    url_dict = load_url_key()

    if selected_key in url_dict:

        combo_script_url.set(url_dict[selected_key])

        return None





def delete_url_key(combo_keys_on, combo_script_url):

    selected_key = combo_keys_on.get()

    if not selected_key:

        messagebox.showwarning('Chưa chọn', 'Bạn chưa chọn key để xóa!')

        return None

    url_dict = load_url_key()

    if selected_key not in url_dict:

        messagebox.showwarning('Không tồn tại', f'''Key \'{selected_key}\' không tồn tại!''')

        return None

# WARNING: Decompyle incomplete





def update_combos(combo_keys_on):

    url_dict = load_url_key()

    keys = list(url_dict.keys())

    combo_keys_on['values'] = keys





def update_combos_ctk(combo_keys_on):

    url_dict = load_url_key()

    keys = list(url_dict.keys())

    combo_keys_on.configure(values = keys)

    combo_keys_on.set('')





def gui_on(root, word):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    from gist import gist_manager as GIST

    word = vbf.khoi_tao_word_2()

    Nut_copy2_link = tk.Button(root, text = 'Trắc nghiệm Online', command = (lambda : gui_on(root, word)), bg = 'blanchedalmond', fg = 'black', width = 20, height = 2)

    Nut_copy2_link.grid(row = 12, column = 0, columnspan = 2, padx = 5, sticky = 'ew', pady = (0, 10))

    Nut_copy2_link = tk.Button(root, text = 'GIST MANAGER PRO', command = (lambda : GIST.gui_gist(root)), bg = 'blanchedalmond', fg = 'black', width = 20, height = 2)

    Nut_copy2_link.grid(row = 12, column = 3, columnspan = 2, padx = 5, sticky = 'ew', pady = (0, 10))

    root.mainloop()

    return None

