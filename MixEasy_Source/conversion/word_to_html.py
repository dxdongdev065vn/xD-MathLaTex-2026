import tkinter as tk

from tkinter import filedialog, messagebox

import mammoth

from pathlib import Path

import re

import os

import zipfile

import base64

import io

from PIL import Image

from lxml import etree

from officemath2latex import process_math_string

from docx import Document

from docx.text.paragraph import Paragraph

from docx.table import Table

from conversion import omml_to_latex as OMML_latex



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





def convert_docx_to_html(input_path):

    '''

    Chuyển đổi một file .docx sang HTML, bao gồm công thức toán, hình ảnh và bảng.

    

    :param docx_file: Đường dẫn đến file docx đầu vào.

    :param output_file: Đường dẫn đến file html đầu ra.

    '''

    document = Document(input_path)

    name_hien_thi = Path(input_path).stem

    html_content = []

    for element in document.element.body:

        tag = element.tag.split('}')[-1]

        if tag == 'p':

            paragraph = Paragraph(element, document)

            html_content.append(para_to_html(paragraph))

            continue

        if not tag == 'tbl':

            continue

        table = Table(element, document)

        html_content.append(table_to_html(table))

    body_html = '\n'.join(html_content)

    html_template = f'''<!DOCTYPE html>\n    <html lang="vi">\n    <head>\n    <meta charset="UTF-8">\n    <meta name="viewport" content="width=device-width, initial-scale=1.0">\n    <title>{name_hien_thi}</title>\n    <style>\n        body {{ font-family: Arial, sans-serif; line-height: 1.6; margin: 20px; }}\n        img {{ max-width: 100%; height: auto; display: block; margin: 10px auto; }}\n        table {{ width: 80%; border-collapse: collapse; margin: 20px auto; }}\n        th, td {{ padding: 8px; text-align: left; border: 1px solid #ddd; }}\n        /* Đảm bảo MathJax block math giữa trang */\n        p[style*=\'text-align:center;\'] {{ text-align: center; }}\n    </style>\n    <script src="https://polyfill.io/v3/polyfill.min.js?features=es6"></script>\n    <script>\n    window.MathJax = {{\n      tex: {{ inlineMath: [[\'$\', \'$\'], [\'\\\\(\', \'\\\\)\']] }}\n    }};\n    </script>\n    <script id="MathJax-script" async\n            src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"\n            onerror="loadLocalMathJax()"></script>\n    <script>\n    function loadLocalMathJax() {{\n      var s = document.createElement(\'script\');\n      s.src = \'./MathJax-3.2.2/es5/tex-mml-chtml.js\';\n      s.async = true;\n      document.head.appendChild(s);\n    }}\n    </script>\n    </head>\n    <body>\n    {body_html}\n    </body>\n    </html>'''

    output_file = Path(input_path).with_suffix('.html')

# WARNING: Decompyle incomplete





def add_files(listbox):

    file_paths = filedialog.askopenfilenames(title = 'Chọn các file Word (.docx)', filetypes = [

        ('Word Documents', '*.docx')])

    for path in file_paths:

        path = os.path.normpath(path)

        if not path not in listbox.get(0, tk.END):

            continue

        listbox.insert(tk.END, path)





def clear_list(listbox):

    listbox.delete(0, tk.END)





def convert_all(listbox):

    files = listbox.get(0, tk.END)

    if not files:

        messagebox.showwarning('Chưa có file', 'Hãy thêm ít nhất một file để chuyển.')

        return None

    for f in files:

        f = os.path.normpath(f)

        convert_docx_to_html(f)

    messagebox.showinfo('Hoàn tất', '✅ Đã chuyển xong tất cả các file!')





def convert_word2html_one():

    docx_path = filedialog.askopenfilename(title = 'Chọn file câu hỏi Word (.docx)', filetypes = [

        ('Word Documents', '*.docx')])

    if not docx_path:

        return None

    f = os.path.normpath(docx_path)

    ok = convert_docx_to_html(f)

    if ok:

        messagebox.showinfo('Hoàn tất', '✅ Đã chuyển xong!')

        return None

    messagebox.showerror('Lỗi', '❌ Hãy tắt file Word trước khi chuyển đổi!')





def gui_convert_word2html(root):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    convert_word2html_one()

    return None

