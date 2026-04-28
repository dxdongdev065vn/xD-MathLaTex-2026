import tkinter as tk

from tkinter import messagebox

import re

import webbrowser



def open_link():

    link = 'https://olmocr.allenai.org'

    webbrowser.open_new(link)





def process_text(text):

    text = text.replace('\\n', '\n')

    text = text.replace('\\(', '$').replace('\\)', '$')

    text = text.replace('\\[', '$').replace('\\]', '$')

    text = text.replace('\\$', '$')

    text = text.replace('\\\\', '\\')

    text = re.sub('!\\[.*?\\]\\(.*?\\)', '', text)

    text = re.sub('\\n{2,}', '\n', text)

    return text.strip()





def convert(txt_input, txt_output):

    input_text = txt_input.get('1.0', tk.END)

    if not input_text.strip():

        messagebox.showwarning('Thiếu', 'Chưa nhập dữ liệu')

        return None

    output = process_text(input_text)

    txt_output.delete('1.0', tk.END)

    txt_output.insert(tk.END, output)





def copy_output(root, txt_output):

    text = txt_output.get('1.0', tk.END)

    root.clipboard_clear()

    root.clipboard_append(text)

    messagebox.showinfo('OK', 'Đã copy')





def gui_convert2pdf_olmocr(root):

    pass

# WARNING: Decompyle incomplete



