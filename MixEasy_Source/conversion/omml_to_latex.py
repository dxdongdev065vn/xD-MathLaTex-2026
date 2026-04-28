import tkinter as tk

from tkinter import filedialog, messagebox

import zipfile

import os

from lxml import etree

from officemath2latex import process_math_string

import re



ElementTree

(lambda omml_xml = None: latex = process_math_string(omml_xml)latex = re.sub('<', '< ', latex)latex = re.sub('>', '> ', latex)latex = re.sub('[\\\\\\s]+$', '', latex)latex = re.sub('\\\\mathbf', '', latex)latex = re.sub('', '', latex)latex = re.sub('', '', latex)latex = re.sub('', '', latex)latex = re.sub('"', "\\,''\\,", latex)latex = re.sub('“', "\\,''\\,", latex)latex = re.sub('\\\\in R', '\\\\in \\\\mathbb{R}', latex)latex = re.sub('\\\\in Z', '\\\\in \\\\mathbb{Z}', latex)latex = re.sub('\\\\in N', '\\\\in \\\\mathbb{N}', latex)latex = re.sub('\\\\([A-Z])(?![A-Za-z])', '\\\\backslash \\1', latex)latex = re.sub('\\\\\\\\left', '\\\\backslash \\\\left', latex)latex = re.sub('%', '\\%', latex)latex = re.sub('[\\\\\\s]+%', '\\%', latex)if not re.search('(\\\\lim(?![A-Za-z])|\\\\operatorname\\{lim\\})', latex) and re.search('\\\\displaystyle', latex):

latex = '\\displaystyle ' + latexif latex:

latex.strip()parser = etree.XMLParser(ns_clean = True, recover = True)omml_element = etree.fromstring(omml_xml.encode('utf-8'), parser = parser)ns_w = {

'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }text_element = omml_element.find('.//w:t', ns_w)# WARNING: Decompyle incomplete

) = None



def replace_equations_using_officemath2latex(docx_file, output_file):

    import zipfile

    etree = etree

    import lxml

    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    ns = {

        'w': W,

        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math' }

# WARNING: Decompyle incomplete





def choose_files(listbox_files):

    paths = filedialog.askopenfilenames(filetypes = [

        ('Word Documents', '*.docx')])

    if paths:

        listbox_files.delete(0, tk.END)

        for path in paths:

            listbox_files.insert(tk.END, path)

        return None





def convert_files(listbox_files, btn_convert, root):

    files = listbox_files.get(0, tk.END)

    if not files:

        messagebox.showerror('Lỗi', 'Vui lòng chọn ít nhất một file .docx!')

        return None

    failed = []

    btn_convert.config(state = tk.DISABLED)

    root.update_idletasks()

    for file_path in files:

        folder = os.path.dirname(file_path)

        filename = os.path.splitext(os.path.basename(file_path))[0]

        output_file = os.path.join(folder, f'''{filename}_latex.docx''')

        replace_equations_using_officemath2latex(file_path, output_file)

    btn_convert.config(state = tk.NORMAL)

    if failed:

        msg = 'Một số file không chuyển đổi được:\n' + '\n'.join(failed)

        messagebox.showwarning('Một số lỗi, kiểm tra xem đã tắt Word chưa', msg)

        return None

    messagebox.showinfo('Thành công', f'''Đã chuyển đổi thành công {len(files)} file.''')

    return None

# WARNING: Decompyle incomplete





def clear_list(listbox_files):

    listbox_files.delete(0, tk.END)





def gui_convert_OMML_tolatex(root):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    gui_convert_OMML_tolatex(root)

    root.mainloop()

    return None

