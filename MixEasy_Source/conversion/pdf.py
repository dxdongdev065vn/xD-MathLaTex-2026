import os

import tkinter as tk

from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox

from tkinter import simpledialog

import win32com.client as win32com

from core import functions as vbf



def save_as_pdf(word):

    doc = word.ActiveDocument

    if doc.Path == '':

        messagebox.showwarning('Chưa lưu tài liệu', '⚠️ Vui lòng lưu tài liệu trước khi chuyển sang PDF!')

        return None

    input_path = doc.FullName

    input_path = vbf.chuan_hoa_duong_dan(input_path)

    base_name = os.path.splitext(os.path.basename(input_path))[0]

    output_path = os.path.join(doc.Path, base_name + '.pdf')

    if hasattr(doc, 'SaveAs2'):

        doc.SaveAs2(output_path, FileFormat = 17)

    else:

        doc.SaveAs(output_path, FileFormat = 17)

    messagebox.showinfo('Hoàn tất', f'''✅ Đã chuyển sang PDF thành công!\n{output_path}''')

    return None

# WARNING: Decompyle incomplete





def print_microsoft_pdf(word):

    doc = word.ActiveDocument

    doc = word.ActiveDocument

    if doc.Path == '':

        messagebox.showwarning('Chưa lưu tài liệu', '⚠️ Vui lòng lưu tài liệu trước khi chuyển sang PDF!')

        return None

    doc_path = doc.FullName

    doc_path = vbf.chuan_hoa_duong_dan(doc_path)

    pdf_path = os.path.splitext(doc_path)[0] + '.pdf'

    word.ActivePrinter = 'Microsoft Print to PDF'

    doc.PrintOut(OutputFileName = pdf_path, Background = False)

    messagebox.showinfo('Hoàn tất', '✅ Đã chuyển sang PDF thành công, file lưu cùng thư mục file word')

    return None

# WARNING: Decompyle incomplete





def convert_to_pdf(input_path, word):

    word.Visible = False

    word.DisplayAlerts = False

    input_path = vbf.chuan_hoa_duong_dan(input_path)

    doc = word.Documents.Open(input_path)

    output_path = os.path.splitext(input_path)[0] + '.pdf'

    doc.SaveAs(output_path, FileFormat = 17)

    doc.Close()

    return None

# WARNING: Decompyle incomplete





def print_to_pdf_word(doc_path, word):

    word.Visible = False

    word.DisplayAlerts = False

    doc_path = vbf.chuan_hoa_duong_dan(doc_path)

    pdf_path = os.path.splitext(doc_path)[0] + '.pdf'

    doc = word.Documents.Open(doc_path)

    word.ActivePrinter = 'Microsoft Print to PDF'

    doc.PrintOut(OutputFileName = pdf_path, Background = False)

    doc.Close()

    print('OK:', pdf_path)

    return None

# WARNING: Decompyle incomplete





def select_files(listbox_pdf):

    files = filedialog.askopenfilenames(filetypes = [

        ('Word Files', '*.doc *.docx')])

    for file in files:

        listbox_pdf.insert(tk.END, file)





def clear_files(listbox_pdf):

    listbox_pdf.delete(0, tk.END)





def convert_to_pdf_batch(word, listbox_pdf):

    word = win32com.client.DispatchEx('Word.Application')

    word.Visible = False

    word.DisplayAlerts = False

    if word.Documents.Count > 0:

        messagebox.showwarning('Word đang mở file', 'Hiện đang có file Word được mở.\nVui lòng đóng tất cả file Word rồi thử lại.')

        return None

    files = listbox_pdf.get(0, tk.END)

    if not files:

        messagebox.showwarning('Chưa có file', 'Danh sách chuyển đổi đang trống.\nVui lòng thêm file Word trước.')

        return None

    for file in files:

        convert_to_pdf(file, word)

    word.Quit()

    messagebox.showinfo('Thông báo', 'Đã hoàn thành việc chuyển các file sang pdf')

    return None

# WARNING: Decompyle incomplete





def run_print_pdf(word, listbox_pdf):

    word = win32com.client.DispatchEx('Word.Application')

    word.Visible = False

    word.DisplayAlerts = False

    if word.Documents.Count > 0:

        messagebox.showwarning('Word đang mở file', 'Hiện đang có file Word được mở.\nVui lòng đóng tất cả file Word rồi thử lại.')

        return None

    files = listbox_pdf.get(0, tk.END)

    if not files:

        messagebox.showwarning('Chưa có file', 'Danh sách chuyển đổi đang trống.\nVui lòng thêm file Word trước.')

        return None

    for file in files:

        print(f'''Đang xử lý: {file}''')

        print_to_pdf_word(file, word)

    word.Quit()

    messagebox.showinfo('Thông báo', 'Đã hoàn thành việc chuyển các file sang pdf')

    return None

# WARNING: Decompyle incomplete





def update_textbox(file_paths):

    listbox_pdf.configure(state = 'normal')

    listbox_pdf.delete('1.0', 'end')

    for path in file_paths:

        listbox_pdf.insert('end', path + '\n')

    listbox_pdf.configure(state = 'disabled')





def select_files_ctk(listbox_pdf):

    files = filedialog.askopenfilenames(filetypes = [

        ('Word Files', '*.doc *.docx')])

    listbox_pdf.configure(state = 'normal')

    listbox_pdf.delete('1.0', 'end')

    for path in files:

        listbox_pdf.insert('end', path + '\n')

    listbox_pdf.configure(state = 'disabled')





def clear_files_ctk(listbox_pdf, success_label_pdf):

    listbox_pdf.configure(state = 'normal')

    listbox_pdf.delete('1.0', 'end')

    listbox_pdf.configure(state = 'disabled')

    success_label_pdf.config(text = '')





def convert_to_pdf_batch_ctk(word, listbox_pdf, success_label_pdf):

    files = listbox_pdf.get('1.0', 'end').strip().splitlines()

    word = vbf.khoi_tao_word_2()

    for file in files:

        converted_path = os.path.normpath(file)

        convert_to_pdf(converted_path, word, success_label_pdf)



