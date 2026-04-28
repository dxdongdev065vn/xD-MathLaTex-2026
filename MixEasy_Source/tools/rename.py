import os
import tkinter as tk
from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox
import unicodedata

def show_frame(frame):
    frame.tkraise()

selected_files = []
entry_pairs = []
result_labels = []

def thong_bao_hoan_thanh():
    messagebox.showinfo('ThÃ´ng bÃ¡o', 'HoÃ n thÃ nh')


def browse_files(frame1):
    for entry_pair in entry_pairs:
        for widget in entry_pair:
            widget.destroy()
        frame = entry_pair[0]
        frame.destroy()
    entry_pairs.clear()
    selected_files.clear()
    result_labels.clear()
    file_paths = filedialog.askopenfilenames()
    for file_path in file_paths:
        entry_pair = []
        framein1 = tk.Frame(frame1)
        framein1.pack()
        entry_pair.append(framein1)
        prefix_var = tk.StringVar()
        after_var = tk.StringVar()
        endfix_var = tk.StringVar()
        tag_var = tk.StringVar()
        (directory, filename) = os.path.split(file_path)
        (filename_no_extension, extension) = os.path.splitext(filename)
        selected_files.append({
            'prefix': prefix_var,
            'after': after_var,
            'endfix': endfix_var,
            'tag': tag_var,
            'path': file_path })
        entry1 = tk.Entry(framein1, textvariable = prefix_var, width = 15)
        entry1.pack(side = tk.LEFT, padx = (5, 5), pady = (0, 5))
        entry2 = tk.Entry(framein1, textvariable = after_var, width = 50)
        entry2.pack(side = tk.LEFT, padx = (5, 10), pady = (0, 5))
        entry2.insert(0, filename_no_extension)
        entry3 = tk.Entry(framein1, textvariable = endfix_var, width = 5)
        entry3.pack(side = tk.LEFT, padx = (5, 5), pady = (0, 5))
        (_, file_extension) = os.path.splitext(file_path)
        tag_var.set(file_extension)
        entry4 = tk.Entry(framein1, textvariable = tag_var, state = 'readonly', width = 5)
        entry4.pack(side = tk.LEFT, padx = (5, 10), pady = (0, 5))
        entry_pair.extend([
            entry1,
            entry2,
            entry3,
            entry4])
        entry_pairs.append(entry_pair)
        result_label = tk.Label(framein1, text = '')
        result_label.pack(side = tk.LEFT, padx = (5, 10), pady = (0, 5))
        result_labels.append(result_label)


def rename_files():
    for i, file_info in enumerate(selected_files):
        prefix = file_info['prefix'].get()
        after = file_info['after'].get()
        endfix = file_info['endfix'].get()
        tag = file_info['tag'].get()
        file_path = file_info['path']
        if not prefix and after and endfix:
            continue
        if not file_path:
            continue
        (directory, filename) = os.path.split(file_path)
        new_filename = os.path.join(directory, prefix + after + endfix + tag)
        os.rename(file_path, new_filename)
        result_labels[i].config(text = 'xong')
    return None
# WARNING: Decompyle incomplete


def clear_entries():
    for entry_pair in entry_pairs:
        for widget in entry_pair:
            widget.destroy()
        frame = entry_pair[0]
        frame.destroy()
    entry_pairs.clear()
    selected_files.clear()
    result_labels.clear()


def select_files(file_list2):
    files = filedialog.askopenfilenames()
    file_list2.delete(0, tk.END)
    for file in files:
        file_list2.insert(tk.END, file)


def replace_in_file_names(files, replace_str, replace_with):
    for file in files:
        file_name = os.path.basename(file)
        new_file_name = file_name.replace(replace_str, replace_with)
        os.rename(file, os.path.join(os.path.dirname(file), new_file_name))


def replace_names(file_list2, replace_entry2, replace_with_entry2):
    files = file_list2.get(0, tk.END)
    replace_str = replace_entry2.get()
    replace_with = replace_with_entry2.get()
    replace_in_file_names(files, replace_str, replace_with)
    thong_bao_hoan_thanh()


def clear_listbox_replace(file_list2):
    file_list2.delete(0, tk.END)


def remove_vietnamese_accents(input_str):
    normalized_str = unicodedata.normalize('NFD', input_str)
# WARNING: Decompyle incomplete


def remove_vietnamese_accents_from_files(file_paths):
    for file_path in file_paths:
        (file_dir, file_name) = os.path.split(file_path)
        new_file_name = remove_vietnamese_accents(file_name)
        new_file_path = os.path.join(file_dir, new_file_name)
        os.rename(file_path, new_file_path)


def remove_vietnamese_accents_from_directory(dir_path):
    new_dir_name = remove_vietnamese_accents(os.path.basename(dir_path))
    new_dir_path = os.path.join(os.path.dirname(dir_path), new_dir_name)
    os.rename(dir_path, new_dir_path)


def select_files_3(selected_files_listbox_3):
    file_paths = filedialog.askopenfilenames()
    if file_paths:
        selected_files_paths = file_paths
        selected_files_listbox_3.delete(0, 'end')
        for file_path in file_paths:
            selected_files_listbox_3.insert('end', file_path)
        return None


def xoa_dau_tieng_viet_file(selected_files_listbox_3):
    selected_files = selected_files_listbox_3.get(0, 'end')
    if selected_files:
        remove_vietnamese_accents_from_files(selected_files)
        thong_bao_hoan_thanh()
        return None


def select_directory(selected_directory_entry3):
    selected_dir = filedialog.askdirectory()
    if selected_dir:
        selected_directory_entry3.delete(0, 'end')
        selected_directory_entry3.insert(0, selected_dir)
        return None


def xoa_dau_tieng_viet_folder(selected_directory_entry3):
    selected_dir = selected_directory_entry3.get()
    if selected_dir:
        remove_vietnamese_accents_from_directory(selected_dir)
        thong_bao_hoan_thanh()
        return None


def gui_rename(root):
    pass
# WARNING: Decompyle incomplete

if __name__ == '__main__':
    root = tk.Tk()
    gui_rename(root)
    root.mainloop()
    return None
