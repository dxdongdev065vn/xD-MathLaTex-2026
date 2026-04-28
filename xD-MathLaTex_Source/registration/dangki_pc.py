import os

import datetime

import wmi

import tkinter as tk

from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox



def get_first_hard_drive_serial():

    c = wmi.WMI()

    hard_drives = c.Win32_DiskDrive()

    if not hard_drives:

        return 'MOD111E2K08A89F'

    first_drive = hard_drives[0]

    serial_number = getattr(first_drive, 'SerialNumber', None)

# WARNING: Decompyle incomplete





def Ma_hoa_huy(text):

    pass

# WARNING: Decompyle incomplete





def Giai_ma_huy(text):

    pass

# WARNING: Decompyle incomplete





def Giai_ma_ngay_huy(text):

    pass

# WARNING: Decompyle incomplete





def chuoi_thanh_ngay(ngay_dung_thu, chuoi):

    '''Chuyển chuỗi định dạng ddMMyyyy thành datetime.date nếu hợp lệ'''

    if not isinstance(chuoi, str):

        return ngay_dung_thu

    if not None(chuoi) != 8 or chuoi.isdigit():

        return ngay_dung_thu

    ngay = int(chuoi[:2])

    thang = int(chuoi[2:4])

    nam = int(chuoi[4:])

    return datetime.date(nam, thang, ngay)

# WARNING: Decompyle incomplete





def write_keyuse(entry_Nhapdk1):

    keyuse = entry_Nhapdk1.get()

    key_kt = Giai_ma_huy(keyuse)

    sr = get_first_hard_drive_serial()

    folder_path = 'C:/MixEasy'

    if not os.path.exists(folder_path):

        os.makedirs(folder_path)

    file_path = os.path.join(folder_path, 'keyuse.ini')

# WARNING: Decompyle incomplete





def read_keyuse():

    file_path = 'C:/MixEasy/keyuse.ini'

# WARNING: Decompyle incomplete





def read_key_use_ngay_het_han(ngay_dung_thu):

    file_path = 'C:/MixEasy/keyuse.ini'

    ngay_het_han = ngay_dung_thu

# WARNING: Decompyle incomplete





def check_sr_and_key():

    sr = get_first_hard_drive_serial()

    key = read_keyuse()

    return sr == key





def check_trial_expiry(ngay_dung_thu):

    current_date = datetime.date.today()

    return current_date <= ngay_dung_thu





def check_het_han_dangki(ngay_het_han):

    current_date = datetime.date.today()

    return current_date <= ngay_het_han





def get_sr_for_dangki(entry_getsr):

    Kt_dangki = check_sr_and_key()

    if Kt_dangki:

        Seridangki = 'Bạn đã đăng kí phần mềm rồi!'

    else:

        Seridangki = get_first_hard_drive_serial()

    entry_getsr.delete(0, tk.END)

    entry_getsr.insert(0, Seridangki)





def giaodien_dangki(Quang_cao, formatted_date_het_han, root):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    Quang_cao = 'Tác giải: Nguyễn Nhật Huy.\nĐiện thoại, Zalo: 0914282232.\nGiá phần mềm: 50 nghìn/1 máy/1 năm.\nSố tài khoản: 0161 000 376 724\n Ngân hàng: Vietcombank.'

    ngay_het_han = datetime.date(2026, 6, 5)

    formatted_date_het_han = ngay_het_han.strftime('%d-%m-%Y')

    giaodien_dangki(Quang_cao, formatted_date_het_han, root)

    root.mainloop()

    return None

