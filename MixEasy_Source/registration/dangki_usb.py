import os

import sys

import datetime

import re

import wmi

import tkinter as tk

from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox



def get_running_drive_serial_USB_xoakituan():

    DEFAULT_SERIAL = 'MOD111E2K08A89F'

    exe_path = os.path.abspath(sys.argv[0])

    drive_letter = os.path.splitdrive(exe_path)[0]

    c = wmi.WMI()

    for logical_disk in c.Win32_LogicalDisk(DeviceID = drive_letter):

        for partition in logical_disk.associators('Win32_LogicalDiskToPartition'):

            for drive in partition.associators('Win32_DiskDriveToDiskPartition'):

                serial = getattr(drive, 'SerialNumber', None)

                if not serial:

                    continue

                cleaned = re.sub('[\\x00-\\x1F\\x7F]', '', serial)

                if len(cleaned) < 10:

                    cleaned += '0914282232'

                

                

                

                return c.Win32_LogicalDisk(DeviceID = drive_letter), logical_disk.associators('Win32_LogicalDiskToPartition'), partition.associators('Win32_DiskDriveToDiskPartition'), cleaned if cleaned else DEFAULT_SERIAL

    return DEFAULT_SERIAL

# WARNING: Decompyle incomplete





def get_first_hard_drive_serial():

    DEFAULT_SERIAL = 'MOD111E2K08A89F'

    exe_path = os.path.abspath(sys.argv[0])

    drive_letter = os.path.splitdrive(exe_path)[0]

    c = wmi.WMI()

    for logical_disk in c.Win32_LogicalDisk(DeviceID = drive_letter):

        for partition in logical_disk.associators('Win32_LogicalDiskToPartition'):

            for drive in partition.associators('Win32_DiskDriveToDiskPartition'):

                serial = getattr(drive, 'SerialNumber', None)

                if not serial:

                    continue

                cleaned = re.sub('[\\x00-\\x1F\\x7F]', '', serial)

                if len(cleaned) < 10:

                    cleaned += 'MOD111E2K08A89F'

                

                

                

                return c.Win32_LogicalDisk(DeviceID = drive_letter), logical_disk.associators('Win32_LogicalDiskToPartition'), partition.associators('Win32_DiskDriveToDiskPartition'), cleaned if cleaned else DEFAULT_SERIAL

    return DEFAULT_SERIAL

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

    full_path = os.path.abspath(sys.argv[0])

    app_path = os.path.splitdrive(full_path)[0]

    folder_path = os.path.join(app_path + os.sep, 'MixEasy_Dang_ki')

    if not os.path.exists(folder_path):

        os.makedirs(folder_path)

    file_path = os.path.join(folder_path, 'keyuse.ini')

# WARNING: Decompyle incomplete





def read_keyuse():

    full_path = os.path.abspath(sys.argv[0])

    app_path = os.path.splitdrive(full_path)[0]

    folder_path = os.path.join(app_path + os.sep, 'MixEasy_Dang_ki')

    file_path = os.path.join(folder_path, 'keyuse.ini')

# WARNING: Decompyle incomplete





def read_key_use_ngay_het_han(ngay_dung_thu):

    full_path = os.path.abspath(sys.argv[0])

    app_path = os.path.splitdrive(full_path)[0]

    folder_path = os.path.join(app_path + os.sep, 'MixEasy_Dang_ki')

    file_path = os.path.join(folder_path, 'keyuse.ini')

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

    ngay_dung_thu = datetime.date(2025, 9, 25)

    ngay_het_han = read_key_use_ngay_het_han(ngay_dung_thu)

    print(ngay_het_han)

    formatted_date_het_han = ngay_het_han.strftime('%d-%m-%Y')

    giaodien_dangki(Quang_cao, formatted_date_het_han, root)

    root.mainloop()

    return None

