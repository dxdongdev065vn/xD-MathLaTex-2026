import os

import datetime

import win32com.client as win32com

from win32com.client import client as win32

import win32api

import re

import sys

import shutil

import pythoncom

import tempfile

import wmi

import random

import tkinter as tk

from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox

from tkinter import simpledialog

import webbrowser

from core import functions as vbf

from formatting import chuan_hoa as chd

from tools import tool_by_docx as docxtool

from mixing import mix_docx_check_data as docxcheck

from mixing import mixeasy_file_mau as fmau



def tim_phan_1(doc, word, find_range):

    myrange = find_range.Duplicate

    find = myrange.Find

    find.ClearFormatting()

    find.Text = '(S1\\@)(*)(E1\\@)'

    find.Replacement.Text = ''

    find.Forward = True

    find.Wrap = 0

    find.MatchWildcards = True

    if find.Execute():

        return myrange





def tim_phan_2(doc, word, find_range):

    myrange = find_range.Duplicate

    find = myrange.Find

    find.ClearFormatting()

    find.Text = '(S2\\@)(*)(E2\\@)'

    find.Replacement.Text = ''

    find.Forward = True

    find.Wrap = 0

    find.MatchWildcards = True

    if find.Execute():

        return myrange





def tim_phan_3(doc, word, find_range):

    myrange = find_range.Duplicate

    find = myrange.Find

    find.ClearFormatting()

    find.Text = '(S3\\@)(*)(E3\\@)'

    find.Replacement.Text = ''

    find.Forward = True

    find.Wrap = 0

    find.MatchWildcards = True

    if find.Execute():

        return myrange





def tim_phan_4(doc, word, find_range):

    myrange = find_range.Duplicate

    find = myrange.Find

    find.ClearFormatting()

    find.Text = '(S4\\@)(*)(E4\\@)'

    find.Replacement.Text = ''

    find.Forward = True

    find.Wrap = 0

    find.MatchWildcards = True

    if find.Execute():

        return myrange





def fix_table_me(word):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def Dap_so_to_Black_Bold(word):

    key = 'ĐS:'

    doc = word.ActiveDocument

    myrange = doc.Content

    for i in range(1, 100):

        myrange_find = myrange.Duplicate

        find = myrange_find.Find

        find.ClearFormatting()

        find.Text = key

        find.MatchWildcards = False

        if find.Execute():

            current_line = myrange_find.Paragraphs(1).Range

            current_line.Font.Color = win32api.RGB(255, 0, 0)

            current_line.Font.Bold = True

            myrange.SetRange(Start = myrange_find.End, End = myrange.End)

            if not myrange_find.Tables.Count > 0:

                continue

            table = myrange_find.Tables(1)

            table_range = table.Range

            table_end = table_range.End

            myrange.SetRange(Start = table_end, End = myrange.End)

            continue

        range(1, 100)

        return None





def xuongdong_phuongan_Me(word):

    doc = word.ActiveDocument

    messages_sentences = []

    dem_so = vbf.check_pic_float(doc, word)

    if dem_so > 0:

        messages_sentences.append(f'''Có {dem_so} hình ảnh không ở chế độ pict in line.\n Tôi đã giúp bạn về dạng pict in line nhưng bạn phải kiểm tra lại vị trí xuất hiện của nó để chỉnh lại phù hợp''')

    vbf.pic_inline_center(word)

    chd.xuongdong_phuongan(word)

    fix_table_me(word)

    vbf.check_sentences_in_tables(word, messages_sentences)

    keywords = [

        'Đáp án:',

        'ĐA:',

        'Đáp số:',

        'TL:',

        'Trả lời:']

    for keyword in keywords:

        find = word.Selection.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = keyword

        find.Replacement.Text = 'ĐS:'

        find.MatchCase = False

        find.MatchWildcards = False

        find.Wrap = 1

        find.Format = False

        find.Execute(Replace = 2)

    vbf.thay_the_replace(word, '(ĐS:)([ ]{1,})', '\\1')

    vbf.thay_the_replace(word, '(ĐS:)([0-9]{1,})(.)([0-9]{1,})', '\\1\\2,\\4')

    vbf.thay_the_replace(word, '(ĐS:)([0-9]{1,})(,)([0-9]{1,})(.)', '\\1\\2\\3\\4')

    vbf.thay_the_replace(word, '(ĐS:)([0-9]{1,})(.)', '\\1\\2')

    Dap_so_to_Black_Bold(word)

    messages_sentences.append('Đã hoàn thành')

    if messages_sentences:

        thong_bao = '\n'.join(messages_sentences)

        messagebox.showinfo('Thông báo', thong_bao)

        return None

    return None

# WARNING: Decompyle incomplete





def xuongdong_phuongan_lite_Me(word):

    doc = word.ActiveDocument

    messages_sentences = []

    dem_so = vbf.check_pic_float(doc, word)

    if dem_so > 0:

        messages_sentences.append(f'''Có {dem_so} hình ảnh không ở chế độ pict in line.\n Tôi đã giúp bạn về dạng pict in line nhưng bạn phải kiểm tra lại vị trí xuất hiện của nó để chỉnh lại phù hợp''')

    vbf.pic_inline_center(word)

    chd.xuongdong_phuongan_lite(word)

    vbf.check_sentences_in_tables(word, messages_sentences)

    messages_sentences.append('Đã hoàn thành')

    if messages_sentences:

        thong_bao = '\n'.join(messages_sentences)

        messagebox.showinfo('Thông báo', thong_bao)

        return None

    return None

# WARNING: Decompyle incomplete





def page_A4_setup_Mix(word):

    doc = word.ActiveDocument

    doc.Content.WholeStory()

    doc.Range().Select()

    page_setup = doc.PageSetup

    page_setup.TopMargin = vbf.InchesToPoints(0.4)

    page_setup.BottomMargin = vbf.InchesToPoints(0.4)

    page_setup.LeftMargin = vbf.InchesToPoints(0.4)

    page_setup.RightMargin = vbf.InchesToPoints(0.4)

    page_setup.Gutter = vbf.InchesToPoints(0)

    page_setup.HeaderDistance = vbf.InchesToPoints(0.24)

    page_setup.FooterDistance = vbf.InchesToPoints(0.24)

    page_setup.PageWidth = vbf.InchesToPoints(8.27)

    page_setup.PageHeight = vbf.InchesToPoints(11.69)

    return None

# WARNING: Decompyle incomplete





def chen_chan_trang(word):

    doc = word.ActiveDocument

    constants = win32.constants

    vbf.thay_the_replace(word, '^m', '')

    sec = doc.Sections(1)

    sec.PageSetup.DifferentFirstPageHeaderFooter = False

    sec.PageSetup.OddAndEvenPagesHeaderFooter = False

    footer = sec.Footers(constants.wdHeaderFooterPrimary)

    footer_range = footer.Range

    footer_range.Text = ''

    base_text = 'Trang 1/2 - mã đề thi <made>'

    footer_range.InsertAfter(base_text)

    txt = footer_range.Text

    idx = txt.find('1/2')

    if idx == -1:

        pos1 = txt.find('1')

        pos2 = txt.find('2', pos1 + 1) if pos1 >= 0 else -1

    else:

        pos1 = idx

        pos2 = idx + 2

    if pos1 < 0 or pos2 < 0:

        raise RuntimeError("Không tìm thấy mẫu '1/2' trong footer.")

    rng_nump = footer_range.Duplicate

    rng_nump.Start = footer_range.Start + pos2

    rng_nump.End = rng_nump.Start + 1

    rng_nump.Fields.Add(rng_nump, constants.wdFieldEmpty, 'NUMPAGES \\* Arabic', True)

    rng_page = footer_range.Duplicate

    rng_page.Start = footer_range.Start + pos1

    rng_page.End = rng_page.Start + 1

    rng_page.Fields.Add(rng_page, constants.wdFieldEmpty, 'PAGE \\* Arabic', True)

    fr = footer.Range

    fr.Font.Italic = True

    fr.Font.Name = 'Times New Roman'

    fr.Font.Size = 11

    fr.ParagraphFormat.Alignment = constants.wdAlignParagraphRight

    top = fr.ParagraphFormat.Borders(constants.wdBorderTop)

    top.LineStyle = constants.wdLineStyleSingle

    top.LineWidth = constants.wdLineWidth050pt

    top.Color = constants.wdColorAutomatic

    fr.ParagraphFormat.Borders.DistanceFromTop = 3

    for side in (constants.wdBorderLeft, constants.wdBorderRight, constants.wdBorderBottom):

        fr.ParagraphFormat.Borders(side).LineStyle = constants.wdLineStyleNone

    doc.Fields.Update()

    return None

# WARNING: Decompyle incomplete





def table_tieu_de(word):

    doc = word.ActiveDocument

    constants = win32.constants

    page_A4_setup_Mix(word)

    if doc.Tables.Count > 0:

        first_table = doc.Tables(1)

        if first_table.Range.Start == 0:

            first_table.Rows(1).Range.Select()

            word.Selection.SplitTable()

    rng = doc.Range(0, 0)

    rng.Collapse(constants.wdCollapseStart)

    table1 = doc.Tables.Add(rng, NumRows = 1, NumColumns = 2)

    table1.Borders.Enable = 0

    table1.AllowAutoFit = False

    table1.Columns(1).Width = vbf.InchesToPoints(3.75)

    table1.Columns(2).Width = vbf.InchesToPoints(3.75)

    c1 = table1.Cell(1, 1).Range

    c1.Text = 'TRƯỜNG THPT CHUYÊN QUỐC HỌC – HUẾ\rTỔ TOÁN\rĐỀ CHÍNH THỨC\r(Đề thi có <sotrang> trang)'

    c1.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter

    c1.ParagraphFormat.SpaceBefore = 0

    c1.ParagraphFormat.SpaceAfter = 0

    c1.Font.Name = 'Times New Roman'

    c1.Font.Size = 12

    c1.Font.Color = win32api.RGB(0, 0, 255)

    c1.Font.Bold = True

    para_last = c1.Paragraphs.Last

    para_last.Range.Font.Italic = True

    para_last.Range.Font.Bold = False

    c2 = table1.Cell(1, 2).Range

    c2.Text = 'ĐỀ KIỂM TRA GIỮA KỲ II\rNĂM HỌC 2025 – 2026\rMôn: Toán – Lớp 10\rThời gian: 90 phút (Không kể thời gian phát đề)'

    c2.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter

    c2.ParagraphFormat.SpaceBefore = 0

    c2.ParagraphFormat.SpaceAfter = 0

    c2.Font.Name = 'Times New Roman'

    c2.Font.Size = 12

    c2.Font.Color = win32api.RGB(0, 0, 255)

    c2.Font.Bold = True

    for para in c2.Paragraphs:

        if not 'Môn:' in para.Range.Text and 'Thời gian:' in para.Range.Text:

            continue

        para.Range.Font.Italic = True

        para.Range.Font.Bold = False

    rng_end = table1.Range

    rng_end.Collapse(constants.wdCollapseEnd)

    rng_end.InsertParagraphAfter()

    new_para = rng_end.Paragraphs(1).Range

    new_para.Font.Name = 'Times New Roman'

    new_para.Font.Size = 5

    new_para.ParagraphFormat.SpaceBefore = 0

    new_para.ParagraphFormat.SpaceAfter = 0

    rng_end = doc.Range(table1.Range.End + 1, table1.Range.End + 1)

    table2 = doc.Tables.Add(rng_end, NumRows = 1, NumColumns = 2)

    table2.Borders.Enable = 0

    table2.AllowAutoFit = False

    table2.Columns(1).Width = vbf.InchesToPoints(5.7)

    table2.Columns(2).Width = vbf.InchesToPoints(1.6)

    c21 = table2.Cell(1, 1).Range

    c21.Text = 'Họ và tên học sinh: ..................................................   Số báo danh: .......................'

    c21.Font.Name = 'Times New Roman'

    c21.ParagraphFormat.SpaceBefore = 0

    c21.ParagraphFormat.SpaceAfter = 0

    c21.Font.Bold = False

    c21.Font.Size = 12

    c22 = table2.Cell(1, 2).Range

    c22.Text = 'Mã đề thi: <made>'

    c22.Font.Name = 'Times New Roman'

    c22.Font.Size = 12

    c22.Font.Bold = True

    c22.ParagraphFormat.Alignment = constants.wdAlignParagraphCenter

    c22.ParagraphFormat.SpaceBefore = 3

    c22.ParagraphFormat.SpaceAfter = 3

    for border_type in (constants.wdBorderTop, constants.wdBorderBottom, constants.wdBorderLeft, constants.wdBorderRight):

        border = table2.Cell(1, 2).Borders(border_type)

        border.LineStyle = constants.wdLineStyleSingle

        border.LineWidth = constants.wdLineWidth150pt

        border.Color = constants.wdColorAutomatic

    return None

# WARNING: Decompyle incomplete





def green_bold_left_a_cong(word):

    doc = word.ActiveDocument

    myrange = doc.Range()

    for i in range(1, 20):

        find = myrange.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = '@'

        find.MatchWildcards = False

        find.MatchCase = True

        find.Forward = True

        if find.Execute():

            para = myrange.Paragraphs(1)

            para.Range.Font.Bold = True

            para.Range.Font.Color = win32api.RGB(0, 128, 0)

            para.Alignment = 3

            myrange.SetRange(myrange.End, doc.Range().End)

            continue

        range(1, 20)

        return None

    return None

# WARNING: Decompyle incomplete





def add_ki_hieu_nhan_dien(word):

    doc = word.ActiveDocument

    vbf.thay_the_replace(word, '(^13)([^9 ]{1,})(PHẦN)', '\\1\\3')

    vbf.thay_the_replace(word, '(^13PHẦN 1.)', '^13PHẦN I.')

    vbf.thay_the_replace(word, '(^13PHẦN 2.)', '^13PHẦN II.')

    vbf.thay_the_replace(word, '(^13PHẦN 3.)', '^13PHẦN III.')

    vbf.thay_the_replace(word, '(^13PHẦN 4.)', '^13PHẦN IV.')

    vbf.STT_2025_new(word)

    vbf.thay_the_replace_1(word, '(^13PHẦN II.)', '^13E1@\\1')

    vbf.thay_the_replace_1(word, '(^13PHẦN III.)', '^13E2@\\1')

    vbf.thay_the_replace_1(word, '(^13PHẦN IV.)', '^13E3@\\1')

    myrange = doc.Range()

    for i in range(1, 3):

        find = myrange.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = '(^13Câu 1[.:])'

        find.Replacement.Text = f'''^13S{i}@\\1'''

        find.MatchWildcards = True

        find.MatchCase = True

        find.Forward = True

        find.Format = True

        if not find.Execute(Replace = 1):

            continue

        myrange = doc.Range(myrange.End, doc.Range().End)

    rng_3_4 = myrange.Duplicate

    find = rng_3_4.Find

    find.ClearFormatting()

    find.Replacement.ClearFormatting()

    find.Text = '(^13PHẦN IV.)'

    find.MatchWildcards = True

    find.MatchCase = True

    find.Forward = True

    find.Format = True

    if find.Execute():

        for i in range(3, 5):

            find = myrange.Find

            find.ClearFormatting()

            find.Replacement.ClearFormatting()

            find.Text = '(^13Câu 1[.:])'

            find.Replacement.Text = f'''^13S{i}@\\1'''

            find.MatchWildcards = True

            find.MatchCase = True

            find.Forward = True

            find.Format = True

            if not find.Execute(Replace = 1):

                continue

            myrange = doc.Range(myrange.End, doc.Range().End)

        find = myrange.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = 'HẾT'

        find.MatchWildcards = False

        find.MatchCase = True

        find.Forward = True

        find.Format = True

        if find.Execute():

            para = myrange.Paragraphs(1)

            para.Range.InsertParagraphBefore()

            new_para = para.Previous()

            new_para.Range.Text = 'E4@\r'

        else:

            selection = word.Selection

            selection.EndKey(Unit = 6)

            word.Selection.TypeParagraph()

            word.Selection.TypeText('E4@')

            word.Selection.TypeParagraph()

    else:

        rng_3_4 = myrange.Duplicate

        find = rng_3_4.Find

        find.ClearFormatting()

        find.Replacement.ClearFormatting()

        find.Text = '(^13PHẦN III.)'

        find.MatchWildcards = True

        find.MatchCase = True

        find.Forward = True

        find.Format = True

        if find.Execute():

            para = rng_3_4.Paragraphs(2)

            text = para.Range.Text.lower()

            if 'tự luận' in text:

                myrange = doc.Range(rng_3_4.End, doc.Range().End)

                find = myrange.Find

                find.ClearFormatting()

                find.Replacement.ClearFormatting()

                find.Text = '(^13Câu 1[.:])'

                find.Replacement.Text = '^13S4@\\1'

                find.MatchWildcards = True

                find.MatchCase = True

                find.Forward = True

                find.Format = True

                if find.Execute(Replace = 1):

                    myrange = doc.Range(myrange.End, doc.Range().End)

                find = myrange.Find

                find.ClearFormatting()

                find.Replacement.ClearFormatting()

                find.Text = 'HẾT'

                find.MatchWildcards = False

                find.MatchCase = True

                find.Forward = True

                find.Format = True

                if find.Execute():

                    para = myrange.Paragraphs(1)

                    rng_before = doc.Range(Start = para.Range.Start, End = para.Range.Start)

                    rng_before.InsertParagraphBefore()

                    rng_new = para.Previous().Range

                    rng_new.Text = 'E4@\r'

                else:

                    selection = word.Selection

                    selection.EndKey(Unit = 6)

                    word.Selection.TypeParagraph()

                    word.Selection.TypeText('E4@')

                    word.Selection.TypeParagraph()

            else:

                find = myrange.Find

                find.ClearFormatting()

                find.Replacement.ClearFormatting()

                find.Text = '(^13Câu 1[.:])'

                find.MatchWildcards = True

                find.Replacement.Text = '^13S3@\\1'

                find.MatchWildcards = True

                find.MatchCase = True

                find.Forward = True

                find.Format = True

                if find.Execute(Replace = 1):

                    myrange = doc.Range(myrange.End, doc.Range().End)

                find = myrange.Find

                find.ClearFormatting()

                find.Replacement.ClearFormatting()

                find.Text = 'HẾT'

                find.MatchWildcards = False

                find.MatchCase = True

                find.Forward = True

                find.Format = True

                if find.Execute():

                    para = myrange.Paragraphs(1)

                    rng_before = doc.Range(Start = para.Range.Start, End = para.Range.Start)

                    rng_before.InsertParagraphBefore()

                    rng_new = para.Previous().Range

                    rng_new.Text = 'E3@\r'

                else:

                    selection = word.Selection

                    selection.EndKey(Unit = 6)

                    word.Selection.TypeParagraph()

                    word.Selection.TypeText('E3@\r')

                    word.Selection.TypeParagraph()

    green_bold_left_a_cong(word)

    return None

# WARNING: Decompyle incomplete





def type_text_xanh_blue(word, text_type, dam_nhat, underline = (False, False)):

    doc = word.ActiveDocument

    word.Selection.Font.Bold = dam_nhat

    word.Selection.Font.Italic = False

    word.Selection.Font.Name = 'Times New Roman'

    word.Selection.Font.Size = 12

    word.Selection.Font.Color = win32api.RGB(0, 0, 255)

    word.Selection.ParagraphFormat.SpaceBefore = 8

    word.Selection.ParagraphFormat.SpaceAfter = 0

    word.Selection.ParagraphFormat.LineSpacing = vbf.LinesToPoints(1.15)

    word.Selection.ParagraphFormat.Alignment = 3

    word.Selection.TypeText(text_type)

    return None

# WARNING: Decompyle incomplete





def type_text(word, text_type, dam_nhat, underline = (False, False)):

    doc = word.ActiveDocument

    word.Selection.Font.Bold = dam_nhat

    word.Selection.Font.Underline = 1 if underline else 0

    word.Selection.Font.Italic = False

    word.Selection.Font.Name = 'Times New Roman'

    word.Selection.Font.Size = 12

    word.Selection.Font.Color = win32api.RGB(0, 0, 0)

    word.Selection.ParagraphFormat.SpaceBefore = 0

    word.Selection.ParagraphFormat.SpaceAfter = 0

    word.Selection.ParagraphFormat.LineSpacing = vbf.LinesToPoints(1.15)

    word.Selection.ParagraphFormat.Alignment = 3

    word.Selection.TypeText(text_type)

    return None

# WARNING: Decompyle incomplete





def mau_tieu_de_phan_I(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    text1 = 'PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.'

    text2 = ' Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án.\r'

    type_text_xanh_blue(word, text1, dam_nhat = True, underline = False)

    type_text_xanh_blue(word, text2, dam_nhat = False, underline = False)

    return None

# WARNING: Decompyle incomplete





def mau_tieu_de_phan_II(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    text1 = 'PHẦN II. Câu trắc nghiệm đúng sai.'

    text2 = ' Thí sinh trả lời từ câu 1 đến câu 3. Trong mỗi ý '

    text3 = 'a) , b) , c) , d)'

    text4 = ' ở mỗi câu, thí sinh chọn đúng hoặc sai.\r'

    type_text_xanh_blue(word, text1, dam_nhat = True, underline = False)

    type_text_xanh_blue(word, text2, dam_nhat = False, underline = False)

    type_text_xanh_blue(word, text3, dam_nhat = True, underline = False)

    type_text_xanh_blue(word, text4, dam_nhat = False, underline = False)

    return None

# WARNING: Decompyle incomplete





def mau_tieu_de_phan_III(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    text1 = 'PHẦN III. Câu trắc nghiệm trả lời ngắn. '

    text2 = ' Thí sinh trả lời từ câu 1 đến câu 4.\r'

    type_text_xanh_blue(word, text1, dam_nhat = True, underline = False)

    type_text_xanh_blue(word, text2, dam_nhat = False, underline = False)

    return None

# WARNING: Decompyle incomplete





def mau_tieu_de_phan_IV(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    text1 = 'PHẦN IV. Câu hỏi tự luận.\r'

    type_text_xanh_blue(word, text1, dam_nhat = True, underline = False)

    return None

# WARNING: Decompyle incomplete





def mau_cau_hoi_ABCD(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    type_text(word, 'Câu 1.', dam_nhat = True, underline = False)

    type_text(word, ' Chọn đáp án đúng.\r', dam_nhat = False, underline = False)

    type_text(word, 'A.', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 1.\r', dam_nhat = False, underline = False)

    type_text(word, 'B', dam_nhat = True, underline = True)

    type_text(word, '.', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 2.\r', dam_nhat = False, underline = False)

    type_text(word, 'C.', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 3.\r', dam_nhat = False, underline = False)

    type_text(word, 'D.', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 4.\r', dam_nhat = False, underline = False)

    return None

# WARNING: Decompyle incomplete





def mau_cau_hoi_abcd(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    type_text(word, 'Câu 1.', dam_nhat = True, underline = False)

    type_text(word, ' Chọn đúng, sai.\r', dam_nhat = False, underline = False)

    type_text(word, 'a', dam_nhat = True, underline = True)

    type_text(word, ')', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 1.\r', dam_nhat = False, underline = False)

    type_text(word, 'b)', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 2.\r', dam_nhat = False, underline = False)

    type_text(word, 'c', dam_nhat = True, underline = True)

    type_text(word, ')', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 3.\r', dam_nhat = False, underline = False)

    type_text(word, 'd)', dam_nhat = True, underline = False)

    type_text(word, ' Đáp án 4.\r', dam_nhat = False, underline = False)

    return None

# WARNING: Decompyle incomplete





def end_de_het(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        messagebox.showinfo('Thông báo', 'Bạn phải đưa con trỏ (dấu nháy) vào vị trí cần đánh')

        return None

    word.Selection.Font.Bold = True

    word.Selection.Font.Italic = False

    word.Selection.Font.Name = 'Times New Roman'

    word.Selection.Font.Size = 12

    word.Selection.Font.Color = win32api.RGB(0, 0, 0)

    word.Selection.ParagraphFormat.SpaceBefore = 8

    word.Selection.ParagraphFormat.SpaceAfter = 0

    word.Selection.ParagraphFormat.LineSpacing = vbf.LinesToPoints(1.15)

    word.Selection.ParagraphFormat.Alignment = 1

    word.Selection.TypeText('------------ HẾT ------------\r')

    word.Selection.Font.Bold = False

    word.Selection.Font.Italic = True

    word.Selection.Font.Name = 'Times New Roman'

    word.Selection.Font.Size = 12

    word.Selection.Font.Color = win32api.RGB(0, 0, 0)

    word.Selection.ParagraphFormat.SpaceBefore = 0

    word.Selection.ParagraphFormat.SpaceAfter = 0

    word.Selection.ParagraphFormat.LineSpacing = vbf.LinesToPoints(1.15)

    word.Selection.ParagraphFormat.Alignment = 3

    word.Selection.ParagraphFormat.FirstLineIndent = vbf.InchesToPoints(0.5)

    word.Selection.TypeText('- Thí sinh không được sử dụng tài liệu;\r- Cán bộ coi thi không giải thích gì thêm.')

    return None

# WARNING: Decompyle incomplete





def chuc_nang_khac_mix(root, word, addin_name):

    pass

# WARNING: Decompyle incomplete





def chuc_nang_khac_EN_mix(root, word, addin_name):

    pass

# WARNING: Decompyle incomplete





def check_ki_hieu(word, messages, text_start, text_end):

    doc = word.ActiveDocument

    indices_S = []

    indices_E = []

    for i, paragraph in enumerate(doc.Paragraphs, start = 1):

        text = paragraph.Range.Text

        if text.startswith(text_start):

            indices_S.append(i)

            continue

        if not text.startswith(text_end):

            continue

        indices_E.append(i)

    if len(indices_S) != len(indices_E):

        messages.append(f'''Số kí hiệu {text_start} và {text_end} không bằng nhau- KHÔNG hợp lệ''')

        return None

    return None

# WARNING: Decompyle incomplete





def check_file_mau(word, messages):

    pass

# WARNING: Decompyle incomplete





def check_file_mau_EN(word, messages):

    pass

# WARNING: Decompyle incomplete





def check_file_mau(word, mode = ('VN',)):

    pass

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P1(word, messages):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P2(word, messages):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P3(word, messages):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def tim_phan_1_EN(doc, word, find_range):

    myrange = find_range.Duplicate

    find = myrange.Find

    find.ClearFormatting()

    find.Text = '(\\<S\\@\\>)(*)(\\<E\\@\\>)'

    find.Replacement.Text = ''

    find.Forward = True

    find.Wrap = 0

    find.MatchWildcards = True

    if find.Execute():

        return myrange





def check_du_lieu_sau_xuongdong_P1_EN(word, messages):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P1_Japan(word, messages):

    doc = word.ActiveDocument

    find_range = doc.Range()

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P123_Me(word):

    doc = word.ActiveDocument

    messages = []

    check_file_mau(word, mode = 'VN')

    for i in range(1, 5):

        check_ki_hieu(word, messages, f'''S{i}@''', f'''E{i}@''')

    check_ki_hieu(word, messages, '<SCHUM@>', '<ECHUM@>')

    check_ki_hieu(word, messages, '<SCHUM@><CHON>', '<ECHUM@><CHON>')

    vbf.add_blank_line_after_table(word)

    check_du_lieu_sau_xuongdong_P1(word, messages)

    check_du_lieu_sau_xuongdong_P2(word, messages)

    check_du_lieu_sau_xuongdong_P3(word, messages)

    chd.xoa_dong_trang_new_Mix(word)

    if messages:

        messages.append('SAU KHI SỬA CHỮA HÃY KIỂM TRA LẠI')

        thong_bao = '\n'.join(messages)

        messagebox.showinfo('Thông báo', thong_bao)

        return None

    messagebox.showinfo('Thông báo', 'Chưa phát hiện lỗi, hãy tiếp tục')

    return None

# WARNING: Decompyle incomplete





def check_du_lieu_sau_xuongdong_P123_Me_EN(word):

    doc = word.ActiveDocument

    messages = []

    check_file_mau(word, mode = 'EN')

    check_ki_hieu(word, messages, '<S@>', '<E@>')

    check_ki_hieu(word, messages, '<SNHOM@>', '<ENHOM@>')

    check_ki_hieu(word, messages, '<SNHOM@><CD>', '<ENHOM@><CD>')

    check_ki_hieu(word, messages, '<SNHOM@><DC>', '<ENHOM@><DC>')

    check_ki_hieu(word, messages, '<SNHOM@><C>', '<ENHOM@><C>')

    check_ki_hieu(word, messages, '<SNHOM@><D>', '<ENHOM@><D>')

    vbf.add_blank_line_after_table(word)

    check_du_lieu_sau_xuongdong_P1_EN(word, messages)

    check_du_lieu_sau_xuongdong_P1_Japan(word, messages)

    chd.xoa_dong_trang_new_Mix(word)

    if messages:

        messages.append('SAU KHI SỬA CHỮA HÃY KIỂM TRA LẠI')

        thong_bao = '\n'.join(messages)

        messagebox.showinfo('Thông báo', thong_bao)

        return None

    messagebox.showinfo('Thông báo', 'Chưa phát hiện lỗi, hãy tiếp tục')

    return None

# WARNING: Decompyle incomplete





def thay_so_trong_phan_duc_lo(word):

    doc = word.ActiveDocument

    if word.Selection.Type != 1:

        myrange = word.Selection.Range

        myrange_find = myrange.Duplicate

        for j in range(1, 20):

            myrange_find.Select()

            vbf.thay_the_replace_stop_1(word, '([\\(])([0-9]{1,2})([\\)])', f'''\\1#{j}\\3''')

        return None

    messagebox.showinfo('Thông báo', 'Chưa chọn (Bôi đen) vùng làm việc')

    return None

# WARNING: Decompyle incomplete





def check_data_nhanh_combine_docx(word):

    doc = word.ActiveDocument

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    bad_part = '\\Python\\VBA\\PC7\\https:\\tcdcnh-my.sharepoint.com\\personal\\hungnn_giahoi_cee_edu_vn\\Documents'

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    if not os.path.exists(doc_path):

        messagebox.showerror('Lỗi', 'File không tồn tại!')

        return None

    doc.Save()

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    bad_part = '\\Python\\VBA\\PC7\\https:\\tcdcnh-my.sharepoint.com\\personal\\hungnn_giahoi_cee_edu_vn\\Documents'

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    doc_path_name = os.path.splitext(doc_path)[0]

    doc_ext = os.path.splitext(doc_path)[1]

    if doc_ext.lower() != '.docx':

        new_file_name = file_name + '.docx'

        doc.SaveAs(FileName = new_file_name, FileFormat = 12)

        doc.Convert()

        doc.Save()

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    doc.Close(SaveChanges = True)

    file_docx = docxtool.open_doc_off_python_docx(doc_path)

    ket_qua = []

    (ok, messages) = docxcheck.check_data_docx_mix(file_docx)

    if ok:

        ket_qua.append('✅ Không phát hiện lỗi')

    else:

        ket_qua.append('❌ Có lỗi')

        for msg in messages:

            ket_qua.append(f'''    - {msg}''')

    thong_bao = '\n'.join(ket_qua)

    word.Documents.Open(doc_path)

    print('Đã xong kiểm tra')

    messagebox.showinfo('KẾT QUẢ KIỂM TRA', thong_bao)

    return None

# WARNING: Decompyle incomplete





def check_data_nhanh_combine_docx_EN(word):

    doc = word.ActiveDocument

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    bad_part = '\\Python\\VBA\\PC7\\https:\\tcdcnh-my.sharepoint.com\\personal\\hungnn_giahoi_cee_edu_vn\\Documents'

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    if not os.path.exists(doc_path):

        messagebox.showerror('Lỗi', 'File không tồn tại!')

        return None

    doc.Save()

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    bad_part = '\\Python\\VBA\\PC7\\https:\\tcdcnh-my.sharepoint.com\\personal\\hungnn_giahoi_cee_edu_vn\\Documents'

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    doc_path_name = os.path.splitext(doc_path)[0]

    doc_ext = os.path.splitext(doc_path)[1]

    if doc_ext.lower() != '.docx':

        new_file_name = file_name + '.docx'

        doc.SaveAs(FileName = new_file_name, FileFormat = 12)

        doc.Convert()

        doc.Save()

    doc_path = doc.FullName

    doc_path = os.path.normpath(os.path.abspath(doc_path))

    if bad_part in doc_path:

        doc_path = doc_path.replace(bad_part, '')

    doc.Close(SaveChanges = True)

    file_docx = docxtool.open_doc_off_python_docx(doc_path)

    ket_qua = []

    (ok, messages) = docxcheck.check_data_docx_mix_EN(file_docx)

    if ok:

        ket_qua.append('✅ Không phát hiện lỗi')

    else:

        ket_qua.append('❌ Có lỗi')

        for msg in messages:

            ket_qua.append(f'''    - {msg}''')

    thong_bao = '\n'.join(ket_qua)

    word.Documents.Open(doc_path)

    print('Đã xong kiểm tra')

    messagebox.showinfo('KẾT QUẢ KIỂM TRA', thong_bao)

    return None

# WARNING: Decompyle incomplete





def chinh_sua_size_font(word):

    pass

# WARNING: Decompyle incomplete





def update_fild(word):

    doc = word.ActiveDocument

    doc.Fields.Update()

    return None

# WARNING: Decompyle incomplete





def update_numpage_py32(word):

    doc = word.ActiveDocument

    num_pages = doc.ComputeStatistics(2)

    num_pages_fix = f'''{num_pages:02}'''

    text = '(Đề thi có )([0-9]{1,2})( trang)'

    text_replace = f'''\\1{num_pages_fix}\\3'''

    myrange = doc.Range()

    find = myrange.Find

    find.ClearFormatting()

    find.Replacement.ClearFormatting()

    find.Text = text

    find.Replacement.Text = text_replace

    find.Forward = True

    find.MatchCase = True

    find.MatchWholeWord = False

    find.MatchWildcards = True

    find.MatchSoundsLike = False

    find.MatchAllWordForms = False

    find.Wrap = 0

    find.Format = False

    if find.Execute(Replace = 1):

        return None

    text = '(Đề có )([0-9]{1,2})( trang)'

    text_replace = f'''\\1{num_pages_fix}\\3'''

    myrange = doc.Range()

    find = myrange.Find

    find.ClearFormatting()

    find.Replacement.ClearFormatting()

    find.Text = text

    find.Replacement.Text = text_replace

    find.Forward = True

    find.MatchCase = True

    find.MatchWholeWord = False

    find.MatchWildcards = True

    find.MatchSoundsLike = False

    find.MatchAllWordForms = False

    find.Wrap = 0

    find.Format = False

    if find.Execute(Replace = 1):

        return None

    text = '<sotrang>'

    text_replace = num_pages_fix

    myrange = doc.Range()

    find = myrange.Find

    find.ClearFormatting()

    find.Replacement.ClearFormatting()

    find.Text = text

    find.Replacement.Text = text_replace

    find.Forward = True

    find.MatchCase = True

    find.MatchWholeWord = False

    find.MatchWildcards = False

    find.MatchSoundsLike = False

    find.MatchAllWordForms = False

    find.Wrap = 0

    find.Format = False

    find.Execute(Replace = 1)

    return None

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    word = vbf.khoi_tao_word_2()

    addin_name = 'HUY'

    chuc_nang_khac_EN_mix(root, word, addin_name)

    return None

