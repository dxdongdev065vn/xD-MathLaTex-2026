import win32com.client as win32com

import win32api

import re

import tkinter as tk

from tkinter import messagebox

from core import functions as vbf



def bo_gach_chan_mu9(word):

    doc = word.ActiveDocument

    selection = word.Selection

    find = selection.Find

    find.ClearFormatting()

    find.Replacement.ClearFormatting()

    find.Font.Underline = True

    find.Replacement.Font.Underline = False

    find.Text = '([^13^9])'

    find.Replacement.Text = '\\1'

    find.Forward = True

    find.MatchCase = True

    find.MatchWholeWord = False

    find.MatchWildcards = True

    find.MatchSoundsLike = False

    find.MatchAllWordForms = False

    find.Wrap = 1

    find.Format = True

    find.Execute(Replace = 2)

    return None

# WARNING: Decompyle incomplete





def tach_cac_loai_dapan_tu_vung_chon(word):

    '''

    Duyệt qua vùng chọn (có thể chứa nhiều bảng hoặc 1 bảng lớn),

    tách đáp án thành 3 nhóm: ABCD, DS, TLN.

    '''

    doc = word.ActiveDocument

    selection = word.Selection

    myrange = selection.Range

    tables = myrange.Tables

    dapan_ABCD = []

    dapan_DS = []

    dapan_TLN = []

    if tables.Count == 0:

        messagebox.showinfo('Thông báo', 'Chưa có bảng trong vùng được chọn.')

        return (dapan_ABCD, dapan_DS, dapan_TLN)

    for t in None(1, tables.Count + 1):

        table = tables(t)

        for row in range(1, table.Rows.Count + 1):

            for col in range(1, table.Columns.Count + 1):

                cell = table.Cell(Row = row, Column = col)

                text = cell.Range.Text.replace('\r', '').replace('\x07', '').replace(' ', '').strip()

                parts = re.split(':', text)

                substring = parts[1].strip() if len(parts) > 1 else text.strip()

                if not substring:

                    continue

                da = substring

                if re.fullmatch('[A-D]', da):

                    dapan_ABCD.append(da)

                    continue

                if re.fullmatch('[DĐS]{4}', da):

                    dapan_DS.append(da)

                    continue

                dapan_TLN.append(da.strip())

    return (dapan_ABCD, dapan_DS, dapan_TLN)

# WARNING: Decompyle incomplete





def danh_dau_dap_an_Table_ABCD_for_One(word, dapan_ABCD):

    doc = word.ActiveDocument

    cell_contents = dapan_ABCD

    socau = len(cell_contents)

    if socau > 0:

        j = 0

        paras = list(doc.Paragraphs)

        total = len(paras)

        i = 0

        if i < total:

            if j < socau:

                text = paras[i].Range.Text.strip()

                if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

                    start = paras[i].Range.Start

                    end = doc.Range().End

                    for k in range(i + 1, total):

                        next_text = paras[k].Range.Text.strip()

                        if not re.match('^(Câu|Question)\\s+[0-9]{1,2}[.:]', next_text):

                            continue

                        end = paras[k].Range.Start

                        range(i + 1, total)

                    cau_range = doc.Range(start, end)

                    myrange_find = cau_range.Duplicate

                    find = myrange_find.Find

                    found = myrange_find.Find.Execute(FindText = '([^13^9]D.)', MatchWildcards = True)

                    if found:

                        dap_an = cell_contents[j]

                        find = cau_range.Find

                        find.ClearFormatting()

                        find.Replacement.ClearFormatting()

                        find.Text = f'''([^13^9])({dap_an}.)'''

                        find.Replacement.Text = '\\1\\2'

                        find.MatchWildcards = True

                        find.MatchCase = True

                        find.Forward = True

                        find.Format = True

                        find.Wrap = 0

                        find.Replacement.Font.Underline = True

                        find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                        if find.Execute(Replace = 1):

                            j += 1

                        else:

                            find = cau_range.Find

                            find.ClearFormatting()

                            find.Replacement.ClearFormatting()

                            find.Text = f'''([^13^9][ ]{{1,}})({dap_an}.)'''

                            find.Replacement.Text = '\\1\\2'

                            find.MatchWildcards = True

                            find.MatchCase = True

                            find.Forward = True

                            find.Format = True

                            find.Wrap = 0

                            find.Replacement.Font.Underline = True

                            find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                            if find.Execute(Replace = 1):

                                j += 1

                            else:

                                myrange_find = cau_range.Duplicate

                                find = myrange_find.Find

                                found = myrange_find.Find.Execute(FindText = '([^13^9][ ]{1,}D.)', MatchWildcards = True)

                                if found:

                                    dap_an = cell_contents[j]

                                    find = cau_range.Find

                                    find.ClearFormatting()

                                    find.Replacement.ClearFormatting()

                                    find.Text = f'''([^13^9][ ]{{1,}})({dap_an}.)'''

                                    find.Replacement.Text = '\\1\\2'

                                    find.MatchWildcards = True

                                    find.MatchCase = True

                                    find.Forward = True

                                    find.Format = True

                                    find.Wrap = 0

                                    find.Replacement.Font.Underline = True

                                    find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                                    if find.Execute(Replace = 1):

                                        j += 1

                                    else:

                                        find = cau_range.Find

                                        find.ClearFormatting()

                                        find.Replacement.ClearFormatting()

                                        find.Text = f'''([^13^9])({dap_an}.)'''

                                        find.Replacement.Text = '\\1\\2'

                                        find.MatchWildcards = True

                                        find.MatchCase = True

                                        find.Forward = True

                                        find.Format = True

                                        find.Wrap = 0

                                        find.Replacement.Font.Underline = True

                                        find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                                        if find.Execute(Replace = 1):

                                            j += 1

                i += 1

                if i < total:

                    if j < socau:

                        continue

                    return None

                return None

            return None

        return None

    return None

# WARNING: Decompyle incomplete





def danh_dau_dap_an_Table_DS_for_One(word, dapan_DS):

    doc = word.ActiveDocument

    cell_contents = dapan_DS

    socau = len(cell_contents)

    if socau > 0:

        paras = list(doc.Paragraphs)

        total = len(paras)

        j = 0

        i = 0

        if i < total:

            if j < socau:

                text = paras[i].Range.Text.strip()

                if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

                    start = paras[i].Range.Start

                    end = doc.Range().End

                    for k in range(i + 1, total):

                        next_text = paras[k].Range.Text.strip()

                        if not re.match('^(Câu|Question)\\s+[0-9]{1,2}[.:]', next_text):

                            continue

                        end = paras[k].Range.Start

                        range(i + 1, total)

                    cau_range = doc.Range(start, end)

                    myrange_find = cau_range.Duplicate

                    found = myrange_find.Find.Execute(FindText = '([^13^9]d[\\)])', MatchWildcards = True)

                    if found:

                        dap_an = cell_contents[j]

                        j += 1

                        findtxt = [

                            '([^13^9]a[\\)])',

                            '([^13^9]b[\\)])',

                            '([^13^9]c[\\)])',

                            '([^13^9]d[\\)])']

                        for idx, txt in enumerate(findtxt):

                            if not idx < len(dap_an):

                                continue

                            if not dap_an[idx] in ('Đ', 'D'):

                                continue

                            myrange_find = cau_range.Duplicate

                            find = myrange_find.Find

                            find.ClearFormatting()

                            find.Replacement.ClearFormatting()

                            find.Text = txt

                            find.Replacement.Text = '\\1'

                            find.MatchWildcards = True

                            find.MatchCase = True

                            find.Forward = True

                            find.Format = True

                            find.Wrap = 0

                            find.Replacement.Font.Underline = True

                            find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                            find.Execute(Replace = 1)

                        findtxt = [

                            '([^13^9][ ]{1,}a[\\)])',

                            '([^13^9][ ]{1,}b[\\)])',

                            '([^13^9][ ]{1,}c[\\)])',

                            '([^13^9][ ]{1,}d[\\)])']

                        for idx, txt in enumerate(findtxt):

                            if not idx < len(dap_an):

                                continue

                            if not dap_an[idx] in ('Đ', 'D'):

                                continue

                            myrange_find = cau_range.Duplicate

                            find = myrange_find.Find

                            find.ClearFormatting()

                            find.Replacement.ClearFormatting()

                            find.Text = txt

                            find.Replacement.Text = '\\1'

                            find.MatchWildcards = True

                            find.MatchCase = True

                            find.Forward = True

                            find.Format = True

                            find.Wrap = 0

                            find.Replacement.Font.Underline = True

                            find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                            find.Execute(Replace = 1)

                    else:

                        myrange_find = cau_range.Duplicate

                        found = myrange_find.Find.Execute(FindText = '([^13^9][ ]{1,}d[\\)])', MatchWildcards = True)

                        if found:

                            dap_an = cell_contents[j]

                            j += 1

                            findtxt = [

                                '([^13^9][ ]{1,}a[\\)])',

                                '([^13^9][ ]{1,}b[\\)])',

                                '([^13^9][ ]{1,}c[\\)])',

                                '([^13^9][ ]{1,}d[\\)])']

                            for idx, txt in enumerate(findtxt):

                                if not idx < len(dap_an):

                                    continue

                                if not dap_an[idx] in ('Đ', 'D'):

                                    continue

                                myrange_find = cau_range.Duplicate

                                find = myrange_find.Find

                                find.ClearFormatting()

                                find.Replacement.ClearFormatting()

                                find.Text = txt

                                find.Replacement.Text = '\\1'

                                find.MatchWildcards = True

                                find.MatchCase = True

                                find.Forward = True

                                find.Format = True

                                find.Wrap = 0

                                find.Replacement.Font.Underline = True

                                find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                                find.Execute(Replace = 1)

                            findtxt = [

                                '([^13^9]a[\\)])',

                                '([^13^9]b[\\)])',

                                '([^13^9]c[\\)])',

                                '([^13^9]d[\\)])']

                            for idx, txt in enumerate(findtxt):

                                if not idx < len(dap_an):

                                    continue

                                if not dap_an[idx] in ('Đ', 'D'):

                                    continue

                                myrange_find = cau_range.Duplicate

                                find = myrange_find.Find

                                find.ClearFormatting()

                                find.Replacement.ClearFormatting()

                                find.Text = txt

                                find.Replacement.Text = '\\1'

                                find.MatchWildcards = True

                                find.MatchCase = True

                                find.Forward = True

                                find.Format = True

                                find.Wrap = 0

                                find.Replacement.Font.Underline = True

                                find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                                find.Execute(Replace = 1)

                i += 1

                if i < total:

                    if j < socau:

                        continue

                    return None

                return None

            return None

        return None

    return None

# WARNING: Decompyle incomplete





def dap_an_Table_TLN_for_One(word, dapan_TLN):

    doc = word.ActiveDocument

    cell_contents = dapan_TLN

    socau = len(cell_contents)

    if socau > 0:

        paras = list(doc.Paragraphs)

        total = len(paras)

        j = 0

        i = 0

        if i < total:

            if j < socau:

                text = paras[i].Range.Text.strip()

                next_start = doc.Range().End

                if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

                    start = paras[i].Range.Start

                    for k in range(i + 1, total):

                        next_text = paras[k].Range.Text.strip()

                        if 'HẾT' in next_text:

                            next_start = paras[k].Range.Start

                            range(i + 1, total)

                        elif not re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', next_text):

                            continue

                        next_start = paras[k].Range.Start

                        range(i + 1, total)

                    cau_range = doc.Range(start, next_start)

                    myrange_find = cau_range.Duplicate

                    found = myrange_find.Find.Execute(FindText = '([^13^9][Dd][.\\)])', MatchWildcards = True)

                    found2 = myrange_find.Find.Execute(FindText = '([^13^9][ ]{1,}[Dd][.\\)])', MatchWildcards = True)

                    if not found and found2:

                        dap_an = cell_contents[j]

                        j += 1

                        rng = doc.Range(next_start, next_start)

                        rng.InsertBefore(f'''ĐS:{dap_an}\r''')

                        para = rng.Paragraphs(1)

                        para_rng = para.Range

                        para_rng.Font.Bold = True

                        para_rng.Font.Color = win32api.RGB(255, 0, 0)

                        para_rng.ParagraphFormat.Alignment = 3

                i += 1

                if i < total:

                    if j < socau:

                        continue

                    return None

                return None

            return None

        return None

    return None

# WARNING: Decompyle incomplete





def danh_dau_tu_dong_all(word):

    doc = word.ActiveDocument

    (dapan_ABCD, dapan_DS, dapan_TLN) = tach_cac_loai_dapan_tu_vung_chon(word)

    print(f'''Loại 1:{dapan_ABCD}, số câu: {len(dapan_ABCD)}''')

    print(f'''Loại 2:{dapan_DS}, số câu: {len(dapan_DS)}''')

    print(f'''Loại 3:{dapan_TLN}), số câu: {len(dapan_TLN)}''')

    if not dapan_ABCD and dapan_DS and dapan_TLN:

        messagebox.showinfo('Thông báo', 'Không phát hiện đáp án hợp lệ trong vùng chọn.')

        return None

    vbf.thay_the_replace(word, '^m', '^13')

    vbf.Convert_Auto_To_Text(word)

    vbf.them_cau_acong_cuoi(word)

    vbf.them_cau_acong_cuoi_En(word)

    vbf.add_blank_line_at_Home(word)

    vbf.add_blank_line_after_table(word)

    if dapan_ABCD:

        danh_dau_dap_an_Table_ABCD_for_One(word, dapan_ABCD)

    if dapan_DS:

        danh_dau_dap_an_Table_DS_for_One(word, dapan_DS)

    if dapan_TLN:

        dap_an_Table_TLN_for_One(word, dapan_TLN)

    bo_gach_chan_mu9(word)

    vbf.xoa_cau_00(word)

    vbf.xoa_dong_trang(word)

    word.Selection.HomeKey(Unit = 6)

    messagebox.showinfo('Hoàn tất', 'Đã đánh dấu xong!')

    return None

# WARNING: Decompyle incomplete





def danh_dau_dap_an_Table_ABCD(word):

    doc = word.ActiveDocument

    selection = word.Selection

    myrange = selection.Range

    table_count = selection.Tables.Count

    if table_count == 0:

        messagebox.showinfo('Thông báo', 'Chưa có bảng trong vùng được chọn')

        return None

    vbf.thay_the_replace(word, '^m', '^13')

    vbf.Convert_Auto_To_Text(word)

    vbf.them_cau_acong_cuoi(word)

    vbf.them_cau_acong_cuoi_En(word)

    vbf.add_blank_line_at_Home(word)

    vbf.add_blank_line_after_table(word)

    table = myrange.Tables(1)

    cell_contents = []

    for row in range(1, table.Rows.Count + 1):

        for col in range(1, table.Columns.Count + 1):

            cell = table.Cell(Row = row, Column = col)

            cell_text = cell.Range.Text.strip()

            cell_text = cell_text.replace('\r', '').replace('\x07', '').replace(' ', '').strip()

            da = ''

            parts = re.split('[:.-]', cell_text)

            if len(substring) >= 1:

                da = substring

            if not len(da) >= 1:

                continue

            if not (lambda .0: pass# WARNING: Decompyle incomplete

)(da()):

                continue

            cell_contents.append(da)

    socau = len(cell_contents)

    if socau == 0:

        messagebox.showinfo('Thông báo', 'Không đọc được đáp án trong bảng')

        return None

    j = 0

    paras = list(doc.Paragraphs)

    total = len(paras)

    i = 0

    if i < total and j < socau:

        text = paras[i].Range.Text.strip()

        if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

            start = paras[i].Range.Start

            end = doc.Range().End

            for k in range(i + 1, total):

                next_text = paras[k].Range.Text.strip()

                if not re.match('^(Câu|Question)\\s+[0-9]{1,2}[.:]', next_text):

                    continue

                end = paras[k].Range.Start

                range(i + 1, total)

            cau_range = doc.Range(start, end)

            myrange_find = cau_range.Duplicate

            find = myrange_find.Find

            found = myrange_find.Find.Execute(FindText = '([^13^9]D.)', MatchWildcards = True)

            if found:

                dap_an = cell_contents[j]

                find = cau_range.Find

                find.ClearFormatting()

                find.Replacement.ClearFormatting()

                find.Text = f'''([^13^9])({dap_an}.)'''

                find.Replacement.Text = '\\1\\2'

                find.MatchWildcards = True

                find.MatchCase = True

                find.Forward = True

                find.Format = True

                find.Wrap = 0

                find.Replacement.Font.Underline = True

                find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                if find.Execute(Replace = 1):

                    j += 1

                else:

                    find = cau_range.Find

                    find.ClearFormatting()

                    find.Replacement.ClearFormatting()

                    find.Text = f'''([^13^9][ ]{{1,}})({dap_an}.)'''

                    find.Replacement.Text = '\\1\\2'

                    find.MatchWildcards = True

                    find.MatchCase = True

                    find.Forward = True

                    find.Format = True

                    find.Wrap = 0

                    find.Replacement.Font.Underline = True

                    find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                    if find.Execute(Replace = 1):

                        j += 1

                    else:

                        myrange_find = cau_range.Duplicate

                        find = myrange_find.Find

                        found = myrange_find.Find.Execute(FindText = '([^13^9][ ]{1,}D.)', MatchWildcards = True)

                        if found:

                            dap_an = cell_contents[j]

                            find = cau_range.Find

                            find.ClearFormatting()

                            find.Replacement.ClearFormatting()

                            find.Text = f'''([^13^9][ ]{{1,}})({dap_an}.)'''

                            find.Replacement.Text = '\\1\\2'

                            find.MatchWildcards = True

                            find.MatchCase = True

                            find.Forward = True

                            find.Format = True

                            find.Wrap = 0

                            find.Replacement.Font.Underline = True

                            find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                            if find.Execute(Replace = 1):

                                j += 1

                            else:

                                find = cau_range.Find

                                find.ClearFormatting()

                                find.Replacement.ClearFormatting()

                                find.Text = f'''([^13^9])({dap_an}.)'''

                                find.Replacement.Text = '\\1\\2'

                                find.MatchWildcards = True

                                find.MatchCase = True

                                find.Forward = True

                                find.Format = True

                                find.Wrap = 0

                                find.Replacement.Font.Underline = True

                                find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                                if find.Execute(Replace = 1):

                                    j += 1

        i += 1

        if i < total and j < socau:

            continue

    bo_gach_chan_mu9(word)

    vbf.xoa_cau_00(word)

    vbf.xoa_dong_trang(word)

    word.Selection.HomeKey(Unit = 6)

    return None

# WARNING: Decompyle incomplete





def danh_dau_dap_an_Table_DS(word):

    doc = word.ActiveDocument

    selection = word.Selection

    myrange = selection.Range

    table_count = selection.Tables.Count

    if table_count == 0:

        messagebox.showinfo('Thông báo', 'Chưa có vùng được chọn')

        return None

    vbf.thay_the_replace(word, '^m', '^13')

    vbf.Convert_Auto_To_Text(word)

    vbf.them_cau_acong_cuoi(word)

    vbf.them_cau_acong_cuoi_En(word)

    vbf.add_blank_line_at_Home(word)

    vbf.add_blank_line_after_table(word)

    table = myrange.Tables(1)

    cell_contents = []

    for row in range(1, table.Rows.Count + 1):

        for col in range(1, table.Columns.Count + 1):

            cell = table.Cell(Row = row, Column = col)

            cell_text = cell.Range.Text.strip()

            cell_text = cell_text.replace('\r', '').replace('\x07', '').replace(' ', '').strip()

            da = ''

            parts = re.split('[:.-]', cell_text)

            if len(substring) >= 1:

                da = substring

            if not len(da) == 4:

                continue

            cell_contents.append(da)

    socau = len(cell_contents)

    if socau == 0:

        messagebox.showinfo('Thông báo', 'Không đọc được đáp án trong bảng')

        return None

    paras = list(doc.Paragraphs)

    total = len(paras)

    j = 0

    i = 0

    if i < total and j < socau:

        text = paras[i].Range.Text.strip()

        if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

            start = paras[i].Range.Start

            end = doc.Range().End

            for k in range(i + 1, total):

                next_text = paras[k].Range.Text.strip()

                if not re.match('^(Câu|Question)\\s+[0-9]{1,2}[.:]', next_text):

                    continue

                end = paras[k].Range.Start

                range(i + 1, total)

            cau_range = doc.Range(start, end)

            myrange_find = cau_range.Duplicate

            found = myrange_find.Find.Execute(FindText = '([^13^9]d[\\)])', MatchWildcards = True)

            if found:

                dap_an = cell_contents[j]

                j += 1

                findtxt = [

                    '([^13^9]a[\\)])',

                    '([^13^9]b[\\)])',

                    '([^13^9]c[\\)])',

                    '([^13^9]d[\\)])']

                for idx, txt in enumerate(findtxt):

                    if not idx < len(dap_an):

                        continue

                    if not dap_an[idx] in ('Đ', 'D'):

                        continue

                    myrange_find = cau_range.Duplicate

                    find = myrange_find.Find

                    find.ClearFormatting()

                    find.Replacement.ClearFormatting()

                    find.Text = txt

                    find.Replacement.Text = '\\1'

                    find.MatchWildcards = True

                    find.MatchCase = True

                    find.Forward = True

                    find.Format = True

                    find.Wrap = 0

                    find.Replacement.Font.Underline = True

                    find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                    find.Execute(Replace = 1)

                findtxt = [

                    '([^13^9][ ]{1,}a[\\)])',

                    '([^13^9][ ]{1,}b[\\)])',

                    '([^13^9][ ]{1,}c[\\)])',

                    '([^13^9][ ]{1,}d[\\)])']

                for idx, txt in enumerate(findtxt):

                    if not idx < len(dap_an):

                        continue

                    if not dap_an[idx] in ('Đ', 'D'):

                        continue

                    myrange_find = cau_range.Duplicate

                    find = myrange_find.Find

                    find.ClearFormatting()

                    find.Replacement.ClearFormatting()

                    find.Text = txt

                    find.Replacement.Text = '\\1'

                    find.MatchWildcards = True

                    find.MatchCase = True

                    find.Forward = True

                    find.Format = True

                    find.Wrap = 0

                    find.Replacement.Font.Underline = True

                    find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                    find.Execute(Replace = 1)

            else:

                myrange_find = cau_range.Duplicate

                found = myrange_find.Find.Execute(FindText = '([^13^9][ ]{1,}d[\\)])', MatchWildcards = True)

                if found:

                    dap_an = cell_contents[j]

                    j += 1

                    findtxt = [

                        '([^13^9][ ]{1,}a[\\)])',

                        '([^13^9][ ]{1,}b[\\)])',

                        '([^13^9][ ]{1,}c[\\)])',

                        '([^13^9][ ]{1,}d[\\)])']

                    for idx, txt in enumerate(findtxt):

                        if not idx < len(dap_an):

                            continue

                        if not dap_an[idx] in ('Đ', 'D'):

                            continue

                        myrange_find = cau_range.Duplicate

                        find = myrange_find.Find

                        find.ClearFormatting()

                        find.Replacement.ClearFormatting()

                        find.Text = txt

                        find.Replacement.Text = '\\1'

                        find.MatchWildcards = True

                        find.MatchCase = True

                        find.Forward = True

                        find.Format = True

                        find.Wrap = 0

                        find.Replacement.Font.Underline = True

                        find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                        find.Execute(Replace = 1)

                    findtxt = [

                        '([^13^9]a[\\)])',

                        '([^13^9]b[\\)])',

                        '([^13^9]c[\\)])',

                        '([^13^9]d[\\)])']

                    for idx, txt in enumerate(findtxt):

                        if not idx < len(dap_an):

                            continue

                        if not dap_an[idx] in ('Đ', 'D'):

                            continue

                        myrange_find = cau_range.Duplicate

                        find = myrange_find.Find

                        find.ClearFormatting()

                        find.Replacement.ClearFormatting()

                        find.Text = txt

                        find.Replacement.Text = '\\1'

                        find.MatchWildcards = True

                        find.MatchCase = True

                        find.Forward = True

                        find.Format = True

                        find.Wrap = 0

                        find.Replacement.Font.Underline = True

                        find.Replacement.Font.Color = win32api.RGB(255, 0, 0)

                        find.Execute(Replace = 1)

        i += 1

        if i < total and j < socau:

            continue

    bo_gach_chan_mu9(word)

    vbf.xoa_cau_00(word)

    vbf.xoa_dong_trang(word)

    word.Selection.HomeKey(Unit = 6)

    return None

# WARNING: Decompyle incomplete





def danh_dau_dap_an_DS_nhieu_cau(word):

    doc = word.ActiveDocument

    

    def danh_dau_dap_an_cau_select(word, doc):

        selection = word.Selection

        myrange = selection.Range

        paras = list(myrange.Paragraphs)

        total = len(paras)

        dap_an = None

        dap_an_para_index = -1

        for i, para in enumerate(paras):

            text = para.Range.Text.strip()

            match = re.search('(ĐS|Đáp\\s*án|Trả\\s*lời)\\s*[:：]\\s*([ĐDS\\s]+)', text, flags = re.IGNORECASE)

            if not match:

                continue

            ds = match.group(2).replace(' ', '').upper()

            if not len(ds) == 4:

                continue

            if not (lambda .0: pass# WARNING: Decompyle incomplete

)(ds()):

                continue

            dap_an = ds

            dap_an_para_index = i

            all

    # WARNING: Decompyle incomplete



    selection = word.Selection

    myrange = selection.Range

    paras = list(myrange.Paragraphs)

    total = len(paras)

    indices = []

    for i, para in enumerate(paras):

        text = para.Range.Text.strip()

        if not re.match('^(Câu|Bài|Question)\\s*\\d+[\\.:]', text, flags = re.IGNORECASE):

            continue

        indices.append(i)

    if not indices:

        messagebox.showinfo('Thông báo', 'Không phát hiện được câu hỏi nào trong vùng bôi đen.')

        return None

    blocks = []

    for i in range(len(indices)):

        start_idx = indices[i]

        blocks.append((start_idx, end_idx))

    dem = 0

    for start, end in blocks:

        cau_range = doc.Range(paras[start].Range.Start, paras[end].Range.End)

        cau_range.Select()

        danh_dau_dap_an_cau_select(word, doc)

        dem += 1

    bo_gach_chan_mu9(word)

    messagebox.showinfo('Hoàn tất', f'''✅ Đã xử lý {dem} câu hỏi''')

    return None

# WARNING: Decompyle incomplete





def dap_an_Table_TLN(word):

    doc = word.ActiveDocument

    selection = word.Selection

    myrange = selection.Range

    table_count = myrange.Tables.Count

    if table_count == 0:

        messagebox.showinfo('Thông báo', 'Chưa có vùng được chọn')

        return None

    vbf.thay_the_replace(word, '^m', '^13')

    vbf.thay_the_replace(word, '(^13Phần )([1234IV]{1,2})([.:])', '^13PHẦN\\2\\3')

    vbf.Convert_Auto_To_Text(word)

    vbf.them_cau_acong_cuoi(word)

    vbf.add_blank_line_at_Home(word)

    vbf.add_blank_line_after_table(word)

    table = myrange.Tables(1)

    cell_contents = []

    for row in range(1, table.Rows.Count + 1):

        for col in range(1, table.Columns.Count + 1):

            cell = table.Cell(Row = row, Column = col)

            cell_text = cell.Range.Text.strip()

            cell_text = cell_text.replace('\r', '').replace('\x07', '').replace(' ', '').strip()

            da = ''

            parts = re.split(':', cell_text)

            if len(substring) >= 1:

                da = substring

            if not len(da) >= 1:

                continue

            cell_contents.append(da)

    socau = len(cell_contents)

    if socau == 0:

        messagebox.showinfo('Thông báo', 'Không đọc được đáp án trong bảng')

        return None

    paras = list(doc.Paragraphs)

    total = len(paras)

    j = 0

    i = 0

    if i < total and j < socau:

        text = paras[i].Range.Text.strip()

        next_start = doc.Range().End

        if re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', text):

            start = paras[i].Range.Start

            for k in range(i + 1, total):

                next_text = paras[k].Range.Text.strip()

                if 'HẾT' in next_text:

                    next_start = paras[k].Range.Start

                    range(i + 1, total)

                elif not re.match('^(Câu|Bài|Question)\\s+[0-9]{1,2}[.:]', next_text):

                    continue

                next_start = paras[k].Range.Start

                range(i + 1, total)

            cau_range = doc.Range(start, next_start)

            myrange_find = cau_range.Duplicate

            found = myrange_find.Find.Execute(FindText = '([^13^9][Dd][.\\)])', MatchWildcards = True)

            if not found:

                dap_an = cell_contents[j]

                j += 1

                rng = doc.Range(next_start, next_start)

                rng.InsertBefore(f'''ĐS:{dap_an}\r''')

                para = rng.Paragraphs(1)

                para_rng = para.Range

                para_rng.Font.Bold = True

                para_rng.Font.Color = win32api.RGB(255, 0, 0)

                para_rng.ParagraphFormat.Alignment = 3

        i += 1

        if i < total and j < socau:

            continue

    vbf.xoa_cau_00(word)

    vbf.xoa_dong_trang(word)

    word.Selection.HomeKey(Unit = 6)

    return None

# WARNING: Decompyle incomplete





def gui_danh_dau_dung_sai_one_cau(root, word):

    pass

# WARNING: Decompyle incomplete





def danh_dau_dap_an(root, word):

    pass

# WARNING: Decompyle incomplete





def danh_dau_dap_an_main(root, word):

    pass

# WARNING: Decompyle incomplete



if __name__ == '__main__':

    root = tk.Tk()

    word = vbf.khoi_tao_word_2()

    danh_dau_dap_an_main(root, word)

    root.mainloop()

    return None

