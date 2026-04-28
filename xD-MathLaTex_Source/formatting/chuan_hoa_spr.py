from spire.doc import Section, Paragraph, Table, TextRange, Shape, FileFormat, HorizontalAlignment, Regex

from spire.doc import UnderlineStyle

from spire.doc import Document as SpireDocument

from spire.doc import Color

from spire.doc import DocumentObjectType

from spire.doc import OfficeMath

import re

import bisect

from mixing import mix_spr_func as mspr_f



def InchesToPoints(inches):

    return inches * 72





def LinesToPoints(lines):

    return lines * 12





def open_doc_off_spire(doc_path):

    doc = SpireDocument()

    doc.LoadFromFile(doc_path)

    return doc





def is_red_color_spr(color):

    if color.R == 255:

        color.R == 255

        if color.G == 0:

            color.G == 0

    return color.B == 0





def remove_last_element(lst):

    return lst[:-1]





def resize_images_to_column_width(doc):

    for i in range(doc.Sections.Count):

        section = doc.Sections[i]

        for j in range(section.Paragraphs.Count):

            paragraph = section.Paragraphs[j]

            for k in range(paragraph.ChildObjects.Count):

                child = paragraph.ChildObjects[k]

                if not isinstance(child, Shape):

                    continue

                picture = child

                if not picture.Width > InchesToPoints(3.5):

                    continue

                aspect_ratio = picture.Height / picture.Width

                picture.Width = InchesToPoints(3.5)

                picture.Height = picture.Width * aspect_ratio





def page_2_cot_Mix(doc):

    for i in range(doc.Sections.Count):

        section = doc.Sections[i]

        page_setup = section.PageSetup

        page_setup.PageWidth = InchesToPoints(8.27)

        page_setup.PageHeight = InchesToPoints(11.69)

        page_setup.Margins.Top = InchesToPoints(0.4)

        page_setup.Margins.Bottom = InchesToPoints(0.4)

        page_setup.Margins.Left = InchesToPoints(0.4)

        page_setup.Margins.Right = InchesToPoints(0.4)

        page_setup.Margins.Gutter = InchesToPoints(0)

        page_setup.HeaderDistance = InchesToPoints(0.24)

        page_setup.FooterDistance = InchesToPoints(0.24)

        total_width = page_setup.PageWidth - page_setup.Margins.Left - page_setup.Margins.Right

        spacing = InchesToPoints(0.1)

        column_width = (total_width - spacing) / 2

        section.AddColumn(column_width, spacing)

        section.PageSetup.ColumnsLineBetween = True





def delete_bookmarks(doc):

    excluded_names = {

        'MDH',

        'num_page'}

    names_to_keep = set()

    bookmark_ends_to_keep = set()

    for i in range(doc.Bookmarks.Count):

        bookmark = doc.Bookmarks[i]

        if not bookmark.Name in excluded_names:

            continue

        names_to_keep.add(bookmark.Name)

        bookmark_ends_to_keep.add(bookmark.BookmarkEnd)

    for i in range(doc.Sections.Count):

        section = doc.Sections[i]

        for j in range(section.Paragraphs.Count):

            para = section.Paragraphs[j]

            index = para.ChildObjects.Count - 1

            if not index >= 0:

                continue

            child = para.ChildObjects[index]

            if isinstance(child, BookmarkStart) or child.Name not in names_to_keep:

                para.ChildObjects.RemoveAt(index)

            elif isinstance(child, BookmarkEnd) and child not in bookmark_ends_to_keep:

                para.ChildObjects.RemoveAt(index)

            index -= 1

            if index >= 0:

                continue

    continue





def Danh_STT_Cau(doc, indices, i):

    section = doc.Sections[0]

    pattern = '^Câu [0-9]{1,}'

    regex = Regex(pattern)

    for p in indices:

        paragraph = doc.Sections[0].Body.ChildObjects[p]

        if not isinstance(paragraph, Paragraph):

            continue

        if not re.match(pattern, paragraph.Text):

            continue

        paragraph.Replace(regex, f'''Câu {i}''')

        i = i + 1





def Danh_STT_Cau_All(doc):

    for k in range(1, 5):

        all_indices_cau = mspr_f.tim_cau_trong_phan(doc, f'''S{k}@''', f'''E{k}@''')

        if not len(all_indices_cau) > 0:

            continue

        i = 1

        for idx in range(0, len(all_indices_cau)):

            cau = all_indices_cau[idx]

            Danh_STT_Cau(doc, cau, i)

            i += len(cau) - 1





def Danh_STT_Cau_EN(doc, indices, i):

    section = doc.Sections[0]

    pattern1 = '^Question [0-9]{1,}'

    pattern2 = '^Câu [0-9]{1,}'

    regex1 = Regex(pattern1)

    regex2 = Regex(pattern2)

    for p in indices:

        paragraph = doc.Sections[0].Body.ChildObjects[p]

        if not isinstance(paragraph, Paragraph):

            continue

        if re.match(pattern1, paragraph.Text):

            paragraph.Replace(regex1, f'''Question {i}''')

            i = i + 1

            continue

        if not re.match(pattern2, paragraph.Text):

            continue

        paragraph.Replace(regex2, f'''Câu {i}''')

        i = i + 1





def Danh_STT_Cau_All_EN(doc):

    all_indices_cau = mspr_f.tim_cau_trong_phan_EN(doc)

    if len(all_indices_cau) > 0:

        i = 1

        for idx in range(0, len(all_indices_cau)):

            cau = all_indices_cau[idx]

            Danh_STT_Cau_EN(doc, cau, i)

            i += len(cau) - 1

        return None





def STT_ABCD_trong_cac_vung(doc, text_start, text_end):

    labels = [

        'A',

        'B',

        'C',

        'D']

    all_indices_cau = mspr_f.tim_cau_trong_phan(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            phuong_an = phuong_an[:-1]

            encounter_count = 0

            for p in phuong_an:

                paragraph = doc.Sections[0].Body.ChildObjects[p]

                if not isinstance(paragraph, Paragraph):

                    continue

                first_child = paragraph.ChildObjects[0]

                current_label = labels[encounter_count % len(labels)]

                first_child.Text = current_label + first_child.Text[1:]

                encounter_count += 1





def STT_ABCD_trong_cac_vung_EN(doc):

    labels = [

        'A',

        'B',

        'C',

        'D']

    all_indices_cau = mspr_f.tim_cau_trong_phan_EN(doc)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            phuong_an = phuong_an[:-1]

            encounter_count = 0

            for p in phuong_an:

                paragraph = doc.Sections[0].Body.ChildObjects[p]

                if not isinstance(paragraph, Paragraph):

                    continue

                first_child = paragraph.ChildObjects[0]

                current_label = labels[encounter_count % len(labels)]

                first_child.Text = current_label + first_child.Text[1:]

                encounter_count += 1





def STT_ABCD(doc):

    STT_ABCD_trong_cac_vung(doc, 'S1@', 'E1@')





def STT_ABCD_EN(doc):

    STT_ABCD_trong_cac_vung_EN(doc)





def STT_abcd_nho(doc):

    labels = [

        'a',

        'b',

        'c',

        'd']

    all_indices_cau = mspr_f.tim_cau_trong_phan(doc, 'S2@', 'E2@')

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_II(doc, cau[k], cau[k + 1])

            phuong_an = phuong_an[:-1]

            encounter_count = 0

            for p in phuong_an:

                paragraph = doc.Sections[0].Body.ChildObjects[p]

                if not isinstance(paragraph, Paragraph):

                    continue

                first_child = paragraph.ChildObjects[0]

                current_label = labels[encounter_count % len(labels)]

                first_child.Text = current_label + first_child.Text[1:]

                encounter_count += 1





def tim_thang_trong_all(doc):

    section = doc.Sections[0]

    indices_cau = []

    for i in range(section.Body.ChildObjects.Count):

        child = section.Body.ChildObjects[i]

        if not isinstance(child, Paragraph):

            continue

        if not re.search('#(\\d+)', child.Text):

            continue

        indices_cau.append(i)

    return indices_cau





def tim_cau_trong_all(doc):

    section = doc.Sections[0]

    pattern1 = '^Question [0-9]{1,}[.:]'

    pattern2 = '^Câu [0-9]{1,}[.:]'

    indices_cau = []

    for i in range(section.Body.ChildObjects.Count):

        child = section.Body.ChildObjects[i]

        if not isinstance(child, Paragraph):

            continue

        if not re.match(pattern1, child.Text) and re.match(pattern2, child.Text):

            continue

        indices_cau.append(i)

    return indices_cau





def STT_thang(doc):

    a = tim_cau_trong_all(doc)

    b = tim_thang_trong_all(doc)

    for bb in b:

        paragraph = doc.Sections[0].Body.ChildObjects[bb]

        if not isinstance(paragraph, Paragraph):

            continue

        matches = re.findall('#(\\d+)', paragraph.Text)

        if not matches:

            continue

        for match in matches:

            current_index = bisect.bisect_left(a, bb)

            new_index = current_index + int(match)

            paragraph.Replace(f'''#{match}''', str(new_index), False, True)





def canh_before_after(doc, p):

    section = doc.Sections[0]

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        paragraph_format = paragraph.Format

        paragraph_format.BeforeAutoSpacing = False

        paragraph_format.AfterAutoSpacing = False

        paragraph_format.BeforeSpacing = 0

        paragraph_format.AfterSpacing = 0

        paragraph_format.LineSpacing = LinesToPoints(1.15)

        return None





def tab_btp_4_at(doc, p):

    default_tab_stop = 0.2

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.21,

        1.97,

        3.73,

        5.49]

    section = doc.Sections[0]

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)

        return None





def tab_btp_2_at(doc, p):

    default_tab_stop = 0.2

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.21,

        3.73]

    section = doc.Sections[0]

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)

        return None





def tab_btp(doc):

    default_tab_stop = 0.2

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.21,

        1.97,

        3.73,

        5.49]

    section = doc.Sections[0]

    for i in range(section.Paragraphs.Count):

        paragraph = section.Paragraphs[i]

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)





def tab_btp_2_cot(doc):

    default_tab_stop = 0.04

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.1,

        0.98,

        1.86,

        2.74]

    section = doc.Sections[0]

    for i in range(section.Paragraphs.Count):

        paragraph = section.Paragraphs[i]

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)





def tab_btp_2_cot_4_at(doc, p):

    default_tab_stop = 0.04

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.1,

        0.98,

        1.86,

        2.74]

    section = doc.Sections[0]

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)

        return None





def tab_btp_2_cot_2_at(doc, p):

    default_tab_stop = 0.04

    doc.DefaultTabStop = InchesToPoints(default_tab_stop)

    tab_stops = [

        0.1,

        1.86]

    section = doc.Sections[0]

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        paragraph.Format.Tabs.Clear()

        for pos in tab_stops:

            tab_stop_position = InchesToPoints(pos)

            paragraph.Format.Tabs.AddTab(tab_stop_position)

        return None





def insert_tab_before_first_child(doc, p):

    paragraph = doc.Sections[0].Body.ChildObjects[p]

    if isinstance(paragraph, Paragraph):

        if paragraph.ChildObjects.Count > 0:

            first_child = paragraph.ChildObjects[0]

            if hasattr(first_child, 'Text'):

                first_child.Text = '\t' + first_child.Text

                return None

            return None

        return None





def insert_tab_before_indices(doc, indices):

    for p in indices:

        insert_tab_before_first_child(doc, p)





def MergeChildObjectsIntoGroup(doc, indices):

    pass

# WARNING: Decompyle incomplete





def option_1(doc, indices):

    pass

# WARNING: Decompyle incomplete





def option_1_2_cot(doc, indices):

    pass

# WARNING: Decompyle incomplete





def option_2_1_1(doc, indices):

    pass

# WARNING: Decompyle incomplete





def option_2_1_1_2_cot(doc, indices):

    pass

# WARNING: Decompyle incomplete





def option_4(doc, indices):

    for p in indices:

        insert_tab_before_first_child(doc, p)

        tab_btp_4_at(doc, p)





def option_4_2_cot(doc, indices):

    for p in indices:

        insert_tab_before_first_child(doc, p)

        tab_btp_2_cot_4_at(doc, p)





def len_option(doc, p):

    body = doc.Sections[0].Body

    if p < 0 or p >= body.ChildObjects.Count:

        raise IndexError('Chỉ số p không hợp lệ.')

    child_object = body.ChildObjects[p]

    text_length = 0

    image_length = 0

    equation_length = 0

    if hasattr(child_object, 'ChildObjects') and child_object.ChildObjects.Count > 0:

        for i in range(child_object.ChildObjects.Count):

            child = child_object.ChildObjects[i]

            if hasattr(child, 'Text'):

                text_length += len(child.Text)

                continue

            if isinstance(child, Shape):

                image_length += int(child.Width // 8.5)

                continue

            if not isinstance(child, OfficeMath):

                continue

            for j in range(child.ChildObjects.Count):

                sub_child = child.ChildObjects[j]

                if not hasattr(sub_child, 'Text'):

                    continue

                equation_length += int(2.5 * len(sub_child.Text))

    elif hasattr(child_object, 'Text'):

        text_length += len(child_object.Text)

    elif isinstance(child_object, Shape):

        image_length += int(child_object.Width // 8.5)

    elif isinstance(child_object, OfficeMath):

        for j in range(child_object.ChildObjects.Count):

            sub_child = child_object.ChildObjects[j]

            if not hasattr(sub_child, 'Text'):

                continue

            equation_length += int(2.5 * len(sub_child.Text))

    total_length = text_length + image_length + equation_length

    return total_length





def max_len_option(doc, indices):

    body = doc.Sections[0].Body

    max_length = 0

    for p in indices:

        if p < 0 or p >= body.ChildObjects.Count:

            raise IndexError(f'''Chỉ số {p} không hợp lệ trong Body.ChildObjects.''')

        current_length = len_option(doc, p)

        if not current_length > max_length:

            continue

        max_length = current_length

    return max_length





def tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end):

    (indices_S, indices_E) = mspr_f.find_phan(doc, text_start, text_end)

    all_indices_cau = []

    if len(indices_S) != len(indices_E):

        messagebox.showinfo('Thông báo', f'''Số kí hiệu {text_start} và {text_end} không bằng nhau, hãy chỉnh sửa phù hợp''')

        return None

    if len(indices_S) != 0:

        section = doc.Sections[0]

        pattern1 = '^Câu [0-9]{1,}[.:]'

        pattern2 = '^Question [0-9]{1,}[.:]'

        pattern3 = '@'

        for start, end in zip(indices_S, indices_E):

            indices_cau = []

            for j in range(start, end):

                child = section.Body.ChildObjects[j]

                if not isinstance(child, Paragraph):

                    continue

                if not re.match(pattern1, child.Text) and re.match(pattern2, child.Text) and re.search(pattern3, child.Text):

                    continue

                indices_cau.append(j)

            indices_cau.append(end)

            all_indices_cau.append(indices_cau)

    return all_indices_cau





def sap_xep_phan_I_trong_cac_vung(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 25:

                option_1(doc, phuong_an)

                continue

            if L1max < 49:

                option_2_1_1(doc, phuong_an)

                continue

            option_4(doc, phuong_an)





def sap_xep_phan_I_trong_cac_vung_font13(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 23:

                option_1(doc, phuong_an)

                continue

            if L1max < 45:

                option_2_1_1(doc, phuong_an)

                continue

            option_4(doc, phuong_an)





def sap_xep_phan_I_trong_cac_vung_font14(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 21:

                option_1(doc, phuong_an)

                continue

            if L1max < 41:

                option_2_1_1(doc, phuong_an)

                continue

            option_4(doc, phuong_an)





def sap_xep_phan_I_trong_cac_vung_Japan(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 9:

                option_1(doc, phuong_an)

                continue

            if L1max < 18:

                option_2_1_1(doc, phuong_an)

                continue

            option_4(doc, phuong_an)





def tim_cau_first(doc):

    section = doc.Sections[0]

    pattern1 = '^Câu [0-9]{1,}[.:]'

    pattern2 = '^Question [0-9]{1,}[.:]'

    indices_cau = []

    for j in range(section.Paragraphs.Count):

        paragraph = section.Paragraphs[j]

        if not isinstance(paragraph, Paragraph):

            continue

        if not re.match(pattern1, paragraph.Text) and re.match(pattern2, paragraph.Text):

            continue

        

        return range(section.Paragraphs.Count), j

    return 0





def lay_size_font(doc, index):

    if doc.Sections.Count > 0 and doc.Sections[0].Paragraphs.Count > 0:

        paragraph = doc.Sections[0].Paragraphs[index]

        for i in range(paragraph.ChildObjects.Count):

            child = paragraph.ChildObjects[i]

            if not isinstance(child, TextRange):

                continue

            font_size = child.CharacterFormat.FontSize

            

            return range(paragraph.ChildObjects.Count), font_size

    return 12





def sap_xep_phan_I(doc):

    index = tim_cau_first(doc)

    size_font = lay_size_font(doc, index)

    if size_font == 13:

        sap_xep_phan_I_trong_cac_vung_font13(doc, 'S1@', 'E1@')

        return None

    if size_font == 14:

        sap_xep_phan_I_trong_cac_vung_font14(doc, 'S1@', 'E1@')

        return None

    sap_xep_phan_I_trong_cac_vung(doc, 'S1@', 'E1@')





def sap_xep_phan_I_EN(doc):

    index = tim_cau_first(doc)

    size_font = lay_size_font(doc, index)

    if size_font == 13:

        sap_xep_phan_I_trong_cac_vung_font13(doc, '<S@>', '<E@>')

        return None

    if size_font == 14:

        sap_xep_phan_I_trong_cac_vung_font14(doc, '<S@>', '<E@>')

        return None

    sap_xep_phan_I_trong_cac_vung(doc, '<S@>', '<E@>')





def sap_xep_phan_I_Japan(doc):

    sap_xep_phan_I_trong_cac_vung_Japan(doc, '<S@>', '<E@>')





def sap_xep_phan_I_2_cot_trong_cac_vung(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 12:

                option_1_2_cot(doc, phuong_an)

                continue

            if L1max < 23:

                option_2_1_1_2_cot(doc, phuong_an)

                continue

            option_4_2_cot(doc, phuong_an)





def sap_xep_phan_I_2_cot_trong_cac_vung_font13(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 11:

                option_1_2_cot(doc, phuong_an)

                continue

            if L1max < 21:

                option_2_1_1_2_cot(doc, phuong_an)

                continue

            option_4_2_cot(doc, phuong_an)





def sap_xep_phan_I_2_cot_trong_cac_vung_font14(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 10:

                option_1_2_cot(doc, phuong_an)

                continue

            if L1max < 19:

                option_2_1_1_2_cot(doc, phuong_an)

                continue

            option_4_2_cot(doc, phuong_an)





def sap_xep_phan_I_2_cot_trong_cac_vung_Japan(doc, text_start, text_end):

    all_indices_cau = tim_cau_trong_phan_for_chuan_hoa(doc, text_start, text_end)

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_I(doc, cau[k], cau[k + 1])

            if not len(phuong_an) > 2:

                continue

            phuong_an = phuong_an[:-1]

            L1max = max_len_option(doc, phuong_an)

            if L1max < 5:

                option_1_2_cot(doc, phuong_an)

                continue

            if L1max < 9:

                option_2_1_1_2_cot(doc, phuong_an)

                continue

            option_4_2_cot(doc, phuong_an)





def sap_xep_phan_I_2_cot(doc):

    index = tim_cau_first(doc)

    size_font = lay_size_font(doc, index)

    if size_font == 13:

        sap_xep_phan_I_2_cot_trong_cac_vung_font13(doc, 'S1@', 'E1@')

        return None

    if size_font == 14:

        sap_xep_phan_I_2_cot_trong_cac_vung_font14(doc, 'S1@', 'E1@')

        return None

    sap_xep_phan_I_2_cot_trong_cac_vung(doc, 'S1@', 'E1@')





def sap_xep_phan_I_2_cot_EN(doc):

    index = tim_cau_first(doc)

    size_font = lay_size_font(doc, index)

    if size_font == 13:

        sap_xep_phan_I_2_cot_trong_cac_vung_font13(doc, '<S@>', '<E@>')

        return None

    if size_font == 14:

        sap_xep_phan_I_2_cot_trong_cac_vung_font14(doc, '<S@>', '<E@>')

        return None

    sap_xep_phan_I_2_cot_trong_cac_vung(doc, '<S@>', '<E@>')





def sap_xep_phan_I_2_cot_Japan(doc):

    sap_xep_phan_I_2_cot_trong_cac_vung_Japan(doc, '<S@>', '<E@>')





def lam_dep_phan_II(doc):

    all_indices_cau = mspr_f.tim_cau_trong_phan(doc, 'S2@', 'E2@')

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_II(doc, cau[k], cau[k + 1])

            phuong_an = phuong_an[:-1]

            for p in phuong_an:

                insert_tab_before_first_child(doc, p)

                tab_btp_4_at(doc, p)





def lam_dep_phan_IV(doc):

    all_indices_cau = mspr_f.tim_cau_trong_phan(doc, 'S4@', 'E4@')

    for idx in range(len(all_indices_cau) - 1, -1, -1):

        cau = all_indices_cau[idx]

        for k in range(len(cau) - 2, -1, -1):

            phuong_an = mspr_f.tim_phuong_an_trong_cau_phan_II(doc, cau[k], cau[k + 1])

            phuong_an = phuong_an[:-1]

            for p in phuong_an:

                insert_tab_before_first_child(doc, p)

                tab_btp_4_at(doc, p)





def copy_char_format(from_fmt, to_fmt):

    to_fmt.FontName = from_fmt.FontName

    to_fmt.FontSize = from_fmt.FontSize

    to_fmt.Bold = from_fmt.Bold

    to_fmt.Italic = from_fmt.Italic

    to_fmt.TextColor = from_fmt.TextColor

    to_fmt.UnderlineStyle = from_fmt.UnderlineStyle





def xoa_gach_chan_tab_trong_spire(doc):

    section = doc.Sections[0]

    labels = [

        'A.',

        'B.',

        'C.',

        'D.',

        'a)',

        'b)',

        'c)',

        'd)']

    for p in range(section.Paragraphs.Count):

        para = section.Paragraphs[p]

        text = para.Text.strip()

        if not text.startswith(tuple(labels)):

            continue

        for i in range(para.ChildObjects.Count):

            obj = para.ChildObjects[i]

            if not isinstance(obj, TextRange):

                continue

            text = obj.Text

            if not '\t' in obj.Text:

                continue

            parts = text.split('\t', 1)

            before_tab = parts[0]

            after_tab = parts[1] if len(parts) > 1 else ''

            if not before_tab == '':

                continue

            para.ChildObjects.RemoveAt(i)

            if after_tab:

                run_text = TextRange(doc)

                run_text.Text = after_tab

                copy_char_format(obj.CharacterFormat, run_text.CharacterFormat)

                para.ChildObjects.Insert(i, run_text)

            run_tab = TextRange(doc)

            run_tab.Text = '\t'

            run_tab.CharacterFormat.UnderlineStyle = UnderlineStyle.none

            para.ChildObjects.Insert(i, run_tab)





def xoa_red_underline_ABCD_abcd(doc):

    labelf = [

        'A',

        'B',

        'C',

        'D',

        'a',

        'b',

        'c',

        'd']

    labels = [

        'A.',

        'B.',

        'C.',

        'D.',

        'a)',

        'b)',

        'c)',

        'd)']

    section = doc.Sections[0]

    for i in range(section.Paragraphs.Count):

        paragraph = section.Paragraphs[i]

        text = paragraph.Text.strip()

        if not text.startswith(tuple(labels)):

            continue

        child = paragraph.ChildObjects[0]

        if not isinstance(child, TextRange):

            continue

        if not child.Text.strip().startswith(tuple(labelf)):

            continue

        child.CharacterFormat.TextColor = Color.get_Blue()

        child.CharacterFormat.UnderlineStyle = UnderlineStyle.none





def xoa_red_underline_ABCD_abcd_sau(doc):

    labelf = [

        'A',

        'B',

        'C',

        'D',

        'a',

        'b',

        'c',

        'd']

    labels = [

        'A.',

        'B.',

        'C.',

        'D.',

        'a)',

        'b)',

        'c)',

        'd)']

    section = doc.Sections[0]

    for i in range(section.Paragraphs.Count):

        paragraph = section.Paragraphs[i]

        text = paragraph.Text.strip()

        if not text.startswith(tuple(labels)):

            continue

        for j in range(paragraph.ChildObjects.Count):

            child = paragraph.ChildObjects[j]

            if not isinstance(child, TextRange):

                continue

            if not child.Text.strip().startswith(tuple(labelf)):

                continue

            if not is_red_color_spr(child.CharacterFormat.TextColor):

                continue

            child.CharacterFormat.TextColor = Color.get_Blue()

            child.CharacterFormat.UnderlineStyle = UnderlineStyle.none





def Xoa_dong_chua_tu(doc, text):

    for p in range(doc.Sections[0].Body.ChildObjects.Count - 1, -1, -1):

        paragraph = doc.Sections[0].Body.ChildObjects[p]

        if not isinstance(paragraph, Paragraph):

            continue

        if not text in paragraph.Text:

            continue

        doc.Sections[0].Body.ChildObjects.RemoveAt(p)





def Xoa_dong_start_with(doc, text):

    for p in range(doc.Sections[0].Body.ChildObjects.Count - 1, -1, -1):

        paragraph = doc.Sections[0].Body.ChildObjects[p]

        if not isinstance(paragraph, Paragraph):

            continue

        if not paragraph.Text.startswith(text):

            continue

        doc.Sections[0].Body.ChildObjects.RemoveAt(p)





def tim_loi_giai_body_all(doc):

    """Tìm tất cả vị trí chứa 'Question X.' hoặc 'Câu X.' trong tài liệu"""

    body = doc.Sections[0].Body

    sum_body = body.ChildObjects.Count

    labels_S = [

        'lời giải',

        'hướng dẫn',

        'giải']

    labels_E = [

        'lời giải',

        'giải',

        'giải:']

    indices_LG = []

    for p in range(sum_body):

        child = body.ChildObjects[p]

        if not isinstance(child, Paragraph):

            continue

        text = child.Text.strip().lower()

        if not text.startswith(tuple(labels_S)) and text.endswith(tuple(labels_E)):

            continue

        indices_LG.append(p)

    return indices_LG





def tim_cau_body_all(doc):

    pass

# WARNING: Decompyle incomplete





def tao_danh_sach_loigiai_cau(a, b):

    pass

# WARNING: Decompyle incomplete





def xoa_loi_giai_lay_de(doc):

    body = doc.Sections[0].Body

    sum_body = body.ChildObjects.Count

    b = tim_loi_giai_body_all(doc)

    c = tao_danh_sach_loigiai_cau(tim_cau_body_all(doc), b)

    if len(b) != len(c):

        print("⚠️ Không khớp số lượng 'lời giải' và 'câu'")

        return None

    to_remove = set()

    for start, end in zip(b, c):

        to_remove.update(range(start, min(end, sum_body)))

    for i in sorted(to_remove, reverse = True):

        if not i < body.ChildObjects.Count:

            continue

        body.ChildObjects.RemoveAt(i)





def update_made(doc, new_text):

    doc.Replace('<made>', new_text, True, True)





def update_so_trang(doc):

    num_pages = doc.GetPageCount()

    new_text = f'''{num_pages:02}'''

    doc.Replace('<sotrang>', new_text, True, False)





def update_so_trang_path(doc_path):

    doc = SpireDocument()

    doc.LoadFromFile(doc_path)

    num_pages = doc.GetPageCount()

    new_text = f'''{num_pages:02}'''

    doc.Replace('<sotrang>', new_text, True, False)

    save_and_close_spire(doc, doc_path)

    return None

# WARNING: Decompyle incomplete





def save_and_close_spire(doc, output_file):

    doc.SaveToFile(output_file)

    doc.Close()

    return None

# WARNING: Decompyle incomplete





def xoa_loi_giai_acong_DS_red(output_file_loi_giai, output_file):

    doc = open_doc_off_spire(output_file_loi_giai)

    Xoa_dong_chua_tu(doc, '@')

    Xoa_dong_start_with(doc, 'ĐS:')

    xoa_loi_giai_lay_de(doc)

    xoa_red_underline_ABCD_abcd_sau(doc)

    update_so_trang(doc)

    save_and_close_spire(doc, output_file)

    return None

# WARNING: Decompyle incomplete





def chuanhoaspr(doc, output_file, new_text, page_kind, bo_loi_giai = ('A4', 'YES')):

    Danh_STT_Cau_All(doc)

    STT_ABCD(doc)

    STT_abcd_nho(doc)

    STT_thang(doc)

    if bo_loi_giai == 'YES':

        xoa_red_underline_ABCD_abcd(doc)

    if page_kind == '2C':

        sap_xep_phan_I_2_cot(doc)

        page_2_cot_Mix(doc)

        resize_images_to_column_width(doc)

    else:

        sap_xep_phan_I(doc)

    lam_dep_phan_II(doc)

    lam_dep_phan_IV(doc)

    if bo_loi_giai == 'YES':

        Xoa_dong_chua_tu(doc, '@')

        Xoa_dong_start_with(doc, 'ĐS:')

        xoa_loi_giai_lay_de(doc)

        update_so_trang(doc)

    else:

        xoa_gach_chan_tab_trong_spire(doc)

    update_made(doc, new_text)

    save_and_close_spire(doc, output_file)

    return None

# WARNING: Decompyle incomplete





def chuanhoaspr_EN(doc, output_file, new_text, page_kind, bo_loi_giai = ('A4', 'YES')):

    Danh_STT_Cau_All_EN(doc)

    STT_ABCD_EN(doc)

    STT_thang(doc)

    if bo_loi_giai == 'YES':

        xoa_red_underline_ABCD_abcd(doc)

    if page_kind == '2C':

        sap_xep_phan_I_2_cot_EN(doc)

        page_2_cot_Mix(doc)

        resize_images_to_column_width(doc)

    else:

        sap_xep_phan_I_EN(doc)

    if bo_loi_giai == 'YES':

        Xoa_dong_chua_tu(doc, '@')

        xoa_loi_giai_lay_de(doc)

        update_so_trang(doc)

    else:

        xoa_gach_chan_tab_trong_spire(doc)

    update_made(doc, new_text)

    save_and_close_spire(doc, output_file)

    return None

# WARNING: Decompyle incomplete





def chuanhoaspr_Japan(doc, output_file, new_text, page_kind, bo_loi_giai = ('A4', 'YES')):

    Danh_STT_Cau_All_EN(doc)

    STT_ABCD_EN(doc)

    STT_thang(doc)

    if bo_loi_giai == 'YES':

        xoa_red_underline_ABCD_abcd(doc)

    if page_kind == '2C':

        sap_xep_phan_I_2_cot_Japan(doc)

        page_2_cot_Mix(doc)

        resize_images_to_column_width(doc)

    else:

        sap_xep_phan_I_Japan(doc)

    if bo_loi_giai == 'YES':

        Xoa_dong_chua_tu(doc, '@')

        xoa_loi_giai_lay_de(doc)

        update_so_trang(doc)

    else:

        xoa_gach_chan_tab_trong_spire(doc)

    update_made(doc, new_text)

    save_and_close_spire(doc, output_file)

    return None

# WARNING: Decompyle incomplete





def chuanhoaspr_ngoai_ngu(doc, output_file, new_text, page_kind, bo_loi_giai, langue = ('A4', 'YES', 'english')):

    if langue == 'japan':

        chuanhoaspr_Japan(doc, output_file, new_text, page_kind, bo_loi_giai)

        return None

    chuanhoaspr_EN(doc, output_file, new_text, page_kind, bo_loi_giai)



