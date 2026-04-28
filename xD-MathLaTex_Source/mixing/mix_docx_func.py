from docx import Document

from docx.text.paragraph import Paragraph

from docx.shared import RGBColor

from docx.shared import Inches, Pt

from docx.oxml import OxmlElement

from docx.oxml.ns import qn

from lxml import etree



ElementTree

from xml.etree.ElementTree import QName

import xml.etree.ElementTree, etree

import os

import re

import random

import zipfile

import bisect

import tkinter as tk

from tkinter import ttk, Label, filedialog, Entry, Button, Frame, Listbox, Scrollbar, messagebox

from tkinter import simpledialog



def find_phan(doc, text_start, text_end):

    '''Tìm chỉ mục của các đoạn văn bắt đầu bằng text_start và text_end'''

    indices_S = []

    indices_E = []

    body = doc.element.body

    paragraphs = list(body.iterchildren())

    for i, para in enumerate(paragraphs):

        if not para.tag.endswith('p'):

            continue

        paragraph = Paragraph(para, doc)

        text = paragraph.text

        if text.startswith(text_start):

            indices_S.append(i)

            continue

        if not text.startswith(text_end):

            continue

        indices_E.append(i)

    return (indices_S, indices_E)





def tim_cau_trong_phan(doc, text_start, text_end):

    '''Tìm danh sách các câu hỏi trong khoảng từ text_start đến text_end'''

    (indices_S, indices_E) = find_phan(doc, text_start, text_end)

    all_indices_cau = []

    if len(indices_S) != len(indices_E):

        messagebox.showinfo('Thông báo', f'''Số kí hiệu {text_start} và {text_end} không bằng nhau, hãy chỉnh sửa phù hợp''')

        return None

    if len(indices_S) > 0:

        pattern1 = '^Câu [0-9]{1,}'

        body = doc.element.body

        paragraphs = list(body.iterchildren())

        for start, end in zip(indices_S, indices_E):

            indices_cau = []

            for j in range(start, end):

                para = paragraphs[j]

                if not para.tag.endswith('p'):

                    continue

                paragraph = Paragraph(para, doc)

                text = paragraph.text

                if not re.match(pattern1, text):

                    continue

                indices_cau.append(j)

            indices_cau.append(end)

            all_indices_cau.append(indices_cau)

    return all_indices_cau





def tim_phuong_an_trong_cau_phan_I(doc, from_a, to_b):

    """Tìm phương án trong câu hỏi theo định dạng 'A.', 'B.', 'C.', 'D.' từ from_a đến to_b"""

    body = doc.element.body

    paragraphs = list(body.iterchildren())

    to_end = to_b

    for j in range(from_a, to_b):

        para = paragraphs[j]

        if not para.tag.endswith('p'):

            continue

        paragraph = Paragraph(para, doc)

        text = paragraph.text.strip()

        if not text == 'Lời giải':

            continue

        to_end = j

        range(from_a, to_b)

    labels = [

        'A.',

        'B.',

        'C.',

        'D.']

    indices_phuong_an = []

    for j in range(from_a, min(to_end, len(paragraphs))):

        para = paragraphs[j]

        if not para.tag.endswith('p'):

            continue

        paragraph = Paragraph(para, doc)

        text = paragraph.text

        if not text.startswith(tuple(labels)):

            continue

        indices_phuong_an.append(j)

    indices_phuong_an.append(to_end)

    return indices_phuong_an





def tim_phuong_an_trong_cau_phan_II(doc, from_a, to_b):

    """Tìm phương án trong câu hỏi theo định dạng 'a)', 'b)', 'c)', 'd)' từ from_a đến to_b"""

    body = doc.element.body

    paragraphs = list(body.iterchildren())

    to_end = to_b

    for j in range(from_a, to_b):

        para = paragraphs[j]

        if not para.tag.endswith('p'):

            continue

        paragraph = Paragraph(para, doc)

        text = paragraph.text.strip()

        if not text == 'Lời giải':

            continue

        to_end = j

        range(from_a, to_b)

    labels = [

        'a)',

        'b)',

        'c)',

        'd)']

    indices_phuong_an = []

    for j in range(from_a, min(to_end, len(paragraphs))):

        para = paragraphs[j]

        if not para.tag.endswith('p'):

            continue

        paragraph = Paragraph(para, doc)

        text = paragraph.text

        if not text.startswith(tuple(labels)):

            continue

        indices_phuong_an.append(j)

    indices_phuong_an.append(to_end)

    return indices_phuong_an





def find_phan_EN(doc):

    text_start = '<S@>'

    text_end = '<E@>'

    (indices_S, indices_E) = find_phan(doc, text_start, text_end)

    return (indices_S, indices_E)





def tim_cau_trong_phan_EN(doc):

    (indices_S, indices_E) = find_phan_EN(doc)

    all_indices_cau = []

    if len(indices_S) != len(indices_E):

        messagebox.showinfo('Thông báo', 'Số kí hiệu <S@> và <E@> không bằng nhau, hãy chỉnh sửa phù hợp')

        return None

    if len(indices_S) > 0:

        pattern1 = '^Question [0-9]{1,}'

        pattern2 = '^Câu [0-9]{1,}'

        body = doc.element.body

        paragraphs = list(body.iterchildren())

        for start, end in zip(indices_S, indices_E):

            indices_cau = []

            for j in range(start, end):

                para = paragraphs[j]

                if not para.tag.endswith('p'):

                    continue

                paragraph = Paragraph(para, doc)

                text = paragraph.text

                if not re.match(pattern1, text) and re.match(pattern2, text):

                    continue

                indices_cau.append(j)

            indices_cau.append(end)

            all_indices_cau.append(indices_cau)

    return all_indices_cau





def tim_phuong_an_trong_cau_phan_I_EN(doc, from_a, to_b):

    indices_phuong_an = tim_phuong_an_trong_cau_phan_I(doc, from_a, to_b)

    return indices_phuong_an



