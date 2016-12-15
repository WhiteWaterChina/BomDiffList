#!/usr/bin/env python
# -*- coding:cp936 -*-
###################################
# Author:yanshuo
# Email:yanshuo5091@163.com
#  Version:1.1

import Tkinter
import tkMessageBox
import ttk
import tkFileDialog
import os
import xlsxwriter
import xlrd

filename_original_bom = unicode()
filename_original_now = unicode()
dir_filename_display = unicode()
data_description = [u'CP', u'CL', u'FN', u'ME', u'MB', u'OM', u'CS', u'IO', u'CD', u'FD', u'KB', u'MS', u'MI', u'LP', u'PB', u'CB',
                    u'VI', u'MD', u'AC', u'RA', u'CT', u'NC', u'HD', u'HM', u'HT', u'HH', u'HF', u'HX', u'MT', u'PS', u'PM', u'PP',
                    u'PF', u'PK', u'GA', u'DS', u'FL', u'TA', u'DG', u'SY', u'CM', u'TJ', u'OS', u'FW', u'SW', u'PU', u'RK', u'SA',
                    u'BB', u'SB', u'BR', u'TC', u'MC', u'PN', u'MY', u'FG', u'FF', u'JS', u'ZJ', u'TM', u'QT', u'WC', u'KO', u'SL',
                    u'HA', u'LK', u'MU', u'HB', u'HC', u'TP', u'DW', u'SD', u'ST', u'SF', u'CR', u'CA', u'PD', u'UP', u'DC', u'CN',
                    u'PJ', u'PG', u'DM', u'MA', u'MM', u'SM', u'SP', u'MZ', u'SC',u'ZZB']

root = Tkinter.Tk()
root.title("BOM转兼容性列表对比工具".decode('gbk'))
root.geometry('800x600')
root.resizable(width=True, height=True)
var_char_entry_filename_bom_need_filter = Tkinter.StringVar()
var_char_entry_filename_now_need_filter = Tkinter.StringVar()
var_char_combox_sheet_now = Tkinter.StringVar()
var_char_entry_filename_after_filter = Tkinter.StringVar()


def get_filename_bom():
    global filename_original_bom
    filename_bom = tkFileDialog.askopenfilename()
    var_char_entry_filename_bom_need_filter.set(filename_bom)
    filename_original_bom = filename_bom


def get_filename_now():
    global filename_original_now
    filename_now = tkFileDialog.askopenfilename()
    var_char_entry_filename_now_need_filter.set(filename_now)
    filename_original_now = filename_now


def set_filename_output():
    global dir_filename_display
    dir_filename_display = tkFileDialog.askdirectory().replace('/', '\\')
    var_char_entry_filename_after_filter.set(dir_filename_display)


def get_data():
    data_bom_pn = []
    data_bom = {}
    data_now_pn = []
    data_display = {}
    filename_output = os.path.join(dir_filename_display, "BOM和兼容性列表的不同处.xlsx".decode('gbk'))
    bom_workbook = xlrd.open_workbook(filename=filename_original_bom)
    sheet_bom = bom_workbook.sheet_by_index(0)
    nrows = sheet_bom.nrows
    for item in range(0, nrows - 1):
        level = unicode(sheet_bom.cell(item, 1).value)
        pn = sheet_bom.cell(item, 8).value
        description = sheet_bom.cell(item, 10).value
        description_number = sheet_bom.cell(item, 7).value
        if level == u'2.0' and description_number in data_description and pn not in data_bom_pn:
            data_bom_pn.append(pn)
            print pn
            data_bom['%s' % pn] = description

    now_workbook = xlrd.open_workbook(filename=filename_original_now)
    sheet_now = now_workbook.sheet_by_index(int(var_char_combox_sheet_now.get()))
    nrows = sheet_now.nrows
    for item in range(1, nrows - 1):
        pn_now = sheet_now.cell(item, 1).value
        if pn_now not in data_now_pn:
            data_now_pn.append(pn_now)

    for item in data_bom_pn:
        if item not in data_now_pn:
            data_display['%s' % item] = data_bom['%s' % item]
    WorkBook = xlsxwriter.Workbook(filename_output)
    sheetone = WorkBook.add_worksheet('sheet1')
    format_workbook = WorkBook.add_format()
    format_workbook.set_border(1)
    count = 0
    for i in data_display.keys():
        sheetone.write(count, 0, i, format_workbook)
        sheetone.write(count, 1, data_display[i], format_workbook)
        count += 1
    WorkBook.close()
    tkMessageBox.showinfo('提示'.decode('gbk'),
                          '处理BOM和兼容性列表的不同处的结果已经生成，请去%s路径查看.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % filename_output)


frame_top = Tkinter.Frame(root, height=20)
frame_top.pack(side=Tkinter.TOP)
frame_top_top = Tkinter.Frame(frame_top, height=40)
frame_top_top.pack()
frame_top_bottom = Tkinter.Frame(frame_top, height=20)
frame_top_bottom.pack()
Tkinter.Label(frame_top_top, text='请在如下选择需要处理的BOM文件'.decode('gbk'), bg='Red').pack()
Tkinter.Entry(frame_top_bottom, textvariable=var_char_entry_filename_bom_need_filter, width=40).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_top_bottom, text='选择文件'.decode('gbk'), command=get_filename_bom, width=20).pack(side=Tkinter.RIGHT)

frame_middle = Tkinter.Frame(root, height=20)
frame_middle.pack(side=Tkinter.TOP)
frame_middle_top = Tkinter.Frame(frame_middle, height=40)
frame_middle_top.pack()
frame_middle_bottom = Tkinter.Frame(frame_middle, height=20)
frame_middle_bottom.pack()
Tkinter.Label(frame_middle_top, text='请在如下选择需要处理的兼容性列表文件'.decode('gbk'), bg='Red').pack()
Tkinter.Entry(frame_middle_bottom, textvariable=var_char_entry_filename_now_need_filter, width=40).pack(
    side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom, text='选择文件'.decode('gbk'), command=get_filename_now, width=20).pack(
    side=Tkinter.RIGHT)

frame_middle_2 = Tkinter.Frame(root, height=20)
frame_middle_2.pack()
frame_middle_top_2 = Tkinter.Frame(frame_middle_2, height=40)
frame_middle_top_2.pack()
frame_middle_bottom_2 = Tkinter.Frame(frame_middle_2, height=20)
frame_middle_bottom_2.pack()
Tkinter.Label(frame_middle_top_2, text='请在如下选择处理结果存放的位置'.decode('gbk'), bg='Red').pack()
Tkinter.Entry(frame_middle_bottom_2, textvariable=var_char_entry_filename_after_filter, width=40).pack(
    side=Tkinter.LEFT)
Tkinter.Button(frame_middle_bottom_2, text='选择文件'.decode('gbk'), command=set_filename_output, width=20).pack(
    side=Tkinter.RIGHT)

frame_middle_3 = Tkinter.Frame(root, height=50)
frame_middle_3.pack()
frame_middle_3_top = Tkinter.Frame(frame_middle_3, height=20)
frame_middle_3_top.pack()
frame_middle_3_bottom = Tkinter.Frame(frame_middle_3, height=20)
frame_middle_3_bottom.pack()
Tkinter.Label(frame_middle_3_top, text='请在如下选择需要处理的兼容性列表的第几个Sheet,从0开始'.decode('gbk'), bg='Red').pack(side=Tkinter.TOP)
box_set_sheet = ttk.Combobox(frame_middle_3_bottom, textvariable=var_char_combox_sheet_now,
                             values=['0', '1', '2', '3', '4', '5', '6'], width=30)
box_set_sheet.pack(side=Tkinter.BOTTOM)
frame_bottom = Tkinter.Frame(root, height=20)
frame_bottom.pack()

Tkinter.Button(frame_bottom, text='GO'.decode('gbk'), width=20, command=get_data).pack(side=Tkinter.LEFT)
Tkinter.Button(frame_bottom, text='退出'.decode('gbk'), width=20, command=root.destroy).pack(side=Tkinter.LEFT)

Tkinter.mainloop()
