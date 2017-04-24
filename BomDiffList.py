#!/usr/bin/env python
# -*- coding:cp936 -*-
###################################
# Author:yanshuo
# Email:yanshuo5091@163.com
#  Version:1.1

import wx
import wx.xrc
import time
import os
import xlsxwriter
import xlrd

filename_original_bom = unicode()
filename_original_now = unicode()
dir_filename_output = unicode()
textCtrl_bom = wx.TextCtrl
textCtrl_list = wx.TextCtrl
textCtrl_dir = wx.TextCtrl
data_description = [u'CP', u'CL', u'FN', u'ME', u'MB', u'OM', u'CS', u'IO', u'CD', u'FD', u'KB', u'MS', u'MI', u'LP',
                    u'PB', u'CB',
                    u'VI', u'MD', u'AC', u'RA', u'CT', u'NC', u'HD', u'HM', u'HT', u'HH', u'HF', u'HX', u'MT', u'PS',
                    u'PM', u'PP',
                    u'PF', u'PK', u'GA', u'DS', u'FL', u'TA', u'DG', u'SY', u'CM', u'TJ', u'OS', u'FW', u'SW', u'PU',
                    u'RK', u'SA',
                    u'BB', u'SB', u'BR', u'TC', u'MC', u'PN', u'MY', u'FG', u'FF', u'JS', u'ZJ', u'TM', u'QT', u'WC',
                    u'KO', u'SL',
                    u'HA', u'LK', u'MU', u'HB', u'HC', u'TP', u'DW', u'SD', u'ST', u'SF', u'CR', u'CA', u'PD', u'UP',
                    u'DC', u'CN',
                    u'PJ', u'PG', u'DM', u'MA', u'MM', u'SM', u'SP', u'MZ', u'SC', u'ZZB']


class BomDiffList(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=wx.EmptyString, pos=wx.DefaultPosition,
                          size=wx.Size(540, 342), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)

        bSizer7 = wx.BoxSizer(wx.VERTICAL)

        bSizer8 = wx.BoxSizer(wx.VERTICAL)

        self.text_1 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择需要处理的BOM文件", wx.DefaultPosition, wx.DefaultSize, 0)
        self.text_1.Wrap(-1)
        self.text_1.SetFont(wx.Font(11, 70, 90, 90, False, "宋体"))
        self.text_1.SetForegroundColour(wx.Colour(255, 255, 0))
        self.text_1.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer8.Add(self.text_1, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7.Add(bSizer8, 1, wx.EXPAND, 5)

        bSizer9 = wx.BoxSizer(wx.HORIZONTAL)

        self.textCtrl_bom = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer9.Add(self.textCtrl_bom, 1, wx.ALL, 5)

        self.button_bom = wx.Button(self, wx.ID_ANY, u"选择BOM文件", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer9.Add(self.button_bom, 0, wx.ALL, 5)

        bSizer7.Add(bSizer9, 1, wx.EXPAND, 5)

        bSizer10 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择需要对比的兼容性列表文件（一定要符合网站上传要求）", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText4.Wrap(-1)
        self.m_staticText4.SetFont(wx.Font(11, 70, 90, 90, False, "宋体"))
        self.m_staticText4.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText4.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer10.Add(self.m_staticText4, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7.Add(bSizer10, 1, wx.EXPAND, 5)

        bSizer11 = wx.BoxSizer(wx.HORIZONTAL)

        self.textCtrl_list = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer11.Add(self.textCtrl_list, 1, wx.ALL, 5)

        self.button_list = wx.Button(self, wx.ID_ANY, u"选择兼容性列表", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer11.Add(self.button_list, 0, wx.ALL, 5)

        bSizer7.Add(bSizer11, 1, wx.EXPAND, 5)

        bSizer12 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText5 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择存放比对结果的文件路径", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText5.Wrap(-1)
        self.m_staticText5.SetFont(wx.Font(11, 70, 90, 90, False, "宋体"))
        self.m_staticText5.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText5.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer12.Add(self.m_staticText5, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7.Add(bSizer12, 1, wx.EXPAND, 5)

        bSizer13 = wx.BoxSizer(wx.HORIZONTAL)

        self.textCtrl_dir = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer13.Add(self.textCtrl_dir, 1, wx.ALL, 5)

        self.button_dir = wx.Button(self, wx.ID_ANY, u"选择存储路径", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer13.Add(self.button_dir, 0, wx.ALL, 5)

        bSizer7.Add(bSizer13, 1, wx.EXPAND, 5)

        bSizer14 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText7 = wx.StaticText(self, wx.ID_ANY, u"请在如下选择要对比兼容性列表的第几个Sheet（从0开始算）", wx.DefaultPosition,
                                           wx.DefaultSize, 0)
        self.m_staticText7.Wrap(-1)
        self.m_staticText7.SetFont(wx.Font(11, 70, 90, 90, False, "宋体"))
        self.m_staticText7.SetForegroundColour(wx.Colour(255, 255, 0))
        self.m_staticText7.SetBackgroundColour(wx.Colour(0, 128, 0))

        bSizer14.Add(self.m_staticText7, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer7.Add(bSizer14, 1, wx.EXPAND, 5)

        bSizer15 = wx.BoxSizer(wx.VERTICAL)

        comboBox_sheetChoices = [u"0", u"1", u"2", u"3", u"4", u"5", u"6", u"7", u"8", u"9"]
        self.comboBox_sheet = wx.ComboBox(self, wx.ID_ANY, u"2", wx.DefaultPosition, wx.DefaultSize,
                                          comboBox_sheetChoices, 0)
        self.comboBox_sheet.SetSelection(10)
        bSizer15.Add(self.comboBox_sheet, 0, wx.ALL | wx.EXPAND, 5)

        bSizer7.Add(bSizer15, 1, wx.EXPAND, 5)

        bSizer16 = wx.BoxSizer(wx.HORIZONTAL)

        self.button_go = wx.Button(self, wx.ID_ANY, u"GO", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer16.Add(self.button_go, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)

        self.button_exit = wx.Button(self, wx.ID_ANY, u"EXIT", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer16.Add(self.button_exit, 0, wx.ALL, 5)

        bSizer7.Add(bSizer16, 1, wx.ALIGN_CENTER, 5)

        self.SetSizer(bSizer7)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.button_bom.Bind(wx.EVT_BUTTON, self.get_filename_bom)
        self.button_list.Bind(wx.EVT_BUTTON, self.get_filename_now)
        self.button_dir.Bind(wx.EVT_BUTTON, self.set_filename_output)
        self.button_go.Bind(wx.EVT_BUTTON, self.get_data)
        self.button_exit.Bind(wx.EVT_BUTTON, self.close)

    def __del__(self):
        pass

    def close(self, event):
        self.Close()

    # Virtual event handlers, overide them in your derived class
    def get_filename_bom(self, event):
        global filename_original_bom
        filename_bom_dialog = wx.FileDialog(self, message=u"选择需要对比的BOM文件", defaultDir=os.getcwd(), defaultFile="", style=wx.OPEN)
        if filename_bom_dialog.ShowModal() == wx.ID_OK:
            filename_bom = filename_bom_dialog.GetPath()
            self.textCtrl_bom.SetValue(filename_bom)
            filename_original_bom = filename_bom
            filename_bom_dialog.Destroy()

    def get_filename_now(self, event):
        global filename_original_now
        filename_now_dialog = wx.FileDialog(self, message=u"选择需要对比的兼容性列表文件", defaultDir=os.getcwd(), defaultFile="", style=wx.OPEN)
        if filename_now_dialog.ShowModal() == wx.ID_OK:
            filename_now = filename_now_dialog.GetPath()
            self.textCtrl_list.SetValue(filename_now)
            filename_original_now = filename_now
            filename_now_dialog.Destroy()

    def set_filename_output(self, event):
        global dir_filename_output
        dir_filename_output_dialog = wx.DirDialog(self, message=u"选择存储路径", style=wx.DD_DEFAULT_STYLE)
        if dir_filename_output_dialog.ShowModal() == wx.ID_OK:
            dir_filename_output = dir_filename_output_dialog.GetPath()
            self.textCtrl_dir.SetValue(dir_filename_output)
            dir_filename_output_dialog.Destroy()

    def get_data(self, event):
        data_bom_pn = []
        data_bom = {}
        data_now_pn = []
        data_display = {}
        timestamp = time.strftime('%Y%m%d', time.localtime())
        filename_output_temp = "BOM和兼容性列表的不同处-%s.xlsx".decode('gbk') % timestamp
        filename_output = os.path.join(dir_filename_output, filename_output_temp)
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
                data_bom['%s' % pn] = description

        now_workbook = xlrd.open_workbook(filename=filename_original_now)
        sheet_now = now_workbook.sheet_by_index(int(self.comboBox_sheet.GetValue()))
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

        diag_finish = wx.MessageDialog(None, '处理BOM和兼容性列表的不同处的结果已经生成，请去%s《%s》路径查看.如果无需其他动作，请点击退出按钮退出程序'.decode('gbk') % (dir_filename_output, filename_output_temp), '提示'.decode('gbk'), wx.OK | wx.ICON_INFORMATION | wx.STAY_ON_TOP)
        diag_finish.ShowModal()


if __name__ == '__main__':
    app = wx.App()
    frame = BomDiffList(None)
    frame.Show()
    app.MainLoop()
