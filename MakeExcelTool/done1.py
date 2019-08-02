# -*-coding:GBK -*-
import wx
import os
import xlrd
import webbrowser as web
import done2
import done3
import done4
import done5
import done6


class SiteLog(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title='Excel制作工具0802版本>>>徐浩峰:13554941602', size=(450, 480))
        self.Center()
        self.OpenFile = wx.Button(self, label='打开', pos=(310, 5), size=(80, 25))
        self.OpenFile.Bind(wx.EVT_BUTTON, self.openfile)
        self.MakeExcel1 = wx.Button(self, label='前台课表目录', pos=(10, 40), size=(120, 25))
        self.MakeExcel1.Bind(wx.EVT_BUTTON, self.readfile)
        self.MakeExcel2 = wx.Button(self, label='制作随材发放条', pos=(140, 40), size=(120, 25))
        self.MakeExcel2.Bind(wx.EVT_BUTTON, self.readfile_a)
        self.MakeExcel3 = wx.Button(self, label='随材需求统计表', pos=(270, 40), size=(120, 25))
        self.MakeExcel3.Bind(wx.EVT_BUTTON, self.readfile_b)
        self.MakeExcel4 = wx.Button(self, label='A5教室门前课表', pos=(10, 75), size=(120, 25))
        self.MakeExcel4.Bind(wx.EVT_BUTTON, self.readfile_c)
        self.MakeExcel5 = wx.Button(self, label='A4教室门前课表', pos=(140, 75), size=(120, 25))
        self.MakeExcel5.Bind(wx.EVT_BUTTON, self.readfile_d)
        self.MakeExcel6 = wx.Button(self, label='排班考勤表制作', pos=(270, 75), size=(120, 25))
        self.MakeExcel6.Bind(wx.EVT_BUTTON, self.readfile_e)

        self.filesFilter = "Excel files(*.xlsx)|*.xlsx"  # |All files (*.*)|*.*
        self.fileDialog = wx.FileDialog(self, message="选择单个文件", wildcard=self.filesFilter, style=wx.FD_OPEN)
        self.FileName = wx.TextCtrl(self, pos=(10, 5), size=(290, 25), style=wx.TE_READONLY | wx.TE_RICH2)

        message_one = u'下载教学视频、最新软件及排班模板>>>>网盘密码: d247'

        self.author = wx.Button(self, label=message_one, pos=(10, 110), size=(380, 25))
        # self.author.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))
        # self.author.SetBackgroundColour(self.BackgroundColour)
        self.author.Bind(wx.EVT_BUTTON, self.OnButton)
        self.FileContent = wx.TextCtrl(self, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def OnButton(self, event):
        web.open("https://pan.baidu.com/s/1VYDtzJpl83JaqhRe0W_zIg")

    def openfile(self, event):
        result = self.fileDialog.ShowModal()
        if result != wx.ID_OK:
            return
        self.FileName.AppendText("%s" % self.fileDialog.GetPath())

    # ~ wx.TextCtrl(self, value=str(self.fileDialog.GetPath()), pos=(5,5), size=(290,25))

    def readfile(self, event):  # 前台表
        try:
            filename = self.fileDialog.GetPath()
            names = done2.wash_data(filename)
            message1 = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '没有打开文件！请点击打开！'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合"前台表课表"要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_a(self, event):  # 制作讲义室随材发放条
        try:
            filename = self.fileDialog.GetPath()
            names = done3.wash_data(filename)
            message2 = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message2, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '请点击打开文件！'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合“制作随材发放条”要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_b(self, event):  # 讲义随材需求量统计表
        try:
            filename = self.fileDialog.GetPath()
            names = done4.wash_data(filename)
            message = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '没有文件被打开'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合“随材需求统计表”要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_c(self, event):  # 教室前课表A5
        try:
            filename = self.fileDialog.GetPath()
            size_a = 'A5'
            names = done5.final_fuc(filename, size_a)
            message = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + 'A5' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '没有文件被打开'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合“教室前课表”要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_d(self, event):  # 教室前课表A4
        try:
            filename = self.fileDialog.GetPath()
            size_a = 'A4'
            names = done5.final_fuc(filename, size_a)
            message = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + 'A4' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '没有文件被打开'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合“教室前课表”要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_e(self, event):  # 排班考勤表制作
        try:
            filename = self.fileDialog.GetPath()
            names = done6.wash_data(filename)
            message = '制作成功！>>>位置：' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '没有文件被打开'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '打开的Excel表不符合要求'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)


if __name__ == '__main__':
    app = wx.App()
    SiteFrame = SiteLog()
    SiteFrame.Show()
    app.MainLoop()
