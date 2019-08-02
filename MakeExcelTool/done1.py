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
        wx.Frame.__init__(self, None, title='Excel��������0802�汾>>>��Ʒ�:13554941602', size=(450, 480))
        self.Center()
        self.OpenFile = wx.Button(self, label='��', pos=(310, 5), size=(80, 25))
        self.OpenFile.Bind(wx.EVT_BUTTON, self.openfile)
        self.MakeExcel1 = wx.Button(self, label='ǰ̨�α�Ŀ¼', pos=(10, 40), size=(120, 25))
        self.MakeExcel1.Bind(wx.EVT_BUTTON, self.readfile)
        self.MakeExcel2 = wx.Button(self, label='������ķ�����', pos=(140, 40), size=(120, 25))
        self.MakeExcel2.Bind(wx.EVT_BUTTON, self.readfile_a)
        self.MakeExcel3 = wx.Button(self, label='�������ͳ�Ʊ�', pos=(270, 40), size=(120, 25))
        self.MakeExcel3.Bind(wx.EVT_BUTTON, self.readfile_b)
        self.MakeExcel4 = wx.Button(self, label='A5������ǰ�α�', pos=(10, 75), size=(120, 25))
        self.MakeExcel4.Bind(wx.EVT_BUTTON, self.readfile_c)
        self.MakeExcel5 = wx.Button(self, label='A4������ǰ�α�', pos=(140, 75), size=(120, 25))
        self.MakeExcel5.Bind(wx.EVT_BUTTON, self.readfile_d)
        self.MakeExcel6 = wx.Button(self, label='�Ű࿼�ڱ�����', pos=(270, 75), size=(120, 25))
        self.MakeExcel6.Bind(wx.EVT_BUTTON, self.readfile_e)

        self.filesFilter = "Excel files(*.xlsx)|*.xlsx"  # |All files (*.*)|*.*
        self.fileDialog = wx.FileDialog(self, message="ѡ�񵥸��ļ�", wildcard=self.filesFilter, style=wx.FD_OPEN)
        self.FileName = wx.TextCtrl(self, pos=(10, 5), size=(290, 25), style=wx.TE_READONLY | wx.TE_RICH2)

        message_one = u'���ؽ�ѧ��Ƶ������������Ű�ģ��>>>>��������: d247'

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

    def readfile(self, event):  # ǰ̨��
        try:
            filename = self.fileDialog.GetPath()
            names = done2.wash_data(filename)
            message1 = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = 'û�д��ļ��������򿪣�'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel������"ǰ̨��α�"Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_a(self, event):  # ������������ķ�����
        try:
            filename = self.fileDialog.GetPath()
            names = done3.wash_data(filename)
            message2 = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message2, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = '�������ļ���'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel�����ϡ�������ķ�������Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_b(self, event):  # �������������ͳ�Ʊ�
        try:
            filename = self.fileDialog.GetPath()
            names = done4.wash_data(filename)
            message = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = 'û���ļ�����'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel�����ϡ��������ͳ�Ʊ�Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_c(self, event):  # ����ǰ�α�A5
        try:
            filename = self.fileDialog.GetPath()
            size_a = 'A5'
            names = done5.final_fuc(filename, size_a)
            message = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + 'A5' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = 'û���ļ�����'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel�����ϡ�����ǰ�α�Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_d(self, event):  # ����ǰ�α�A4
        try:
            filename = self.fileDialog.GetPath()
            size_a = 'A4'
            names = done5.final_fuc(filename, size_a)
            message = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + 'A4' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = 'û���ļ�����'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel�����ϡ�����ǰ�α�Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)

    def readfile_e(self, event):  # �Ű࿼�ڱ�����
        try:
            filename = self.fileDialog.GetPath()
            names = done6.wash_data(filename)
            message = '�����ɹ���>>>λ�ã�' + str(os.getcwd()) + '\\' + names
            wx.TextCtrl(self, value=message, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except FileNotFoundError:
            message1 = 'û���ļ�����'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)
        except xlrd.biffh.XLRDError:
            message1 = '�򿪵�Excel������Ҫ��'
            wx.TextCtrl(self, value=message1, pos=(5, 140), size=(430, 480), style=wx.TE_MULTILINE)


if __name__ == '__main__':
    app = wx.App()
    SiteFrame = SiteLog()
    SiteFrame.Show()
    app.MainLoop()
