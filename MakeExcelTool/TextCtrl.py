import wx
import done1


class MyDialog(wx.Dialog):
	def __init__(self, parent, text):
		wx.Dialog.__init__(self, parent, -1, u'版本信息', pos = wx.DefaultPosition, size=(500, 300))
		sizer = wx.GridSizer(rows=5, cols=1)
		label_1 = wx.StaticText(self, -1, text)
		label_1.SetFont(wx.Font(14, wx.DEFAULT, wx.NORMAL, wx.BOLD))
		label_2 = wx.StaticText(self, -1, u'软件版本：V1.0(2018.03.01')
		label_2.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))
		label_3 = wx.StaticText(self, -1, u'版权所有：maydolly')
		label_3.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))
		label_4 = wx.TextCtrl(self, -1, u'联系作者：www.classnotes.cn', size=(300, -1), style=wx.TE_READONLY | wx.TE_AUTO_URL | wx.TE_RICH | wx.BORDER_NONE)
		label_4.SetFont(wx.Font(12, wx.DEFAULT, wx.NORMAL, wx.NORMAL))
		label_4.SetBackgroundColour(self.BackgroundColour)
		label_4.Bind(wx.EVT_TEXT_URL, self.OnButton)
		self.number = 1
		okbtn = wx.Button(self, wx.ID_OK,'OK')
		okbtn.SetDefault()
		sizer.Add(label_1, flag=wx.ALIGN_CENTER)
		sizer.Add(label_2, flag=wx.ALIGN_CENTER)
		sizer.Add(label_3, flag=wx.ALIGN_CENTER)
		sizer.Add(label_4, flag=wx.ALIGN_CENTER)
		sizer.Add(okbtn, flag=wx.ALIGN_CENTER)
		self.SetSizer(sizer)

	def OnButton(self, evt):
		if evt.GetMouseEvent().LeftIsDown():
			webbrowser.open('http://www.classnotes.cn')


if __name__ == '__main__':
	app = wx.App()
	SiteFrame = done1.SiteLog()
	MyDialog = MyDialog()
	SiteFrame.Show(MyDialog)
	app.MainLoop()
