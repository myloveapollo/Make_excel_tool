import wx
import os
import function


class SiteLog(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self, None, title='新诊制作08月07日版本>>>徐浩峰:13554941602', size=(520, 480))
		self.Center()
		self.OpenFile = wx.Button(self, label='打开考场信息表', pos=(305, 5), size=(130, 25))
		self.OpenFile.Bind(wx.EVT_BUTTON, self.OnOpenFile_one)
		self.FileName_one = wx.TextCtrl(self, pos=(5, 5), size=(290, 25), style=wx.TE_READONLY | wx.TE_RICH2)

		self.OpenFile = wx.Button(self, label='打开完成表的文件夹', pos=(305, 45), size=(130, 25))
		self.OpenFile.Bind(wx.EVT_BUTTON, self.OnOpenFile_two)

		self.MakeExcel1 = wx.Button(self, label='制作', pos=(440, 5), size=(65, 65))
		self.MakeExcel1.Bind(wx.EVT_BUTTON, self.ReadFile)

		self.filesFilter = "Excel files(*.xlsx)|*.xlsx"  # |All files (*.*)|*.*
		self.fileDialog_one = wx.FileDialog(self, message="选择单个文件", wildcard=self.filesFilter, style=wx.FD_OPEN)
		self.fileDialog_two = wx.FileDialog(self, message="选择单个文件", wildcard=self.filesFilter, style=wx.FD_OPEN)

		message = '注意：打开的新诊表必须按指定要求处理后才可以制作。' + '\n' + '欢迎咨询：徐浩峰' + '\n' + '-----------------------------------------------------------'
		self.FileContent = wx.TextCtrl(self, value=message, pos=(5, 130), size=(500, 480),
									   style=(wx.TE_MULTILINE | wx.TE_READONLY))

	def OnOpenFile_one(self, event):
		fileResult = self.fileDialog_one.ShowModal()
		if fileResult != wx.ID_OK:
			return
		self.FileName_one.AppendText("%s" % self.fileDialog_one.GetPath())

	def OnOpenFile_two(self, event):
		os.startfile(str(os.getcwd()))

	def ReadFile(self, event):  # 前台表
		try:
			filename = self.fileDialog_one.GetPath()
			finish_name = '新诊制作完成表.xlsx'
			function.read_data(filename, finish_name)
			message1 = '制作成功！可点击“打开完成表的文件夹”转到该位置>>>位置：' + os.path.abspath(path=finish_name)
			wx.TextCtrl(self, value=message1, pos=(5, 130), size=(500, 480), style=wx.TE_MULTILINE | wx.TE_READONLY)
		except FileNotFoundError:
			message1 = '没有打开文件！请点击打开！'
			wx.TextCtrl(self, value=message1, pos=(5, 130), size=(500, 480), style=wx.TE_MULTILINE)
		except NameError:
			message1 = '打开的表有误！请再次确认打开的excel表,列名的顺序是：年级/学科/班次/教师/辅导老师/教学点/教室/上课时间。且不能有隐藏列。'
			wx.TextCtrl(self, value=message1, pos=(5, 130), size=(500, 480), style=wx.TE_MULTILINE)
		except AssertionError:
			message1 = '打开的表有误！请再次确认打开的excel表,列名的顺序是：年级/学科/班次/教师/辅导老师/教学点/教室/上课时间。且不能有隐藏列。'
			wx.TextCtrl(self, value=message1, pos=(5, 130), size=(500, 480), style=wx.TE_MULTILINE)


if __name__ == '__main__':
	app = wx.App()
	SiteFrame = SiteLog()
	SiteFrame.Show()
	app.MainLoop()