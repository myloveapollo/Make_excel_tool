#保洁开门图
import numpy as np#导入nump数据模块
import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill,Border,Side,Alignment,Protection,Font
pd.options.mode.chained_assignment = None

# ~ def twonum(nums, target):
	# ~ d = {}
	# ~ f =[]
	# ~ for i,item in enumerate(nums):
		# ~ tmp + item = target
		# ~ d[i] = item
		# ~ for k,v in d.items():
			# ~ if tmp == v:
				# ~ f.append((i,k))
	# ~ return f 


def read_data(file_name, finish_name):
	data = pd.read_excel(file_name, sheet_name='Sheet0', usecols=[3, 4, 17, 18, 22])#读取表的内容
	writer = pd.ExcelWriter(finish_name)
	writer1 = pd.ExcelWriter("生成的表.xlsx")
	data.教室 = data.教室.str.replace('.*[\u4e00-\u9fa5]|([).*(])|(【).*(】)','')#保留教室号座位第一列
	data.上课时间 = data.上课时间.str.replace('.*[\u4e00-\u9fa5]', '')#保留上课时间作为行
	
	data_room = sorted(set(list(data.教室)))#教室作为第一列
	data_time_one = sorted(set(list(data.上课时间)))
	
	data_time = []
	for n in data_time_one:
		a = n.split('-')
		data_time.append(a[0])
		data_time.append(a[1])
	data_time = sorted(data_time)#时间作为行
	
	nan_15 = list(np.nan for i in range(0,15))
	makedata = pd.DataFrame(np.zeros(shape=(14,len(data_time))), index=data_room,columns=data_time)
	
	data_list_location = []
	for index in data_room:
		data_chroom = data.loc[data.教室.str.contains(index)]
		for n in data_chroom.上课时间:
			a = n.split('-')
			if a[0] in data_time:

	# ~ data.to_excel(writer, index=False, header=True)
	# ~ makedata.to_excel(writer1, index=True, header=True)
	# ~ writer.save()
	# ~ writer1.save()
	
filename = '桂庙暑假.xlsx'
finishname = '桂庙暑假时间图.xlsx'
read_data(filename, finishname)
