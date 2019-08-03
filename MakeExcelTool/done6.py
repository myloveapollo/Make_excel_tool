# -*-coding:GBK -*-
# 运营考勤排班导入表
import re
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, numbers
from openpyxl.worksheet.datavalidation import DataValidation

pd.options.mode.chained_assignment = None


def cell_style(ws, len_index):
    width_dict = {'A': 15.13, 'B': 15.13, 'C': 19.28, 'D': 15.13, 'E': 15.13, 'F': 15.13}
    font = Font(name='宋体', size=11, bold=False)
    font_red = Font(color='FF0000')

    dv_type = DataValidation(type="list", formula1='"年假,工作,休息,入离职缺勤,培训,病假,医疗期,事假,婚假,产假,产检,哺乳假,丧假,陪产假"', allow_blank=False)
    ws.add_data_validation(dv_type)
    type_index = 'D2:D' + str(len_index + 1)
    dv_type.add(type_index)

    for row in ws.iter_rows(min_row=2, max_row=len_index + 1, min_col=1, max_col=1):  # 第一列格式并对不是6和8字符的格式判定
        for cell in row:
            if len(cell.value) == 6 or len(cell.value) == 8:
                cell.font = font
            else:
                cell.font = font_red
            cell.number_format = '@'

    for row in ws.iter_rows(min_row=2, max_row=len_index + 1, min_col=2, max_col=3):  # 2/3文本格式
        for cell in row:
            cell.font = font
            cell.number_format = '@'  # 'yyyy-mm-dd'

    for row in ws.iter_rows(min_row=2, max_row=len_index + 1, min_col=4, max_col=4):  # 4文本格式
        for cell in row:
            cell.font = font

    for row in ws.iter_rows(min_row=2, max_row=len_index + 1, min_col=5, max_col=6):  # 5/6列格式
        for cell in row:
            cell.number_format = '@'
            if len(str(cell.value)) != 5:
                cell.font = font_red
            else:
                cell.font = font

    for k, v in width_dict.items():
        ws.column_dimensions[k].width = v


def wash_data(file_name_name):
    names = ['工号', '姓名', '周一', '周二', '周三', '周四', '周五', '周六', '周日']

    # 制作表名
    data_name = pd.read_excel(file_name_name, sheet_name=0, header=0, names=None, usecols=[0])
    column = []
    for col in data_name:
        column.append(col)
    column_name = column[0].replace(' ', '').replace("服务中心", '').replace("排班表", '')
    in_excel = column_name + '考勤导入表.xlsx'  # 命名最后生成的表  str(data_name.columns)

    data = pd.read_excel(file_name_name, sheet_name=0, header=None, names=names,
                         usecols=[1, 2, 3, 6, 9, 12, 15, 18, 21])  # 读取表,
    data.replace('：', ':', regex=True, inplace=True)  # 将：替换为:
    data_time = data.iloc[1]  # 提取日期已备用,格式要正常的

    data.drop(axis=0, index=[1, 0, 2, 3, 4], inplace=True)  # 删除没有的行信息
    data.dropna(axis=0, how='all', inplace=True)  # 删除全空的行信息
    work_type = ['OFF', '年假', '入离职缺勤',  '病假', '医疗期', '事假', '婚假', '产假', '产检', '哺乳假', '丧假', '陪产假' ]
    for dub in work_type:
        data.replace(dub, dub+'-'+dub, regex=True, inplace=True)  # OFF项等于空 正则表达式子，不区分大小写
    data.replace(np.nan, '此处排班不能为空的', regex=True, inplace=True)
    data = data.T  # 倒置

    # 起草一个表
    data_make = pd.DataFrame(np.full([7 * len(data.columns), 6], np.nan),
                             columns=['工号', '姓名', '日期(YYYY-MM-DD)', '类型', '上班时间', '下班时间'])

    data_job_num = []
    for nam in list(data.iloc[0]):
        for i in range(1, 8):
            data_job_num.append(nam)
    data_make['工号'] = data_job_num

    # 获取姓名
    data_col_name = []
    for nam in list(data.iloc[1]):
        for i in range(1, 8):
            data_col_name.append(nam)
    data_make['姓名'] = pd.Series(data_col_name)  # 导入姓名

    # 获取时间
    data.drop(axis=0, index='姓名', inplace=True)
    data.drop(axis=0, index='工号', inplace=True)
    data_col_arrivetime = []  # 上班时间
    data_col_leavetime = []  # 下班时间
    data_type = []  # 类型

    for col in list(data.columns):
        data_split_f = data[col].str.split('-').str[0]  # 上班时间用前面的
        data_split_s = data[col].str.split('-').str[1]  # 下班时间用后面的
        for f in list(data_split_f):
            if len(str(f)) == 4:  # len(f) == 4:
                f = '0' + f
            data_col_arrivetime.append(f)
        for s in list(data_split_s):
            data_col_leavetime.append(s)
            if s == 'OFF':
                data_type.append('休息')
            elif s =='年假':
                data_type.append('年假')
            elif s == '入离职缺勤':
                data_type.append('入离职缺勤')
            # elif s == '培训':
            #     data_type.append('培训')
            elif s == '病假':
                data_type.append('病假')
            elif s== '医疗期':
                data_type.append('医疗期')
            elif s == '事假':
                data_type.append('事假')
            elif s == '婚假':
                data_type.append('婚假')
            elif s == '产假':
                data_type.append('产假')
            elif s == '产检':
                data_type.append('产检')
            elif s == '哺乳假':
                data_type.append('哺乳假')
            elif s == '丧假':
                data_type.append('丧假')
            elif s == '陪产假':
                data_type.append('陪产假')
            else:
                data_type.append('工作')

    data_make['类型'] = data_type  # 导入类型数据
    # print(data_type)
    data_make['上班时间'] = data_col_arrivetime  # 导入上班时间
    data_make['下班时间'] = data_col_leavetime  # 导入下班时间
    data_make.上班时间 = data_make.上班时间.str.replace('OFF|年假|入离职缺勤|病假|医疗期|事假|婚假|产假|产检|哺乳假|丧假|陪产假', '', regex=True)#|年假|入离职缺勤|病假|医疗期|事假|婚假|产假|产检|哺乳假|丧假|陪产假
    data_make.下班时间 = data_make.上班时间.str.replace('OFF|年假|入离职缺勤|病假|医疗期|事假|婚假|产假|产检|哺乳假|丧假|陪产假', '', regex=True)

    # ~ #获取'日期(YYYY-MM-DD)'--------方法1
    # ~ data_time = list(map(lambda x: '2019-'+x, data_time))#用公式在每个字符前加2019-

    data_yyyy = []
    for i in range(0, len(list(data.columns))):
        for data_t in list(data_time[2:]):
            data_yyyy.append(str(data_t)[:-9])  # 按每个col，datatime次添加
    data_make['日期(YYYY-MM-DD)'] = data_yyyy  # 赋值给col日期

    # 用openpyxl设置格式
    wb = Workbook()
    ws = wb.create_sheet('Sheet1', -1)
    for r in dataframe_to_rows(data_make, index=False, header=True):
        ws.append(r)
    cell_style(ws, len(data_make.index))
    wb.save(in_excel)
    return in_excel


file_name = '桂庙8.5-8.11.xlsx'
wash_data(file_name)
