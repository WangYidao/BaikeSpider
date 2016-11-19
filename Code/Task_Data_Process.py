# 处理百度系统后台导出的数据
# 一个垂类一个文件，一个任务一个sheet

from openpyxl import load_workbook,Workbook                                        # 导入Excel文件读取模块

import re

from time import strftime

WB_Source = load_workbook('/Users/bin/Desktop/1114 bbck任务.xlsx')

WS_source = WB_Source.active

Entry_Name_Index = -1
Type_Index = -1
Entry_State_Index = -1
Editor_Index = -1
Task_Name_Index = -1

WB_Dest = Workbook()

HKHT = "航空航天_"+strftime("%Y%m%d")
XXKX = "信息科学_"+strftime("%Y%m%d")
QHHJ = "气候环境_"+strftime("%Y%m%d")
NYLY = "能源利用_"+strftime("%Y%m%d")
TWX = "天文学_"+strftime("%Y%m%d")

# 新建垂类sheet
WS_HKHT = WB_Dest.active
WS_HKHT.title = HKHT

WS_XXKX = WB_Dest.create_sheet(XXKX)
WS_QHHJ = WB_Dest.create_sheet(QHHJ)
WS_NYLY = WB_Dest.create_sheet(NYLY)
WS_TWX = WB_Dest.create_sheet(TWX)

for column_index in range(1, WS_source.max_column+1):

    if WS_source.cell(row=1, column=column_index).value == "名称":
        Task_Name_Index = column_index
    elif WS_source.cell(row=1, column=column_index).value == "所属垂类":
        Type_Index = column_index
    elif WS_source.cell(row=1, column=column_index).value == "所属任务":
        Task_Name_Index = column_index
    elif WS_source.cell(row=1, column=column_index).value == "状态":
        Entry_State_Index = column_index
    elif WS_source.cell(row=1, column=column_index).value == "编辑记录":
        Editor_Index = column_index
    else:
        pass

for row_index in range(2, WS_source.max_row+1):

    if WS_source.cell(row=row_index, column=Editor_Index).value == " -":
        continue
    else:
        pass

    if WS_source.cell(row=row_index, column=Entry_State_Index) == "已达标":
        if WS_source.cell(row=row_index, column=Type_Index) == "航空航天":

        elif WS_source.cell(row=row_index, column=Type_Index) == "信息科学":
        elif WS_source.cell(row=row_index, column=Type_Index) == "气候环境":
        elif WS_source.cell(row=row_index, column=Type_Index) == "能源利用":
        elif WS_source.cell(row=row_index, column=Type_Index) == "天文学":
        else:
            pass
