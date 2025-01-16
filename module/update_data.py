import xlwings as xw 
from win32com.client import Dispatch 
import ctypes
import pandas as pd
from openpyxl import load_workbook
import time

import os,sys
sys.path.append(os.getcwd())

from module.cal_data import cal_result


def is_file_open(file_path):
    '''
    检查试算文件是否被打开，提示用户关闭文件。
    '''
    try:
        with open(file_path, 'r+'):
            pass
    except (IOError, PermissionError):
        return True
    return False


def batch_update_excel_openpyxl(path,sheet_name, update_details):
    """
    使用 openpyxl 批量更新指定Excel文件中的指定工作表和单元格的值。
    :param path: str, Excel文件路径
    :param sheet_name: str, 工作表名称
    :param updates: dataframe, 包含单元格地址和更新值的数据框
    """
    path=path
    updates=update_details.set_index('单元格')['金额'].to_dict()

    if is_file_open(path):
        raise Exception(f"{path}文件已打开，请关闭文件后重试。")

    try:
        # 加载工作簿
        workbook = load_workbook(path)
        # 获取工作表
        sheet = workbook[sheet_name]
        # 遍历更新任务字典
        for key, value in updates.items():
            sheet[key] = value
        # 保存更改
        workbook.save(path)
    except Exception as e:
        print(f"更新失败: {e}")


def xlwings_update_data(path,sheet_name:str,update_details,engine,visible,auto_close):
    '''
    使用xlwings批量更新数据 
    :param path: str, Excel文件路径
    :param sheet_name: str, 工作表名称
    :param updates: dataframe, 包含单元格地址和更新值的数据框
    '''
    path=path 
    data=update_details.set_index('单元格')['金额'].to_dict()

    if is_file_open(path):
        raise Exception(f"{path}文件已打开，请关闭文件后重试。")

    if engine is None or engine=='wps':
        # xlwings 打开 wps 
        xl = xw._xlwindows.COMRetryObjectWrapper(Dispatch("Ket.Application")) 
        impl = xw._xlwindows.App(visible=False, add_book=False, xl=xl)     
        with xw.App(visible=visible, add_book=False, impl=impl) as app:
            wb=app.books.open(path)
            sheet = wb.sheets[sheet_name]
            for cell, value in data.items():
                sheet.range(cell).value = value
            wb.save()
            if auto_close==True or auto_close is None:
                wb.close()    
    else:
        # excel 引擎
        with xw.App(visible=visible) as app:
            wb=app.books.open(path)
            sheet = wb.sheets[sheet_name]
            for cell, value in data.items():
                sheet.range(cell).value = value
            wb.save()
            if auto_close==True or auto_close is None:
                wb.close()


def VBA_update_data(path,sheet_name,update_details,engine,visible,auto_close):
    '''
    使用VBA代码批量更新数据 
    :param path: str, Excel文件路径
    :param sheet_name: str, 工作表名称
    :param updates: dataframe, 包含单元格地址和更新值的数据框
    '''

    path=path 
    data=update_details[['单元格','金额']].values.tolist()
    visible=visible

    if is_file_open(path):
        raise Exception(f"{path}文件已打开，请关闭文件后重试。")

    #VBA代码
    code_vba='''
    Sub BatchUpdate(updates As Variant, sheetName As String)
    Dim i As Long
    Dim targetSheet As Worksheet
    Dim cellAddress As String
    Dim cellValue As Variant

    ' 获取目标工作表
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    ' 遍历数组，更新单元格
    For i = LBound(updates, 1) To UBound(updates, 1)
        cellAddress = updates(i, 0)
        cellValue = updates(i, 1)
        targetSheet.Range(cellAddress).Value = cellValue
    Next i
    End Sub

    '''
    if engine is None or engine=='wps':     # xlwings 打开 wps 
        xl = xw._xlwindows.COMRetryObjectWrapper(Dispatch("Ket.Application")) 
        impl = xw._xlwindows.App(visible=False, add_book=False, xl=xl)     
        with xw.App(visible=visible, add_book=False, impl=impl) as app:
            wb=app.books.open(path)
            vb_object = wb.api.VBProject.VBComponents.Add(1)  # 1 表示标准模块
            vb_object.CodeModule.AddFromString(code_vba)
            wb.macro("BatchUpdate")(data,sheet_name)
            wb.api.VBProject.VBComponents.Remove(vb_object)
            wb.save()
            if auto_close==True or auto_close is None:
                wb.close()
    else:     # xlwings 打开 excel
        with xw.App(visible=visible) as app:
            wb=app.books.open(path)
            vb_object = wb.api.VBProject.VBComponents.Add(1)  # 1 表示标准模块
            vb_object.CodeModule.AddFromString(code_vba)
            wb.macro("BatchUpdate")(data,sheet_name)
            wb.api.VBProject.VBComponents.Remove(vb_object)
            wb.save()
            if auto_close==True or auto_close is None:
                wb.close()



if __name__ == '__main__':

    path_map=r"D:\audit_project\AUTO_TB\映射模板设计.xlsx"
    path_acct=r"D:\audit_project\AUTO_TB\科目余额表示例.xlsx"
    path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"

    # path_map=r"D:\audit_project\AUTO_TB\映射模板设计.xlsx"
    # path_acct=r"D:\audit_project\AUTO_TB\东方生物\科目余额表\北京博朗生.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.33 北京博朗生科技有限公司 2024.xlsx"
    
    result=cal_result(path_acct_balance=path_acct,path_mapping=path_map)
    start=time.time()
    for k,v in result.items():
        # xlwings_update_data(path=path_workingpaper,sheet_name=k,update_details=v,engine='excel',visible=False,auto_close=True)#很慢
        # VBA_update_data(path=path_workingpaper,sheet_name=k,update_details=v,engine='excel',visible=False,auto_close=True)
        VBA_update_data(path=path_workingpaper,sheet_name=k,update_details=v,engine='wps',visible=False,auto_close=True)
        # batch_update_excel_openpyxl(path=path_workingpaper,sheet_name=k,update_details=v)#有点慢
    end=time.time()
    print(f"更新耗时{round(end-start,2)}'秒")
    print('成功')