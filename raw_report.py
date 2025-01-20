import pandas as pd 
import duckdb 
from multiprocessing.dummy import Pool as ThreadPool
import os 
import xlwings as xw
import numpy as np
from win32com.client import Dispatch 
from module.tool_fun import get_file_list,write_data_toduckdb
from module.read_raw_report import read_balance_sheet,read_income_satement

'''读数据'''
def fast_read_fun(folder_path,func):
    pool = ThreadPool(10)
    #获取文件列表
    file_list = [os.path.join(folder_path,file) for file in os.listdir(folder_path) if file.endswith('.xlsx')]

    data_list = pool.map(func, file_list)
    pool.close()
    re=pd.concat(data_list,ignore_index=True)
    return re




if __name__ == '__main__':
    # '''资产负债表入库'''
    # folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\1、财务账套\2024财务报表"
    # db_path=r"D:\audit_project\AUTO_TB\DB\东方生物_FY24.duckdb"

    # re_b=[]
    # for file in os.listdir(folder):
    #     df=read_balance_sheet(os.path.join(folder,file))
    #     print(file,df.shape)
    #     re_b.append(df)
    # df_b=pd.concat(re_b,ignore_index=True)
    # write_data_toduckdb(df_b,'资产负债表',db_path)

    df=read_balance_sheet(r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\1、财务账套\2024财务报表\2024年10期月报_1.16.1_海南启悟私募基金管理有限公司_人民币(元)_财务报表_2025011715000522.xlsx")
    print(df.shape)

    # df_b = fast_read_fun(folder,read_balance_sheet)
    # write_data_toduckdb(df_all,'资产负债表',db_path)
    # df_i = fast_read_fun(folder,read_income_statement)
    # write_data_toduckdb(df_all,'利润表',db_path)
    # print('done')

    '''批量修改8_费用 按照母公司的格式'''

    ######################################################################################

    
