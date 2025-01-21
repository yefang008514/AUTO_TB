import pandas as pd 
import duckdb 
from multiprocessing.dummy import Pool as ThreadPool
# from multiprocessing import Pool 
import asyncio
import os 
import xlwings as xw
import numpy as np
from win32com.client import Dispatch 
from module.tool_fun import *
from module.read_raw_report import read_report


'''处理原报表'''
if __name__ == '__main__':
    folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\1、财务账套\2024财务报表"
    db_path=r"D:\audit_project\AUTO_TB\DB\东方生物_FY24.duckdb" 

    re_b=[]
    re_i=[]
    for file in os.listdir(folder):
        if file.endswith('.xlsx') and '杭州公健知识产权' not in file and '财务分析报表' not in file:
            file_path=os.path.join(folder,file)
            df_b,df_i=read_report(file_path)
            re_b.append(df_b)
            re_i.append(df_i)
    df_b=pd.concat(re_b,ignore_index=True)
    df_i=pd.concat(re_i,ignore_index=True)

    write_data_toduckdb(df_b,'资产负债表',db_path)
    write_data_toduckdb(df_i,'利润表',db_path)
    print('done')


    ######################################################################################

    

