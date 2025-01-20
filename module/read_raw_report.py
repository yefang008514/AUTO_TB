import pandas as pd
import duckdb 

import os,sys
sys.path.append(os.getcwd())

from module.tool_fun import read_data_xlwings
import xlwings as xw



'''读取一个财务报告'''
def read_report(file_path):
    file_path=file_path
    with xw.App(visible=False) as app:
        book=app.books.open(file_path)
        
        data_range=book.sheets['资产负债表'].used_range.options(pd.DataFrame, index=False).value
        header_row_index=3
        header = data_range[header_row_index - 1]  # 第三行表头
        print(header)
        data = data_range[header_row_index:]  # 表头以下的数据
        df_balance = pd.DataFrame(data, columns=header)

        data_range=book.sheets['利润表'].used_range.options(pd.DataFrame, header=3, index=False).value
        header_row_index=3
        header = data_range[header_row_index - 1]  # 第三行表头
        data = data_range[header_row_index:]  # 表头以下的数据
        df_income = pd.DataFrame(data, columns=header)

        book.close()
    return df_balance, df_income


if __name__ == '__main__':
    file_path = r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\1、财务账套\2024财务报表\2024年12期月报_1.14_北京汉同生物科技有限公司_人民币(元)_财务报表_2025011714362366.xlsx"
    df_balance, df_income = read_report(file_path)
    print(df_balance.columns)
    print(df_income.columns) 
