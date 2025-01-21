import pandas as pd 
import duckdb 
from multiprocessing.dummy import Pool as ThreadPool
from module.read_data import Acct_Reader
import os 
import xlwings as xw
import numpy as np
from win32com.client import Dispatch 
from module.tool_fun import get_file_list,write_data_toduckdb


'''读科目余额表数据'''
def fast_read_data_acct(folder_path):
    pool = ThreadPool(10)
    file_list = get_file_list(folder_path)
    record_list=[True for i in range(len(file_list))]
    args=zip(file_list,file_list,record_list)
    data_list = pool.starmap(Acct_Reader.read_account_balance, args)
    pool.close()
    re=pd.concat(data_list,ignore_index=True)
    return re

'''！！！！！！！
此代码用于统一<8_费用>格式和<1.00>母公司的<8_费用>
！！！！！！
'''    


if __name__ == '__main__':
    '''科目余额表入库'''
    # folder=r"D:\audit_project\AUTO_TB\东方生物\FY24\科目余额表_FY24_树状展开"
    # db_path=r"D:\audit_project\AUTO_TB\DB\东方生物_FY24.duckdb"

    # df_all = fast_read_data_acct(folder)
    # write_data_toduckdb(df_all,'科目余额表',db_path)
    # print('done')

    '''批量修改8_费用 按照母公司的格式'''

    ######################################################################################

    wb_template_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\2、试算\2024年试算-最新\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    with xw.App(visible=False, add_book=False) as app:
        wb_template=app.books.open(wb_template_path)
        df_cost_1=pd.DataFrame(wb_template.sheets['8_费用'].range("B2:D58").value,dtype=object) #转换成字符串的df
        df_cost_1=df_cost_1.applymap(lambda x:"'"+str(x) if x is not None else None)
        df_cost_2=pd.DataFrame(wb_template.sheets['8_费用'].range("B68:D113").value,dtype=object)
        df_cost_2=df_cost_2.applymap(lambda x:"'"+str(x) if x is not None else None)
        wb_template.close()

    path_workingpaper=r"D:\audit_project\AUTO_TB\试算单元格映射表\试算科余路径关系表_250121_temp.xlsm"
    df_path=pd.read_excel(path_workingpaper,sheet_name='匹配结果',header=0,dtype=object)
    path_list=df_path['试算底稿路径'].tolist()
    
    for path in path_list:
        xl = xw._xlwindows.COMRetryObjectWrapper(Dispatch("Ket.Application")) 
        impl = xw._xlwindows.App(visible=False, add_book=False, xl=xl)     
        with xw.App(visible=False, add_book=False, impl=impl) as app:
            print(f'正在处理{path}')
            wb=app.books.open(path)
            wb.sheets['8_费用'].range("B2:D58").value=df_cost_1.values #销售费用
            wb.sheets['8_费用'].range("B68:D113").value=df_cost_2.values #管理费用
            wb.save()
            wb.close()
    print('done')


