import streamlit as st
import pandas as pd

import os,sys
sys.path.append(os.getcwd())

from module.read_data import MappingReader, Acct_Reader
from module.cal_data import Verify_Statement, unpivot_df_account_balance, cal_cell_amount
from module.update_data import VBA_update_data



#替换路径名称最后一段
def replace_last_segment(file_path, new_segment):
    # 使用 os.path.split 将路径分割为目录和文件名
    directory, filename = os.path.split(file_path)
    
    # 将文件名替换为新的字符串
    new_file_path = os.path.join(directory, new_segment)
    
    return new_file_path

def main_flow(df_mapping,path_account_balance,path_workingpaper,single_save=True):
    '''
    输入：
    df_mapping: 映射表字典
    path_account_balance: 存放账户余额表的路径
    path_workingpaper: 存放试算底稿的路径
    path_log_save: 存放<原报表>日志的路径
    输出：
    1.更新试算底稿
    2.针对<原报表>返回校验结果提示用户
    '''
    dfs=df_mapping
    path_account_balance=path_account_balance
    path_workingpaper=path_workingpaper

    '''调试用'''
    # print(path_account_balance)

    #1.读科目余额表
    df_account_balance=Acct_Reader(path=path_account_balance).read_account_balance()

    #2.处理科目余额表数据
    df_acct_2d=unpivot_df_account_balance(df_account_balance)

    #3.校验[原报表数据]
    result_verify=Verify_Statement(dfs['原报表'],df_acct_2d).verify_pre_result()

    #4.生成附注数据
    result_updates={key:cal_cell_amount(value,df_acct_2d) for key,value in dfs.items()}

    #5.更新试算底稿
    for k,v in result_updates.items():
        para_update={'path':path_workingpaper,
                    'sheet_name':k,
                    'update_details':v,
                    'engine':'excel',
                    'visible':False,
                    'auto_close':True}
        VBA_update_data(**para_update)

    #6.保存日志
    result_verify['底稿路径']=path_workingpaper

    if single_save==True or single_save is None:
        save_path = replace_last_segment(file_path=path_workingpaper,new_segment='原文件合并日志_auto')
        os.makedirs(save_path, exist_ok=True)
        index=path_workingpaper.split('\\')[-1]
        path_log_save=os.path.join(save_path,f'{index}_原报表_日志.xlsx')
        result_verify.to_excel(path_log_save,index=False)
    else:
        path_log_save=None
        pass

    return result_verify,path_log_save

# def muti_main_flow(df_mapping,list_acct_path,list_workingpaper_path):
#     '''
#     多线程处理
#     输入：
#     df_mapping: 映射表字典
#     list_acct_path: 存放账户余额表的路径列表
#     list_workingpaper_path: 存放试算底稿的路径列表
#     输出：原报表日志
#     '''
#     df_mapping=df_mapping
#     list_acct_path=list_acct_path
#     list_workingpaper_path=list_workingpaper_path

#     #多线程
#     pool = ThreadPool(6)
#     df_list=[df_mapping for i in range(len(list_acct_path))]
#     arguments=list(zip(df_list,list_acct_path,list_workingpaper_path))
#     result=pool.starmap(main_flow,arguments)

#     path_workingpaper=list_workingpaper_path[0]
#     save_path = replace_last_segment(file_path=path_workingpaper,new_segment='原文件合并日志_auto')
#     path_log_save=os.path.join(save_path,f'原报表批量处理_日志.xlsx')

#     df_log=pd.concat(result)
#     df_log.to_excel(path_log_save,index=False)


    #多线程
    # path_relation=r"D:\audit_project\AUTO_TB\试算科余映射.xlsm"
    # df_relation=pd.read_excel(path_relation,sheet_name='匹配结果',header=0)

    # list_acct_path=df_relation['科目余额表路径'].to_list()
    # list_workingpaper_path=df_relation['试算底稿路径'].to_list()

    # df_mapping=MappingReader(path=path_mapping,header=1).read_mapping_table()
    # muti_main_flow(df_mapping,list_acct_path,list_workingpaper_path) 
    # print('Done')

#循环
def loop_main_flow(df_mapping,list_acct_path,list_workingpaper_path):
    '''
    循环处理
    输入：
    df_mapping: 映射表字典
    list_acct_path: 存放账户余额表的路径列表
    list_workingpaper_path: 存放试算底稿的路径列表
    输出：原报表日志
    '''
    df_mapping=df_mapping
    list_acct_path=list_acct_path
    list_workingpaper_path=list_workingpaper_path

    for i in range(len(list_acct_path)):
        main_flow(df_mapping,list_acct_path[i],list_workingpaper_path[i]) 



if __name__ == '__main__':

    
    path_mapping=r"D:\audit_project\AUTO_TB\映射模板设计.xlsx"


    #单线程
    # path_account_balance=r"D:\audit_project\AUTO_TB\科目余额表示例.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    
    path_account_balance=r"D:\audit_project\AUTO_TB\东方生物\科目余额表\北京博朗生.xlsx"
    path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.33 北京博朗生科技有限公司 2024.xlsx"

    df_mapping=MappingReader(path=path_mapping,header=1).read_mapping_table()
    main_flow(df_mapping,path_account_balance,path_workingpaper) 
    print('Done')

    #循环 先走遍历的逻辑
    # path_relation=r"D:\audit_project\AUTO_TB\试算科余映射.xlsm"
    # df_relation=pd.read_excel(path_relation,sheet_name='匹配结果',header=0)

    # list_acct_path=df_relation['科目余额表路径'].to_list()
    # list_workingpaper_path=df_relation['试算底稿路径'].to_list()

    # df_mapping=MappingReader(path=path_mapping,header=1).read_mapping_table()

    # loop_main_flow(df_mapping,list_acct_path,list_workingpaper_path) 
    # print('Done')




