import streamlit as st
import pandas as pd
import xlwings as xw
import os,sys
sys.path.append(os.getcwd())

from module.read_data import MappingReader, Acct_Reader
from module.cal_data import Verify_Statement, unpivot_df_account_balance, cal_cell_amount
from module.update_data import VBA_update_data,batch_update_excel_openpyxl



#替换路径名称最后一段
def replace_last_segment(file_path, new_segment):
    # 使用 os.path.split 将路径分割为目录和文件名
    directory, filename = os.path.split(file_path)
    
    # 将文件名替换为新的字符串
    new_file_path = os.path.join(directory, new_segment)
    
    return new_file_path

def main_flow(df_mapping,path_account_balance,path_workingpaper,single_save,engine,project,exchange_rate):
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
    if project=='新纪元':
        df_account_balance=Acct_Reader(path=path_account_balance).read_account_balance()
    elif project=='SAP_华峰':
        df_account_balance=Acct_Reader(path=path_account_balance).read_account_balance_HF()

    #2.处理科目余额表数据
    df_acct_2d=unpivot_df_account_balance(df_account_balance,project)


    #3.校验[原报表数据] 如果模板里面没有就不处理，!!华峰不处理原报表!!
    judge_flag=('原报表' not in dfs.keys() or project=='SAP_华峰')
    if judge_flag:
        pass
    else:
        result_verify=Verify_Statement(dfs['原报表'],df_acct_2d).verify_pre_result()

    #4.生成试算数据
    #添加汇率

    result_updates={key:cal_cell_amount(value,df_acct_2d,key,exchange_rate) for key,value in dfs.items()}
    # xw.view(result_updates['8_费用'])

    #5.更新试算底稿
    for k,v in result_updates.items():
        if k=='8_费用':#费用的特殊判断
            new_v=v
            #如果E31->E49单元格不为0，则更新E30为0
            cell_list_1=['E'+str(m) for m in range(31,50)]
            flag_1=new_v[new_v['单元格'].isin(cell_list_1)]['金额'].sum()!=0
            cell_list_2=['E'+str(l) for l in range(142,148)]
            #如果E142->E147单元格不为0，则更新E141为0
            flag_2=new_v[new_v['单元格'].isin(cell_list_2)]['金额'].sum()!=0
            if flag_1:
                idx_update=new_v.query("单元格=='E30'").index
                new_v.loc[idx_update,'金额']=0
            elif flag_2:
                idx_update=new_v.query("单元格=='E141'").index
                new_v.loc[idx_update,'金额']=0  
        else:
            new_v=v

        #更新试算底稿
        try:
            para_update={'path':path_workingpaper,
                            'sheet_name':k,
                            'update_details':new_v,
                            'engine':engine,
                            'visible':False,# False True
                            'auto_close':True}
            if engine in ['wps','excel']:
                VBA_update_data(**para_update)
            elif engine=='openpyxl':
                para_update={'path':path_workingpaper,
                            'sheet_name':k,
                            'update_details':new_v}
                batch_update_excel_openpyxl(**para_update)
        except Exception as e:
            raise ValueError(f'更新{path_workingpaper}失败,错误信息：{e}')

    #6.保存日志
    if judge_flag:
        result_verify=pd.DataFrame()
        path_log_save=None
    else:
        result_verify['底稿路径']=path_workingpaper
        if single_save==True or single_save is None:
            save_path = replace_last_segment(file_path=path_workingpaper,new_segment='原文件合并日志_auto')
            os.makedirs(save_path, exist_ok=True)
            index=path_workingpaper.split('\\')[-1].replace('.xlsx','')
            path_log_save=os.path.join(save_path,f'{index}_原报表_日志.xlsx')
            result_verify.to_excel(path_log_save,index=False)
        else:
            path_log_save=None
            pass

    return result_verify,path_log_save


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

    
    path_mapping=r"D:\audit_project\AUTO_TB\试算单元格映射表\试算单元格映射表_东方基因_v20250118.xlsx"


    #单线程
    # path_mapping=r"D:\audit_project\AUTO_TB\试算单元格映射表\试算单元格映射表_东方基因_费用.xlsx"
    # path_account_balance=r"D:\audit_project\AUTO_TB\科目余额表示例.xlsx"
    # path_account_balance=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024科目余额表\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    
    # path_account_balance=r"D:\audit_project\AUTO_TB\东方生物\科目余额表\北京博朗生.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.33 北京博朗生科技有限公司 2024.xlsx"

    # path_account_balance=r"D:\audit_project\AUTO_TB\东方生物\FY24\科目余额表_FY24_树状展开\1.02 南京长健生物科技有限公司 2024.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\2、试算\2024年试算-最新\1.02 南京长健生物科技有限公司 2024.xlsx"

    # path_account_balance=r"D:\audit_project\AUTO_TB\东方生物\FY24\科目余额表_FY24_树状展开\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    # path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\2、试算\2024年试算-最新\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"

    # df_mapping=MappingReader(path=path_mapping,header=1).read_mapping_table()
    # main_flow(df_mapping=df_mapping,path_account_balance=path_account_balance,path_workingpaper=path_workingpaper,single_save=True,engine='excel') 
    # print('Done')

    #循环 先走遍历的逻辑
    # path_relation=r"D:\audit_project\AUTO_TB\试算科余映射.xlsm"
    # df_relation=pd.read_excel(path_relation,sheet_name='匹配结果',header=0)

    # list_acct_path=df_relation['科目余额表路径'].to_list()
    # list_workingpaper_path=df_relation['试算底稿路径'].to_list()

    # df_mapping=MappingReader(path=path_mapping,header=1).read_mapping_table()

    # loop_main_flow(df_mapping,list_acct_path,list_workingpaper_path) 
    # print('Done')




