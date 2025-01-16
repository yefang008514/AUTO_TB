import os,sys
sys.path.append(os.getcwd())
from module.read_data import MappingReader,Acct_Reader
from module.cal_data import Verify_Statement,unpivot_df_account_balance,cal_cell_amount
from module.update_data import VBA_update_data


def main_flow(path_mapping,path_account_balance,path_workingpaper,path_log_save):
    '''
    输入：
    path_mapping: 存放映射表的路径
    path_account_balance: 存放账户余额表的路径
    path_workingpaper: 存放试算底稿的路径
    path_log_save: 存放<原报表>日志的路径
    输出：
    1.更新试算底稿
    2.针对<原报表>返回校验结果提示用户
    '''
    path_mapping=path_mapping
    path_account_balance=path_account_balance
    path_workingpaper=path_workingpaper
    path_log_save=path_log_save

    # 1.读取数据 
    dfs=MappingReader(path=path_mapping,header=1).read_mapping_table()
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
                    'engine':'wps',
                    'visible':False,
                    'auto_close':True}
        VBA_update_data(**para_update)

    #6.保存日志
    result_verify['底稿路径']=path_workingpaper
    result_verify.to_excel(path_log_save,index=False)
    return result_verify


if __name__ == '__main__':
    path_mapping=r"D:\audit_project\AUTO_TB\映射模板设计.xlsx"
    path_account_balance=r"D:\audit_project\AUTO_TB\科目余额表示例.xlsx"
    path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.00 浙江东方基因生物制品股份有限公司 2024.xlsx"
    path_log_save=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.00_原报表.xlsx"
    main_flow(path_mapping,path_account_balance,path_workingpaper,path_log_save) 



