import duckdb 
import pandas as pd
import xlwings as xw
import os,sys
sys.path.append(os.getcwd())

# from module.read_data import MappingReader, Acct_Reader
from HF_SAP import paste_report_data_HF,paste_cost_data_HF
from main_flow import main_flow
from module.main_flow import main_flow
from module.read_data import MappingReader,clean_start_value
from module.read_raw_report import main_flow_report
from module.workingpapaer_cost import gen_cost_workingpaper,custom_read_and_paste_main,read_excel_multi
from module.extract_inter import main_merge_raw_wb


#贴原报表
def paste_report_loop(type=None,sheet_name=None):
    #生成底稿和报表路径对应关系
    save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算底稿'
    report_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\财务报表"
    scope_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx"

    df=pd.read_excel(scope_path,header=0,dtype=object)
    if type=='外币':
        df=df[df['是否外币']=='是'].copy() #USD
    else:
        df=df[df['是否外币']!='是'].copy() #RMB

    df['path']=df['24年试算序号'].astype(str)+'-'+df['公司名称']+'2024.xlsx'
    path_list=df['path'].tolist()
    path_list=[os.path.join(save_folder,path) for path in path_list] #底稿路径

    report_list=df['报表名称'].tolist()
    report_list=[os.path.join(report_folder,report+'.xlsx') for report in report_list] #报表路径
    for i in range(10,len(path_list)):
        engine='excel'
        report_path=report_list[i]
        file_path=path_list[i]
        print(f'正在粘贴试算：{file_path}') #打印底稿位置
        paste_report_data_HF(report_path,file_path,engine,sheet_name)

#贴费用明细
def paste_cost_loop(type=None,exchange_rate=None):
    #生成底稿和报表路径对应关系
    save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算底稿'
    report_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\FBL3H"
    scope_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx"

    if type is None:
        final_rate=1
    else:
        final_rate=exchange_rate

    df=pd.read_excel(scope_path,header=0,dtype=object)
    if type=='外币':
        df=df[df['是否外币']=='是'].copy() #USD
    else:
        df=df[df['是否外币']!='是'].copy() #RMB

    df['path']=df['24年试算序号'].astype(str)+'-'+df['公司名称']+'2024.xlsx'
    path_list=df['path'].tolist()
    path_list=[os.path.join(save_folder,path) for path in path_list] #底稿路径

    report_list=df['报表名称'].tolist()
    report_list=[os.path.join(report_folder,report+'.csv') for report in report_list] #报表路径 用csv后缀
    for i in range(10,len(path_list)):
        engine='wps'
        report_path=report_list[i]
        file_path=path_list[i]
        print(f'正在粘贴试算：{file_path}') #打印底稿位置
        paste_cost_data_HF(report_path,file_path,engine,final_rate)

def run_main_flow_loop(type=None,exchange_rate=None,uploaded_mapping=None):
    #生成底稿和报表路径对应关系
    save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算底稿'
    # save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \华峰集团2024年报\1、报表及试算\4-华峰集团试算-2024.12.31'
    report_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\科目余额表"
    scope_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx"

    df=pd.read_excel(scope_path,header=0,dtype=object)
    if type=='外币':
        df=df[df['是否外币']=='是'].copy() #USD
    else:
        df=df[df['是否外币']!='是'].copy() #RMB

    if exchange_rate is None:
        exchange_rate=[1,1]

    df['path']=df['24年试算序号'].astype(str)+'-'+df['公司名称']+'2024.xlsx'
    path_list=df['path'].tolist()
    path_list=[os.path.join(save_folder,path) for path in path_list] #底稿路径

    report_list=df['报表名称'].tolist()
    report_list=[os.path.join(report_folder,report+'.xlsx') for report in report_list] #科目余额表表路径 

    for i in range(10,len(path_list)):
        print(f'正在处理试算：{path_list[i]}') #打印底稿位置
        path_account_balance = report_list[i]
        path_workingpaper = path_list[i]

        df_mapping = MappingReader(path=uploaded_mapping, header=1).read_mapping_table()# uploaded_mapping 映射表路径
        df_mapping = clean_start_value(df_mapping) #不需要期初
        # sheet_selected=['原报表','1','2','4','4.1_递延','5','6','8'] #更新制定sheet
        sheet_selected=['8'] #更新制定sheet
        df_mapping={ele:df_mapping[ele] for ele in sheet_selected} #更新制定sheet
        print(df_mapping.keys())
        single_save=False
        engine='wps'
        project='SAP_华峰'

        result,log_file_path = main_flow(df_mapping, path_account_balance, path_workingpaper,single_save,engine,project,exchange_rate)

    






if __name__ == '__main__':

    # paste_report_loop()#贴原报表

    paste_cost_loop()#贴费用明细

    # path_mapping=r"D:\wps_cloud_sync\339514258\WPS云盘\华峰集团_FY24\【单元格映射表】华峰_【科余】_all.xlsx"
    # run_main_flow_loop(uploaded_mapping=path_mapping)#科目余额表贴试算

    # paste_report_loop(type='外币',sheet_name='原报表USD')#贴原报表

    # paste_cost_loop(type='外币',exchange_rate=7.1140)#贴费用明细

    # path_mapping=r"D:\wps_cloud_sync\339514258\WPS云盘\华峰集团_FY24\【单元格映射表】华峰_【科余】_外币.xlsx"
    # run_main_flow_loop(type='外币',exchange_rate=[7.1884,7.1140],uploaded_mapping=path_mapping)#科目余额表贴试算



    #USD 期末汇率 7.1884  本期平均汇率 7.1140
    
    # exchange_rate=['7.1884','7.1140']

    # 
 


    print('done')
