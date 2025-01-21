import pandas as pd
import duckdb 
import xlwings as xw
from warnings import filterwarnings

import os,sys
sys.path.append(os.getcwd())

from module.read_data import MappingReader
from module.tool_fun import read_data_xlwings
from module.update_data import VBA_update_data,batch_update_excel_openpyxl



'''读取一个财务报告'''
def read_report(file_path,show_log=None):
    file_path=file_path
    with xw.App(visible=False) as app:
        book=app.books.open(file_path)

        ##########资产负债表#############
        # data_range=book.sheets['资产负债表'].used_range.value
        data_range=book.sheets['资产负债表'].range('A4:F68').value
        # print(data_range)
        header_row_index=1
        # header = data_range[header_row_index - 1]  # 表头
        header=['资产','资产_期末余额','资产_上年年末余额','负债权益','负债权益_期末余额','负债权益_上年年末余额']
        data = data_range[header_row_index:]  # 表头以下的数据
        df_balance = pd.DataFrame(data, columns=header)

        df_balance = df_balance[df_balance['资产'].notnull()]
        df_balance = df_balance[df_balance['资产'].apply(lambda x: 1 if '制表人' in str(x) else 0)==0]
        
        ##########利润表################
        # data_range=book.sheets['利润表'].used_range.value
        data_range=book.sheets['利润表'].range('A4:E79').value
        header_row_index=1
        header = data_range[header_row_index - 1]  # 表头
        data = data_range[header_row_index:]  # 表头以下的数据
        df_income = pd.DataFrame(data, columns=header)

        df_income = df_income[df_income['项目名称'].notnull()]
        df_income = df_income[df_income['项目名称'].apply(lambda x: 1 if '制表人' in str(x) else 0)==0]

        df_balance['file_path']=file_path
        df_income['file_path']=file_path

        book.close()
        if show_log==True or show_log is None: 
            print(file_path, df_balance.shape,df_income.shape)
    
    return df_balance, df_income


'''清洗资产负债表 逆透视'''
def clean_balance(data_balance):

    filterwarnings('ignore')

    df_balance = data_balance.copy()

    #1.拆分 资产、负债权益,填充空值 去掉项目前后的空格
    df_assets = df_balance[['资产','资产_期末余额','资产_上年年末余额','file_path']]
    df_assets.columns = ['项目','期末余额','期初余额','file_path']
    df_assets['项目'].fillna('',inplace=True)
    df_assets['项目']=df_assets['项目'].apply(lambda x: str(x).strip())

    df_liab = df_balance[['负债权益','负债权益_期末余额','负债权益_上年年末余额','file_path']]
    df_liab.columns = ['项目','期末余额','期初余额','file_path']
    df_liab['项目'].fillna('',inplace=True)
    df_liab['项目']=df_liab['项目'].apply(lambda x: str(x).strip())

    #2.资产负债表 拆分流动资产,非流动资产，流动负债，非流动负债,所有者权益
    idx_1 = df_assets[df_assets['项目'].str.contains('流动资产合计')].index[0]
    df_liq_assets = df_assets.iloc[:idx_1,:].copy()
    df_liq_assets['分类']= '流动资产'

    idx_2 = df_assets[df_assets['项目'].str.contains('非流动资产合计')].index[0]
    df_nonliq_assets = df_assets.iloc[idx_1+1:idx_2,:].copy()
    df_nonliq_assets['分类']= '非流动资产'

    # 负债
    idx_3 = df_liab[df_liab['项目'].str.contains('流动负债合计')].index[0]
    df_cur_liab = df_liab.iloc[:idx_3,:].copy()
    df_cur_liab['分类']= '流动负债'

    idx_4 = df_liab[df_liab['项目'].str.contains('非流动负债:')].index[0]
    idx_5 = df_liab[df_liab['项目'].str.contains('非流动负债合计')].index[0]
    df_noncur_liab = df_liab.iloc[idx_4:idx_5,:].copy()
    df_noncur_liab['分类']= '非流动负债'

    idx_6 = df_liab[df_liab['项目'].str.contains('所有者权益（或股东权益）')].index[0]
    idx_7 = df_liab[df_liab['项目'].str.contains('所有者权益（或股东权益）合计')].index[0]
    df_equity = df_liab.iloc[idx_6:idx_7,:].copy()
    df_equity['分类']= '所有者权益'

    df_s=pd.concat([df_assets,df_liab]).copy()
    df_summary=df_s[df_s['项目'].str.contains('合计')]
    df_summary['分类']= '合计'

    #3.合并资产负债表
    re_balance = pd.concat([df_liq_assets,df_nonliq_assets,df_cur_liab,df_noncur_liab,df_equity,df_summary],ignore_index=True)
    re_balance['索引号']=re_balance['项目']+'_'+re_balance['分类']

    #4.逆透视拆分出[金额分类,金额]
    re_balance_melt=pd.melt(re_balance,id_vars=['项目','分类','索引号','file_path'],
    value_vars=['期末余额','期初余额'],
    var_name='金额类型',
    value_name='金额')

    # return re_balance,re_balance_melt
    return re_balance_melt 

'''清洗利润表 逆透视'''
def clean_income(data_income):

    df_income = data_income.copy()
    df_income['分类']='损益'

    # 1.清洗[项目名称]前后空格
    df_income.rename(columns={'项目名称':'项目'},inplace=True)
    df_income['项目']=df_income['项目'].str.strip()
    
    # 2.添加索引号 需要用自增id标记唯一值
    df_income['索引号']=df_income.index.astype(str)+'_'+df_income['项目']

    # 3.逆透视拆分出[金额分类,金额]
    df_income_melt=pd.melt(df_income,id_vars=['项目','分类','索引号','file_path'],
    value_vars=['本期金额','本年金额','上年金额','上年同期金额'],
    var_name='金额类型',
    value_name='金额')

    # return df_income,df_income_melt
    return df_income_melt 


'''拼接资产负债表、利润表'''
def concat_report(data_balance,data_income):

    df_balance = data_balance.copy()
    df_income = data_income.copy()

    df_balance['表名']='资产负债表'
    df_income['表名']='利润表'

    re_col=['项目','分类','金额类型','金额','索引号','file_path','表名']

    result=pd.concat([df_balance[re_col],df_income[re_col]],axis=0,ignore_index=True)

    return result


'''计算结果'''
def cal_result(data_mapping,data_report):
    '''
    #关联映射表和科目余额表计算【单元格金额】
    返回 [单元格,金额]
    '''
    df_mapping = data_mapping.copy()
    df_report = data_report.copy()

    duckdb.register('df_mapping',df_mapping)
    duckdb.register('df_report',df_report)
    query='''
    select 单元格,round(sum(b.金额*a.运算符),2) as 金额
    from df_mapping a
    left join df_report b 
    on a.账户代码=b.索引号 and a.金额列=b.金额类型
    group by 单元格;
    ''' 
    result=duckdb.sql(query).df()
    result['金额'].fillna(0,inplace=True)

    return result

'''
主函数 
输入:
映射表字典
财务报表路径、试算底稿路径
更新试算底稿结果
'''
def main_flow_report(data_mapping,path_report,path_workingpaper,engine):
    df_mapping=data_mapping['原报表'] 
    path_report=path_report
    path_workingpaper=path_workingpaper
    engine=engine

    #1.读取财务报告
    df_balance, df_income=read_report(file_path=path_report,show_log=False)

    #2.清洗数据
    df_balance_melt = clean_balance(df_balance)
    df_income_melt = clean_income(df_income)

    #3.拼接数据
    df_report = concat_report(df_balance_melt,df_income_melt)

    #4.计算结果
    df_result = cal_result(df_mapping,df_report)

    #5.更新试算底稿
    para_update={'path':path_workingpaper,
                        'sheet_name':'原报表',
                        'update_details':df_result,
                        'engine':engine,
                        'visible':False,
                        'auto_close':True}
    if engine in ['wps','excel']:
        VBA_update_data(**para_update)
    elif engine=='openpyxl':
        para_update={'path':path_workingpaper,
                    'sheet_name':'原报表',
                    'update_details':df_result}
        batch_update_excel_openpyxl(**para_update)

    #返回经过处理的财务报告(科目余额表、利润表拼接在一起了)
    return df_report
    



if __name__ == '__main__':

    dfs=MappingReader(path=r"D:\audit_project\AUTO_TB\试算单元格映射表\【原报表】单元格映射表_东方基因.xlsx",header=1).read_mapping_table()
    path_report=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024财务报表\2024年12期月报_1.19_万子健检测技术（上海）有限公司_人民币(元)_财务报表_2025011714360623.xlsx"
    path_workingpaper=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算-最新\1.19 万子健检测技术（上海）有限公司2024.xlsx"
    engine='excel'
    main_flow_report(dfs,path_report,path_workingpaper,engine)
    print('done')






    #######################草稿############################
    # file_path = r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \东方生物2024年年审\1、财务账套\2024财务报表\2024年12期月报_1.14_北京汉同生物科技有限公司_人民币(元)_财务报表_2025011714362366.xlsx"
    # df_balance, df_income = read_report(file_path)
    # print(df_balance.columns)
    # xw.view(df_balance)
    # print(df_income.columns) 
    # xw.view(df_income)

    # df_balance_cleaned,df_balance_melt = clean_balance(df_balance)
    # df_income_cleaned,df_income_melt = clean_income(df_income)

    # xw.view(df_balance_cleaned)
    # xw.view(df_income_cleaned)
    # xw.view(df_balance_melt)
    # xw.view(df_income_melt)

