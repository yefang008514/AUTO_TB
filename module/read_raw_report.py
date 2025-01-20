import pandas as pd
import duckdb 




'''1.读取资产负债表(特定)'''
def read_balance_sheet(file_path):
    
    # 1.读文件
    df=pd.read_excel(file_path,sheet_name='资产负债表',dtype=object,header=3)
    df=df.iloc[:-2,:].copy()
    df['file_name']=file_path # 增加文件路径列
    # 2.校验表头第一个字段
    if df.columns[0]!='资产':
        raise ValueError('资产负债表的表头第一个字段必须为"资产"')
        
    return df


'''2.读取利润表(特定)'''

def read_income_satement(file_path):
    
    # 1.读文件
    df=pd.read_excel(file_path,sheet_name='利润表',dtype=object,header=3)
    df['file_name']=file_path # 增加文件路径列
    df=df.iloc[:-2,:].copy()
    # 2.校验表头第一个字段
    if df.columns[0]!='项目名称':
        raise ValueError('利润表的表头第一个字段必须为"项目名称"')
        
    return df