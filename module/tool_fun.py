import os 
import duckdb
import xlwings as xw
import pandas as pd

'''1.获取文件夹下所有文件'''
def get_file_list(folder_path):
    file_list = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if (file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.csv')) and '~$' not in file:
                file_list.append(os.path.join(root, file))
            else:
                continue
    return file_list

'''2.向duckdb写入数据'''
def write_data_toduckdb(data,table_name,db_path):

    data=data.copy()
    with duckdb.connect(db_path) as con:
        con.register('data',data)
        con.sql(f'''
        drop table if exists {table_name};
        create table if not exists {table_name} as select * from data;
        ''')


''''3.处理定位表头，去掉干扰数据'''
def df_auto_header(data):
    df=data.copy()

    #####寻找表头 表头不能是最后一行#####
    missing_values_count = df.iloc[:-1,:].isnull().sum(axis=1)

    # 找到空值最少的行（表头）
    header_row_index = missing_values_count.idxmin()

    # 提取表头行
    header_row = df.iloc[header_row_index]

    #去掉表头字符串中的空格
    header_row = header_row.apply(lambda x:x.strip() if isinstance(x,str) else x)

    # 重新整理 DataFrame，将表头设置为空值最少的行，并删除该行
    df_body = df.iloc[header_row_index+1:,:]
    # 将表头行设置为 DataFrame 的列名
    df_body.columns = header_row
    # 重新设置索引
    df_body.reset_index(drop=True, inplace=True)

    return df_body


'''4.使用xlwings读取数据 返回df  auto_header 如果随便填就不识别表头'''
def read_data_xlwings(path,sheet_name=None,header=None,auto_header=None):
    mypath=path
    header_final=header if header is not None else 0
    with xw.App(visible=False) as app:
        book=app.books.open(mypath)
        if sheet_name is not None:
            table=book.sheets[sheet_name].used_range
            df=table.options(pd.DataFrame, header=header_final, index=False).value
        else:
            table=book.sheets[0].used_range
            df=table.options(pd.DataFrame, header=header_final, index=False).value
        book.close()
    #默认自动识别表头
    if auto_header is None or auto_header==True:
        df=df_auto_header(data=df).copy()
    else:
        pass 
    return df