import os 
import dukckdb


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