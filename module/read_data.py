import pandas as pd
import xlwings as xw 
import os 
import duckdb 


#读取映射表 
class MappingReader:
    def __init__(self,path,header):
        self.path=path
        self.header=header

    # 读取映射表，整理对应数据到字典 后期需要根据映射表字典进行数据提取
    def read_mapping_table(self):
        dfs = pd.read_excel(self.path,sheet_name=None,header=self.header,dtype=object)
        
        #清洗竖线数据
        result={key:self.extract_data(value) for key,value in dfs.items()}

        return result

    #处理带竖线|的数据,返回一个df
    def clean_split_df(self,data,index):
        df=data.copy()
        temp_dict={}
        temp_dict['账户代码']=df['账户代码'].iloc[index].split('|')
        temp_dict['运算符']=df['运算符'].iloc[index].split('|')
        temp_dict['金额列']=df['金额列'].iloc[index].split('|')
        temp_dict['单元格']=[df['单元格'].iloc[index] for n in range(len(temp_dict['账户代码']))]

        result=pd.DataFrame(temp_dict)
        result=result[['单元格','账户代码','运算符','金额列']]
        # result=result[['行次','项目名称','单元格','账户代码','账户名称','运算符','金额列']] #不需要那么多项目，后期根据需要再修改
        return result


    # 把竖线清洗合并提取数据
    def extract_data(self,data):
        df=data.copy()
        must_col=['单元格','账户代码','运算符','金额列']
        # must_col=['行次','项目名称','单元格','账户代码','账户名称','运算符','金额列'] #不需要那么多项目，后期根据需要再修改
        df=df[must_col].copy()

        #提取行中有"|"的数据 和没有"|"的数据
        mask=(df['账户代码'].apply(lambda x: 1 if'|' in str(x) else 0)==1)
        df_split = df[mask].copy()

        if len(df_split)>0:
            df_split=df_split.applymap(lambda x: x.replace('\n','')).copy()#清洗\n
            list_split=[self.clean_split_df(df_split,n) for n in range(len(df_split))]
            re_split=pd.concat(list_split).copy()
        else:
            re_split=pd.DataFrame()

        #提取行中没有"|"的数据
        df_no_split = df[~mask].copy()

        result=pd.concat([df_no_split,re_split],ignore_index=True)

        #把运算符变成1和-1
        result['运算符'].fillna('',inplace=True)
        result['运算符']=result['运算符'].map({'+':1,'-':-1,'_':-1,'':0})#下划线防止错打字

        #清洗前后空格
        result=result.applymap(lambda x: x.strip() if isinstance(x,str) else x).copy()

        return result

    

class Acct_Reader:
    def __init__(self,path):
        self.path=path

    # 读取新纪元导出的单个科目余额表
    def read_account_balance(self,path=None,record_path=None):

        if path is None:
            file_path=self.path
        else:
            file_path=path

        #读取科目余额表所有项目
        dfs = pd.read_excel(file_path,sheet_name=None,header=None,dtype=object)
        #遍历每个sheet，选取行数>0的,如果有2个行数大于0的df就报错
        re=[]
        for sheet_name,df in dfs.items():
            if len(df)>0:
                re.append(df)
        if len(re)>1:
            raise ValueError(f'{file_path}包含多个sheet，请检查导出文件是否正确')
        
        df=re[0].copy()
        #表体从第四行开始 不要最后一行
        body=df.iloc[4:-1,:].copy()
        col_names=['账户代码',
                    '账户名称',
                    '期初余额_方向',
                    '期初余额_金额',
                    '期初余额_借方金额',
                    '期初余额_贷方金额',
                    '期间发生额_借方金额',
                    '期间发生额_贷方金额',
                    '累计发生额_借方金额',
                    '累计发生额_贷方金额',
                    '期末余额_方向',
                    '期末余额_金额',
                    '期末余额_借方金额',
                    '期末余额_贷方金额']
        result=body.copy()
        result.columns=col_names
        #带金额的列保留两位小数

        for col in col_names:
            if '金额' in col: 
                result[col]=result[col].astype(float).round(2)
        if record_path is not None:
            result['file_path']=file_path
        else:
            pass

        return result

def Data_Loader(path_mapping,path_account_balance):

    '''
    加载基础数据
    :param path_mapping: 映射表路径
    :param path_account_balance: 科目余额表路径
    :return: 映射表：dict{key:sheet名称，value:单元格映射规则的df} 和 科目余额表数据(未经过处理)
    '''

    path_mapping=path_mapping
    path_account_balance=path_account_balance

    #读取映射表
    dfs=MappingReader(path=path_mapping,header=1).read_mapping_table()
    #读取科目余额表
    df_account_balance=Acct_Reader(path=path_account_balance).read_account_balance()

    return dfs,df_account_balance

'''去掉含期初余额的行'''
def clean_start_value(df_mapping):
    #去掉含期初余额的行
    result_1={k:v[~v['金额列'].fillna('').str.contains('期初余额')] for k,v in df_mapping.items()}
    #去掉没有账户代码的行 原报表特殊处理不要去掉金额列
    result={k:v[(v['账户代码'].notnull())&(v['金额列'].notnull())] for k,v in result_1.items() if k!='原报表'}
    if '原报表' in result_1.keys():
        result['原报表']=result_1['原报表']
    return result



if __name__ == '__main__':

    path = r"D:\audit_project\AUTO_TB\试算单元格映射表\试算单元格映射表_东方基因_v20250118.xlsx"
    

    dfs = MappingReader(path=path,header=1).read_mapping_table()
    dfs=clean_start_value(dfs)
    xw.view(dfs['原报表'])

    # path_account_balance = r"D:\audit_project\AUTO_TB\东方生物\科目余额表\北京博朗生.xlsx"
    # df=Acct_Reader(path=path_account_balance).read_account_balance()
    # xw.view(df)
