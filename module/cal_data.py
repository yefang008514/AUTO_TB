import pandas as pd
import duckdb 
import xlwings as xw
import os,sys
sys.path.append(os.getcwd())

from module.read_data import MappingReader,Acct_Reader

#######################################结果计算####################################

def unpivot_df_account_balance(data,project):

    '''
    用duckdb的unpivot把[科目余额表]的各金额转换成二维表格格式
    返回[账户代码,账户名称,期初余额_方向,期末余额_方向,项目,金额]
    '''
    df=data.copy()
    duckdb.register('df',df)
    if project=='新纪元':
        query='''
        unpivot df
        on 
        "期初余额_金额",
        "期初余额_借方金额",
        "期初余额_贷方金额",
        "期间发生额_借方金额",
        "期间发生额_贷方金额",
        "累计发生额_借方金额",
        "累计发生额_贷方金额",
        "期末余额_金额",
        "期末余额_借方金额",
        "期末余额_贷方金额"
        into 
        name 项目
        value 金额
        '''
    elif project=='SAP_华峰':
        query='''
        unpivot df
        on 
        "本位币货币期初",
        "本位货币借方",
        "本位货币贷方",
        "本位货币期末",
        "外币期初",
        "外币借方",
        "外币贷方",
        "外币期末"
        into 
        name 项目
        value 金额
        '''

    result=duckdb.sql(query).df()
    #兼容适配账户代码
    if project=='SAP_华峰':
        result['账户代码']=result['科目代码']
        result['账户名称']=result['科目名称']
    return result


def cal_cell_amount(data_map,data_acctount_2d):
    '''
    #关联映射表和科目余额表计算【单元格金额】
    返回 [单元格,金额]
    '''
    df_map=data_map.copy()
    df_acct_2d=data_acctount_2d.copy()

    duckdb.register('df_map',df_map)
    duckdb.register('df_acct_2d',df_acct_2d)
    query='''
    select 单元格,round(sum(b.金额*a.运算符),2) as 金额
    from df_map a
    left join df_acct_2d b 
    on a.账户代码=b.账户代码 and a.金额列=b.项目
    group by 单元格;
    ''' 
    result=duckdb.sql(query).df()
    result['金额'].fillna(0,inplace=True)

    return result

def cal_result(path_acct_balance,path_mapping):

    '''
    计算每个单元格需要的金额
    返回一个字典，key是sheet名称 value是[单元格,金额]为表头的dataframe
    '''

    #赋值路径
    path_acct=path_acct_balance
    path_map=path_mapping

    #读取映射表到字典key是sheet名称 value是对应的映射表
    dfs = MappingReader(path=path_map,header=1).read_mapping_table()

    #读取科目余额表
    df_acct_balance=Acct_Reader(path=path_acct).read_account_balance()

    #处理科目余额表，转换成二维表格格式
    df_acct_2d=unpivot_df_account_balance(df_acct_balance)

    #######附注数据生成######
    result={key:cal_cell_amount(value,df_acct_2d) for key,value in dfs.items()}

    return result

######################################结果校验 仅针对sheet【原报表】########################################

class Verify_Statement:

    def __init__(self,data_map,data_acctount):
        self.df_map=data_map.copy()
        self.df_acct=data_acctount.copy()

    def cal_acct_amount(self):
        '''
        关联映射表和科目余额表计算【科目金额】

        返回 [账户代码,金额,账户分类]
        '''
        df_map=self.df_map
        df_acct_2d=self.df_acct

        duckdb.register('df_map',df_map)
        duckdb.register('df_acct_2d',df_acct_2d)
        query='''
        select *,
        case
        when left(账户代码,1)='1' then '资产'
        when left(账户代码,1)='2' then '负债'
        when left(账户代码,1)='4' then '权益'
        when left(账户代码,1)='5' then '成本'
        when left(账户代码,1)='6' then '损益' end 账户分类
        from 
        (
        select 
        left(a.账户代码,4) 账户代码,
        round(sum(b.金额*a.运算符),2) as 金额
        from df_map a left join df_acct_2d b 
        on a.账户代码=b.账户代码 and a.金额列=b.项目
        where b.金额 is not null
        group by left(a.账户代码,4)
        )t;
        ''' 
        result=duckdb.sql(query).df()
        result['金额'].fillna(0,inplace=True)

        return result


    def verify_pre_result(self):
        '''
        !!!初步结果校验：
        用于校验自动从科目余额表中获取的数据和科目余额表当前存在的数据
        提示用户还有一些科目在科目余额表中，需要手工调整<原报表>

        返回[账户代码,账户名称,期末余额_方向,期末余额_金额,abs_差异_科余-试算,账户类型]
        注：<账户代码为一级科目>
        '''
        df_result=self.cal_acct_amount()
        df_acct_2d=self.df_acct

        duckdb.register('df_result',df_result)
        duckdb.register('df_acct_2d',df_acct_2d)
        query='''
        select *,
        case
        when left(账户代码,1)='1' then '资产'
        when left(账户代码,1)='2' then '负债'
        when left(账户代码,1)='4' then '权益'
        when left(账户代码,1)='5' then '成本'
        when left(账户代码,1)='6' then '损益' end 账户分类
        from 
        (
        select 
        a.账户代码,
        a.账户名称,
        a.期末余额_方向,
        a.金额 期末余额_金额,
        abs(a.金额)-abs(b.金额) "ABS_差异_[科余-试算]"
        from df_acct_2d a left join df_result b 
        on a.账户代码=b.账户代码 
        where 
        a.项目='期末余额_金额' and len(a.账户代码)=4 --一级科目的期末余额
        and (b.账户代码 is null or abs(a.金额)-abs(b.金额)!=0) --寻找多的和有差异的科目
        -- and a.账户代码 not in ('4103','4104') 把本年利润和未分配利润留着 方便用户校验
        )t
        ''' 
        result=duckdb.sql(query).df()

        return result




if __name__ == '__main__':

    path_acct=r"D:\audit_project\AUTO_TB\科目余额表示例.xlsx"
    path_map=r"D:\audit_project\AUTO_TB\映射模板设计.xlsx"

    #读取映射表到字典key是sheet名称 value是对应的映射表
    dfs = MappingReader(path=path_map,header=1).read_mapping_table()
    df_map=dfs['原报表']

    #读取科目余额表
    df_acct_balance=Acct_Reader(path=path_acct).read_account_balance()

    #处理科目余额表，转换成二维表格格式
    df_acct_2d=unpivot_df_account_balance(df_acct_balance)

    df_ver=Verify_Statement(data_map=df_map,data_acctount=df_acct_2d).verify_pre_result()

    xw.view(df_ver)
