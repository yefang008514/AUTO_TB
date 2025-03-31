import streamlit as st
import pandas as pd
import xlwings as xw
import duckdb
import os,sys
import time
sys.path.append(os.getcwd())


# 华峰集团定制


# 功能一：从财务报表粘贴数据到试算底稿
def paste_report_data_HF(report_path,file_path):
    #读财务报表
    dfs=pd.read_excel(report_path,sheet_name=None,header=0)
    df_balance_sheet=dfs['资产负债表']
    df_income_statement=dfs['利润表']

    #把资产负债表的期末余额取过来
    df_balance_sheet_1=df_balance_sheet.iloc[:,2]
    df_balance_sheet_2=df_balance_sheet.iloc[:,6]

    #把利润表期末余额取过来 本年累计
    df_income_statement_1=df_income_statement.iloc[0:26,3]

    #把数据写入试算底稿
    with xw.App(visible=False,add_book=False) as app:
        wb=app.books.open(file_path)
        sht=wb.sheets['原报表']
        #从N8单元格开始粘贴资产负债表的期末余额 竖向粘贴
        sht.range('N8').options(transpose=True).value=df_balance_sheet_1.values
        #从P8单元格开始粘贴资产负债表的期初余额
        sht.range('P8').options(transpose=True).value=df_balance_sheet_2.values
        #从N84单元格开始粘贴利润表的期末余额
        sht.range('N84').options(transpose=True).value=df_income_statement_1.values
        wb.save()
        wb.close()
    print('粘贴数据成功！')





# 功能二：粘贴费用明细到8_费用
def paste_cost_data_HF(cost_path,file_path):

    #读费用明细
    if cost_path.endswith('.csv'):
        df_cost=pd.read_csv(cost_path,header=0,dtype=object)
    elif cost_path.endswith('.xlsx'):
        df_cost=pd.read_excel(cost_path,header=0,engine='openpyxl')
    else:
        print('费用明细文件格式不支持！')
    df_cost['凭证货币价值']=df_cost['凭证货币价值'].apply(lambda x:float(x.replace(',',''))).round(2)

    df_cost_summary=duckdb.sql('''
    select "功能范围：文本" 科目名称,"总账科目：短文本" 底稿科目,
    round(sum("凭证货币价值"),2) 金额
    from df_cost
    where "功能范围：文本" in ('销售费用','管理费用','研发费用')
    group by "功能范围：文本","总账科目：短文本";
    ''').df()

    df_sa_1=df_cost_summary[df_cost_summary['科目名称']=='销售费用'].loc[:,'底稿科目']
    df_sa_2=df_cost_summary[df_cost_summary['科目名称']=='销售费用'].loc[:,'金额']
    len_sa=len(df_sa_1)

    df_ad_1=df_cost_summary[df_cost_summary['科目名称']=='管理费用'].loc[:,'底稿科目']
    df_ad_2=df_cost_summary[df_cost_summary['科目名称']=='管理费用'].loc[:,'金额']
    len_ad=len(df_ad_1)

    df_rd_1=df_cost_summary[df_cost_summary['科目名称']=='研发费用'].loc[:,'底稿科目']
    df_rd_2=df_cost_summary[df_cost_summary['科目名称']=='研发费用'].loc[:,'金额']
    len_rd=len(df_rd_1)

    #把费用明细数据写入试算底稿
    with xw.App(visible=False,add_book=False) as app:
        wb=app.books.open(file_path)
        sheet=wb.sheets['8_费用']

        #计算需要添加的行
        #统计A列销售费用、管理费用、研发费用的出现频率
        list_A=sheet.range('A1:A500').value
        Series_A=pd.Series(list_A)
        count_sa=len(Series_A[Series_A=='销售费用'])-2
        count_ad=len(Series_A[Series_A=='管理费用'])-2
        count_rd=len(Series_A[Series_A=='研发费用'])-2
        

        # #老代码
        # sa_add=len_sa-60 if len_sa-60>0 else 0
        # ad_add=len_ad-74 if len_ad-74>0 else 0
        # rd_add=len_rd-64 if len_rd-64>0 else 0

        #新代码 动态更新更稳健  需要添加几行
        sa_add=len_sa-count_sa if len_sa-count_sa>0 else 0
        ad_add=len_ad-count_ad if len_ad-count_ad>0 else 0
        rd_add=len_rd-count_rd if len_rd-count_rd>0 else 0

        #如果要加行需要复制哪行
        sa_copy=count_sa # 60
        ad_copy=(sa_copy+count_ad+6)+sa_add #140 +sa_add
        rd_copy=(ad_copy+count_rd+6)+sa_add+ad_add # 210+sa_add+ad_add

        #计算开始粘贴的行
        sa_start=2 # 2
        ad_start=(sa_start+count_sa+6)+sa_add # 68+sa_add
        rd_start=(ad_start+count_ad+6)+sa_add+ad_add # 148+sa_add+ad_add

        def copy_row_paste(sheet, source_row, insert_times):
        # 在指定工作表中复制源行并多次插入
            # 插入指定次数的空行
            for _ in range(insert_times):
                sheet.api.Rows(source_row).Insert(Shift=-4121, CopyOrigin=1)
            # 获取原始行的内容
            target_range = sheet.range(f"{source_row + insert_times}:{source_row + insert_times}")
            # 复制原始行的内容到插入的空行
            for i in range(insert_times):
                target_range.copy(sheet.range(f"{source_row + i}:{source_row + i}"))

        #复制行
        if sa_add>0:
            source_row = sa_copy #哪行要复制
            insert_times = sa_add #复制几次
            copy_row_paste(sheet, source_row, insert_times)
        if ad_add>0:
            source_row = ad_copy #哪行要复制
            insert_times = ad_add #复制几次
            copy_row_paste(sheet, source_row, insert_times)
        if rd_add>0:
            source_row = rd_copy #哪行要复制
            insert_times = rd_add #复制几次
            copy_row_paste(sheet, source_row, insert_times)

        #销售费用
        sheet.range(f'C{sa_start}').options(transpose=True).value=df_sa_1.values
        sheet.range(f'E{sa_start}').options(transpose=True).value=df_sa_2.values
        #管理费用
        sheet.range(f'C{ad_start}').options(transpose=True).value=df_ad_1.values
        sheet.range(f'E{ad_start}').options(transpose=True).value=df_ad_2.values
        #研发费用
        sheet.range(f'C{rd_start}').options(transpose=True).value=df_rd_1.values
        sheet.range(f'E{rd_start}').options(transpose=True).value=df_rd_2.values
        wb.save()
        wb.close()
    
    print('粘贴费用明细成功！')





if __name__ == '__main__':
    #功能一 贴财报
    # start_time = time.time()
    # report_path=r'D:\audit_project\AUTO_TB\华峰化学\测试试算\b3-重庆化工：财务报表-2024.13.XLSX'
    # file_path=r'D:\audit_project\AUTO_TB\华峰化学\测试试算\【试算】b2-重庆化工.xlsx'
    # paste_report_data_HF(report_path,file_path)
    # end_time = time.time()
    # print('运行时间：', round(end_time - start_time,2))

    #功能二 贴费用
    start_time = time.time()
    cost_path=r'D:\audit_project\AUTO_TB\华峰化学\测试试算\FBL3H-费用明细账-b2-重庆化工.csv'
    file_path=r'D:\audit_project\AUTO_TB\华峰化学\测试试算\【试算】b2-重庆化工.xlsx'
    paste_cost_data_HF(cost_path,file_path)
    end_time = time.time()
    print('运行时间：', round(end_time - start_time,2))
