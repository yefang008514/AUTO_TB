import pandas as pd 
import os 


# 拆分财务报表
def depart_report(path,save_folder):
    dfs=pd.read_excel(path,sheet_name=None,header=0,dtype=object)
    df_balance=dfs['资产负债表']
    df_income=dfs['利润表']


    df_balance['期末余额']=df_balance['期末余额'].astype(float).round(2)
    df_balance['年初余额']=df_balance['年初余额'].astype(float).round(2)
    df_balance['期末余额.1']=df_balance['期末余额.1'].astype(float).round(2)
    df_balance['年初余额.1']=df_balance['年初余额.1'].astype(float).round(2)

    df_income['本期发生额']=df_income['本期发生额'].astype(float).round(2)
    df_income['本年累计发生额']=df_income['本年累计发生额'].astype(float).round(2)



    company_list=df_balance['公司名称'].unique()
    for company in company_list:
        print(company)
        df_1=df_balance[df_balance['公司名称']==company]
        df_2=df_income[df_income['公司名称']==company]
        df_1=df_1.drop(columns=['公司名称'])
        df_2=df_2.drop(columns=['公司名称','上年同期累计发生额'])

        df_1.columns=[i.replace('.1','') for i in df_1.columns]

        with pd.ExcelWriter(os.path.join(save_folder,company+'.xlsx')) as writer:
            df_1.to_excel(writer,sheet_name='资产负债表',index=False)
            df_2.to_excel(writer,sheet_name='利润表',index=False)

#拆分费用
def depart_cost(path,save_folder,path_scope):
    df=pd.read_excel(path,sheet_name='FBL3H费用',header=0,dtype=object)
    df['凭证货币价值']=df['凭证货币价值'].astype(float).round(2)

    df_mapping=pd.read_excel(path_scope,header=0,dtype=object)
    mapping_dict=df_mapping.set_index('SAP代码')['报表名称'].to_dict()

    company_list=list(mapping_dict.keys())
    for company in company_list:
        company_name=mapping_dict[company]
        print(company,mapping_dict[company])
        df_1=df[df['公司代码']==company]
        df_1.to_csv(os.path.join(save_folder,company_name+'.csv'),index=False)

        # with pd.ExcelWriter(os.path.join(save_folder,company_name+'.xlsx')) as writer:
        #     df_1.to_excel(writer,sheet_name='Sheet1',index=False)


#拆分科目余额表,清洗负数在数字后面的情况
def depart_balance(path,save_folder):
    df=pd.read_excel(path,sheet_name='科目余额表',header=0,dtype=object)
    def clean_x(x):
        str_x=str(x)
        num=str_x.replace(',','').replace('-','')
        if '-' in str_x:
            result=-1*round(float(num),2)
        else:
            result=round(float(num),2)

        return result
    # 外币期初	外币借方	外币贷方	外币期末	本位币货币期初	本位货币借方	本位货币贷方	本位货币期末
    clean_list=['外币期初','外币借方','外币贷方','外币期末','本位币货币期初','本位货币借方','本位货币贷方','本位货币期末']
    for col in clean_list:
        df[col]=df[col].apply(clean_x)
    company_list=df['公司名称'].unique()
    for company in company_list:
        print(company)
        df_1=df[df['公司名称']==company]
        df_1=df_1.drop(columns=['公司名称']) #删除公司名称列
        with pd.ExcelWriter(os.path.join(save_folder,company+'.xlsx')) as writer:
            df_1.to_excel(writer,sheet_name='Sheet1',index=False)




if __name__=='__main__':

    #拆财务报表
    # path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\多公司报表及FBL3H.XLSX"
    # save_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\财务报表"
    # os.makedirs(save_folder,exist_ok=True)

    #拆费用
    # path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\多公司报表及FBL3H.XLSX"
    # save_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\FBL3H"
    # path_scope=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx"
    # os.makedirs(save_folder,exist_ok=True)
    # depart_cost(path,save_folder,path_scope)

    #拆科目余额表
    path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\多公司报表及FBL3H.XLSX"
    save_folder=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\科目余额表"
    os.makedirs(save_folder,exist_ok=True)
    depart_balance(path,save_folder)