import streamlit as st
import pandas as pd
import time
from multiprocessing import freeze_support
import pathlib

import os,sys
sys.path.append(os.getcwd())

from module.main_flow import main_flow
from module.read_data import MappingReader,clean_start_value
from module.read_raw_report import main_flow_report
from module.workingpapaer_cost import gen_cost_workingpaper,custom_read_and_paste_main,read_excel_multi
from module.extract_inter import main_merge_raw_wb
                    

if __name__ == '__main__':
    
    freeze_support()
    # 页面配置
    st.set_page_config(page_title="试算辅助工具", page_icon="📋", layout="wide")


    # 主侧边栏导航
    st.sidebar.title("请选择功能")
    main_section = st.sidebar.radio(" ", ["1.写入数据到试算底稿", "2.从试算底稿提取数据"])
    

    # 页面逻辑
    if main_section == "1.写入数据到试算底稿":
        # 页面标题
        st.title("1.写入数据到试算底稿")
        # 模拟子侧边栏
        with st.sidebar.expander("请选择子功能"):
            mode = st.radio(" ", ["1.单文件执行", "2.批量循环执行", "3.从财务报告更新试算<原报表>"])
        #提示
        st.markdown('''
        !!!!强烈建议使用本功能前备份原始文件!!!!  
        !!!!强烈建议使用本功能前备份原始文件!!!!  
        !!!!强烈建议使用本功能前备份原始文件!!!!''')

        ##################初始化参数#################
        single_save=True
        uploaded_mapping = st.file_uploader("请上传【试算单元格映射表】", type=['xlsx','xlsm'])
        engine = st.selectbox("选择引擎", ["excel", "wps","openpyxl"])
        mode_start = st.selectbox("是否需要期初", ["否", "是"])
        project = st.selectbox("请选择项目类型",["新纪元","SAP_华峰"])

        if uploaded_mapping:
            df_mapping = MappingReader(path=uploaded_mapping, header=1).read_mapping_table()
            #1.如果不需要期初，更新df_mapping
            if mode_start=="否":
                df_mapping=clean_start_value(df_mapping)
            else:
                pass
            #2.如果需要特定sheet执行，更新df_mapping
            sheet_list = ['否']+list(df_mapping.keys())
            sheet_selected = st.selectbox("执行特定sheet?", sheet_list)
            if sheet_selected!='否':
                df_mapping={sheet_selected:df_mapping[sheet_selected]}
            else:
                pass   
            ####################[子功能模块[(上传了【试算单元格映射表】才出现)######################
            if mode == "1.单文件执行":
                st.subheader("单文件执行模式")

                path_account_balance =st.text_input("请输入科目余额表文件路径:")
                path_workingpaper = st.text_input("请输入试算底稿文件路径:")

                if st.button("执行"):
                    if path_account_balance is not None and path_workingpaper is not None:
                        try:
                            result,log_file_path = main_flow(df_mapping, path_account_balance, path_workingpaper,single_save,engine,project)
                            if len(result)>0:
                                st.success("处理完成! 日志保存在: " + log_file_path)
                                st.dataframe(result)
                            else:
                                st.success("处理完成!")
                        except Exception as e:
                            st.error(f"执行失败！错误信息：{e}")
                    else:
                        st.error("请输入所有必要的路径！")

            elif mode == "2.批量循环执行":
                st.subheader("批量循环执行模式")
                uploaded_relation = st.file_uploader("请上传【试算科余路径关系表】", type=['xlsx','xlsm'])
                if st.button("执行"):
                    if uploaded_relation:
                        df_relation = pd.read_excel(uploaded_relation, sheet_name='匹配结果', header=0)
                        list_acct_path = df_relation['科目余额表路径'].tolist()
                        list_workingpaper_path = df_relation['试算底稿路径'].tolist()

                        for i in range(len(list_acct_path)):
                            try:
                                path_account_balance = list_acct_path[i]
                                path_workingpaper = list_workingpaper_path[i]
                                result,log_file_path=main_flow(df_mapping, path_account_balance, path_workingpaper,single_save,engine)
                                #显示进度条
                                file_name_TB=list_workingpaper_path[i].split('\\')[-1]
                                st.write(f'''正在处理文件：{file_name_TB},执行进度：{i+1}/{len(list_acct_path)}''')
                                st.progress((i+1) / len(list_acct_path))

                                #若返回空result不显示日志信息
                                if len(result)>0:
                                    st.success("处理完成! 日志保存在: " + log_file_path)
                            except Exception as e:
                                st.error(f"执行失败！错误信息：{e}")
                    else:
                        st.error("请上传映射关系文件！")
            
            elif mode == "3.从财务报告更新试算<原报表>":
                st.subheader("从财务报告更新试算<原报表>")
                uploaded_finance_report = st.file_uploader("请上传【试算财务报告关系表】", type=['xlsx','xlsm'])
                if st.button("执行"):
                    if uploaded_finance_report:
                        df_relation_report = pd.read_excel(uploaded_finance_report, sheet_name='匹配结果', header=0)
                        list_finance_report_path = df_relation_report['财务报告路径'].tolist()
                        list_workingpaper_path = df_relation_report['试算底稿路径'].tolist()
                        for i in range(len(list_finance_report_path)):
                            try:
                                path_report = list_finance_report_path[i]
                                path_workingpaper = list_workingpaper_path[i]
                                result=main_flow_report(df_mapping,path_report,path_workingpaper,engine)
                                #显示进度条
                                st.write(f'''正在处理文件：{path_workingpaper},执行进度：{i+1}/{len(list_finance_report_path)}''')
                                st.progress((i+1) / len(list_finance_report_path))
                            except Exception as e:
                                st.error(f"执行失败！错误信息：{e}")
                    else:
                        st.error("请上传映射关系文件！")

    elif main_section == "2.从试算底稿提取数据":
        # 页面标题
        st.title("从试算底稿提取数据")
        # 模拟子侧边栏
        with st.sidebar.expander("请选择子功能"):
            mode = st.radio(" ", ["1.导出[销售、管理、研发费用底稿]", "2.自定义批量导出数据","3.导出往来数据"])

        if mode == "1.导出[销售、管理、研发费用底稿]":
            st.subheader("导出[销售、管理、研发费用底稿]")
            #初始化路径
            st.subheader("!!!该功能尚不成熟，请谨慎使用!!!")
            
            path_data = st.text_input("请输入【试算底稿文件夹】路径:")
            path_app = pathlib.Path(__file__).parent.resolve()#获取当前文件所在目录
            path_paper = os.path.join(path_app,r'期间费用模板_empty.xlsx')#相对路径转换成绝对路径
            path_save = st.text_input("请输入底稿保存路径:")
            read_mode = st.selectbox("选择模式", ["非穿透","穿透文件夹"])
            st.markdown('''若选择：'穿透文件夹',会读取【试算底稿文件夹】下所有文件及其子文件夹所有文件  
            若选择：'非穿透',则只读取【试算底稿文件夹】下的文件''')

            if st.button("执行"):
                try:
                    start_time = time.time()
                    gen_cost_workingpaper(path_data,path_paper,path_save,read_mode)
                    end_time = time.time()
                    st.success(f"导出完成！耗时：{round(end_time-start_time,2)}秒,详见{path_save}")
                except Exception as e:
                    st.error(f"执行失败！错误信息：{e}")
        
        elif mode == "2.自定义批量导出数据":
            st.subheader("自定义批量导出数据")
            st.markdown('''
            该功能默认将[文件夹]中所有**xlsx文件**合并，文件名称会在合并表格的**第一列**,  
            若某个xlsx文件没有某个sheet会自动跳过，  
            请注意要导出的表头的格式是否一致，不一致时会导致合并结果不符合预期
            默认保存路径在输入的文件夹路径的[合并结果.xlsx]下  
            ''')
            #初始化路径
            path = st.text_input("请输入【文件夹】路径:")
            sheet_name = st.text_input("请输入sheet名称:")
            start_cell = st.text_input("请输入开始单元格:")
            end_cell = st.text_input("请输入结束单元格:")
            time_stamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
            path_save = os.path.join(path,f'合并结果_sheet【{sheet_name}】_{time_stamp}.xlsx')
            read_mode = st.selectbox("选择模式", ["非穿透","穿透文件夹"])
            st.markdown('''若选择：'穿透文件夹',会读取【试算底稿文件夹】下所有文件及其子文件夹所有文件  
            若选择：'非穿透',则只读取【试算底稿文件夹】下的文件''')

            engine = 'openpyxl'
            header = None
            if st.button("执行"):
                try:
                    start_time = time.time()
                    #这里有默认筛选条件 可以根据条件修改
                    df=read_excel_multi(path,sheet_name,start_cell,end_cell,engine,header,read_mode)
                    df.to_excel(path_save,index=False)
                    end_time = time.time()
                    st.success(f"导出完成！耗时：{round(end_time-start_time,2)}秒,详见{path_save}")
                except Exception as e:
                    st.error(f"执行失败！错误信息：{e}")

        elif mode == "3.导出往来数据":
            st.subheader("导出往来数据")
            st.markdown('''
            该功能能够自动提取试算底稿的[应收账款、应付账款、CF2、预收账款、预付账款、其他非流动资产]的所有明细,  
            自动剔除0余额列和不相关列,将文件夹下所有公司的底稿合并输出到同一个excel文件中,  
            !!!默认不读取带**合并**字样的文件!!!
            !!!该功能目前仅针对新版试算底稿（参考FY24华峰化学、FY24东方生物）!!!  
            同一个保存路径文件夹默认覆盖之前的合并结果  
            ''')
            source_path=st.text_input("请输入【试算底稿文件夹】路径:")
            save_folder=st.text_input("请输入保存文件夹:")
            mode=st.selectbox("选择模式", ["非穿透","穿透文件夹"])
            st.markdown('''若选择：'穿透文件夹',会读取【试算底稿文件夹】下所有文件及其子文件夹所有文件  
            若选择：'非穿透',则只读取【试算底稿文件夹】下的文件''')
            if st.button("执行"):
                try:
                    main_merge_raw_wb(source_path,save_folder,mode)
                except Exception as e:
                    st.error(f"执行失败！错误信息：{e}")


    # # 添加版权信息
    st.sidebar.write("---")
    st.sidebar.write('''
    copyright
    © [20250221] [立信会计师事务所浙江分所 21部]。保留所有权利。
    使用本工具遇到任何问题，请联系：[yefang@bdo.com.cn]
    ''')

