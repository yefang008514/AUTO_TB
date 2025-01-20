import streamlit as st
import pandas as pd

import os,sys
sys.path.append(os.getcwd())

from module.main_flow import main_flow
from module.read_data import MappingReader,clean_start_value




def main_streamlit():
    st.title("试算填写辅助工具")

    st.markdown(
    '''
    copyright
    © [20250110] [立信会计师事务所浙江分所 21部]。保留所有权利。
    使用本工具遇到任何问题，请联系：[yefang@bdo.com.cn]
    !!!!强烈建议使用本工具前备份原始文件!!!!  
    !!!!强烈建议使用本工具前备份原始文件!!!!  
    !!!!强烈建议使用本工具前备份原始文件!!!!  
    ''')

    mode = st.sidebar.selectbox("选择模式", ["单文件执行", "批量循环执行"])
    engine = st.selectbox("选择引擎", ["excel", "wps","openpyxl"])
    mode_start = st.selectbox("是否需要期初", ["是", "否"])

    single_save=True

    uploaded_mapping = st.file_uploader("请上传【试算单元格映射表】", type=['xlsx','xlsm'])

    if uploaded_mapping:
        df_mapping = MappingReader(path=uploaded_mapping, header=1).read_mapping_table()
        if mode_start=="否":
            df_mapping=clean_start_value(df_mapping)
        else:
            pass

        if mode == "单文件执行":
            st.subheader("单文件执行模式")

            path_account_balance =st.text_input("请输入科目余额表文件路径:")
            path_workingpaper = st.text_input("请输入试算底稿文件路径:")

            if st.button("执行"):
                if path_account_balance is not None and path_workingpaper is not None:
                    try:
                        result,log_file_path = main_flow(df_mapping, path_account_balance, path_workingpaper,single_save,engine)
                        st.success("执行完成！日志已生成。")
                        st.dataframe(result)
                    except Exception as e:
                        st.error(f"执行失败！错误信息：{e}")
                else:
                    st.error("请输入所有必要的路径！")

        elif mode == "批量循环执行":
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

if __name__ == '__main__':

    main_streamlit()

    # D:\audit_project\AUTO_TB\DATA\科目余额表\1.33 北京博朗生科技有限公司 2024.xlsx
    # C:\Users\yefan\WPSDrive\339514258\WPS云盘\东方基因\2024年试算\1.33 北京博朗生科技有限公司 2024.xlsx