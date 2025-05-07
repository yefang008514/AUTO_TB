import shutil
import os
import pandas as pd

# 此代码用于生成空TB模板，批量重命名

def copy_and_rename_excel(source_file_path, destination_file_path):
    # 复制文件
    shutil.copy2(source_file_path, destination_file_path)
    # 重命名文件（如果需要的话，这里destination_file_path已经包含了新文件名）
    os.rename(destination_file_path, destination_file_path)

def gen_path_list_from_scope(path,save_folder):
    df=pd.read_excel(path,header=0,dtype=object)
    df['path']=df['24年试算序号'].astype(str)+'-'+df['公司名称']+'2024.xlsx'

    path_list=df['path'].tolist()
    path_list=[os.path.join(save_folder,path) for path in path_list]

    return path_list

def gen_path_list_from_scope_USD(path,save_folder):
    df=pd.read_excel(path,header=0,dtype=object)
    df['path']=df['24年试算序号'].astype(str)+'-'+df['公司名称']+'2024.xlsx'
    df=df[df['是否外币']=='是']

    path_list=df['path'].tolist()
    path_list=[os.path.join(save_folder,path) for path in path_list]

    return path_list


# # 使用示例
# source_path = r'd:\audit_project\AUTO_TB\module\original.xlsx'
# save_folder = r'd:\audit_project\AUTO_TB\module\copied'
# new_file_name = 'renamed.xlsx'

# # 确保保存文件夹存在
# # os.makedirs(save_folder, exist_ok=True)

# # 构造新的文件路径
# new_file_path = os.path.join(save_folder, new_file_name)

# # 复制并重命名文件
# copy_and_rename_excel(source_path, new_file_path)


if __name__ == '__main__':

    #本位币
    # path=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx'
    # save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算底稿'
    # path_list=gen_path_list_from_scope(path,save_folder)
    # for path in path_list:
    #     source_file_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\共享文件夹 \华峰集团2024年报\1、报表及试算\4-华峰集团试算-2024.12.31\0-单体试算（华峰集团单体专用）.xlsx"
    #     copy_and_rename_excel(source_file_path,path)

    #外币
    path=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算范围.xlsx'
    save_folder=r'C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算底稿_bk'
    path_list=gen_path_list_from_scope_USD(path,save_folder)
    for path in path_list:
        source_file_path=r"C:\Users\yefan\WPSDrive\339514258\WPS云盘\华峰集团_FY24\试算模板\外币模板.xlsx"
        copy_and_rename_excel(source_file_path,path)


    print('Done!')
