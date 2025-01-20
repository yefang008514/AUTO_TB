import xlwings as xw
from multiprocessing import Pool
import pandas as pd 

import os

def gen_excel(file_path):
    # 创建 Excel 应用实例
    app = xw.App(visible=False)
    try:
        # 打开 Excel 文件
        wb = app.books.add()
        # 写入示例数据
        df = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6]})
        wb.sheets[0].range('A1').value = df
        # 保存并关闭 Excel 文件
        wb.save(file_path)
        wb.close()
    finally:
        app.quit()

def process_excel(file_info):
    file_path, process_id = file_info
    # 每个进程创建自己的 Excel 应用实例
    app = xw.App(visible=False)
    try:
        wb = app.books.open(file_path)
        # 在这里执行需要的操作，例如读取 A1 单元格的值
        value = wb.sheets[0].range("A1").value
        print(f"Process {process_id}: {value}")
        # 写入示例
        wb.sheets[0].range("B1").value = f"Processed by {process_id}"
        wb.save()
        wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    os.chdir(r"D:\audit_project\AUTO_TB\test")

    file_paths = [f"example{i}.xlsx" for i in range(0, 100)]
    # for i in file_paths:
    #     gen_excel(i)

    # 假设需要处理的文件列表
    # files = [(file_path, i) for i, file_path in enumerate(file_paths)]
    
    # # 使用多进程池
    # with Pool(processes=10) as pool:  # 创建进程池
    #     pool.map(process_excel, files)

    # 假设需要处理的文件列表
    files = [(file_path, i) for i, file_path in enumerate(file_paths)]
    
    pool=Pool(processes=20)
    pool.map(gen_excel,file_paths)
    pool.close()


    # # 使用多进程池
    # with Pool(processes=20) as pool:  # 创建进程池
    #     pool.map(gen_excel, files)
