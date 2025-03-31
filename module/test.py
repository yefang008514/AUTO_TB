import xlwings as xw

# def batch_insert_row(file_path, sheet_name, row_num, insert_count):
#     # 打开Excel文件和指定工作表
#     with xw.App(visible=False) as app:
#         wb = app.books.open(file_path)
#         sheet = wb.sheets[sheet_name]
        
#         # 计算要插入的位置和范围
#         last_row = sheet.cells.last_cell.row
#         target_range = sheet.range((row_num, 1), (row_num, sheet.cells(row_num, sheet.cells.last_cell.column).end('left').column))
        
#         # 循环插入指定次数
#         for _ in range(insert_count):
#             # 向下移动现有行
#             sheet.api.Rows(row_num).Insert()
#             # 复制原始行数据到新插入的行
#             target_range.copy(sheet.range((row_num, 1)))
        
#         # 保存并关闭
#         wb.save()
#         wb.close()
#     print(f"成功插入 {insert_count} 次行 {row_num} 到文件：{file_path}")


# def batch_insert_row(file_path, sheet_name, row_num, insert_count):
#     with xw.App(visible=False) as app:
#         wb = xw.Book(file_path)
#         sheet = wb.sheets[sheet_name]
        
#         # 选择要复制的整行
#         target_range = sheet.range(f"{row_num}:{row_num}")
        
#         # 批量插入空行
#         sheet.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=1)  # -4121表示向下移动
        
#         # 复制原始行到插入区域
#         target_range.copy(sheet.range(f"{row_num}:{row_num + insert_count - 1}"))
        
#         # 保存并关闭
#         wb.save()
#         wb.close()
#     print(f"成功插入 {insert_count} 次行 {row_num} 到文件：{file_path}")


def batch_insert_row(file_path, sheet_name, row_num, insert_count):
    with xw.App(visible=False) as app:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        
        # 插入指定次数的空行
        for _ in range(insert_count):
            sheet.api.Rows(row_num).Insert(Shift=-4121, CopyOrigin=1)
        
        # 获取原始行的内容
        target_range = sheet.range(f"{row_num + insert_count}:{row_num + insert_count}")
        
        # 复制原始行的内容到插入的空行
        for i in range(insert_count):
            target_range.copy(sheet.range(f"{row_num + i}:{row_num + i}"))
        
        # 保存并关闭
        wb.save()
        wb.close()
    print(f"成功插入 {insert_count} 次行 {row_num} 到文件：{file_path}")


def create_example_excel(file_path='example.xlsx'):
    # 创建一个新的Excel工作簿
    with xw.App(visible=False) as app:
        wb = app.books.add()
        sheet = wb.sheets[0]
        sheet.name = "Sheet1"
        
        # 填充示例数据
        data = [
            ["序号", "姓名", "年龄", "职位"],
            [1, "张三", 28, "工程师"],
            [2, "李四", 32, "经理"],
            [3, "王五", 29, "分析师"],
            [4, "赵六", 35, "主管"],
            [5, "钱七", 30, "销售"]
        ]
        sheet.range("A1").value = data
        
        # 保存Excel文件
        wb.save(file_path)
        wb.close()
    print(f"示例Excel文件已生成：{file_path}")

# 生成示例Excel文件
# create_example_excel()
# batch_insert_row('example.xlsx', 'Sheet1', 3, 5)



# batch_insert_row(r'华峰化学/测试试算/【试算】b2-重庆化工.xlsx', '8_费用', 3, 5)

import pandas as pd

a=[1,2,3,1,1,1]
temp=[i for i in a if i==1]
series_a=pd.Series(a)
print(len(series_a[series_a==1]))
# print(temp)

