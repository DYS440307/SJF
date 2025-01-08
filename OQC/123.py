import pandas as pd

# 询问用户输入Excel文件的路径
file_path = input("请输入Excel文件的路径: ")

# 读取Excel文件中的“7月份”和“8月份”工作表
try:
    # 读取“7月份”工作表
    sheet_july = pd.read_excel(file_path, sheet_name='7月份')
    # 读取“8月份”工作表
    sheet_august = pd.read_excel(file_path, sheet_name='8月份')

    # 提取第三列和第四列，并删除第三列的重复项
    unique_july = sheet_july.iloc[:, [2, 3]].drop_duplicates(subset=[sheet_july.columns[2]])
    unique_august = sheet_august.iloc[:, [2, 3]].drop_duplicates(subset=[sheet_august.columns[2]])

    # 将唯一项和对应的第四列值写入新的Excel文件中
    with pd.ExcelWriter('unique_items_with_values_july_august.xlsx') as writer:
        unique_july.to_excel(writer, sheet_name='7月份_唯一项', index=False)
        unique_august.to_excel(writer, sheet_name='8月份_唯一项', index=False)

    print("唯一项及对应数值已成功提取并保存到'unique_items_with_values_july_august.xlsx'文件中！")
except FileNotFoundError:
    print("文件未找到，请检查文件路径是否正确。")
except ValueError as e:
    print(f"读取工作表时发生错误: {e}")
except Exception as e:
    print(f"处理文件时发生错误: {e}")
