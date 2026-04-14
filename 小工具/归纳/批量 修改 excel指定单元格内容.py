import os
from openpyxl import load_workbook


def batch_modify_excel_cell():
    print("=== Excel 批量修改单元格工具 ===")

    # 1. 获取用户输入
    folder_path = input("请输入Excel所在文件夹路径：").strip()
    target_cell = input("请输入要修改的单元格（例如 A1、B3）：").strip()
    new_content = input("请输入单元格的新内容：")

    # 检查文件夹是否存在
    if not os.path.isdir(folder_path):
        print("错误：输入的文件夹路径不存在！")
        return

    # 遍历文件夹中的所有文件
    excel_count = 0
    success_count = 0

    for filename in os.listdir(folder_path):
        # 只处理 .xlsx 格式的Excel文件
        if filename.lower().endswith(".xlsx"):
            excel_count += 1
            file_path = os.path.join(folder_path, filename)

            try:
                # 打开Excel文件
                wb = load_workbook(file_path)
                # 默认修改第一个工作表
                ws = wb.active

                # 修改指定单元格
                ws[target_cell] = new_content

                # 保存文件
                wb.save(file_path)
                wb.close()

                success_count += 1
                print(f"✅ 成功修改：{filename}")

            except Exception as e:
                print(f"❌ 修改失败：{filename}，原因：{str(e)}")

    # 最终结果
    print("\n=== 执行完成 ===")
    print(f"找到Excel文件：{excel_count} 个")
    print(f"成功修改：{success_count} 个")


if __name__ == "__main__":
    batch_modify_excel_cell()