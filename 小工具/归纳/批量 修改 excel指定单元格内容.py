import os
from openpyxl import load_workbook

def batch_modify_or_check_excel():
    print("=== Excel 批量工具：修改 / 核查单元格 ===")
    print("1. 批量修改指定单元格内容")
    print("2. 批量核查指定单元格内容\n")

    # 选择模式
    mode = input("请选择功能（输入 1 或 2）：").strip()
    if mode not in ["1", "2"]:
        print("错误：请输入正确的数字 1 或 2！")
        return

    # 统一获取输入
    folder_path = input("\n请输入Excel所在文件夹路径：").strip()
    target_cell = input("请输入要操作的单元格（例如 A1、B3）：").strip()

    # 检查文件夹
    if not os.path.isdir(folder_path):
        print("错误：文件夹路径不存在！")
        return

    excel_count = 0
    success_count = 0

    print("\n========================================")
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".xlsx"):
            excel_count += 1
            file_path = os.path.join(folder_path, filename)

            try:
                wb = load_workbook(file_path, read_only=(mode == "2"))  # 核查时只读打开，更快
                ws = wb.active

                if mode == "1":
                    # ========== 模式1：修改 ==========
                    new_content = input("\n请输入单元格新内容：")
                    ws[target_cell] = new_content
                    wb.save(file_path)
                    print(f"✅ 已修改：{filename} | {target_cell} = {new_content}")

                else:
                    # ========== 模式2：核查 ==========
                    cell_value = ws[target_cell].value
                    print(f"📄 {filename} | {target_cell} = {cell_value}")

                wb.close()
                success_count += 1

            except Exception as e:
                print(f"❌ 处理失败：{filename}，原因：{str(e)}")

    print("\n========================================")
    print(f"总Excel文件：{excel_count} 个")
    print(f"成功处理：{success_count} 个")
    print("任务完成！")

if __name__ == "__main__":
    batch_modify_or_check_excel()