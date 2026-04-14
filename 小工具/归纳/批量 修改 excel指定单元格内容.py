import os
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

# --------------------------
# 【批量核查：支持 单个单元格 / 区域范围】
# --------------------------
def batch_check_fast_compare():
    print("=== Excel 快速对比工具 ===")
    folder_path = input("请输入Excel文件夹路径：").strip()

    if not os.path.isdir(folder_path):
        print("❌ 文件夹不存在！")
        return

    # 用户可自由输入：单个单元格（A1）或 范围（A2:L4）
    target_area = input("请输入要对比的单元格/区域（如 A1 或 A2:L4）：").strip()

    try:
        min_col, min_row, max_col, max_row = range_boundaries(target_area)
    except:
        print("❌ 区域格式错误！")
        return

    all_files_data = []
    file_names = []

    print("\n正在读取文件...")

    # 遍历读取所有Excel
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)
            try:
                wb = load_workbook(file_path, read_only=True, data_only=True)
                ws = wb.active

                file_data = []
                for row in ws.iter_rows(min_row=min_row, max_row=max_row,
                                        min_col=min_col, max_col=max_col):
                    row_values = [str(cell.value) if cell.value is not None else "" for cell in row]
                    file_data.append(row_values)

                wb.close()
                all_files_data.append(file_data)
                file_names.append(filename)
                print(f"✅ {filename}")

            except Exception as e:
                print(f"❌ 读取失败：{filename} | {e}")

    if not all_files_data:
        print("\n未找到Excel文件")
        return

    # --------------------------
    # 【清晰对比展示：完整显示，不截断】
    # --------------------------
    print("\n" + "=" * 160)
    print(f"📊 对比区域：{target_area}")
    print(f"文件总数：{len(file_names)}")
    print("=" * 160)

    for row_idx in range(len(all_files_data[0])):
        current_row = min_row + row_idx
        print(f"\n📌 第 {current_row} 行 对比")
        print("-" * 160)

        for file_idx, data in enumerate(all_files_data):
            fname = f"[{file_names[file_idx]:<25}]"
            line = " | ".join(data[row_idx])
            print(f"{fname} {line}")

    print("\n" + "=" * 160)
    print("✅ 对比完成！所有内容完整显示，可直接肉眼对比")

# --------------------------
# 【批量修改单个单元格】
# --------------------------
def batch_modify_cell():
    print("\n=== 批量修改单个单元格 ===")
    folder_path = input("文件夹路径：").strip()
    cell = input("要修改的单元格（如 A1）：").strip()
    new_val = input("新内容：").strip()

    count = 0
    for f in os.listdir(folder_path):
        if f.lower().endswith(".xlsx"):
            try:
                file_path = os.path.join(folder_path, f)
                wb = load_workbook(file_path)
                wb.active[cell] = new_val
                wb.save(file_path)
                wb.close()
                count += 1
            except:
                pass
    print(f"\n✅ 修改完成！共处理 {count} 个文件")

# --------------------------
# 主菜单
# --------------------------
if __name__ == "__main__":
    print("请选择功能：")
    print("1 → 【核查对比】支持单个单元格 / 区域范围（A2:L4）")
    print("2 → 【批量修改】单个单元格内容")
    choice = input("\n输入 1 或 2：").strip()

    if choice == "1":
        batch_check_fast_compare()
    elif choice == "2":
        batch_modify_cell()
    else:
        print("输入错误！")