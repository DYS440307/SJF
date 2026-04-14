import os
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import csv


# 统一格式化，确保 PyCharm 绝对对齐
def format_cell(value, width=10):
    s = str(value).strip() if value is not None else ""
    return s.ljust(width)[:width]


# ============================
# 【核查对比：绝对对齐版】
# ============================
def batch_check_fast_compare():
    print("=" * 80)
    print("📊 Excel 批量对比工具（PyCharm 对齐版）")
    print("=" * 80)

    folder_path = input("请输入文件夹路径：").strip()
    if not os.path.isdir(folder_path):
        print("❌ 路径错误")
        return

    target_area = input("请输入区域（如 A2:L4）：").strip()

    try:
        min_col, min_row, max_col, max_row = range_boundaries(target_area)
    except:
        print("❌ 格式错误")
        return

    # 读取所有文件
    files_data = []
    filenames = []

    for f in os.listdir(folder_path):
        if f.lower().endswith(".xlsx"):
            try:
                wb = load_workbook(os.path.join(folder_path, f), read_only=True, data_only=True)
                ws = wb.active
                rows = []
                for row in ws.iter_rows(min_row, max_row, min_col, max_col):
                    rows.append([cell.value for cell in row])
                wb.close()
                files_data.append(rows)
                filenames.append(f)
                print(f"✅ {f}")
            except:
                print(f"❌ {f}")

    if not files_data:
        print("\n未找到Excel")
        return

    # ============================
    # 核心：PyCharm 绝对对齐输出
    # ============================
    print("\n" + "=" * 120)
    print(f"🔍 对比区域：{target_area}")
    print("=" * 120)

    for row_idx in range(len(files_data[0])):
        current_row_num = min_row + row_idx
        print(f"\n📌 第 {current_row_num} 行")
        print("-" * 120)

        for file_idx, data in enumerate(files_data):
            # 文件名固定宽度
            name = f"[{filenames[file_idx]:<25}]"

            # 每一列强制等宽
            cells = [format_cell(v) for v in data[row_idx]]
            line = " │ ".join(cells)

            print(f"{name} │ {line}")

    print("\n" + "=" * 120)
    print("✅ 对比完成！PyCharm 控制台完美对齐")


# ============================
# 【批量修改】
# ============================
def batch_modify_cell():
    print("\n=== 批量修改单元格 ===")
    folder = input("路径：").strip()
    cell = input("单元格（A1）：").strip()
    content = input("新内容：").strip()

    count = 0
    for f in os.listdir(folder):
        if f.lower().endswith(".xlsx"):
            try:
                wb = load_workbook(os.path.join(folder, f))
                wb.active[cell] = content
                wb.save(os.path.join(folder, f))
                wb.close()
                count += 1
            except:
                pass
    print(f"\n✅ 完成！修改 {count} 个文件")


# ============================
# 【新增：批量替换文件名称】
# ============================
def batch_rename_files():
    print("\n" + "=" * 50)
    print("📝 批量替换文件名称（只修改Excel文件）")
    print("=" * 50)

    folder = input("请输入文件夹路径：").strip()
    if not os.path.isdir(folder):
        print("❌ 路径错误")
        return

    old_str = input("请输入要替换的文字：").strip()
    new_str = input("请输入替换后的新文字：").strip()
    confirm = input(f"\n确认替换：「{old_str}」→「{new_str}」？(y/n)：").strip().lower()

    if confirm != "y":
        print("❌ 已取消重命名")
        return

    count = 0
    for filename in os.listdir(folder):
        # 只处理Excel文件
        if not filename.lower().endswith((".xlsx", ".xls")):
            continue

        # 包含要替换的文字才修改
        if old_str in filename:
            old_path = os.path.join(folder, filename)
            new_filename = filename.replace(old_str, new_str)
            new_path = os.path.join(folder, new_filename)

            try:
                os.rename(old_path, new_path)
                print(f"✅ {filename} → {new_filename}")
                count += 1
            except Exception as e:
                print(f"❌ 重命名失败：{filename}，原因：{str(e)}")

    print(f"\n🎉 重命名完成！成功修改 {count} 个Excel文件")


# ============================
# 主程序
# ============================
if __name__ == "__main__":
    print("===== Excel 批量工具合集 =====")
    print("1 → 核查对比（对齐版）")
    print("2 → 批量修改单元格")
    print("3 → 批量替换文件名称")  # 新增选项
    print("==============================")
    c = input("请选择功能序号：").strip()
    if c == "1":
        batch_check_fast_compare()
    elif c == "2":
        batch_modify_cell()
    elif c == "3":
        batch_rename_files()  # 调用新增功能
    else:
        print("❌ 输入错误，请输入1/2/3")