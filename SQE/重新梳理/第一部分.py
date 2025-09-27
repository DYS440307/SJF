import openpyxl
import os
import time


def process_excel_columns(file_path):
    """
    处理Excel文件的A列，将有内容的单元格值向下复制直到遇到下一个有内容的单元格
    同时处理A/B列组合去重（保留首次出现的组合）
    仅删除因去重操作而产生的空行（优化大量行删除效率）
    处理后直接覆盖原文件，并在控制台显示处理进度

    参数:
        file_path: Excel文件的路径
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误: 文件 '{file_path}' 不存在")
            return

        print("开始加载Excel文件...")
        # 加载工作簿（read_only=False确保可写，data_only=True可减少内存占用）
        workbook = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
        # 获取第一个工作表
        sheet = workbook.active
        print("Excel文件加载完成")

        # 获取最大行数
        max_row = sheet.max_row
        print(f"检测到文件共有 {max_row} 行数据")

        # --------------------------
        # 第一步：处理A列的填充逻辑
        # --------------------------
        print("开始处理A列填充...", end="")
        current_a_value = None
        for row in range(1, max_row + 1):
            cell_value = sheet[f'A{row}'].value
            if cell_value is not None and str(cell_value).strip() != '':
                current_a_value = cell_value
            elif current_a_value is not None:
                sheet[f'A{row}'].value = current_a_value

            if row % 1000 == 0:  # 降低打印频率，减少IO开销
                print(".", end="", flush=True)

        print("\nA列填充处理完成")

        # 重新获取最大行数
        max_row = sheet.max_row

        # --------------------------
        # 第二步：处理A/B列组合去重
        # --------------------------
        print("开始处理A/B列组合去重...", end="")
        seen_pairs = set()
        rows_to_clear = []

        for row in range(1, max_row + 1):
            a_val = str(sheet[f'A{row}'].value).strip() if sheet[f'A{row}'].value is not None else ''
            b_val = str(sheet[f'B{row}'].value).strip() if sheet[f'B{row}'].value is not None else ''
            pair = (a_val, b_val)

            if not a_val and not b_val:
                continue

            if pair in seen_pairs:
                rows_to_clear.append(row)
            else:
                seen_pairs.add(pair)

            if row % 1000 == 0:  # 降低打印频率
                print(".", end="", flush=True)

        print("\nA/B列组合去重检查完成")

        # 清空重复的组合
        print("开始清除重复行...", end="")
        for i, row in enumerate(rows_to_clear):
            sheet[f'A{row}'].value = None
            sheet[f'B{row}'].value = None

            if (i + 1) % 1000 == 0:  # 降低打印频率
                print(".", end="", flush=True)

        print("\n重复行清除完成")

        # --------------------------
        # 第三步：优化删除去重产生的空行（核心优化点）
        # --------------------------
        # 筛选出去重后A、B列都为空的行
        duplicate_empty_rows = []
        for row in rows_to_clear:
            a_val = str(sheet[f'A{row}'].value).strip() if sheet[f'A{row}'].value is not None else ''
            b_val = str(sheet[f'B{row}'].value).strip() if sheet[f'B{row}'].value is not None else ''
            if not a_val and not b_val:
                duplicate_empty_rows.append(row)

        # 按行号从大到小排序（必须保持）
        duplicate_empty_rows.sort(reverse=True)
        total_to_delete = len(duplicate_empty_rows)
        print(f"检测到 {total_to_delete} 个因去重产生的空行，开始删除...", end="")

        if total_to_delete == 0:
            print("\n没有需要删除的空行")
        else:
            # 核心优化：将连续的行分组，批量删除
            batches = []
            current_batch_start = duplicate_empty_rows[0]
            current_batch_end = duplicate_empty_rows[0]

            for row in duplicate_empty_rows[1:]:
                # 检查当前行是否与上一行连续（因已倒序，连续行是row = current_batch_end - 1）
                if row == current_batch_end - 1:
                    current_batch_end = row
                else:
                    # 新的不连续行，结束上一个批次
                    batches.append((current_batch_end, current_batch_start - current_batch_end + 1))
                    current_batch_start = row
                    current_batch_end = row
            # 添加最后一个批次
            batches.append((current_batch_end, current_batch_start - current_batch_end + 1))

            # 批量删除每个批次的行
            deleted_count = 0
            for start_row, row_count in batches:
                sheet.delete_rows(start_row, row_count)
                deleted_count += row_count

                # 每删除1000行打印一个点，减少IO操作
                if deleted_count % 1000 == 0:
                    print(".", end="", flush=True)

        print("\n去重产生的空行删除完成")

        # 保存文件（移除了错误的参数）
        print("正在保存文件...")
        workbook.save(file_path)  # 修正：移除了不支持的参数
        print(f"处理完成，已覆盖原文件: {file_path}")
        print(f"共处理 {len(rows_to_clear)} 行重复的A/B列组合")
        print(f"共删除 {len(duplicate_empty_rows)} 个因去重产生的空行")

    except Exception as e:
        print(f"\n处理文件时发生错误: {str(e)}")


if __name__ == "__main__":
    excel_path = r"E:\System\desktop\PY\SQE\关系梳理\1_采购入库单副本副本.xlsx"
    process_excel_columns(excel_path)
