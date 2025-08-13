import pandas as pd
import os
import time


def process_supplier_data():
    # 文件路径
    file1_path = r"E:\System\desktop\PY\SQE\2025年.xlsx"
    file2_path = r"E:\System\desktop\PY\SQE\声乐QCDS综合评分表 - 副本.xlsx"

    # 检查文件是否存在
    for path in [file1_path, file2_path]:
        if not os.path.exists(path):
            print(f"错误：文件不存在 - {path}")
            return

    try:
        # 读取文件1数据（无表头）
        df1 = pd.read_excel(file1_path, header=None)
        print(f"调试信息：文件1总行数为 {len(df1)}")  # 新增调试信息

        # 获取月份行(第2行，索引1)，找到7月对应的列索引
        # 先检查月份行是否存在
        if len(df1) < 2:
            print("错误：文件1数据不足，无法找到月份行")
            return

        month_row = df1.iloc[1]  # 第2行是月份行
        july_col_index = None
        for idx, value in month_row.items():
            if str(value).strip() in ['7', '7月', '七月']:
                july_col_index = idx
                break

        if july_col_index is None:
            print("错误：在文件1中未找到7月的数据列")
            return

        # 读取文件1中的供应商名称和对应的7月合格率
        supplier_data = {}
        row_count = len(df1)
        start_row = 2  # B3对应的索引（Excel行号3 → 索引2）

        # 检查起始行是否存在
        if start_row >= row_count:
            print("错误：文件1数据不足，无法找到供应商起始行")
            return

        # 遍历所有供应商（每5行一个供应商）
        for i in range(start_row, row_count, 5):
            # 供应商名称在B列（索引1）
            # 检查当前行是否存在
            if i >= row_count:
                print(f"警告：供应商行索引 {i} 超出范围，停止处理")
                break

            supplier_short = str(df1.iloc[i, 1]).strip()
            if not supplier_short or supplier_short == 'nan':
                continue

            # 合格率行在名称行下方第3行（索引i+3）
            pass_rate_row = i + 3
            # 增加严格的边界检查
            if pass_rate_row >= row_count:
                print(
                    f"警告：供应商 {supplier_short} 的合格率行索引 {pass_rate_row} 超出范围（总行数：{row_count}），跳过处理")
                continue

            # 检查7月列是否存在
            if july_col_index >= len(df1.columns):
                print(f"警告：7月列索引 {july_col_index} 超出文件1的列范围，停止处理")
                break

            # 获取7月合格率
            pass_rate_cell = df1.iloc[pass_rate_row, july_col_index]

            try:
                # 处理百分比格式
                pass_rate_str = str(pass_rate_cell).strip().replace('%', '')
                pass_rate = float(pass_rate_str)

                # 转换为小数形式
                if pass_rate > 1:
                    pass_rate = pass_rate / 100

                supplier_data[supplier_short] = pass_rate
                print(f"已读取供应商：{supplier_short}，7月合格率：{pass_rate * 100:.2f}%，计算值：{pass_rate * 45:.2f}")
            except ValueError:
                print(f"警告：供应商 {supplier_short} 的7月合格率不是有效数字，原始值：{pass_rate_cell}")
            except Exception as e:
                print(f"处理供应商 {supplier_short} 时发生错误：{str(e)}")

        if not supplier_data:
            print("警告：未从文件1中读取到任何供应商数据")
            return

        # 读取文件2数据
        df2 = pd.read_excel(file2_path)

        # 确保D列存在，不存在则创建
        if 3 >= len(df2.columns):
            df2.insert(3, '计算结果', None)
            print("警告：文件2中未找到D列，已自动创建")

        # 批量处理写入逻辑
        def match_supplier(full_name):
            full_name_str = str(full_name).strip()
            if not full_name_str or full_name_str == 'nan':
                return None

            for short_name, pass_rate in supplier_data.items():
                if short_name in full_name_str:
                    # 计算：合格率 * 45
                    return round(pass_rate * 45, 2)

            return None

        # 应用匹配函数，将计算结果写入D列
        df2.iloc[:, 3] = df2.iloc[:, 2].apply(match_supplier)

        # 统计结果
        total = len(df2)
        updated = df2.iloc[:, 3].notna().sum()
        no_matches = total - updated

        # 输出匹配报告
        print("\n" + "=" * 50)
        print(f"匹配报告:")
        print(f"总记录数: {total}")
        print(f"成功匹配并更新: {updated}")
        print(f"未找到匹配: {no_matches}")
        print("=" * 50 + "\n")

        # 打印未匹配的供应商名称
        if no_matches > 0:
            print("未匹配的供应商名称:")
            for idx, name in enumerate(df2.iloc[:, 2]):
                if pd.isna(df2.iloc[idx, 3]):
                    print(f"- {name}")

        # 尝试关闭文件（防止文件被锁定）
        time.sleep(2)

        # 直接覆盖写入原文件
        try:
            df2.to_excel(file2_path, index=False)
            print(f"\n处理完成，已直接更新原文件：{file2_path}")
            print(f"成功更新 {updated} 条记录，均为 7月合格率×45 的计算结果")
        except PermissionError:
            print(f"错误：没有权限写入文件 {file2_path}，请关闭可能打开的文件后重试")
        except Exception as e:
            print(f"保存文件时发生错误：{str(e)}")

    except Exception as e:
        print(f"处理过程中发生错误：{str(e)}")


if __name__ == "__main__":
    process_supplier_data()
