import pandas as pd
import random
import os
import numpy as np

# ====================== 文件路径 ======================
file_path = r"E:\System\desktop\C1S\对外发送\C1S_箱体数据原档.xlsx"
output_path = os.path.splitext(file_path)[0] + "_随机打乱列_去重修复.xlsx"

# ====================== 扰动精度控制 ======================
decimal_places = 4   # 👈 控制小数位数（重点）
# ======================================================


def col_idx_to_excel(col_idx):
    col_idx += 1
    col_str = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str


def perturb_value(val, decimal_places=3):
    """
    按指定小数精度做扰动
    - 扰动范围与精度绑定
    """
    try:
        num = float(val)

        # 根据小数位自动计算扰动量级
        scale = 10 ** (-decimal_places)

        noise = num * random.uniform(-scale, scale)
        new_val = num + noise

        # 控制精度
        return round(new_val, decimal_places)

    except:
        return val


def fix_row_duplicates(df):
    print("\n🧠 开始修复行内重复值...")

    for row_idx in range(df.shape[0]):
        row = df.iloc[row_idx].tolist()

        seen = {}
        new_row = []

        for col_idx, val in enumerate(row):
            if pd.isna(val):
                new_row.append(val)
                continue

            key = str(val)

            if key not in seen:
                seen[key] = 1
                new_row.append(val)
            else:
                seen[key] += 1

                # ===== 数值扰动 =====
                if isinstance(val, (int, float)) or str(val).replace('.', '', 1).isdigit():
                    new_val = perturb_value(val, decimal_places)
                else:
                    new_val = f"{val}_{seen[key]}"

                new_row.append(new_val)

                print(f"⚠️ 第{row_idx+2}行重复 {val} -> {new_val}")

        df.iloc[row_idx] = new_row

    return df


def check_row_duplicates(df, sheet_name):
    print(f"\n🔍 检查重复值: {sheet_name}")

    for row_idx, row in df.iterrows():
        values = [v for v in row.tolist() if pd.notna(v)]
        duplicates = set([v for v in values if values.count(v) > 1])

        if duplicates:
            for dup_val in duplicates:
                cols = [i for i, v in enumerate(row.tolist()) if v == dup_val]
                cell_positions = [
                    f"{col_idx_to_excel(col)}{row_idx + 2}"
                    for col in cols
                ]
                print(f"⚠️ 仍重复 [{dup_val}] -> {', '.join(cell_positions)}")


# ====================== 主流程 ======================
excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

    for sheet in sheet_names:
        print(f"\n====== 处理 sheet: {sheet} ======")

        df = pd.read_excel(file_path, sheet_name=sheet)

        if df.shape[1] <= 1:
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # ===== 随机列 =====
        first_col = df.iloc[:, 0]
        other_cols = df.iloc[:, 1:]

        cols_list = other_cols.columns.tolist()
        random.shuffle(cols_list)
        shuffled_cols = other_cols[cols_list]

        new_df = pd.concat([first_col, shuffled_cols], axis=1)

        # ===== 去重修复 =====
        new_df = fix_row_duplicates(new_df)

        # ===== 检查 =====
        check_row_duplicates(new_df, sheet)

        # ===== 保存 =====
        new_df.to_excel(writer, sheet_name=sheet, index=False)

print(f"\n✅ 完成！输出文件：\n{output_path}")