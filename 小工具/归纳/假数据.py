import pandas as pd
import random
import os

# ====================== 你的文件路径（已直接填好）======================
file_path = r"E:\System\desktop\C1S\对外发送\C1S_箱体数据.xlsx"
# 输出文件名（自动保存在同一文件夹，不会覆盖原文件）E:\System\desktop\C1S箱体 - 副本.xlsx，这是我的文件路径，里面包含三个sheet，帮我写个python要求 如下，保证每个sheet的第一列不动，其他列随机交换，对三个sheet都这么 操作
output_path = os.path.splitext(file_path)[0] + "_随机打乱列.xlsx"

# 读取所有sheet
excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names  # 获取所有sheet名称

# 创建一个写入器，用于保存多个sheet
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    for sheet in sheet_names:
        print(f"正在处理 sheet: {sheet}")

        # 读取当前sheet
        df = pd.read_excel(file_path, sheet_name=sheet)

        if df.shape[1] <= 1:
            # 如果只有1列，直接保存不处理
            df.to_excel(writer, sheet_name=sheet, index=False)
            continue

        # ============== 核心逻辑 ==============
        # 第一列不动
        first_col = df.iloc[:, 0]
        # 剩下的列
        other_cols = df.iloc[:, 1:]

        # 随机打乱剩下的列顺序
        cols_list = other_cols.columns.tolist()
        random.shuffle(cols_list)
        shuffled_cols = other_cols[cols_list]

        # 合并：第一列 + 打乱后的其他列
        new_df = pd.concat([first_col, shuffled_cols], axis=1)

        # 保存到新文件
        new_df.to_excel(writer, sheet_name=sheet, index=False)

print(f"\n✅ 处理完成！文件已保存到：\n{output_path}")