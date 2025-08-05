import pandas as pd
import os

def merge_suppliers(excel_path):
    # 检查文件是否存在
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到文件: {excel_path}")

    # 读取全部为字符串，并用空串填充缺失值
    df = pd.read_excel(excel_path, dtype=str).fillna('')

    # 确保必要列存在
    cols = ['子项物料编码1', '供应商名称1', '物料编码2', '供应商名称2']
    for c in cols:
        if c not in df.columns:
            raise KeyError(f"缺少列: {c}")

    # 清洗：去除前后空格
    for c in cols:
        df[c] = df[c].astype(str).str.strip()

    # 构建 “物料编码2” → [供应商名称2...] 的映射
    code2_to_sup2 = {}
    for _, row in df.iterrows():
        code2 = row['物料编码2']
        sup2 = row['供应商名称2']
        if code2 and sup2:
            code2_to_sup2.setdefault(code2, []).append(sup2)

    # 去重每个列表中的供应商
    for code2, sup_list in code2_to_sup2.items():
        code2_to_sup2[code2] = sorted(set(sup_list))

    # 遍历每行，根据子项物料编码1 找到匹配的供应商2 并合并到供应商1
    updated = 0
    for idx, row in df.iterrows():
        code1 = row['子项物料编码1']
        old_sups = [s for s in row['供应商名称1'].split(',') if s]
        matched = code2_to_sup2.get(code1, [])

        # 合并、去重、排序
        all_sups = sorted(set(old_sups + matched))
        new_val = ', '.join(all_sups)

        # 写回，如果变化则计数
        if new_val != row['供应商名称1']:
            df.at[idx, '供应商名称1'] = new_val
            updated += 1
            # 可选：打印前几条示例
            if updated <= 5:
                print(f"更新第 {idx+2} 行: 编码={code1} -> 供应商1='{new_val}'")

    # 输出结果
    out_path = os.path.splitext(excel_path)[0] + '_merged.xlsx'
    df.to_excel(out_path, index=False)
    print(f"\n共更新 {updated} 行；结果已保存到：{out_path}")

if __name__ == "__main__":
    file_path = r"E:\System\download\1.xlsx"
    merge_suppliers(file_path)
