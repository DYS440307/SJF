import pandas as pd
import os

def process_excel(file_path):
    # 读取文件
    df = pd.read_excel(file_path, dtype=str)

    # 去除空格，处理 NaN（用 applymap 会有警告，所以改成 apply + map）
    df = df.fillna('').apply(lambda col: col.map(lambda x: str(x).strip()))

    # 遍历每一行，根据 A 列去 C 列匹配
    for idx, row in df.iterrows():
        code_a = row['物料编码1']
        if not code_a:
            continue

        # 找到 C 列等于 A 列值的行
        matched_suppliers = df.loc[df['物料编码2'] == code_a, '供应商名称2']

        # 去重并合并成分号分隔的字符串
        supplier_list = sorted(set(matched_suppliers) - {''})
        if supplier_list:
            df.at[idx, '供应商名称1'] = ';'.join(supplier_list)

    # 保存结果
    dir_name, file_name = os.path.split(file_path)
    base_name, ext = os.path.splitext(file_name)
    output_file = os.path.join(dir_name, f"{base_name}_processed{ext}")
    df.to_excel(output_file, index=False)
    print(f"处理完成，已保存到: {output_file}")

if __name__ == '__main__':
    excel_path = r"E:\System\download\1.xlsx"
    if os.path.exists(excel_path):
        process_excel(excel_path)
    else:
        print(f"文件不存在: {excel_path}")
