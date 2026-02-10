import pandas as pd
import os
import time

# 配置项（请根据实际Excel列名/路径修改）
FILE_PATH = r"E:\System\download\采购入库单_2026021010395077_236281.xlsx"
SUPPLIER_COL = "供应商"  # Excel中供应商列名
MATERIAL_COL = "物料编码"  # Excel中物料编码列名
OUTPUT_DIR = r"E:\System\download"  # 结果保存目录


def main():
    # 1. 检查文件是否存在
    if not os.path.exists(FILE_PATH):
        print(f"❌ 错误：文件 {FILE_PATH} 不存在，请检查路径！")
        return

    # 2. 读取并清洗数据
    df = pd.read_excel(FILE_PATH)
    df = df.dropna(subset=[SUPPLIER_COL, MATERIAL_COL])  # 删空值行
    df[SUPPLIER_COL] = df[SUPPLIER_COL].astype(str).str.strip()
    df[MATERIAL_COL] = df[MATERIAL_COL].astype(str).str.strip()

    # 3. 展示选择菜单
    print("请选择处理方式：")
    print("1. 按【供应商】去重 → 聚合对应唯一物料编码")
    print("2. 按【物料编码】去重 → 聚合对应唯一供应商")
    choice = input("输入序号（1/2）：").strip()

    # 4. 根据选择执行对应逻辑
    if choice == "1":
        # 按供应商去重，聚合物料编码
        def agg_materials(mats):
            all_mats = []
            for mat in mats:
                all_mats.extend([m.strip() for m in mat.split(";") if m.strip()])
            return ";".join(sorted(list(set(all_mats))))

        result = df.groupby(SUPPLIER_COL, as_index=False)[MATERIAL_COL].apply(agg_materials)
        result.columns = ["供应商（唯一）", "对应唯一物料编码"]
        filename = f"供应商_物料编码_去重结果_{time.strftime('%Y%m%d%H%M%S')}.xlsx"

    elif choice == "2":
        # 按物料编码去重，聚合供应商
        def agg_suppliers(sups):
            return ";".join(sorted(list(set(sups))))

        result = df.groupby(MATERIAL_COL, as_index=False)[SUPPLIER_COL].apply(agg_suppliers)
        result.columns = ["物料编码（唯一）", "对应唯一供应商"]
        filename = f"物料编码_供应商_去重结果_{time.strftime('%Y%m%d%H%M%S')}.xlsx"

    else:
        print("❌ 输入错误！仅支持输入1或2")
        return

    # 5. 直接输出结果（展示全部）
    print("\n✅ 处理结果：")
    print(result.to_string(index=False))

    # 6. 自动导出结果到Excel
    output_path = os.path.join(OUTPUT_DIR, filename)
    result.to_excel(output_path, index=False, engine="openpyxl")
    print(f"\n📁 结果已自动保存至：{output_path}")


if __name__ == "__main__":
    main()