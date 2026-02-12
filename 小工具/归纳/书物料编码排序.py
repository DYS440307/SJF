import pandas as pd

# 定义文件路径
file_path = r"E:\System\desktop\PY\飞书&金蝶\物料供应商转化 - 副本.xlsx"

try:
    # 读取Excel文件（默认读取第一个工作表）
    df = pd.read_excel(file_path, engine="openpyxl")

    # 1. 建立「金蝶物料编码」→「金蝶供应商全称」的映射字典
    # 去除空值行，避免匹配出错
    valid_data = df.dropna(subset=["金蝶物料编码", "金蝶供应商全称"])
    k3_code_to_supplier = dict(zip(valid_data["金蝶物料编码"], valid_data["金蝶供应商全称"]))


    # 2. 遍历每行，根据「飞书物料编码」匹配并填充「飞书供应商」
    # 定义填充逻辑的函数
    def fill_feishu_supplier(row):
        feishu_code = row["飞书物料编码"]
        # 如果飞书物料编码在金蝶编码映射中，返回对应供应商，否则保留空值
        return k3_code_to_supplier.get(feishu_code, "")


    # 应用函数填充「飞书供应商」列
    df["飞书供应商"] = df.apply(fill_feishu_supplier, axis=1)

    # 3. 保存修改后的文件（覆盖原文件，如需备份可修改文件名）
    df.to_excel(file_path, index=False, engine="openpyxl")

    print("数据匹配填充完成！文件已保存至：", file_path)

except FileNotFoundError:
    print(f"错误：未找到文件，请检查路径是否正确 → {file_path}")
except Exception as e:
    print(f"处理过程中出现错误：{str(e)}")