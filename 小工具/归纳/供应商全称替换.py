import pandas as pd

# 定义文件路径
# 供应商全称-简称映射表路径
mapping_file = r"E:\System\desktop\PY\飞书&金蝶\供应商名称转化表.xlsx"
# 需要替换供应商名称的物料表路径
material_file = r"E:\System\desktop\PY\飞书&金蝶\物料供应商转化 - 副本.xlsx"

try:
    # ---------------------- 第一步：读取供应商全称-简称映射表 ----------------------
    # 读取转化表，默认读取第一个工作表
    df_mapping = pd.read_excel(mapping_file, engine="openpyxl")
    # 建立「供应商全称」→「简称」的映射字典（去除空值，避免匹配出错）
    df_mapping = df_mapping.dropna(subset=["供应商全称", "简称"])
    full_to_short = dict(zip(df_mapping["供应商全称"], df_mapping["简称"]))
    print("成功读取供应商映射关系，共匹配到 {} 组全称-简称对应关系".format(len(full_to_short)))

    # ---------------------- 第二步：读取物料表并替换供应商名称 ----------------------
    # 读取物料表
    df_material = pd.read_excel(material_file, engine="openpyxl")


    # 定义替换函数：将全称替换为简称，无匹配则保留原内容
    def replace_supplier_name(full_name):
        # 如果是空值/NaN，直接返回空；否则查找简称，找不到则返回原名称
        if pd.isna(full_name):
            return ""
        return full_to_short.get(full_name, full_name)


    # 替换「金蝶供应商全称」列
    if "金蝶供应商全称" in df_material.columns:
        df_material["金蝶供应商全称"] = df_material["金蝶供应商全称"].apply(replace_supplier_name)
    # 替换「飞书供应商」列
    if "飞书供应商" in df_material.columns:
        df_material["飞书供应商"] = df_material["飞书供应商"].apply(replace_supplier_name)

    # ---------------------- 第三步：保存修改后的文件 ----------------------
    # 覆盖原文件保存（如需备份，可修改文件名，比如加后缀 _已替换简称）
    df_material.to_excel(material_file, index=False, engine="openpyxl")

    print("✅ 供应商名称替换完成！修改后的文件已保存至：", material_file)

except FileNotFoundError as e:
    print(f"❌ 错误：未找到指定文件，请检查路径是否正确 → {e.filename}")
except KeyError as e:
    print(f"❌ 错误：Excel文件中缺少关键列 → {e}，请检查列名是否为「供应商全称」「简称」「金蝶供应商全称」「飞书供应商」")
except Exception as e:
    print(f"❌ 处理过程中出现未知错误：{str(e)}")