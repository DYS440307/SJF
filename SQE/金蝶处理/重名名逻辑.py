import os
import re

# ===== 配置路径 =====
folder_path = r"Z:\\3-品质部\\实验室\\邓洋枢\\3-规格书\\新建文件夹\\新建文件夹 (2)\\07.受控图纸"

# ===== 文件名处理函数 =====
def normalize_filename(name):
    """
    智能分隔文件名：
    1. 空格、下划线 → “；”
    2. 数字、字母、中文之间自动加“；”
    3. 保留“字母+数字”整体（如 A0、B10、C05）
    4. 若文件名前缀为 纯数字 + 字母（如 110300006L），自动改成 110300006-L
    """
    base, ext = os.path.splitext(name)

    # 1️⃣ 替换空格、下划线为“；”
    base = base.replace(" ", "；").replace("_", "；")

    # 2️⃣ 如果前缀是数字+字母（如110300006L），加上连字符
    base = re.sub(r'^(\d+)([A-Za-z])', r'\1-\2', base)

    # 3️⃣ 临时保护字母+数字组合，如 A0、B10、C05、AA10 等
    base = re.sub(r'([A-Za-z]+[0-9]+)', r'[\1]', base)

    # 4️⃣ 在数字、字母、中文之间插入“；”
    base = re.sub(r'(?<=[0-9])(?=[A-Za-z\u4e00-\u9fff])', '；', base)
    base = re.sub(r'(?<=[A-Za-z])(?=[0-9\u4e00-\u9fff])', '；', base)
    base = re.sub(r'(?<=[\u4e00-\u9fff])(?=[A-Za-z0-9])', '；', base)

    # 5️⃣ 合并多余“；”
    base = re.sub(r'；{2,}', '；', base)

    # 6️⃣ 恢复保护的组合 [A0] → A0
    base = re.sub(r'\[([A-Za-z0-9]+)\]', r'\1', base)

    return base + ext


# ===== 遍历文件并重命名 =====
for root, dirs, files in os.walk(folder_path):
    for file in files:
        old_path = os.path.join(root, file)
        new_name = normalize_filename(file)
        if new_name != file:
            new_path = os.path.join(root, new_name)
            try:
                os.rename(old_path, new_path)
                print(f"✅ {file} → {new_name}")
            except Exception as e:
                print(f"❌ 重命名失败 {file}: {e}")
