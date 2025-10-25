import os
import re

# =========================
# 配置区域
# =========================
folder_path = r"E:\System\desktop\PY\BOMM"

# 删除关键词（文件名中包含这些词就删除）
delete_keywords = ["SOP", "控制", "承认书", "报告", "外形图", "定位治具", "变更"]

# =========================
# 文件重命名逻辑
# =========================
def clean_filename(filename):
    name, ext = os.path.splitext(filename)

    # === 删除不需要的文件 ===
    for word in delete_keywords:
        if word in name:
            return None

    # === 特殊清理逻辑 ===
    patterns = [
        (r"（\d+）", ""),        # 中文括号数字
        (r"\(\d+\)", ""),        # 英文括号数字
        (r"Model\s*\(\d+\)", ""),  # Model (1)
        (r"\s*\d+-\d+-\d+", ""),   # 例如 23-9-9
    ]
    for pattern, repl in patterns:
        name = re.sub(pattern, repl, name)

    # === 提取第一个9位数字主型号 ===
    match = re.search(r"(\d{9}(?:-[a-d])?)", name)
    if match:
        name = match.group(1)
    else:
        name = re.sub(r"[^0-9A-Za-z]", "", name)  # 兜底清理

    # 如果清理后为空，返回 None（删除）
    if not name.strip():
        return None

    return name.strip("，") + ext



# =========================
# 遍历文件夹并处理
# =========================
for root, _, files in os.walk(folder_path):
    for file in files:
        old_path = os.path.join(root, file)
        new_name = clean_filename(file)

        # 删除关键词或空名文件
        if new_name is None:
            try:
                os.remove(old_path)
                print(f"❌ 删除: {file}")
            except PermissionError:
                print(f"⚠️ 无法删除（权限或占用）：{file}")
            continue

        new_path = os.path.join(root, new_name)

        # 如果重名，进行内容比较
        if os.path.exists(new_path) and new_path != old_path:
            old_size = os.path.getsize(old_path)
            new_size = os.path.getsize(new_path)

            if old_size == new_size:
                print(f"🟡 跳过重复（内容相同）：{file}")
                try:
                    os.remove(old_path)
                except PermissionError:
                    print(f"⚠️ 无法删除重复文件（权限）：{file}")
                continue
            else:
                print(f"⚠️ 删除旧重名文件（不同内容）：{new_path}")
                try:
                    os.remove(new_path)
                except PermissionError:
                    print(f"⚠️ 无法删除旧文件（权限）：{new_path}")
                    continue

        # 执行重命名
        if new_name != file:
            try:
                os.rename(old_path, new_path)
                print(f"✅ 重命名: {file} → {new_name}")
            except PermissionError:
                print(f"⚠️ 无法重命名（被占用或权限不足）：{file}")
