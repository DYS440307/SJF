import os
import re

# =========================
# 配置区域
# =========================
folder_path = r"E:\System\desktop\PY\图纸归档系统\Attachment_7a75566b-741c-44f1-8b4f-346d17656c1f"

# =========================
# 只保留 PDF 文件，删除其它格式
# =========================
for root, dirs, files in os.walk(folder_path):
    for file in files:
        file_path = os.path.join(root, file)
        if not file.lower().endswith('.pdf'):
            os.remove(file_path)
            print(f"❌ 删除非PDF文件：{file_path}")

# =========================
# 删除重名的 PDF 文件（保留最新修改的）
# =========================
pdf_files = {}
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.lower().endswith('.pdf'):
            base_name = file.lower()
            file_path = os.path.join(root, file)
            mtime = os.path.getmtime(file_path)
            if base_name not in pdf_files or mtime > pdf_files[base_name][1]:
                pdf_files[base_name] = (file_path, mtime)

# 删除重复文件
all_pdf_paths = [os.path.join(root, f) for root, _, files in os.walk(folder_path) for f in files if f.lower().endswith('.pdf')]
unique_files = {v[0] for v in pdf_files.values()}
for f in all_pdf_paths:
    if f not in unique_files:
        os.remove(f)
        print(f"🗑️ 删除重复PDF：{f}")

# =========================
# 格式化文件名：仅保留前面的物料号部分
# =========================
pattern = re.compile(r'^(\d{6,})')  # 匹配以6位以上数字开头的物料号

for root, dirs, files in os.walk(folder_path):
    for file in files:
        if not file.lower().endswith('.pdf'):
            continue

        old_path = os.path.join(root, file)
        match = pattern.match(file)
        if match:
            new_name = match.group(1) + ".pdf"
            new_path = os.path.join(root, new_name)

            # 如果存在同名文件，删除旧的再改名
            if os.path.exists(new_path) and new_path != old_path:
                os.remove(new_path)

            os.rename(old_path, new_path)
            print(f"✅ 重命名：{file} → {new_name}")
        else:
            print(f"⚠️ 未匹配物料号：{file}")
