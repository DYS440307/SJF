import os
import shutil

# ====================== 路径配置（直接用，无需修改）======================
source_path = r"E:\System\desktop\PY\实验室\PDF输出"
target_path = r"E:\System\desktop\PY\实验室\汇总"

# 创建目标文件夹
os.makedirs(target_path, exist_ok=True)

# 统计变量
moved_count = 0
found_count = 0

# ====================== 递归遍历【所有子文件夹】======================
print("=" * 50)
print(f"开始扫描文件夹：{source_path}")
print("正在遍历所有子文件夹...\n")

# os.walk 强制遍历所有层级的子文件夹，绝对不会漏
for root, dirs, files in os.walk(source_path):
    # 调试打印：告诉你当前正在扫描哪个文件夹（关键！）
    print(f"📂 正在扫描：{root}")

    for filename in files:
        # 筛选：文件名包含【合并报告】 + 是PDF文件（大小写都支持）
        if "合并报告" in filename and filename.lower().endswith(".pdf"):
            found_count += 1
            old_path = os.path.join(root, filename)
            new_path = os.path.join(target_path, filename)

            # 重名处理：自动加序号，不覆盖
            counter = 1
            while os.path.exists(new_path):
                name, ext = os.path.splitext(filename)
                new_path = os.path.join(target_path, f"{name}_{counter}{ext}")
                counter += 1

            # 移动文件
            try:
                shutil.move(old_path, new_path)
                print(f"✅ 已转移：{filename}")
                moved_count += 1
            except Exception as e:
                print(f"❌ 转移失败：{filename} | 原因：{e}")

# ====================== 最终结果打印 ======================
print("\n" + "=" * 50)
print(f"扫描完成！共找到符合条件文件：{found_count} 个")
print(f"成功转移文件：{moved_count} 个")
print("=" * 50)