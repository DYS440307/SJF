import os
import shutil

# 定义源目录和目标目录
source_dir = r"Z:\公共文件夹\新建文件夹 (2)\DYS\01 项目承认发放资料"
target_dir = r"Z:\公共文件夹\新建文件夹 (2)\DYS\提取"  # 目标目录与源目录相同

# 遍历源目录下的所有子文件夹和文件
for root, dirs, files in os.walk(source_dir):
    # 只处理子文件夹中的文件，跳过源目录本身
    if root != source_dir:
        for file in files:
            # 检查文件是否为PDF
            if file.lower().endswith('.pdf'):
                # 构建完整的源文件路径和目标文件路径
                src_path = os.path.join(root, file)
                dest_path = os.path.join(target_dir, file)

                # 处理文件名重复的情况
                counter = 1
                while os.path.exists(dest_path):
                    # 生成新的文件名
                    name, ext = os.path.splitext(file)
                    new_file = f"{name}_{counter}{ext}"
                    dest_path = os.path.join(target_dir, new_file)
                    counter += 1

                # 移动文件
                try:
                    shutil.move(src_path, dest_path)
                    print(f"已移动: {src_path} -> {dest_path}")
                except Exception as e:
                    print(f"移动文件时出错: {src_path} - {str(e)}")

print("PDF文件提取完成！")
