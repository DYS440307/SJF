import os
import shutil


def copy_files_from_subfolders(src_dir, dest_dir):
    # 创建目标目录（如果不存在）
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)

    # 遍历源目录下的所有子文件夹
    for root, dirs, files in os.walk(src_dir):
        for file in files:
            src_file = os.path.join(root, file)
            dest_file = os.path.join(dest_dir, file)

            # 处理文件名冲突
            if os.path.exists(dest_file):
                base, extension = os.path.splitext(file)
                counter = 1
                new_dest_file = os.path.join(dest_dir, f"{base}_{counter}{extension}")
                while os.path.exists(new_dest_file):
                    counter += 1
                    new_dest_file = os.path.join(dest_dir, f"{base}_{counter}{extension}")
                dest_file = new_dest_file

            # 复制文件
            shutil.copy2(src_file, dest_file)
            print(f"Copied {src_file} to {dest_file}")


# 源目录和目标目录
src_directory = r"E:\System\desktop\PY\OQC\Input"
dest_directory = r"E:\System\desktop\PY\OQC\Input"

# 执行复制操作
copy_files_from_subfolders(src_directory, dest_directory)
