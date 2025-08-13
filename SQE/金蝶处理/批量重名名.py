import os


def batch_rename_files():
    # 文件所在目录
    directory = r"E:\System\download\附件\Attachment_d8121f73-c28f-4eee-8942-e65773454464"

    # 获取目录中所有文件
    for filename in os.listdir(directory):
        # 构建完整文件路径
        old_path = os.path.join(directory, filename)

        # 只处理文件，不处理目录
        if not os.path.isfile(old_path):
            continue

        # 寻找第二个下划线的位置
        first_underscore = filename.find("_")
        if first_underscore == -1:
            continue  # 没有下划线，跳过

        second_underscore = filename.find("_", first_underscore + 1)
        if second_underscore != -1:
            # 从第二个下划线后开始截取
            new_filename = filename[second_underscore + 1:]

            # 构建新文件路径
            new_path = os.path.join(directory, new_filename)

            # 避免重名覆盖
            if os.path.exists(new_path):
                print(f"警告：文件 '{new_filename}' 已存在，跳过重命名 '{filename}'")
                continue

            # 执行重命名
            os.rename(old_path, new_path)
            print(f"已重命名：{filename} -> {new_filename}")
        # else: 不处理不符合格式的文件，静默跳过

    print("批量重命名操作完成")


if __name__ == "__main__":
    batch_rename_files()
