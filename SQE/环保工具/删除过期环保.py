import os
import re


def delete_expired_files_recursive(target_dir="."):
    """
    递归遍历指定目录下的所有子目录，删除文件名包含"过期时间(2025-xx-xx)"的文件

    参数:
        target_dir: 要检查的根目录路径，默认是当前目录（.）
    """
    # 正则表达式匹配 "过期时间(2025-xx-xx)" 格式
    pattern = re.compile(r'过期时间\((2025)-\d{2}-\d{2}\)')

    # 递归遍历根目录下的所有目录、子目录、文件
    # os.walk返回：当前目录路径、子目录列表、文件列表
    for root_dir, sub_dirs, files in os.walk(target_dir):
        for filename in files:
            # 拼接当前文件的完整路径
            file_path = os.path.join(root_dir, filename)

            # 只处理文件（os.walk已过滤文件夹，此处做双重校验）
            if os.path.isfile(file_path):
                # 检查文件名是否匹配正则
                if pattern.search(filename):
                    try:
                        # ========== 安全提示 ==========
                        # 首次运行建议先注释下面的 os.remove 行，只保留 print 测试
                        os.remove(file_path)
                        # print(f"【待删除】{root_dir}\\{filename}")
                        # 测试无误后，取消注释上面的 os.remove(file_path)，并注释掉这行提示
                        # print(f"已删除过期文件: {root_dir}\\{filename}")
                    except Exception as e:
                        # 捕获删除失败的异常（比如文件被占用、权限不足）
                        print(f"删除文件失败 {root_dir}\\{filename}: {str(e)}")


if __name__ == "__main__":
    # ******** 指定根目录 ********
    # 厂商ROHS、REACH 根文件夹路径（加r避免转义，保留中文/特殊符号）
    ROOT_DIRECTORY = r"E:\System\download\厂商ROHS、REACH"

    # 第一步：检查根目录是否存在
    if not os.path.exists(ROOT_DIRECTORY):
        print(f"错误：指定的根目录不存在 → {ROOT_DIRECTORY}")
    else:
        # 第二步：执行递归删除操作
        print(f"开始递归检查根目录（含所有子目录）: {os.path.abspath(ROOT_DIRECTORY)}")
        delete_expired_files_recursive(ROOT_DIRECTORY)
        print("所有目录检查完成！")