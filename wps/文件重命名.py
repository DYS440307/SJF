import os
import re
from datetime import datetime

def rename_first_level(path, old_char, new_char):
    try:
        # 列出指定路径下的第一层级文件和文件夹
        for item in os.listdir(path):
            old_item_path = os.path.join(path, item)
            new_item = item.replace(old_char, new_char)
            new_item_path = os.path.join(path, new_item)
            # 重命名文件和文件夹
            os.rename(old_item_path, new_item_path)
        print("第一层级替换完成！")
    except Exception as e:
        print(f"出现错误: {e}")

def rename_all_levels(path, old_char, new_char):
    try:
        # 使用os.walk遍历路径下的所有文件和文件夹
        for root, dirs, files in os.walk(path):
            for name in dirs + files:
                old_item_path = os.path.join(root, name)
                new_name = name.replace(old_char, new_char)
                new_item_path = os.path.join(root, new_name)
                # 重命名文件和文件夹
                os.rename(old_item_path, new_item_path)
        print("所有层级替换完成！")
    except Exception as e:
        print(f"出现错误: {e}")

def rename_folders(directory):
    # 定义正则表达式模式，提取类似 "SYS202403-012" 的部分
    pattern = re.compile(r"SYS\d{6}\d{3}")

    # 提取排序键的函数
    def extract_sort_key(folder_name):
        match = pattern.search(folder_name)
        if match:
            return match.group(), ""  # 优先级高的排序键
        else:
            # 获取文件夹的修改时间
            folder_path = os.path.join(directory, folder_name)
            modification_time = os.path.getmtime(folder_path)
            return "", modification_time  # 优先级低的排序键

    # 获取目录中的所有文件夹
    folders = [f for f in os.listdir(directory) if os.path.isdir(os.path.join(directory, f))]

    # 对文件夹按提取出的部分和修改时间进行排序
    folders.sort(key=lambda f: extract_sort_key(f))

    # 遍历文件夹并重命名
    for i, folder in enumerate(folders, start=1):
        new_name = f"{i}_{folder}"
        old_path = os.path.join(directory, folder)
        new_path = os.path.join(directory, new_name)

        # 重命名文件夹
        os.rename(old_path, new_path)

    print("文件夹已重命名完成")

def main():
    # 选择操作类型
    print("请选择操作类型:")
    print("1: 对第一层级文件和文件夹进行替换")
    print("2: 对路径下所有层级文件和文件夹进行替换")
    print("3: 对文件夹进行重命名")
    choice = input("请输入操作类型编号 (1、2 或 3): ")

    # 根据选择调用相应的函数
    if choice in ['1', '2']:
        # 手动输入路径、要替换的字符和新字符
        path = input("请输入路径: ")
        old_char = input("请输入要替换的字符: ")
        new_char = input("请输入新的字符: ")

        if choice == '1':
            rename_first_level(path, old_char, new_char)
        elif choice == '2':
            rename_all_levels(path, old_char, new_char)
    elif choice == '3':
        path = input("请输入路径: ")
        if os.path.exists(path):
            rename_folders(path)
        else:
            print("输入的路径不存在，请检查后再试。")
    else:
        print("无效的选择，请输入 1、2 或 3")

if __name__ == "__main__":
    main()
