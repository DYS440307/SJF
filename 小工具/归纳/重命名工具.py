import os
import shutil
import sys


def input_directory(prompt_text):
    """
    交互式获取用户输入的目录路径，并验证合法性
    """
    while True:
        # 获取用户输入（兼容Python2/3）
        if sys.version_info >= (3, 0):
            dir_path = input(prompt_text).strip()
        else:
            dir_path = raw_input(prompt_text).strip()

        # 处理空输入
        if not dir_path:
            print("⚠️ 路径不能为空，请重新输入！")
            continue

        # 标准化路径（处理斜杠、引号等问题）
        dir_path = os.path.normpath(dir_path)

        # 验证路径是否存在
        if os.path.exists(dir_path):
            # 验证是否是目录
            if os.path.isdir(dir_path):
                print(f"✅ 路径验证通过：{dir_path}")
                return dir_path
            else:
                print(f"❌ 错误：输入的路径不是目录（是文件），请重新输入！")
        else:
            # 询问是否创建目录
            create_confirm = input(f"❓ 目录 {dir_path} 不存在，是否创建？(y/n)：").strip().lower()
            if create_confirm in ['y', 'yes']:
                try:
                    os.makedirs(dir_path)
                    print(f"✅ 目录创建成功：{dir_path}")
                    return dir_path
                except Exception as e:
                    print(f"❌ 目录创建失败：{str(e)}，请重新输入！")
            else:
                print("⚠️ 请重新输入有效的目录路径！")


def input_suffix_text(prompt_text):
    """
    获取用户输入的要插入的尾部文字，处理空输入
    """
    while True:
        if sys.version_info >= (3, 0):
            suffix = input(prompt_text).strip()
        else:
            suffix = raw_input(prompt_text).strip()

        if suffix:
            return suffix
        else:
            print("⚠️ 插入的文字不能为空，请重新输入！")


def confirm_operation(target_dir, suffix_text, exclude_extensions):
    """
    展示操作信息，让用户确认是否执行
    """
    print("\n" + "=" * 60)
    print("📋 即将执行的操作信息：")
    print(f"   目标目录：{target_dir}")
    print(f"   插入到文件名尾部的文字：「{suffix_text}」")
    print(f"   排除的文件格式：{exclude_extensions}")
    print("=" * 60)

    confirm = input("\n❓ 确认执行重命名操作？(y/n，输入n取消)：").strip().lower()
    if confirm not in ['y', 'yes']:
        print("🚫 用户取消操作，程序退出！")
        return False
    return True


def add_suffix_to_files(target_dir, suffix_text, exclude_extensions=['.txt']):
    """
    给指定目录下非排除格式的文件尾部插入指定文字
    (注：仅修改文件名，不修改文件内容，会跳过目录/隐藏文件)

    参数:
        target_dir (str): 目标目录路径
        suffix_text (str): 要插入到文件名尾部的文字（扩展名前）
        exclude_extensions (list): 排除的文件扩展名列表，默认排除.txt
    """
    # 标准化扩展名（统一转小写，确保匹配准确）
    exclude_extensions = [ext.lower() for ext in exclude_extensions]

    # 统计变量
    total_files = 0
    renamed_files = 0
    skipped_files = 0

    print(f"\n📁 开始处理目录：{target_dir}")
    print(f"🔍 排除扩展名：{exclude_extensions}")
    print(f"✏️ 要插入的尾部文字：「{suffix_text}」\n")

    try:
        # 遍历目录下所有文件（仅一级，不递归子目录）
        for filename in os.listdir(target_dir):
            file_path = os.path.join(target_dir, filename)

            # 跳过目录、隐藏文件
            if os.path.isdir(file_path) or filename.startswith('.'):
                skipped_files += 1
                continue

            total_files += 1

            # 分离文件名和扩展名
            file_base, file_ext = os.path.splitext(filename)
            file_ext_lower = file_ext.lower()

            # 跳过排除的扩展名文件
            if file_ext_lower in exclude_extensions:
                print(f"⏩ 跳过排除格式：{filename}")
                skipped_files += 1
                continue

            # 构造新文件名（尾部插入指定文字）
            new_filename = f"{file_base}{suffix_text}{file_ext}"
            new_file_path = os.path.join(target_dir, new_filename)

            # 避免文件名重复（如果已存在则加数字后缀）
            counter = 1
            temp_new_path = new_file_path
            while os.path.exists(temp_new_path):
                temp_new_path = os.path.join(target_dir, f"{file_base}{suffix_text}_{counter}{file_ext}")
                counter += 1
            new_file_path = temp_new_path

            # 重命名文件（使用shutil.move兼容跨文件系统）
            shutil.move(file_path, new_file_path)
            renamed_files += 1
            print(f"✅ 重命名成功：{filename} -> {os.path.basename(new_file_path)}")

        # 输出汇总信息
        print("\n" + "=" * 50)
        print(f"📊 处理完成汇总：")
        print(f"   目录总文件数：{total_files}")
        print(f"   成功重命名：{renamed_files} 个")
        print(f"   跳过/排除：{skipped_files} 个")
        print("=" * 50)
        return True

    except Exception as e:
        print(f"\n❌ 处理过程中出错：{str(e)}")
        return False


# 主函数（交互式操作）
if __name__ == "__main__":
    print("🎉 文件名尾部插入文字工具 v1.0")
    print("🔔 说明：仅处理指定目录下的文件，不递归子目录，默认排除.txt文件\n")

    # 1. 交互式获取目标目录
    target_dir = input_directory(
        "📌 请输入要处理的目录路径（可直接粘贴）："
    )

    # 2. 交互式获取要插入的尾部文字
    insert_suffix = input_suffix_text(
        "✏️ 请输入要插入到文件名尾部的文字（如：模板）："
    )

    # 3. 确认排除的扩展名（默认.txt，可自定义）
    exclude_ext_input = input("🔧 请输入要排除的文件扩展名（多个用逗号分隔，默认.txt）：").strip()
    if exclude_ext_input:
        exclude_ext = [ext.strip().lower() for ext in exclude_ext_input.split(',')]
    else:
        exclude_ext = ['.txt']

    # 4. 用户确认操作
    if not confirm_operation(target_dir, insert_suffix, exclude_ext):
        sys.exit(0)

    # 5. 执行重命名
    add_suffix_to_files(target_dir, insert_suffix, exclude_ext)

    # 6. 结束提示
    input("\n🎊 操作完成！按回车键退出程序...")