import os
import shutil

# 配置参数
txt_path = r"E:\System\download\厂商ROHS、REACH\处理失败文件.txt"  # 失败文件列表路径
target_dir = r"E:\System\download\失效pdf"  # 目标复制文件夹


def copy_failed_files():
    # 1. 创建目标文件夹（不存在则创建）
    if not os.path.exists(target_dir):
        os.makedirs(target_dir)
        print(f"已创建目标文件夹: {target_dir}")

    # 2. 读取txt文件并解析路径
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"错误：未找到文件 {txt_path}")
        return
    except Exception as e:
        print(f"读取txt文件失败: {str(e)}")
        return

    # 3. 遍历每行解析源文件路径并复制
    success_count = 0
    fail_list = []

    for line in lines:
        line = line.strip()
        if not line:
            continue  # 跳过空行

        # 按 "->" 分割，提取源文件路径
        if "->" in line:
            source_path = line.split("->")[0].strip()
            # 验证源文件是否存在
            if os.path.exists(source_path):
                try:
                    # 获取文件名，拼接目标路径
                    file_name = os.path.basename(source_path)
                    target_path = os.path.join(target_dir, file_name)
                    # 复制文件（覆盖已存在的同名文件）
                    shutil.copy2(source_path, target_path)
                    success_count += 1
                    print(f"成功复制: {file_name}")
                except Exception as e:
                    fail_list.append(f"{source_path} | 复制失败: {str(e)}")
            else:
                fail_list.append(f"{source_path} | 源文件不存在")

    # 4. 输出执行结果
    print("\n" + "=" * 50)
    print(f"执行完成 | 成功复制: {success_count} 个文件")
    if fail_list:
        print(f"复制失败: {len(fail_list)} 个文件")
        for fail_info in fail_list:
            print(f"  - {fail_info}")


if __name__ == "__main__":
    copy_failed_files()
    input("按回车键退出...")