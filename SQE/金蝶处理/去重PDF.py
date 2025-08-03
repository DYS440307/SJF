import os
import re
from tqdm import tqdm
from datetime import datetime


def extract_code(filename):
    """从文件名中提取编码（连续的数字序列，优先取最长的数字串）"""
    # 提取所有数字序列
    number_sequences = re.findall(r'\d+', filename)
    if not number_sequences:
        return None
    # 取最长的数字序列作为编码（通常编码是最长的数字串）
    return max(number_sequences, key=len)


def deduplicate_pdf_files(folder_path):
    """
    去重指定文件夹中的PDF文件，支持两种场景：
    1. 保留原始文件，删除带有_1、_2等后缀的重复文件
    2. 对于含相同编码的文件（如100500006相关文件），保留修改日期最近的

    参数:
    folder_path: 包含PDF文件的文件夹路径
    """
    try:
        # 检查文件夹是否存在
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"文件夹不存在: {folder_path}")

        # 获取文件夹中所有PDF文件及完整路径
        pdf_files = [
            os.path.join(folder_path, f)
            for f in os.listdir(folder_path)
            if f.lower().endswith('.pdf')
        ]

        if not pdf_files:
            print("文件夹中没有找到PDF文件")
            return

        # 第一步：按编码分组文件（核心去重逻辑）
        print("开始按编码分组处理...")
        code_groups = {}
        for file_path in pdf_files:
            filename = os.path.basename(file_path)
            # 提取编码
            code = extract_code(filename)
            if code:
                if code not in code_groups:
                    code_groups[code] = []
                # 存储(文件路径, 修改时间)
                mtime = os.path.getmtime(file_path)
                code_groups[code].append((file_path, mtime))

        # 处理每个编码组：保留修改时间最新的文件
        total_deleted = 0
        with tqdm(total=len(code_groups), desc="编码组处理进度") as pbar:
            for code, files in code_groups.items():
                # 只有组内文件数>1时才需要处理
                if len(files) > 1:
                    # 按修改时间排序（最新的在最后）
                    files_sorted = sorted(files, key=lambda x: x[1])
                    # 最新的文件（最后一个）
                    latest_file = files_sorted[-1][0]
                    # 要删除的文件（除了最新的）
                    to_delete = [f[0] for f in files_sorted[:-1]]

                    print(f"\n编码 {code} 找到 {len(files)} 个相关文件，保留最新的:")
                    print(f"  保留: {os.path.basename(latest_file)} "
                          f"(修改时间: {datetime.fromtimestamp(files_sorted[-1][1])})")

                    # 删除旧文件
                    for file_path in to_delete:
                        try:
                            os.remove(file_path)
                            total_deleted += 1
                            print(f"  删除: {os.path.basename(file_path)} "
                                  f"(修改时间: {datetime.fromtimestamp(os.path.getmtime(file_path))})")
                        except Exception as e:
                            print(f"  删除文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                pbar.update(1)

        # 第二步：处理带_数字后缀的残留重复文件（补充逻辑）
        print("\n处理带数字后缀的重复文件...")
        remaining_pdfs = [
            f for f in os.listdir(folder_path)
            if f.lower().endswith('.pdf')
        ]

        # 匹配带有_数字后缀的文件名
        pattern = re.compile(r'^(.+?)_(\d+)\.pdf$', re.IGNORECASE)
        suffix_groups = {}

        for file in remaining_pdfs:
            match = pattern.match(file)
            if match:
                base_name = f"{match.group(1)}.pdf"
                if base_name not in suffix_groups:
                    suffix_groups[base_name] = []
                suffix_groups[base_name].append(file)

        # 删除带后缀的重复文件
        with tqdm(total=len(suffix_groups), desc="后缀文件处理进度") as pbar:
            for base_file, duplicates in suffix_groups.items():
                base_file_path = os.path.join(folder_path, base_file)
                if os.path.exists(base_file_path):
                    for dup_file in duplicates:
                        dup_file_path = os.path.join(folder_path, dup_file)
                        try:
                            os.remove(dup_file_path)
                            total_deleted += 1
                            print(f"已删除后缀重复文件: {dup_file}")
                        except Exception as e:
                            print(f"删除文件 {dup_file} 时出错: {str(e)}")
                pbar.update(1)

        print(f"\n去重完成，共删除 {total_deleted} 个重复文件")

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")


if __name__ == "__main__":
    # 指定PDF文件所在的文件夹路径
    folder_path = r"Z:\公共文件夹\新建文件夹 (2)\07.受控图纸"

    # 调用去重函数
    deduplicate_pdf_files(folder_path)
