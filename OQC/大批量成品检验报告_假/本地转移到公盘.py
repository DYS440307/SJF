import os
import shutil

# 源目录
source_directory = r"F:\system\Desktop\PY\OQC\大批量转移\Temp"

# 目标目录
target_directory = r"Z:\3-品质部\实验室\邓洋枢工作文件夹\1-实验室相关文件\1-消音室测试报告-OQC出货性能测试\2023年"

# 获取源目录下的所有文件
files = os.listdir(source_directory)

# 遍历每个文件
for file_name in files:
    # 构造完整路径
    source_file_path = os.path.join(source_directory, file_name)

    # 检查文件是否为xlsx格式
    if file_name.endswith(".xlsx"):
        # 提取文件夹名字，即"_"和".xlsx"之间的字符串
        folder_name = file_name.split("_")[1].split(".xlsx")[0]

        # 构造目标文件夹路径
        target_folder_path = os.path.join(target_directory, folder_name)

        # 如果目标文件夹不存在，创建它
        if not os.path.exists(target_folder_path):
            os.makedirs(target_folder_path)

        # 构造目标文件路径
        target_file_path = os.path.join(target_folder_path, file_name)

        # 复制文件到目标文件夹
        shutil.copy2(source_file_path, target_file_path)

print("文件复制完成")
