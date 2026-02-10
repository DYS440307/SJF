import openpyxl
import os

# 定义目标Excel文件路径
file_path = r"E:\System\download\2026年1月对账.xlsx"

# 检查文件是否存在
if not os.path.exists(file_path):
    print(f"错误：文件 {file_path} 不存在，请检查路径是否正确！")
else:
    # 打开Excel文件（read_only=False才能修改）
    wb = openpyxl.load_workbook(file_path)

    # 遍历所有工作表
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]

        # 1. 设置纸张方向为横向（直接用字符串，兼容所有版本）
        ws.page_setup.orientation = "landscape"  # landscape=横向，portrait=纵向

        # 2. 设置纸张大小为A5（A5对应的标准标识是9，直接写更兼容）
        ws.page_setup.paperSize = 9  # A5=9，A4=8，可根据需要调整

        # 3. 设置缩放打印，整个工作表在一页
        ws.page_setup.fitToPage = True  # 启用适配页面模式
        ws.page_setup.fitToWidth = 1  # 宽度适配1页
        ws.page_setup.fitToHeight = 1  # 高度适配1页

    # 保存修改后的文件（建议先备份原文件，这里演示另存为新文件，避免覆盖）
    new_file_path = r"E:\System\download\2026年1月对账_调整后.xlsx"
    wb.save(new_file_path)

    # 关闭工作簿
    wb.close()

    print(f"成功！已生成调整后的文件：{new_file_path}")
    print("打印设置调整内容：")
    print("- 纸张方向：横向")
    print("- 纸张大小：A5")
    print("- 打印缩放：整个工作表适配一页")