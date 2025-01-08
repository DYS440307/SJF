import os
import xlwings as xw
#奥克斯的出货模板
# 源文件夹路径
source_folder = "F:\\system\\Desktop\\Temp\\2023"
# 目标模板文件路径
template_file = "F:\\system\\Desktop\\Temp\\奥克斯语音模组检验报告单模板.xlsx"
# 新文件保存路径
new_folder = "F:\\system\\Desktop\\Temp\\2023新"

# 确保新文件夹存在
if not os.path.exists(new_folder):
    os.makedirs(new_folder)

# 遍历源文件夹中的所有Excel文件
for file_name in os.listdir(source_folder):
    if file_name.endswith('.xlsx'):
        # 加载源Excel文件
        source_path = os.path.join(source_folder, file_name)
        source_wb = xw.Book(source_path)
        source_ws = source_wb.sheets['SY4175B4-01-X1']  # 假设数据在第一个工作表

        # 加载目标模板文件
        target_wb = xw.Book(template_file)
        target_ws = target_wb.sheets['奥克语音']  # 假设模板的数据也在第一个工作表

        # 复制单元格
        target_ws.range('J4').value = source_ws.range('J4').value
        target_ws.range('E6').value = source_ws.range('E6').value
        target_ws.range('E7').value = source_ws.range('E7').value
        target_ws.range('K7').value = source_ws.range('K7').value
        target_ws.range('K8').value = source_ws.range('K8').value

        # 生成新文件名，去除日期中的冒号和时分秒部分
        date_str = str(source_ws.range('K8').value)[:10]  # 截取日期部分，格式为 YYYY-MM-DD
        new_file_name = "奥克斯语音模组_" + date_str.replace('-', '') + ".xlsx"
        new_file_path = os.path.join(new_folder, new_file_name)

        # 保存新文件
        target_wb.save(new_file_path)
        target_wb.close()

        print(f"文件 {file_name} 已处理并保存为 {new_file_name}")

        # 关闭源文件
        source_wb.close()

print("所有文件处理完毕。")
