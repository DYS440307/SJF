# AP/主程序.py
from AP.配置文件.基础参数写入 import process_data
from AP.配置文件.层级选择 import user_choice
from AP.配置文件.路径配置 import experiment_record_path, experiment_report_path

# 用户输入的ID
user_id = input("请输入ID: ")
# # 调用共用.py中的处理函数
process_data(user_id, experiment_record_path, experiment_report_path)
user_choice()
