import os
import openpyxl


def modify_excel_files(root_dir):
    """
    遍历指定目录及其子目录，根据配置修改符合条件的Excel文件
    同时将所有Excel文件的B2单元格设置为"品质部"
    :param root_dir: 根目录路径
    """
    # 配置列表：可动态修改的规则都在这里定义
    template_configs = [
        {
            "keyword": "常温连续负荷模板",
            "b5": "扬声器寿命测试系统-精深PS5018S"
        },
        {
            "keyword": "低温存储模板",
            "b5": "恒温恒湿箱-NTH-225C"
        },
        {
            "keyword": "低温额定功率模板",
            "b5": "扬声器寿命测试系统-精深1000A",
            "d5": "恒温恒湿箱-NTH-225C"
        },
        {
            "keyword": "高温存储模板",
            "b5": "恒温恒湿箱-HE-WS-576C9"
        },
        {
            "keyword": "高温高湿额定功率模板",
            "b5": "扬声器寿命测试系统-精深1000A",
            "d5": "恒温恒湿箱-HE-WS-576C9",
            "formula_cell": "B4",
            "formula": "=G2+8"
        }
    ]

    # 遍历目录及其子目录
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            # 检查是否为Excel文件
            if filename.endswith(('.xlsx', '.xlsm')):
                # 查找匹配的配置
                matched_config = None
                for config in template_configs:
                    if config["keyword"] in filename:
                        matched_config = config
                        break

                file_path = os.path.join(dirpath, filename)
                try:
                    # 打开Excel文件
                    workbook = openpyxl.load_workbook(file_path)
                    # 获取第一个工作表
                    sheet = workbook.active

                    # 记录修改内容，用于日志输出
                    changes = []

                    # 所有Excel文件的B2单元格都设置为"品质部"
                    sheet['B2'] = "品质部"
                    changes.append("B2设置为 品质部")

                    # 根据配置修改其他单元格（如果有匹配的配置）
                    if matched_config:
                        if "b5" in matched_config:
                            sheet['B5'] = matched_config["b5"]
                            changes.append(f"B5设置为 {matched_config['b5']}")

                        if "d5" in matched_config:
                            sheet['D5'] = matched_config["d5"]
                            changes.append(f"D5设置为 {matched_config['d5']}")

                        if "formula_cell" in matched_config and "formula" in matched_config:
                            sheet[matched_config["formula_cell"]] = matched_config["formula"]
                            changes.append(f"{matched_config['formula_cell']}设置公式为 {matched_config['formula']}")

                    # 保存修改
                    workbook.save(file_path)
                    print(f"已修改: {file_path} -> {', '.join(changes)}")

                except Exception as e:
                    print(f"处理文件 {file_path} 时出错: {str(e)}")


if __name__ == "__main__":
    # 目标目录路径
    target_directory = r"Z:\3-品质部\实验室\邓洋枢\1-实验室相关文件\3-周期验证\2025年\小米"

    # 检查目录是否存在
    if not os.path.exists(target_directory):
        print(f"错误: 目录 {target_directory} 不存在")
    else:
        print(f"开始处理目录: {target_directory}")
        modify_excel_files(target_directory)
        print("处理完成")
