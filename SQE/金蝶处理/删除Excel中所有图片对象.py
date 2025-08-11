import os
import win32com.client
import pythoncom


def delete_images_in_excel(file_path):
    """删除Excel文件中的所有图片"""
    try:
        # 初始化COM
        pythoncom.CoInitialize()

        # 创建Excel应用对象
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # 不显示Excel窗口
        excel.DisplayAlerts = False  # 不显示警告信息

        # 打开工作簿
        workbook = excel.Workbooks.Open(file_path)

        # 遍历所有工作表
        for worksheet in workbook.Worksheets:
            # 检查是否有图片
            if worksheet.Shapes.Count > 0:
                # 从后往前删除，避免索引问题
                for i in range(worksheet.Shapes.Count, 0, -1):
                    shape = worksheet.Shapes(i)
                    # 检查是否为图片类型
                    if shape.Type == 13:  # 13 表示图片类型
                        shape.Delete()

        # 保存并关闭工作簿
        workbook.Save()
        workbook.Close()

        print(f"已处理: {file_path}")

    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {str(e)}")
    finally:
        # 退出Excel应用
        if 'excel' in locals():
            excel.Quit()
        # 释放COM资源
        pythoncom.CoUninitialize()


def process_directory(root_dir):
    """处理目录下的所有Excel文件"""
    if not os.path.exists(root_dir):
        print(f"目录不存在: {root_dir}")
        return

    # 遍历目录及其子目录
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for filename in filenames:
            # 检查文件是否为Excel文件
            if filename.lower().endswith(('.xlsx', '.xls')):
                file_path = os.path.join(dirpath, filename)
                delete_images_in_excel(file_path)


if __name__ == "__main__":
    target_directory = r"E:\System\download\8月份"
    print(f"开始处理目录: {target_directory}")
    process_directory(target_directory)
    print("处理完成")
