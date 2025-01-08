import os
import xlwings as xw

def main():
    # 指定目录路径
    directory = r"F:\system\Desktop\奥克斯语音模组"

    # 获取目录下的所有文件
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

    for file in excel_files:
        file_path = os.path.join(directory, file)
        process_excel(file_path)

def process_excel(file_path):
    # 连接到现有的 Excel 应用程序或者启动一个新的
    app = xw.App(visible=False)

    # 打开工作簿
    wb = app.books.open(file_path)

    # 选择第一个工作表
    sheet = wb.sheets[0]

    # 获取单元格E7的值
    e7_value = sheet.range("E7").value

    # 根据条件写入K7单元格的值
    k7_value = get_k7_value(e7_value)
    sheet.range("K7").value = k7_value

    # 保存并关闭工作簿
    wb.save()
    wb.close()

    # 关闭 Excel 应用程序
    app.quit()

def get_k7_value(value):
    if 2 <= value <= 8:
        return 2
    elif 9 <= value <= 15:
        return 3
    elif 16 <= value <= 25:
        return 5
    elif 26 <= value <= 50:
        return 8
    elif 51 <= value <= 90:
        return 13
    elif 91 <= value <= 150:
        return 20
    elif 151 <= value <= 280:
        return 32
    elif 281 <= value <= 500:
        return 50
    elif 501 <= value <= 1200:
        return 80
    elif 1201 <= value <= 3200:
        return 125
    elif 3201 <= value <= 10000:
        return 200
    elif 10001 <= value <= 35000:
        return 315
    elif 35001 <= value <= 150000:
        return 500
    elif 150001 <= value <= 500000:
        return 800
    elif value > 500001:
        return 1250
    else:
        return None

if __name__ == "__main__":
    main()


# 下面的操作针对"F:\system\Desktop\奥克斯语音模组"目录下所有的Excel文件，
# 如果E7单元格的数据在2~8，则在K7中写入2、如果E7单元格的数据在9~15，则在K7中写入3、如果E7单元格的数据在16~25，则在K7中写入5、如果E7单元格的数据在26~50，
# 则在K7中写入8、如果E7单元格的数据在51~90，则在K7中写入13、如果E7单元格的数据在91~150，则在K7中写入20、如果E7单元格的数据在151~280，则在K7中写入32、
# 如果E7单元格的数据在281~500，则在K7中写入50、如果E7单元格的数据在501~1200，则在K7中写入80、如果E7单元格的数据在1201~3200，则在K7中写入125、如果E7单元格的数据在3201~10000，
# 则在K7中写入200、如果E7单元格的数据在10001~35000，则在K7中写入315、如果E7单元格的数据在35001~150000，则在K7中写入500、如果E7单元格的数据在150001~500000，则在K7中写入800、如果E7单元格的数据大于500001，则在K7中写入1250