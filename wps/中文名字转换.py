import openpyxl
from pypinyin import pinyin, Style

# 打开Excel文件
wb = openpyxl.load_workbook(r'F:\system\Desktop\PY\小东西\团队人员清单.xlsx')
sheet = wb.active
#  邓洋枢 Deng Yangshu

# 循环遍历B2:B17单元格中的中文名字，并将其转换成中文拼音写入C2:C17单元格中
for row in range(2, 18):
    chinese_name = sheet[f'B{row}'].value
    pinyin_name = ''
    if len(chinese_name) == 2:
        pinyin_name = ' '.join([py[0].capitalize() for py in pinyin(chinese_name, style=Style.NORMAL)])
    elif len(chinese_name) == 3:
        pinyin_name = ' '.join([py[0].capitalize() for py in pinyin(chinese_name[:2], style=Style.NORMAL)])
        pinyin_name += ''.join([py[0] for py in pinyin(chinese_name[2], style=Style.NORMAL)])
    else:
        pinyin_name = ''.join([py[0].capitalize() for py in pinyin(chinese_name, style=Style.NORMAL)])
    sheet[f'C{row}'] = pinyin_name

# 保存修改后的Excel文件
wb.save(r'F:\system\Desktop\PY\小东西\团队人员清单.xlsx')
