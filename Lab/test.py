import openpyxl
import secrets


def generate_random_numbers(existing_values, value_range):
    min_val, max_val, min_diff = value_range
    max_attempts = 100
    for _ in range(max_attempts):
        # 生成较小值（B列或D列）
        smaller_value = round(secrets.SystemRandom().uniform(min_val, max_val - min_diff), 3)
        # 生成较大值（C列或E列），比前者大至少min_diff
        larger_value = round(secrets.SystemRandom().uniform(smaller_value + min_diff, max_val), 3)

        # 检查是否有重复
        if smaller_value not in existing_values and larger_value not in existing_values:
            return smaller_value, larger_value

    raise Exception("无法在100次尝试内生成不重复的随机数")


def write_to_excel(file_path):
    try:
        # 打开工作簿
        workbook = openpyxl.load_workbook(file_path)
        # 获取第一个工作表
        sheet = workbook.active

        # 存储已存在的值
        existing_values = set()

        # 使用通用变量名，不关联特定业务含义
        RANGE_1 = (130, 142, 2)  # 范围1配置：最小值，最大值，最小差值
        RANGE_2 = (5.8, 6.125, 0.12)  # 范围2配置：最小值，最大值，最小差值

        # 对B12:C16和D12:E16的每个单元格对进行操作
        for row in range(12, 17):
            # 生成B和C列的数据（B较小，C较大）
            value_b, value_c = generate_random_numbers(existing_values, RANGE_1)

            # 生成D和E列的数据（D较小，E较大）
            value_d, value_e = generate_random_numbers(existing_values, RANGE_2)

            # 写入数据（注意列顺序：B<C, D<E）
            sheet[f'B{row}'] = value_b
            sheet[f'C{row}'] = value_c
            sheet[f'D{row}'] = value_d
            sheet[f'E{row}'] = value_e

            # 更新已存在的值集合
            existing_values.update([value_b, value_c, value_d, value_e])

        # 保存工作簿
        workbook.save(file_path)
        print("成功写入数据到B12:E16区域，所有数值唯一")
    except Exception as e:
        print(f"发生错误：{e}")


if __name__ == "__main__":
    # 指定Excel文件路径
    file_path = r"E:\System\pic\A报告\S003-高温存储模板.xlsx"
    write_to_excel(file_path)