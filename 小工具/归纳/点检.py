import pandas as pd
from datetime import date, timedelta
# 导入中国法定节假日判断库
from chinese_calendar import is_holiday, is_workday

# -------------------------- 核心配置 --------------------------
# 1. 完整设备组（成组出现，顺序不变）
equipment_group = [
    "冷热冲击箱_62530003",
    "恒温恒湿箱_62530004",
    "高温烤箱_62530005",
    "恒温恒湿箱_62530002",
    "扬声器老化系统_62620011",
    "扬声器老化系统_62620001",
    "扬声器老化系统_62620006",
    "扬声器老化系统_62620016",
    "高温老化房_61270005",
    "紫外线加速耐候试验机_62530006",
    "水平垂直燃烧试验机_62670001",
    "单臂跌落试验机_62570001",
    "模拟汽车运输振动台_62560002",
    "线材摇摆试验机_62560001",
    "多功能耐摩擦试验机_62560003",
    "耐压仪/绝缘测试仪_62070001",
    "FO测试仪_62010002",
    "扫频信号发生器_62060001",
    "KLippeL_62200004"
]

# 2. 输出文件路径（按你指定的路径）
output_file = r"E:\System\download\设备点检表.xlsx"


# -------------------------------------------------------------

def get_date_range(choice):
    """
    根据选择生成对应时间段的有效日期列表
    有效日期：周一至周六 + 非中国法定节假日
    choice: 1=当月, 2=当年
    """
    today = date.today()
    valid_dates = []

    # 确定日期范围的起始和终止
    if choice == 1:
        # 当月：从当月1号到当月最后一天
        start_day = date(today.year, today.month, 1)
        if today.month == 12:
            end_day = date(today.year + 1, 1, 1)
        else:
            end_day = date(today.year, today.month + 1, 1)
    elif choice == 2:
        # 当年：从当年1月1号到次年1月1号
        start_day = date(today.year, 1, 1)
        end_day = date(today.year + 1, 1, 1)
    else:
        print("输入错误，默认生成当月数据")
        start_day = date(today.year, today.month, 1)
        end_day = date(today.year, today.month + 1, 1)

    # 遍历日期，筛选有效日期
    current_day = start_day
    while current_day < end_day:
        # 筛选条件：
        # 1. 周一至周六（weekday() 0=周一，5=周六，6=周日）
        # 2. 非法定节假日（is_holiday返回True则是节假日，需跳过）
        if current_day.weekday() < 6 and not is_holiday(current_day):
            formatted_date = current_day.strftime("%Y/%m/%d")
            valid_dates.append(formatted_date)
        current_day += timedelta(days=1)

    return valid_dates


def generate_excel(date_list):
    """生成Excel：每个日期对应完整的设备组"""
    # 构建最终数据：每个日期匹配整组设备
    final_data = []
    for dt in date_list:
        for equip in equipment_group:
            final_data.append({"日期": dt, "设备名称": equip})

    # 导出Excel
    df = pd.DataFrame(final_data)
    df.to_excel(output_file, index=False)

    # 打印生成信息
    print(f"\n✅ 生成成功！")
    print(f"📅 有效日期数量：{len(date_list)} 天（周一至周六 + 非法定节假日）")
    print(f"📝 总数据行数：{len(final_data)} 行（{len(date_list)} 个日期 × {len(equipment_group)} 个设备）")
    print(f"📁 文件保存路径：{output_file}")


# -------------------------- 交互选择 --------------------------
if __name__ == "__main__":
    print("===== 设备点检表生成工具（含节假日过滤） =====")
    print("请选择生成的日期范围：")
    print("1 - 生成当月（周一至周六+非法定节假日）的点检表")
    print("2 - 生成当年（周一至周六+非法定节假日）的点检表")

    # 获取用户选择（防输入错误）
    while True:
        try:
            user_choice = int(input("\n请输入数字（1/2）："))
            if user_choice in [1, 2]:
                break
            else:
                print("输入错误！请只输入1或2")
        except ValueError:
            print("输入错误！请输入数字1或2")

    # 生成日期列表
    print(f"\n正在生成{'当月' if user_choice == 1 else '当年'}的有效日期数据（过滤周日和法定节假日）...")
    date_list = get_date_range(user_choice)

    # 生成Excel
    generate_excel(date_list)