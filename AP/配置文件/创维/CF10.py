import openpyxl

from AP.配置文件.IMP数据迁移到报告 import write_data_to_report
from AP.配置文件.SPL实验数据处理 import process_spl_data
from AP.配置文件.THD实验数据处理 import process_thd_data
from AP.配置文件.路径配置 import experiment_report_path, IMP_path
from AP.配置文件.IMP实验数据处理 import process_excel
target_value = 1350
range_min = 200
range_max = 500
mode = 1 # 1最大值 2求最小值

SPL1 = 2000
SPL2 = 2500
SPL3 = 3000
SPL4 = 4000

THD1 = 1000
THD2 = 10000

def BnO_main_speaker_bass():
    Fb_nominal = 380 # 低音箱体的调谐频率 (Hz)
    ACR_nominal = 4  # 音响的ACR值
    SPL_nominal = 80  # 音响的声压级 (dB)
    THD_nominal = 5  # 总谐波失真 (百分比)
    B9 = "Fb(Hz)"

    # 定义波动百分比
    Fb_percentage = 20  # ±20%
    ACR_percentage = 20  # ±15%
    SPL_percentage = 3  # ±3 dB
    THD_max = 10  # ≤10%

    # 计算上下限
    Fb_min = Fb_nominal * (1 - Fb_percentage / 100)  # Fb 最小值
    Fb_max = Fb_nominal * (1 + Fb_percentage / 100)  # Fb 最大值

    ACR_min = ACR_nominal * (1 - ACR_percentage / 100)  # ACR 最小值
    ACR_max = ACR_nominal * (1 + ACR_percentage / 100)  # ACR 最大值

    SPL_min = SPL_nominal - SPL_percentage  # SPL 最小值
    SPL_max = SPL_nominal + SPL_percentage  # SPL 最大值

    # THD 为 ≤ 10%，表示最大值为 10，最小值为 0
    THD_min = 0
    THD_max = THD_nominal

    # 输出结果
    print(f"Fb范围: {Fb_min} Hz ~ {Fb_max} Hz")
    print(f"ACR范围: {ACR_min} ~ {ACR_max}")
    print(f"SPL范围: {SPL_min} dB ~ {SPL_max} dB")
    print(f"THD范围: {THD_min}% ~ {THD_max}%")

    # 将常量和波动百分比拼接成字符串
    Fb_with_percentage = f"{Fb_nominal}±{Fb_percentage}%"
    ACR_with_percentage = f"{ACR_nominal}±{ACR_percentage}%"
    SPL_with_percentage = f"{SPL_nominal}±{SPL_percentage}"
    THD_with_max = f"≤{THD_nominal}"

    # 写入实验报告文件
    try:
        # 打开实验报告的工作簿
        experiment_report_wb = openpyxl.load_workbook(experiment_report_path)

        # 检查是否有名为“实验报告”的工作表
        if "实验报告" in experiment_report_wb.sheetnames:
            experiment_report_ws = experiment_report_wb["实验报告"]
        else:
            raise ValueError("实验报告文件中没有名为‘实验报告’的工作表！")

        # 将拼接后的数据写入到指定单元格
        experiment_report_ws["B10"] = Fb_with_percentage
        experiment_report_ws["D10"] = ACR_with_percentage
        experiment_report_ws["F10"] = SPL_with_percentage
        experiment_report_ws["H10"] = THD_with_max
        experiment_report_ws["B9"] = B9

        # 保存工作簿
        experiment_report_wb.save(experiment_report_path)
        print(f"数据已成功写入：")
        print(f"Fc: {Fb_with_percentage} 到 B10")
        print(f"ACR: {ACR_with_percentage} 到 D10")
        print(f"SPL: {SPL_with_percentage} 到 F10")
        print(f"THD: {THD_with_max} 到 H10")

    except Exception as e:
        print(f"发生错误: {e}")
    # 对实验数据处理
    process_excel(target_value, range_min, range_max, mode)
    process_spl_data([SPL1, SPL2, SPL3, SPL4])
    process_thd_data(THD1, THD2)
    # 开始迁移
    write_data_to_report(IMP_path, experiment_report_path)
    print("迁移完成")



