import os
import re
import datetime
import pandas as pd
from openpyxl import load_workbook  # 用于写入Excel（保留原有格式）


def find_shipping_plan_excel(folder_path):
    matched_files = []
    if not os.path.exists(folder_path):
        print(f"错误：文件夹 {folder_path} 不存在！")
        return matched_files
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            if "出货计划" in file_name and file_name.endswith(('.xlsx', '.xls')):
                matched_files.append(file_path)
    return matched_files


def find_inspection_template(folder_path):
    template_files = []
    if not os.path.exists(folder_path):
        print(f"错误：模板文件夹 {folder_path} 不存在！")
        return ""
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path) and file_name.endswith(('.xlsx', '.xls')):
            if re.search(r'成品检验报告单模板', file_name, re.IGNORECASE):
                template_files.append(file_path)
    if template_files:
        template_files.sort(key=lambda x: 0 if x.endswith('.xlsx') else 1)
        return template_files[0]
    return ""


def clean_sheet_name_for_material(sheet_name):
    if pd.isna(sheet_name) or sheet_name == "":
        return ""
    s = str(sheet_name).strip()
    s = re.sub(r'\s+', '', s)
    remove_words = ["成品", "检验", "报告", "单", "模板", "版", "新版", "旧版",
                    "sheet", "工作表", "测试", "副本", "最终", "正式",
                    "质检", "品质", "检测", "报告单"]
    for word in remove_words:
        s = s.replace(word, "")
    full2half = {
        '－': '-', '＿': '_', '／': '/', '＼': '\\',
        '（': '(', '）': ')', '【': '[', '】': ']', '：': ':'
    }
    for cn, en in full2half.items():
        s = s.replace(cn, en)
    s = re.sub(r'[^A-Za-z0-9\-]', '', s)
    return s.strip()


def get_inspection_sheet_names(template_file):
    sheet_names = []
    if not template_file or not os.path.exists(template_file):
        print(f"错误：模板文件 {template_file} 不存在！")
        return []
    try:
        xl_file = pd.ExcelFile(template_file)
        for raw_sheet in xl_file.sheet_names:
            cleaned_sheet = clean_sheet_name_for_material(raw_sheet)
            if cleaned_sheet:
                sheet_names.append(cleaned_sheet)
        return sheet_names
    except Exception as e:
        print(f"读取检验模板工作表失败：{e}")
        if "xlrd" in str(e):
            print("提示：读取.xls文件需要安装xlrd库，执行命令：pip install xlrd==1.2.0")
        return []


def clean_sheet_name(sheet_name):
    if not isinstance(sheet_name, str):
        return ""
    cleaned = re.sub(r'\s+', '', sheet_name)
    cleaned = re.sub(r'[^0-9一二三四五六七八九十\-]', '', cleaned)
    num_map = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5', '六': '6', '七': '7', '八': '8', '九': '9',
               '十': '10'}
    for cn_num, ar_num in num_map.items():
        cleaned = cleaned.replace(cn_num, ar_num)
    cleaned = re.sub(r'[^\d]+', '-', cleaned)
    cleaned = re.sub(r'-+', '-', cleaned)
    cleaned = cleaned.strip('-')
    return cleaned


def extract_date_sheets(file_path):
    date_sheets = []
    sheet_name_map = {}
    try:
        xl_file = pd.ExcelFile(file_path)
        all_sheets = xl_file.sheet_names
        for raw_sheet in all_sheets:
            cleaned_sheet = clean_sheet_name(raw_sheet)
            sheet_name_map[cleaned_sheet] = raw_sheet
            parts = cleaned_sheet.split('-')
            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                try:
                    month = int(parts[0])
                    day = int(parts[1])
                    if 1 <= month <= 12 and 1 <= day <= 31:
                        date_sheets.append(raw_sheet)
                except ValueError:
                    continue
        return date_sheets, sheet_name_map
    except Exception as e:
        print(f"读取文件工作表失败：{e}")
        return [], {}


def get_specified_header(file_path, sheet_name):
    specified_cols = ["客户", "交货日期", "销售订单", "物料编码", "叫料数量", "交货地点"]
    col_mapping = {}
    try:
        df_header = pd.read_excel(file_path, sheet_name=sheet_name, nrows=1, header=None)
        header_raw = df_header.iloc[0].tolist()
        for idx, raw_col in enumerate(header_raw):
            if pd.isna(raw_col):
                continue
            cleaned_raw = str(raw_col).strip()
            for spec_col in specified_cols:
                if spec_col in cleaned_raw:
                    col_mapping[cleaned_raw] = spec_col
                    break
        existing_spec_cols = list(col_mapping.values())
        return existing_spec_cols, col_mapping
    except Exception as e:
        print(f"读取表头失败：{e}")
        return [], {}


def clean_and_filter_data(file_path, sheet_name, col_mapping):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
        df.rename(columns=col_mapping, inplace=True)
        specified_cols = list(col_mapping.values())
        df_spec = df[specified_cols].copy()
        for col in df_spec.columns:
            if df_spec[col].dtype == 'object':
                df_spec[col] = df_spec[col].astype(str).str.strip()
                df_spec[col] = df_spec[col].replace("nan", "")
        df_spec.fillna("", inplace=True)
        return df_spec
    except Exception as e:
        print(f"清洗数据失败：{e}")
        return pd.DataFrame()


def filter_data_by_location(df_spec, selected_location):
    if selected_location == "全选" or "交货地点" not in df_spec.columns:
        return df_spec.copy()
    filtered_df = df_spec[df_spec["交货地点"] == selected_location].copy()
    return filtered_df


# ===================== 核心修改：写入检验单号到原Excel文件 =====================
def write_inspection_order_to_excel(file_path, sheet_name, date_str):
    """
    直接写入检验单号到原Excel文件的指定工作表
    :param file_path: 原Excel文件路径
    :param sheet_name: 目标工作表名称
    :param date_str: 日期字符串（YYYYMMDD）
    """
    try:
        # 1. 读取原工作表数据
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # 2. 生成检验单号（如果已有该列，先删除再重新生成）
        if "检验单号" in df.columns:
            df = df.drop(columns=["检验单号"])
        # 3. 生成3位递增序号的检验单号
        df['检验单号'] = [f"{date_str}{str(i + 1).zfill(3)}" for i in range(len(df))]
        # 4. 调整列顺序：检验单号放在第一列
        cols = ['检验单号'] + [col for col in df.columns if col != '检验单号']
        df = df[cols]

        # 5. 写入原Excel文件（保留其他工作表，仅更新目标工作表）
        # 方案1：xlsx格式（推荐，保留格式）
        if file_path.endswith('.xlsx'):
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        # 方案2：xls格式（兼容旧格式）
        elif file_path.endswith('.xls'):
            # 先读取所有工作表
            xl_file = pd.ExcelFile(file_path)
            all_sheets = {s: pd.read_excel(file_path, sheet_name=s) for s in xl_file.sheet_names}
            # 更新目标工作表
            all_sheets[sheet_name] = df
            # 重新写入所有工作表（xls不支持追加，只能覆盖）
            with pd.ExcelWriter(file_path, engine='xlwt') as writer:
                for s_name, s_data in all_sheets.items():
                    s_data.to_excel(writer, sheet_name=s_name, index=False)

        print(f"\n✅ 检验单号已成功写入文件：{file_path} → 工作表：{sheet_name}")
        return True
    except Exception as e:
        print(f"\n❌ 写入检验单号失败：{e}")
        return False


def match_material_code_with_inspection_sheets(filtered_df, template_file):
    if "物料编码" not in filtered_df.columns:
        print("\n⚠️ 无物料编码列，无法匹配")
        return {}

    material_codes = filtered_df["物料编码"].replace("", pd.NA).dropna().unique().tolist()
    material_codes_cleaned = []
    for code in material_codes:
        cleaned_code = clean_sheet_name_for_material(code)
        if cleaned_code:
            material_codes_cleaned.append(cleaned_code)

    # 仅保留物料编码的DEBUG输出
    print("\n===== DEBUG: 物料编码清洗后 ======")
    for i, (raw, cleaned) in enumerate(zip(material_codes, material_codes_cleaned)):
        print(f"原始：|{raw}| → 清洗后：|{cleaned}|")
        if "G0202-000464" in str(raw) or "G0202-000464" in str(cleaned):
            print(f"⚠️ 找到目标编码：原始={raw}，清洗后={cleaned}")

    if not material_codes_cleaned:
        print("\n⚠️ 无有效物料编码")
        return {}

    inspection_sheets_cleaned = get_inspection_sheet_names(template_file)

    match_result = {}
    for raw_code, cleaned_code in zip(material_codes, material_codes_cleaned):
        matched = [sheet for sheet in inspection_sheets_cleaned if sheet == cleaned_code]
        if matched:
            match_result[raw_code] = matched

    return match_result


def main():
    try:
        from tabulate import tabulate
    except ImportError:
        print("正在安装表格库...")
        os.system("pip install tabulate -q openpyxl xlwt")  # 新增写入所需库
        from tabulate import tabulate

    target_folder = r"E:\System\desktop\OQC出货报表"

    # 1. 查找出货计划文件
    excel_files = find_shipping_plan_excel(target_folder)
    if not excel_files:
        print("未找到出货计划文件")
        return

    if len(excel_files) > 1:
        print("\n找到多个文件：")
        for i, f in enumerate(excel_files, 1):
            print(f"{i}. {f}")
        c = input("请选择序号（默认1）：") or "1"
        selected_file = excel_files[int(c) - 1]
    else:
        selected_file = excel_files[0]
    print(f"\n当前文件：{selected_file}")

    # 2. 提取日期格式的工作表
    date_sheets, sheet_name_map = extract_date_sheets(selected_file)
    if not date_sheets:
        print("无有效日期工作表")
        return

    print("\n工作表清洗结果：")
    for cleaned, raw in sheet_name_map.items():
        if cleaned != raw:
            print(f"  {raw} → {cleaned}")

    print(f"\n有效日期表：{date_sheets}")

    # 3. 功能选择（选中日期）
    print("\n===== 功能选择 =====")
    print("1. 今日日期表")
    print("2. 手动输入月日")
    print("3. 当月全部表")
    func = input("序号（默认1）：") or "1"

    today = datetime.datetime.now()
    selected_sheet = None
    selected_date = today

    # 4. 选中日期后立即写入检验单号
    if func == "1":
        key = f"{today.month}-{today.day}"
        selected_sheet = sheet_name_map.get(key)
        if selected_sheet:
            print(f"\n✅ 选中今日表：{key} / {selected_sheet}")
            # 生成日期字符串并写入检验单号
            date_str = today.strftime("%Y%m%d")
            write_inspection_order_to_excel(selected_file, selected_sheet, date_str)
        else:
            print(f"\n❌ 无今日表")
            return

    elif func == "2":
        while True:
            m = input("月份：")
            d = input("日期：")
            try:
                month = int(m)
                day = int(d)
                selected_date = datetime.datetime(today.year, month, day)
                key = f"{month}-{day}"
                selected_sheet = sheet_name_map.get(key)
                if selected_sheet:
                    print(f"\n✅ 选中表：{key} / {selected_sheet}")
                    # 生成日期字符串并写入检验单号
                    date_str = selected_date.strftime("%Y%m%d")
                    write_inspection_order_to_excel(selected_file, selected_sheet, date_str)
                else:
                    print("\n❌ 无此表")
                    return
                break
            except:
                print("输入无效")

    elif func == "3":
        month_sheets = []
        for cleaned, raw in sheet_name_map.items():
            parts = cleaned.split('-')
            if len(parts) == 2 and parts[0].isdigit() and int(parts[0]) == today.month:
                month_sheets.append((int(parts[1]), raw, cleaned))
        month_sheets.sort()
        if not month_sheets:
            print("\n❌ 当月无表")
            return
        print(f"\n{today.month}月表格：")
        for i, (day, raw, cleaned) in enumerate(month_sheets, 1):
            print(f"{i}. {cleaned}")
        while True:
            c = input("选择序号：")
            try:
                day = month_sheets[int(c) - 1][0]
                selected_date = datetime.datetime(today.year, today.month, day)
                selected_sheet = month_sheets[int(c) - 1][1]
                print(f"\n✅ 选中：{selected_sheet}")
                # 生成日期字符串并写入检验单号
                date_str = selected_date.strftime("%Y%m%d")
                write_inspection_order_to_excel(selected_file, selected_sheet, date_str)
                break
            except:
                print("无效")

    # 5. 后续逻辑：读取写入后的数据，筛选交货地点，匹配模板
    # 读取写入后的工作表数据
    cols, mapping = get_specified_header(selected_file, selected_sheet)
    if not cols:
        print("\n❌ 无指定列")
        return
    df_spec = clean_and_filter_data(selected_file, selected_sheet, mapping)

    # 交货地点筛选
    location_selected = "全选"
    if "交货地点" in df_spec.columns:
        loc_list = df_spec["交货地点"].replace("", pd.NA).dropna().unique().tolist()
        if loc_list:
            show = ["全选"] + loc_list
            print("\n===== 交货地点 =====")
            for i, l in enumerate(show, 1):
                print(f"{i}. {l}")
            while True:
                c = input("选择序号（默认1）：") or "1"
                try:
                    location_selected = show[int(c) - 1]
                    print(f"\n✅ 已选交货地点：{location_selected}")
                    break
                except:
                    print("无效")

    # 筛选后数据
    final_df = filter_data_by_location(df_spec, location_selected)
    if final_df.empty:
        print("\n⚠️ 无数据")
        return

    # 输出带检验单号的数据预览（验证写入结果）
    print(f"\n===== 数据预览（{location_selected}） =====")
    print(tabulate(final_df.head(), headers='keys', tablefmt='grid', showindex=False))

    print(f"\n📊 总行数：{len(final_df)}")
    if "叫料数量" in final_df.columns:
        total = final_df["叫料数量"].replace("", 0).astype(float).sum()
        print(f"叫料总数：{total:.0f}")

    # 匹配检验模板
    print("\n===== 匹配成品检验报告单模板 =====")
    template = find_inspection_template(target_folder)
    if not template:
        print("❌ 未找到：成品检验报告单模板")
    else:
        print(f"✅ 已找到模板：{os.path.basename(template)}")
        match = match_material_code_with_inspection_sheets(final_df, template)
        if match:
            print("\n🎉 匹配成功：")
            for code, sheets in match.items():
                print(f"🔹 物料 {code} → 匹配工作表：{sheets}")
        else:
            print("\n⚠️ 无匹配的工作表（看上方DEBUG信息找原因）")


if __name__ == "__main__":
    main()