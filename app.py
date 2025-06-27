from datetime import datetime, timedelta
import io
import re
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side

st.set_page_config(
    page_title="台灣大車隊報表整理 App",
    page_icon="🍁",
    layout="wide",
)


def parse_extension_input(extension_input):
    extension = {}
    for line in extension_input.split('\n'):
        parts = line.strip().split(':')
        if len(parts) == 2:
            employee_id, ext = parts
            extension[employee_id.strip()] = ext.strip()
    return extension


def process_dataframe(df):
    # 找到 "旅次明細表" 所在的行
    start_row = df[df.apply(lambda row: '旅次明細表' in str(row.values), axis=1)].index

    # 找到 "總共：" 所在的行
    end_row = df[df.apply(lambda row: str(row.values[0]).startswith('總共：'), axis=1)].index

    # 尋找列帳期間
    billing_period = ''
    for _, row in df.iterrows():
        if '列帳期間：' in str(row.values):
            billing_period = row.values[1]  # 假設列帳期間在最後一列
            break

    if len(start_row) > 0 and len(end_row) > 0:
        # 如果找到 "旅次明細表" 和 "總共："，則取中間的數據
        start_row = start_row[0] + 1
        end_row = end_row[0]
        df = df.iloc[start_row:end_row].reset_index(drop=True)

        # 將第一行設為列標題
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header

        # 重置索引
        df = df.reset_index(drop=True)
    elif len(start_row) > 0:
        # 如果只找到 "旅次明細表"，則從該行之後開始取數據
        start_row = start_row[0] + 1
        df = df.iloc[start_row:].reset_index(drop=True)

        # 將第一行設為列標題
        new_header = df.iloc[0]
        df = df[1:]
        df.columns = new_header

        # 重置索引
        df = df.reset_index(drop=True)
    else:
        st.warning("未找到 '旅次明細表' 行，顯示原始數據。")

    return df, billing_period


def display_employee_data(df):
    employee_column = '員工編號'
    name_column = '員工姓名'

    if employee_column not in df.columns and name_column not in df.columns:
        st.error(f"找不到 '{employee_column}' 或 '{name_column}' 列。請確保數據中包含這些列。")
        return

    # 獲取所有唯一的員工編號和姓名
    employees = df[[employee_column, name_column]].drop_duplicates()

    # 創建一個包含員工編號和姓名的選項列表
    employee_options = [f"{row[employee_column]} - {row[name_column]}" for _, row in employees.iterrows()]

    # 創建一個選擇框讓用戶選擇要查看的員工
    selected_employee = st.selectbox("選擇員工", employee_options)

    # 從選擇的選項中提員工編號
    selected_employee_id = selected_employee.split(' - ')[0]

    # 顯示選中員工的數據
    employee_data = df[df[employee_column] == selected_employee_id]
    st.subheader(f"員工 {selected_employee} 的數據")
    st.dataframe(employee_data)


def create_employee_sheets(df, billing_period, original_file, grouped_employees, extension):
    employee_column = '員工編號'
    name_column = '員工姓名'

    if employee_column not in df.columns or name_column not in df.columns:
        st.error(f"找不到 '{employee_column}' 或 '{name_column}' 列。請確保數據中包含這些列。")
        return None

    # 加載原始工作簿
    workbook = load_workbook(original_file)

    summary_sheet = workbook.create_sheet("總表")
    end_date_str = billing_period.split('~')[1].strip()
    end_date = datetime.strptime(end_date_str, "%Y 年 %m 月 %d 日")

    current_year_month = end_date.strftime("%Y/%m")
    today = (end_date + timedelta(days=1)).strftime("%Y/%m/%d")

    fixed_rows = [
        ['台灣大車隊乘車費總表', '', '', current_year_month],
        ['列帳期間：', '', '', billing_period],
        ['收據日期', '', '', today],
    ]

    for i, row in enumerate(fixed_rows):
        summary_sheet.append(row)
        summary_sheet.merge_cells(f'A{i+1}:C{i+1}')
        summary_sheet.merge_cells(f'D{i+1}:G{i+1}')

    summary_sheet.append(["NO", "員工姓名", "工號", "聯絡電話", "筆數", "折扣後車資", "ACK"])

    summary_data = []

    # 處理分組的員工
    for group, employees in grouped_employees.items():
        if not employees:
            continue

        # 使用 grouped_employees 中的第一個員工編號作為代表
        first_employee_id = employees[0]
        group_df = df[df[employee_column].isin(employees)]

        if group_df.empty:
            continue

        # 獲取第一個員工的資訊
        first_employee_data = group_df[group_df[employee_column] == first_employee_id].iloc[0]
        first_employee_name = first_employee_data[name_column]

        total_count = len(group_df)
        total_amount = group_df['折扣後車資'].sum() if '折扣後車資' in group_df.columns else 0

        summary_data.append(
            [
                len(summary_data) + 1,
                first_employee_name,
                first_employee_id,
                extension.get(first_employee_id, ""),
                total_count,
                total_amount,
                "",
            ]
        )

        sheet_name = f'{first_employee_id} {first_employee_name}'
        sheet_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)
        worksheet = workbook.create_sheet(sheet_name)

        # 创建固定的行内容
        fixed_rows = [
            ['企業會員乘車服務電子對帳單'],
            ['客戶名稱：', '', '友訊科技股份有限公司'],
            ['列帳期間：', '', billing_period],
        ]

        # 将固定行内容写入工作表
        for row in fixed_rows:
            worksheet.append(row)

        # 將組內所有員工的數據寫入工作表，按照員工編號排序
        sorted_group_df = group_df.sort_values(by=employee_column)
        worksheet.append(sorted_group_df.columns.tolist())
        for _, row in sorted_group_df.iterrows():
            worksheet.append(row.tolist())

        # 计算统计数据
        total_count = len(sorted_group_df)
        total_amount = sorted_group_df['折扣後車資'].sum() if '折扣後車資' in sorted_group_df.columns else 0

        # 创建统计数据行
        stats_rows = [
            ['總筆數', total_count, '', '', '', '', '折扣後：', total_amount],
            [],
            ['*車資總計(運送服務費)：', '', '', '', '', '', f"{total_amount}元"],
            ['乘車券印製費：', '', '', '', '', '', '0元'],
            ['滯納金：', '', '', '', '', '', '0元'],
            ['其它費用：', '', '', '', '', '', '0元'],
            ['本期應繳帳款：', '', '', '', '', '', f"{total_amount}元"],
            ['特殊費用：', '', '', '', '', '', '0元'],
        ]

        # 写入统计数据行
        for row in stats_rows:
            worksheet.append(row)

        # 设置字体大小和调整列宽
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(size=12)

        for idx, column in enumerate(sorted_group_df.columns):
            column_letter = get_column_letter(idx + 1)
            if column in ['上車地點', '下車地點']:
                worksheet.column_dimensions[column_letter].width = 12
            else:
                max_length = max(sorted_group_df[column].astype(str).map(len).max() + 4, len(str(column)) + 6)
                worksheet.column_dimensions[column_letter].width = max_length

        # 合併第一行單元格並置中
        max_col = len(sorted_group_df.columns)
        worksheet.merge_cells(f'A1:{get_column_letter(max_col)}1')
        title_cell = worksheet['A1']
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.merge_cells(f'C2:{get_column_letter(max_col)}2')
        worksheet.merge_cells(f'C3:{get_column_letter(max_col)}3')

        # 设第一行为粗体
        bold_font = Font(size=12, bold=True)
        title_cell.font = bold_font
        start_row = len(fixed_rows) + len(sorted_group_df) + 2
        total_count_cell = worksheet.cell(row=start_row, column=1)
        total_count_cell.font = bold_font
        total_count_cell = worksheet.cell(row=start_row, column=2)
        total_count_cell.font = bold_font
        total_amount_cell = worksheet.cell(row=start_row, column=7)
        total_amount_cell.font = bold_font
        total_amount_cell = worksheet.cell(row=start_row, column=8)
        total_amount_cell.font = bold_font
        payable_amount_cell = worksheet.cell(row=start_row + 6, column=1)
        payable_amount_cell.font = bold_font
        payable_amount_value_cell = worksheet.cell(row=start_row + 6, column=7)
        payable_amount_value_cell.font = bold_font

        # 添加外框线
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )

        # 为数据部分添加全部框线
        for row in worksheet[f'A1' :f'{get_column_letter(max_col)}{worksheet.max_row}']:
            for cell in row:
                cell.border = thin_border

    # 處理未分組的員工
    ungrouped_employees = set(df[employee_column]) - set(employee for group in grouped_employees.values() for employee in group)
    for employee in ungrouped_employees:
        employee_df = df[df[employee_column] == employee]
        employee_name = employee_df[name_column].iloc[0]

        total_count = len(employee_df)
        total_amount = employee_df['折扣後車資'].sum() if '折扣後車資' in employee_df.columns else 0

        summary_data.append(
            [
                len(summary_data) + 1,
                employee_name,
                employee,
                extension.get(employee, ""),
                total_count,
                total_amount,
                "",
            ]
        )

        sheet_name = f'{employee} {employee_name}'
        sheet_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)
        worksheet = workbook.create_sheet(sheet_name)

        # 创建固定的行内容
        fixed_rows = [
            ['企業會員乘車服務電子對帳單'],
            ['客戶名稱：', '', '友訊科技股份有限公司'],
            ['列帳期間：', '', billing_period],
        ]

        # 将固定行内容写入工作表
        for row in fixed_rows:
            worksheet.append(row)

        # 重置 employee_df 的索引并保留原始列名
        employee_df_reset = employee_df.reset_index(drop=True)

        # 将员工数据包括标题）写入工作表
        worksheet.append(employee_df_reset.columns.tolist())
        for _, row in employee_df_reset.iterrows():
            worksheet.append(row.tolist())

        # 计算统计数据
        total_count = len(employee_df_reset)
        total_amount = employee_df_reset['折扣後車資'].sum() if '折扣後車資' in employee_df_reset.columns else 0

        # 创建统计数据行
        stats_rows = [
            ['總筆數', total_count, '', '', '', '', '折扣後：', total_amount],
            [],
            ['*車資總計(運送服務費)：', '', '', '', '', '', f"{total_amount}元"],
            ['乘車券印製費：', '', '', '', '', '', '0元'],
            ['滯納金：', '', '', '', '', '', '0元'],
            ['其它費用：', '', '', '', '', '', '0元'],
            ['本期應繳帳款：', '', '', '', '', '', f"{total_amount}元"],
            ['特殊費用：', '', '', '', '', '', '0元'],
        ]

        # 写入统计数据行
        for row in stats_rows:
            worksheet.append(row)

        # 设置字体大小和调整列宽
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(size=12)

        for idx, column in enumerate(employee_df_reset.columns):
            column_letter = get_column_letter(idx + 1)
            if column in ['上車地點', '下車地點']:
                worksheet.column_dimensions[column_letter].width = 12
            else:
                max_length = max(employee_df_reset[column].astype(str).map(len).max() + 4, len(str(column)) + 6)
                worksheet.column_dimensions[column_letter].width = max_length

        # 合併第一行單元格並置中
        max_col = len(employee_df_reset.columns)
        worksheet.merge_cells(f'A1:{get_column_letter(max_col)}1')
        title_cell = worksheet['A1']
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.merge_cells(f'C2:{get_column_letter(max_col)}2')
        worksheet.merge_cells(f'C3:{get_column_letter(max_col)}3')

        # 设置第一行为粗体
        bold_font = Font(size=12, bold=True)
        title_cell.font = bold_font
        start_row = len(fixed_rows) + len(employee_df_reset) + 2
        total_count_cell = worksheet.cell(row=start_row, column=1)
        total_count_cell.font = bold_font
        total_count_cell = worksheet.cell(row=start_row, column=2)
        total_count_cell.font = bold_font
        total_amount_cell = worksheet.cell(row=start_row, column=7)
        total_amount_cell.font = bold_font
        total_amount_cell = worksheet.cell(row=start_row, column=8)
        total_amount_cell.font = bold_font
        payable_amount_cell = worksheet.cell(row=start_row + 6, column=1)
        payable_amount_cell.font = bold_font
        payable_amount_value_cell = worksheet.cell(row=start_row + 6, column=7)
        payable_amount_value_cell.font = bold_font

        # 添加外框线
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )

        # 为数据部分添加全部框线
        for row in worksheet[f'A1' :f'{get_column_letter(max_col)}{worksheet.max_row}']:
            for cell in row:
                cell.border = thin_border

    # 根據員工編號排序 summary_data
    summary_data.sort(key=lambda x: x[2])  # x[2] 是員工編號

    # 將排序後的 summary_data 寫入總表
    for i, row in enumerate(summary_data, start=1):
        row[0] = i  # 更新序號
        summary_sheet.append(row)

    # 计算总计
    total_count = sum(row[4] for row in summary_data)
    total_amount = sum(row[5] for row in summary_data)
    summary_sheet.append(["合計", "", "", "", total_count, total_amount, ""])
    summary_sheet.merge_cells(f"A{len(summary_data) + 5}:D{len(summary_data) + 5}")

    # 设置总表格式
    for row in summary_sheet.iter_rows(min_row=1, max_row=len(summary_data) + 5, min_col=1, max_col=7):
        for cell in row:
            cell.font = Font(size=12)
            cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin'),
            )

    for row in summary_sheet.iter_rows(min_row=1, max_row=3):
        for col, cell in enumerate(row, start=1):
            if col == 1:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col == 4:
                cell.alignment = Alignment(horizontal='left', vertical='center')

    for row in summary_sheet.iter_rows(min_row=4):
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 在處理完所有工作表後，重新排序
    sheets = workbook.sheetnames
    # 只對員工工作表進行排序，保留前兩個工作表不變
    employee_sheets = sheets[2:]

    # 根據工作表名稱（員工編號）進行排序
    sorted_sheets = sorted(employee_sheets, key=lambda x: x.split()[0])

    # 重新排列工作表
    for i, sheet_name in enumerate(sorted_sheets, start=2):
        workbook.move_sheet(sheet_name, offset=i - workbook.index(workbook[sheet_name]))

    # 将修改后的工作保存到内存中
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output


def get_all_employee_ids(df):
    employee_column = '員工編號'
    if employee_column not in df.columns:
        st.error(f"找不到 '{employee_column}' 列。請確保數據中包含此列。")
        return []
    return sorted(df[employee_column].unique().tolist())


def main():
    st.title("Excel數據整理工具(2025-06)")

    # 上传Excel文件
    uploaded_file = st.file_uploader("請上傳Excel文件", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 獲取上傳文件的原始名稱
        original_filename = uploaded_file.name

        # 讀取所有工作表
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        # 讓用戶選擇工作表
        selected_sheet = st.selectbox("請選擇要處理的工作表", sheet_names)

        # 讀取選定的工作表
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=None)

        # 顯示原始數據
        st.subheader(f"原始數據 - {selected_sheet}")
        st.dataframe(df)

        # 處理數據
        processed_df, billing_period = process_dataframe(df)

        # 顯示處理後的數據
        st.subheader(f"處理後的數據 - {selected_sheet}")
        st.dataframe(processed_df)

        # 顯示每個員工的數據
        display_employee_data(processed_df)

        # 获取所有员工编号
        all_employee_ids = get_all_employee_ids(processed_df)

        # 添加输入框让用户输入员工编号和分机的对应关系
        default_extension_input = "00156: [24104] 101\n00451: [24104] 106\n00605: [24103] 26\n00640: 8800\n01037: 8887\n01746: 8801\n02100: [24104] 152\n02177: 6524\n02230: [24103] 25\n02345: 2257\n02435: [24104] 154\n03057: 6429\n04352: 6412\n04655: 6714\n04779: 6522\n04826: 2347\n05055: 5313\n05259: 2340\n05350: 5815\n05667: 6723\n05709: 5683\n05834: 8233\n05838: 6762\n05839: 2222\n05870: 8238\n05871: 8237\n05876: 5314\n05931: 5234\n05959: 6113\n06000: 6814\n06075: 5684\n06076: 5719\n06101: 6618\n06183: 5663\n06265: 2254\n06267: 6432\n06290: 8235\n06294: 8254\n06306: 6621\n06309: 2377\n06312: 2321\n06403: 2253\n06404: 5227\n06408: 6813\n06462: 0\n06466: 5691\n06468: 2218\n06629: 5674\n06632: 5688\n06634: 2258\n06723: 8228\n06725: 5248\n06744: 5312\n06802: 8664\n06807: [24104] 109\n06814: 8666\n06815: 8808\n06817: 1201\n06819: 1300\n06822: 8853\n06825: 5851\n06828: 8603\n06857: 8700\n06862: 8854\n06864: 8835\n06869: 8703\n06901: 5687\n06929: 3278\n06945: 5323\n06962: 5751\n06983: 5817\n06984: 6718\n06987: [24103] 24\n06989: 6812\n06990: 6747\n07010: 5736\n07020: 3336\n07030: 6414\n07049: 8212\n07067: 5213\n07086: 8672\n07098: [24103] 23\n07116: 8639\n07145: 8248\n07147: 5523\n07203: 6434\n07207: 2332\n07215: 2339\n07224: 5692\n07255: 6715\n07259: 6736\n07352: 8624\n07412: 2271\n07468: 6417\n07477: 8247\n07478: 8804\n07543: 6826\n07567: 5664\n07572: 1301\n07577: 5665\n07599: 2375\n07620: 8222\n07630: 6766\n07633: 6426\n07701: 5245\n07703: 5259\n07707: 2361\n07712: 3228\n07727: 6754\n07737: 3289\n07743: 5258\n07785: 6821\n07797: 6748\n07811: 2261\n07828: 3264\n07831: 1202\n07843: 5681\n07846: 5853\n07847: 5271\n07877: 2536\n07894: 8809\n07901: 2371\n07924: 8246\n07969: 6521\n08063: 3272\n08075: 2537\n08093: 8679\n08094: 2374\n08119: 5268\n08165: 2349\n08182: 5554\n08192: 8221\n08196: 3314\n08204: 2518\n08212: 8625\n08227: 2272\n08229: 5266\n08243: [24103] 18\n08254: 8500\n08263: 8635\n08269: [24104] 105\n08274: 8819\n08281: 8802\n08285: 8701\n08292: 8873\n08304: 3465\n08326: [24104] 103\n08336: 8709\n08337: 6735\n08346: 2526\n08352: 3315\n08371: 5223\n08384: 2224\n08416: 2323\n08419: 6514\n08435: 8275\n08446: 2345\n08456: 2366\n08457: 2215\n08479: 5267\n08493: 5225\n08508: 8813\n08518: 6518\n08522: 3275\n08525: 8838\n08527: 3286\n08529: 2554\n08537: 6764\n08557: 8815\n08561: 2534\n08568: 6815\n08570: 3312\n08573: 3313\n08574: 5318\n08579: 2557\n08597: 5269\n08599: 5256\n08606: 6672\n08636: 3461\n08640: 3322\n08643: 6665\n08667: 3324\n08671: 5564\n08672: 3325\n08684: 3225\n08688: 8256\n08689: 3468\n08702: 5235\n08711: 3319\n08725: 2565\n08736: 6667\n08737: 3328\n08742: 3474\n08746: 5218\n08749: 6436\n08752: 3329\n08758: 2567\n08759: 5238\n08777: 5286\n08781: 3331\n08784: 5685\n08787: 8258\n08791: 2334\n08794: 6767\n08807: 2363\n08837: 6120\n08848: 2313\n08854: 2315\n08856: 5734\n08859: 2575\n08861: 2275\n08863: 6121\n08872: 3487\n08875: 6824\n08879: 2524\n08882: 2525\n08888: 2278\n08896: 3488\n08897: 6675\n08935: 8613\n08940: 5232\n08942: 2378\n08943: 6000\n08951: 6300\n08956: 6312\n08974: 8265\n08979: 8272\n08980: 6111\n08981: 8271\n08986: 3564\n08992: 6751\n08995: 6772\n08996: 6423\n08998: 8119\n09007: 6503\n09021: 6413\n09022: 2317\n09023: 5818\n09025: 6716\n09037: 5695\n09048: 5661\n09051: 5885\n09052: 6817\n09060: 5819\n09070: 2216\n09079: 6411\n09081: 5220\n09083: 6719\n09086: 5221\n09087: 6100\n09092: 6112\n09096: 6425\n09097: 6511\n09098: 6513\n09100: 5231\n09108: 5319\n09109: 6763\n09110: 6427\n09113: 8632\n09114: 8637\n09116: 8225\n09117: 5662\n09125: 8826\n09127: 3316\n09129: 8291\n09131: 8266\n09133: 8273\n09134: 3621\n09135: 8267\n09137: 6834\n09139: 5553\n09140: 8274\n09142: 6315\n09148: 5321\n09151: 5324\n09153: 6110\n09156: [24103] 19\n09157: 8621\n09160: 5556\n09161: 8616\n09162: 6314\n09166: 3622\n09174: 5666\n09180: 2338\n09181: 8829\n09182: 2337\n09183: 3625\n09185: 1204\n09186: 3473\n09187: 1\n09188: 8660\n09189: 0\n09190: 8239\n09191: 6726\n09194: 8224\n09199: 2523\n09200: 6651\n09205: 2563\n09206: 6676\n09209: 2312\n09210: 8673\n09213: 2576\n09214: 5738\n09215: 6717\n09217: 8668\n09223: 6655\n09225: 6525\n09227: 8681\n09230: 6418\n09236: 0\n09239: 2539\n09241: 6532\n09242: 2535\n09248: 2274\n09249: 5222\n09251: 5229\n09253: 3475\n09254: 3476\n09256: 1\n09257: 3323\n09258: 2568\n09263: 5243\n09265: 5315\n09279: 8667\n09281: 8627\n09284: 8236\n09289: 5559\n09291: 8674\n09292: 6424\n09296: 8676\n09301: 5753\n09302: 6419\n09303: 8665\n09314: 5215\n09316: 2268\n09317: 6659\n09320: 5557\n09321: 6653\n09322: 6661\n09326: 3624\n09328: 8705\n09329: 5669\n09330: [24103] 12\n09334: 6774\n09340: 8252\n09341: 5524\n09342: 2331\n09343: 6663\n09344: 8638\n09345: 8836\n09347: 2516\n09348: 2517\n09349: 6114\n09350: 3464\n09351: 2372\n09353: 6776\n09354: 5525\n09356: 6438\n09358: 2562\n09359: 5881\n09360: 5739\n09363: 6507\n09364: 6825\n09365: 2569\n09366: 6664\n09367: 6508\n09368: 2541\n09369: 5735\n09372: 5671\n09374: 8278\n09375: [24103] 16\n09376: 6652\n09378: 5667\n09379: 5224\n09380: 5672\n09381: 6656\n09382: 3623\n09383: 5682\n09384: 6115\n09386: 5882\n09387: 6421\n09388: 6827\n09390: 6828\n09391: 3466\n09392: 6829\n09393: 6657\n09394: 6673\n09395: 6658\n09396: 3628\n09398: 6831\n09399: 6666\n09400: 6668\n09401: 6516\n09402: 8617\n09403: 8259\n09404: 8675\n09405: 6713\n09406: 6761\n09407: 8612\n09408: 6439\n07030: 6412\n08332: 6654\n09324: 2531\n09335: 6416"
        extension_input = st.text_area(
            "請輸入員工編號和分機的對應關係（每行一個，格式為 '員工編號: 分機'）",
            default_extension_input,
        )

        # 解析用户输入的分机对应关系
        extension = parse_extension_input(extension_input)

        # 找出缺漏的员工编号
        missing_ids = [emp_id for emp_id in all_employee_ids if emp_id not in extension]

        # 如果有缺漏的员工编号，更新输入框
        if missing_ids:
            missing_input = "\n".join([f"{emp_id}: " for emp_id in missing_ids])
            updated_extension_input = missing_input + "\n" + extension_input
            st.warning(f"以下員工編號缺少分機對應關係：{', '.join(missing_ids)}")
            extension_input = st.text_area(
                "更新後的員工編號和分機對應關係（請為缺漏的編號添加分機）",
                updated_extension_input,
            )
            extension = parse_extension_input(extension_input)

        # 添加输入框让用户输入要分组的员工编号
        grouped_employees_input = st.text_area(
            "請輸入要一起分組的員工編號（每組一行，用逗號分隔）",
            "",
        )

        # 處理用戶輸入的分組信息
        grouped_employees = {}
        if grouped_employees_input.strip():  # 只有當輸入不為空時才處理
            for i, group in enumerate(grouped_employees_input.split('\n')):
                employees = [emp.strip() for emp in group.split(',') if emp.strip()]
                if employees:
                    grouped_employees[f'Group_{i+1}'] = employees

        # 創建包含每個員工數據的Excel文件
        output = create_employee_sheets(processed_df, billing_period, uploaded_file, grouped_employees, extension)

        if output:
            # 提供下载按钮，使用原始文件名
            st.download_button(
                label="下載修改後的Excel文件",
                data=output.getvalue(),
                file_name=original_filename.replace('.xlsx', '_更新.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
