from datetime import datetime
import io
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
    employee_options = [
        f"{row[employee_column]} - {row[name_column]}" for _, row in employees.iterrows()
    ]

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
    current_date = datetime.now()
    current_year_month = current_date.strftime("%Y/%m")
    today = current_date.strftime("%Y/%m/%d")
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
        total_amount = (
            sorted_group_df['折扣後車資'].sum() if '折扣後車資' in sorted_group_df.columns else 0
        )

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
                max_length = max(
                    sorted_group_df[column].astype(str).map(len).max() + 4, len(str(column)) + 6
                )
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
        for row in worksheet[f'A1':f'{get_column_letter(max_col)}{worksheet.max_row}']:
            for cell in row:
                cell.border = thin_border

    # 處理未分組的員工
    ungrouped_employees = set(df[employee_column]) - set(
        employee for group in grouped_employees.values() for employee in group
    )
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
        total_amount = (
            employee_df_reset['折扣後車資'].sum()
            if '折扣後車資' in employee_df_reset.columns
            else 0
        )

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
                max_length = max(
                    employee_df_reset[column].astype(str).map(len).max() + 4, len(str(column)) + 6
                )
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
        for row in worksheet[f'A1':f'{get_column_letter(max_col)}{worksheet.max_row}']:
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
    for row in summary_sheet.iter_rows(
        min_row=1, max_row=len(summary_data) + 5, min_col=1, max_col=7
    ):
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
    st.title("Excel數據整理工具")

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
        default_extension_input = "08956: 6312\n07030: 6412\n05259: 2340\n06294: 8254\n08332: 6654\n09025: 6716\n09092: 6112\n09137: 6834\n09214: 5738\n09324: 2531\n07468: 6417\n08951: 6300\n09021: 6413\n09335: 6416"
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
        output = create_employee_sheets(
            processed_df, billing_period, uploaded_file, grouped_employees, extension
        )

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
