import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side


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

    # 從選擇的選項中提取員工編號
    selected_employee_id = selected_employee.split(' - ')[0]

    # 顯示選中員工的數據
    employee_data = df[df[employee_column] == selected_employee_id]
    st.subheader(f"員工 {selected_employee} 的數據")
    st.dataframe(employee_data)


def create_employee_sheets(df, billing_period, original_file):
    employee_column = '員工編號'
    name_column = '員工姓名'

    if employee_column not in df.columns or name_column not in df.columns:
        st.error(f"找不到 '{employee_column}' 或 '{name_column}' 列。請確保數據中包含這些列。")
        return None

    # 加載原始工作簿
    workbook = load_workbook(original_file)

    for employee, group in df.groupby(employee_column):
        employee_name = group[name_column].iloc[0]
        sheet_name = f'{employee}_{employee_name}'

        # 如果工作表已存在，則刪除它
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

        # 創建新的工作表
        worksheet = workbook.create_sheet(sheet_name)

        # 創建固定的行內容
        fixed_rows = [
            ['企業會員乘車服務電子對帳單'],
            ['客戶名稱：', '', '友訊科技股份有限公司'],
            ['列帳期間：', '', billing_period],
        ]

        # 將固定行內容寫入工作表
        for row in fixed_rows:
            worksheet.append(row)

        # 重置 group 的索引並保留原始列名
        group_reset = group.reset_index(drop=True)

        # 將員工數據（包括標題行）寫入工作表
        worksheet.append(group_reset.columns.tolist())
        for _, row in group_reset.iterrows():
            worksheet.append(row.tolist())

        # 計算統計數據
        total_count = len(group_reset)
        total_amount = group_reset['折扣後車資'].sum() if '折扣後車資' in group_reset.columns else 0

        # 創建統計數據行
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

        # 寫入統計數據行
        for row in stats_rows:
            worksheet.append(row)

        # 設置字體大小和調整列寬
        for row in worksheet.iter_rows():
            for cell in row:
                cell.font = Font(size=12)

        for idx, column in enumerate(group_reset.columns):
            column_letter = get_column_letter(idx + 1)
            if column in ['上車地點', '下車地點']:
                worksheet.column_dimensions[column_letter].width = 12
            else:
                max_length = max(
                    group_reset[column].astype(str).map(len).max() + 4, len(str(column)) + 6
                )
                worksheet.column_dimensions[column_letter].width = max_length

        # 合併第一行單元格並置中
        max_col = len(group_reset.columns)
        worksheet.merge_cells(f'A1:{get_column_letter(max_col)}1')
        title_cell = worksheet['A1']
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        worksheet.merge_cells(f'C2:{get_column_letter(max_col)}2')
        worksheet.merge_cells(f'C3:{get_column_letter(max_col)}3')

        # 設置第一行為粗體
        bold_font = Font(size=12, bold=True)
        title_cell.font = bold_font
        start_row = len(fixed_rows) + len(group_reset) + 2
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

        # 添加外框線
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
        )

        # 為數據部分添加全部框線
        for row in worksheet[f'A1':f'{get_column_letter(max_col)}{worksheet.max_row}']:
            for cell in row:
                cell.border = thin_border

    # 將修改後的工作簿保存到內存中
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output


def main():
    st.title("Excel數據整理工具")

    # 上傳Excel文件
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

        # 創建包含每個員工數據的Excel文件
        output = create_employee_sheets(processed_df, billing_period, uploaded_file)

        if output:
            # 提供下載按鈕，使用原始文件名
            st.download_button(
                label="下載修改後的Excel文件",
                data=output.getvalue(),
                file_name=original_filename.replace('.xlsx', '_更新.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
