import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font


# ================= 核心处理函数 =================
def merge_excel_sheets(uploaded_file):
    # ---------- 1. 读取所有 sheet ----------
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

    result_df = None
    sheet_order = []

    for sheet_name, df in all_sheets.items():
        if df.empty:
            continue

        sheet_order.append(sheet_name)

        # 统一列名（源文件第一行是表头）
        df = df.rename(columns={
            df.columns[0]: '名字',
            df.columns[1]: 'list',
            df.columns[2]: '点数',
            df.columns[3]: '金额'
        })

        # sheet 内汇总
        sheet_summary = df.groupby('名字').agg({
            'list': lambda x: '，'.join(x.astype(str)),
            '点数': 'sum',
            '金额': 'sum'
        }).reset_index()

        # sheet 专属列名
        sheet_summary = sheet_summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '金额': f'{sheet_name}_金额'
        })

        # 横向合并
        if result_df is None:
            result_df = sheet_summary
        else:
            result_df = result_df.merge(
                sheet_summary,
                on='名字',
                how='outer'
            )

    # ---------- 2. 汇总金额 ----------
    amount_cols = [c for c in result_df.columns if c.endswith('_金额')]
    result_df['汇总金额'] = result_df[amount_cols].sum(axis=1, skipna=True)

    # ---------- 3. 写入内存 Excel（无表头） ----------
    output = BytesIO()
    result_df.to_excel(output, index=False, header=False)
    output.seek(0)

    # ---------- 4. openpyxl 处理表头 ----------
    wb = load_workbook(output)
    ws = wb.active

    ws.insert_rows(1, amount=2)

    # cn 列
    ws['A1'] = 'cn'
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    col = 2
    for sheet_name in sheet_order:
        ws.merge_cells(
            start_row=1,
            start_column=col,
            end_row=1,
            end_column=col + 2
        )
        ws.cell(row=1, column=col).value = sheet_name

        ws.cell(row=2, column=col).value = 'list'
        ws.cell(row=2, column=col + 1).value = '点数'
        ws.cell(row=2, column=col + 2).value = '金额'

        col += 3

    # 汇总金额
    ws.cell(row=1, column=col).value = '汇总'
    ws.cell(row=2, column=col).value = '金额'

    # ---------- 5. 汇总金额加粗 ----------
    bold_font = Font(bold=True)
    for row in range(3, ws.max_row + 1):
        ws.cell(row=row, column=col).font = bold_font

    # 保存到 BytesIO
    final_output = BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return final_output


# ================= Streamlit UI =================
st.set_page_config(page_title="Excel 多 Sheet 汇总工具", layout="wide")

st.title("📊 Excel 多 Sheet 汇总工具")
st.markdown(
    """
**功能说明：**
- 源文件：每个 Sheet 为 4 列（名字 / list / 点数 / 金额）
- 自动按名字汇总
- 输出为双行表头（cn + 各 Sheet + 汇总金额）
"""
)

uploaded_file = st.file_uploader(
    "📂 请上传 Excel 文件（.xlsx）",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        result_file = merge_excel_sheets(uploaded_file)

        st.success("✅ 处理完成，可以下载结果文件")

        st.download_button(
            label="⬇️ 下载合并后的 Excel",
            data=result_file,
            file_name="合并结果_双行表头.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ 处理失败，请检查文件格式")
        st.exception(e)
