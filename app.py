import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile
import os


st.set_page_config(page_title="Excel 合并工具", layout="wide")
st.title("📊 Excel 多 Sheet 合并工具")

st.markdown("""
**功能说明：**
- 上传一个 Excel（每个 sheet：名字 / list / 点数 / 金额）
- 自动生成：
  - 双行表头
  - sheet 分组
  - 分隔空列
  - 汇总金额加粗
""")


uploaded_file = st.file_uploader(
    "请上传 Excel 文件（.xlsx）",
    type=["xlsx"]
)

if uploaded_file is not None:
    if st.button("🚀 开始处理"):
        with st.spinner("正在处理，请稍候..."):
            # 保存上传文件到临时目录
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.read())
                input_path = tmp.name

            output_path = input_path.replace(".xlsx", "_output.xlsx")

            # ================== 原来的核心逻辑 ==================
            all_sheets = pd.read_excel(input_path, sheet_name=None)

            result_df = None
            sheet_order = []

            for sheet_name, df in all_sheets.items():
                if df.empty:
                    continue

                sheet_order.append(sheet_name)

                df = df.rename(columns={
                    df.columns[0]: '名字',
                    df.columns[1]: 'list',
                    df.columns[2]: '点数',
                    df.columns[3]: '金额'
                })

                sheet_summary = df.groupby('名字').agg({
                    'list': lambda x: '，'.join(x.astype(str)),
                    '点数': 'sum',
                    '金额': 'sum'
                }).reset_index()

                sheet_summary = sheet_summary.rename(columns={
                    'list': f'{sheet_name}_list',
                    '点数': f'{sheet_name}_点数',
                    '金额': f'{sheet_name}_金额'
                })

                if result_df is None:
                    result_df = sheet_summary
                else:
                    result_df = result_df.merge(
                        sheet_summary,
                        on='名字',
                        how='outer'
                    )

            amount_cols = [c for c in result_df.columns if c.endswith('_金额')]
            result_df['汇总金额'] = result_df[amount_cols].sum(axis=1, skipna=True)

            result_df.to_excel(output_path, index=False, header=False)

            wb = load_workbook(output_path)
            ws = wb.active
            ws.insert_rows(1, amount=2)

            ws['A1'] = 'cn'
            ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

            col = 2
            col += 1

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

                col += 4

            ws.cell(row=1, column=col).value = '汇总'
            ws.cell(row=2, column=col).value = '金额'

            bold_font = Font(bold=True)
            for row in range(3, ws.max_row + 1):
                ws.cell(row=row, column=col).font = bold_font

            wb.save(output_path)
            # ================== 处理结束 ==================

            with open(output_path, "rb") as f:
                st.success("处理完成！")
                st.download_button(
                    label="⬇️ 下载结果文件",
                    data=f,
                    file_name="合并结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
