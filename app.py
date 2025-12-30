import streamlit as st
import pandas as pd
import io

def merge_excel_with_international_amount(
    excel_file,
    total_international_amount
):
    all_sheets = pd.read_excel(excel_file, sheet_name=None)

    result_df = None
    sheet_total_weights = {}

    for sheet_name, df in all_sheets.items():
        if df.empty:
            continue

        cols = list(df.columns)
        
        # 按照新要求修改表头映射：cn, list, 点数, 重量, 金额
        # 假设顺序是固定的
        rename_map = {
            cols[0]: 'cn',
            cols[1]: 'list',
            cols[2]: '点数',
            cols[3]: '重量',
            cols[4]: '金额'
        }

        df = df.rename(columns=rename_map)

        # Sheet 内汇总
        sheet_summary = df.groupby('cn').agg({
            'list': lambda x: '，'.join(x.astype(str)),
            '点数': 'sum',
            '重量': 'sum',
            '金额': 'sum'
        }).reset_index()

        # 重命名各列，带上 Sheet 前缀
        sheet_summary = sheet_summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '重量': f'{sheet_name}_重量',
            '金额': f'{sheet_name}_金额'
        })

        # 记录该 Sheet 的总重量用于分摊计算
        sheet_total_weights[sheet_name] = sheet_summary[f'{sheet_name}_重量'].sum(skipna=True)

        if result_df is None:
            result_df = sheet_summary
        else:
            result_df = result_df.merge(
                sheet_summary,
                on='cn',
                how='outer'
            )

    # 计算国际金额分摊
    for sheet_name, total_weight in sheet_total_weights.items():
        weight_col = f'{sheet_name}_重量'
        intl_col = f'{sheet_name}_国际金额'

        if weight_col not in result_df.columns or total_weight == 0:
            result_df[intl_col] = 0.0
        else:
            # 按重量占比分摊
            result_df[intl_col] = (
                result_df[weight_col].fillna(0) / total_weight * total_international_amount
            )

    # --- 新增全局汇总逻辑 ---
    
    # 查找所有相关列
    list_cols = [c for c in result_df.columns if c.endswith('_list')]
    point_cols = [c for c in result_df.columns if c.endswith('_点数')]
    weight_cols = [c for c in result_df.columns if c.endswith('_重量')]
    amount_cols = [c for c in result_df.columns if c.endswith('_金额') and '国际' not in c]
    intl_cols = [c for c in result_df.columns if c.endswith('_国际金额')]

    # 1. 汇总所有 List (合并字符串)
    result_df['总List'] = result_df[list_cols].apply(
        lambda row: '，'.join([str(val) for val in row if pd.notna(val) and str(val).strip() != '']), 
        axis=1
    )
    
    # 2. 汇总各项数值
    result_df['总点数'] = result_df[point_cols].sum(axis=1, skipna=True)
    result_df['总重量'] = result_df[weight_cols].sum(axis=1, skipna=True)
    result_df['汇总金额'] = result_df[amount_cols].sum(axis=1, skipna=True)
    result_df['总国际金额'] = result_df[intl_cols].sum(axis=1, skipna=True)

    # 整理列顺序
    new_cols = ['cn']
    sheet_names = sorted(sheet_total_weights.keys())

    for s in sheet_names:
        for suffix in ['_list', '_点数', '_重量', '_金额', '_国际金额']:
            col = f'{s}{suffix}'
            if col in result_df.columns:
                new_cols.append(col)

    # 将汇总列放在最后
    new_cols.extend(['总List', '总点数', '总重量', '汇总金额', '总国际金额'])
    result_df = result_df[new_cols]

    # 小数控制
    final_weight_cols = [c for c in result_df.columns if '重量' in c]
    result_df[final_weight_cols] = result_df[final_weight_cols].round(2)

    final_money_cols = [c for c in result_df.columns if '金额' in c]
    result_df[final_money_cols] = result_df[final_money_cols].round(3)

    return result_df


# ================= Streamlit UI =================

st.set_page_config(page_title="国际金额分摊工具", layout="wide")

st.title("📊 国际金额按重量分摊（新表头版）")
st.markdown("输入要求列序：`cn, list, 点数, 重量, 金额`")

uploaded_file = st.file_uploader(
    "上传 Excel 文件",
    type=["xlsx"]
)

total_international_amount = st.number_input(
    "输入总国际金额",
    min_value=0.0,
    step=100.0,
    value=0.0
)

if uploaded_file and total_international_amount > 0:
    if st.button("🚀 生成汇总表"):
        with st.spinner("正在计算，请稍候..."):
            try:
                result_df = merge_excel_with_international_amount(
                    uploaded_file,
                    total_international_amount
                )

                st.success("✅ 生成完成")
                st.dataframe(result_df, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    result_df.to_excel(writer, index=False, sheet_name="汇总")

                st.download_button(
                    label="⬇ 下载 Excel 汇总表",
                    data=buffer.getvalue(),
                    file_name="国际汇总结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"处理出错，请检查文件格式是否正确。错误信息: {e}")
else:
    st.info("📌 请上传文件并输入大于 0 的总国际金额")
