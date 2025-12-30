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
        
        # 严格对应用户要求的索引：cn(0), list(1), 点数(2), 重量(3), 金额(4)
        rename_map = {
            cols[0]: 'cn',
            cols[1]: 'list',
            cols[2]: '点数',
            cols[3]: '重量',
            cols[4]: '金额'
        }

        df = df.rename(columns=rename_map)
        
        # 预处理：将重量和金额的空值填补为 0
        df['重量'] = pd.to_numeric(df['重量'], errors='coerce').fillna(0)
        df['金额'] = pd.to_numeric(df['金额'], errors='coerce').fillna(0)
        df['点数'] = pd.to_numeric(df['点数'], errors='coerce').fillna(0)

        # Sheet 内汇总：按人名合并
        sheet_summary = df.groupby('cn').agg({
            'list': lambda x: '，'.join([str(i) for i in x if pd.notna(i) and str(i).strip() != '']),
            '点数': 'sum',
            '重量': 'sum',
            '金额': 'sum'
        }).reset_index()

        # 重命名带上前缀
        sheet_summary = sheet_summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '重量': f'{sheet_name}_重量',
            '金额': f'{sheet_name}_原始金额'
        })

        # 记录该 Sheet 总重量
        sheet_total_weights[sheet_name] = sheet_summary[f'{sheet_name}_重量'].sum()

        if result_df is None:
            result_df = sheet_summary
        else:
            result_df = result_df.merge(sheet_summary, on='cn', how='outer')

    # 计算国际分摊及叠加金额
    for sheet_name, total_weight in sheet_total_weights.items():
        weight_col = f'{sheet_name}_重量'
        orig_amount_col = f'{sheet_name}_原始金额'
        intl_col = f'{sheet_name}_分摊国际'
        combined_col = f'{sheet_name}_单表总计' # 国际金额 + 原始金额

        if weight_col not in result_df.columns:
            continue

        # 1. 计算分摊的国际金额
        if total_weight > 0:
            result_df[intl_col] = (result_df[weight_col].fillna(0) / total_weight) * total_international_amount
        else:
            result_df[intl_col] = 0.0
            
        # 2. 叠加金额：国际分摊 + 原始金额
        result_df[combined_col] = result_df[intl_col] + result_df[orig_amount_col].fillna(0)

    # --- 全局大汇总 ---
    list_cols = [c for c in result_df.columns if c.endswith('_list')]
    point_cols = [c for c in result_df.columns if c.endswith('_点数')]
    weight_cols = [c for c in result_df.columns if c.endswith('_重量')]
    orig_amt_cols = [c for c in result_df.columns if c.endswith('_原始金额')]
    intl_amt_cols = [c for c in result_df.columns if c.endswith('_分摊国际')]
    combined_amt_cols = [c for c in result_df.columns if c.endswith('_单表总计')]

    result_df['总List汇总'] = result_df[list_cols].apply(
        lambda row: '；'.join([str(val) for val in row if pd.notna(val) and str(val).strip() != '']), axis=1
    )
    result_df['总点数'] = result_df[point_cols].sum(axis=1)
    result_df['总重量'] = result_df[weight_cols].sum(axis=1)
    result_df['总原始金额'] = result_df[orig_amt_cols].sum(axis=1)
    result_df['总国际金额'] = result_df[intl_amt_cols].sum(axis=1)
    result_df['总支出(含国际)'] = result_df[combined_amt_cols].sum(axis=1)

    # 排序：名字 -> 各Sheet详情 -> 总汇总
    new_cols = ['cn']
    sheet_names = sorted(sheet_total_weights.keys())
    for s in sheet_names:
        for sfx in ['_list', '_点数', '_重量', '_原始金额', '_分摊国际', '_单表总计']:
            c_name = f'{s}{sfx}'
            if c_name in result_df.columns:
                new_cols.append(c_name)
    
    new_cols.extend(['总List汇总', '总点数', '总重量', '总原始金额', '总国际金额', '总支出(含国际)'])
    result_df = result_df[new_cols]

    # 精度格式化
    res_cols = result_df.columns
    result_df[[c for c in res_cols if '重量' in c]] = result_df[[c for c in res_cols if '重量' in c]].round(2)
    result_df[[c for c in res_cols if '金额' in c or '总计' in c or '支出' in c]] = \
        result_df[[c for c in res_cols if '金额' in c or '总计' in c or '支出' in c]].round(3)

    return result_df

# ================= Streamlit UI =================
st.set_page_config(page_title="金额叠加分摊工具", layout="wide")
st.title("💰 国际金额分摊与叠加工具")
st.info("请确保 Excel 前五列顺序为：cn, list, 点数, 重量, 金额")

uploaded_file = st.file_uploader("上传 Excel", type=["xlsx"])
total_intl = st.number_input("输入需分摊的总国际金额", min_value=0.0, value=0.0, step=10.0)

if uploaded_file and total_intl >= 0:
    if st.button("🚀 开始计算并生成汇总"):
        try:
            res = merge_excel_with_international_amount(uploaded_file, total_intl)
            st.success("计算完成！")
            st.dataframe(res, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                res.to_excel(writer, index=False, sheet_name='汇总结果')
            
            st.download_button(
                label="⬇️ 下载汇总 Excel",
                data=output.getvalue(),
                file_name="分摊汇总表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"发生错误：{e}")
