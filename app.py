import streamlit as st
import pandas as pd
import io

def process_excel_global_weight(uploaded_file, total_intl_amount):
    # 1. 读取所有工作表
    all_sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)
    
    # --- 第一步：计算全文件的“全局总重量” ---
    global_total_weight = 0
    valid_sheets = {}
    
    for sheet_name, df in all_sheets_dict.items():
        if df.empty:
            continue
        # 强制定位第4列（索引3）为重量
        weight_col = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
        global_total_weight += weight_col.sum()
        valid_sheets[sheet_name] = df

    if global_total_weight == 0:
        st.error("警告：全文件总重量为 0，国际金额将无法分摊。")
        unit_weight_cost = 0
    else:
        # 计算每一单位重量应分摊的金额
        unit_weight_cost = total_intl_amount / global_total_weight
        st.info(f"📊 全局总重量: {global_total_weight:.2f} | 单位重量分摊成本: {unit_weight_cost:.4f}")

    # --- 第二步：处理每个 Sheet 并应用全局分摊公式 ---
    result_df = None
    all_summaries = []

    for sheet_name, df in valid_sheets.items():
        temp_df = pd.DataFrame()
        temp_df['cn'] = df.iloc[:, 0].astype(str)
        temp_df['list'] = df.iloc[:, 1].fillna('')
        temp_df['点数'] = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(0)
        temp_df['重量'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
        temp_df['原始金额'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)

        # 【核心公式】：叠加后的金额 = 原始金额 + (本行重量 * 全局单位成本)
        temp_df['叠加金额'] = temp_df['原始金额'] + (temp_df['重量'] * unit_weight_cost)

        summary = temp_df.groupby('cn').agg({
            'list': lambda x: '，'.join([str(i) for i in x if str(i).strip() != '']),
            '点数': 'sum',
            '重量': 'sum',
            '叠加金额': 'sum'
        }).reset_index()

        summary = summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '重量': f'{sheet_name}_重量',
            '叠加金额': f'{sheet_name}_最终金额'
        })
        all_summaries.append(summary)

    # --- 第三步：合并并进行大汇总 ---
    for s in all_summaries:
        if result_df is None:
            result_df = s
        else:
            result_df = result_df.merge(s, on='cn', how='outer')

    # 汇总全局列
    list_cols = [c for c in result_df.columns if '_list' in c]
    point_cols = [c for c in result_df.columns if '_点数' in c]
    weight_cols = [c for c in result_df.columns if '_重量' in c]
    final_amt_cols = [c for c in result_df.columns if '_最终金额' in c]

    result_df['【所有List汇总】'] = result_df[list_cols].apply(
        lambda row: '；'.join([str(v) for v in row if pd.notna(v) and str(v).strip() != '']), axis=1
    )
    result_df['【总点数】'] = result_df[point_cols].sum(axis=1)
    result_df['【总重量】'] = result_df[weight_cols].sum(axis=1)
    result_df['【全表最终支出(含分摊)】'] = result_df[final_amt_cols].sum(axis=1)

    # 格式化
    for c in result_df.columns:
        if '重量' in c:
            result_df[c] = result_df[c].round(2)
        elif '金额' in c or '最终' in c:
            result_df[c] = result_df[c].round(3)

    return result_df

# ================= Streamlit UI =================

st.set_page_config(page_title="国际均摊计算", layout="wide")

st.title("国际金额分摊工具")
st.markdown("""
**计算逻辑：**
1. 扫描所有 Sheet，算出整个文件的**总重量**。
2. 用“总国际金额 / 总重量”得出**单位重量成本**。
3. **叠加金额** = 原始金额 + (重量 × 单位成本)。
4. 如果重量为 0，则该项只保留原始金额。
""")

# 侧边栏配置
st.sidebar.header("配置参数")
uploaded_file = st.sidebar.file_uploader("上传 Excel 文件", type=["xlsx"])
total_intl = st.sidebar.number_input("输入总国际金额", min_value=0.0, value=950.0, step=10.0)

if uploaded_file:
    if st.sidebar.button("🚀 开始处理"):
        with st.spinner("正在计算中..."):
            try:
                # 执行逻辑
                final_df = process_excel_global_weight(uploaded_file, total_intl)
                
                st.success("✅ 处理完成！")
                
                # 数据预览
                st.dataframe(final_df, use_container_width=True)
                
                # 导出 Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='汇总结果')
                
                st.download_button(
                    label="⬇️ 下载汇总结果 Excel",
                    data=output.getvalue(),
                    file_name="国际表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"处理过程中出错: {e}")
else:
    st.info("💡 请先在左侧上传 Excel 文件并点击开始按钮。")
