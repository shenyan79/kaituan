import streamlit as st
import pandas as pd
import io

def merge_excel_with_international_amount(
    excel_file,
    total_international_amount
):
    # 读取所有 Sheet
    all_sheets = pd.read_excel(excel_file, sheet_name=None)

    result_df = None
    sheet_total_weights = {}

    for sheet_name, df in all_sheets.items():
        if df.empty:
            continue

        # 获取当前 sheet 的列名列表
        cols = list(df.columns)
        
        # --- 强制映射逻辑 ---
        # 0:cn, 1:list, 2:点数, 3:重量, 4:金额
        # 如果列数不足，填充空列以防报错
        new_df = pd.DataFrame()
        new_df['cn'] = df.iloc[:, 0]
        new_df['list'] = df.iloc[:, 1]
        new_df['点数'] = df.iloc[:, 2] if len(cols) > 2 else 0
        new_df['重量'] = df.iloc[:, 3] if len(cols) > 3 else 0
        new_df['金额'] = df.iloc[:, 4] if len(cols) > 4 else 0

        # 数据清洗：强制转换为数值，空值转为 0
        new_df['重量'] = pd.to_numeric(new_df['重量'], errors='coerce').fillna(0)
        new_df['金额'] = pd.to_numeric(new_df['金额'], errors='coerce').fillna(0)
        new_df['点数'] = pd.to_numeric(new_df['点数'], errors='coerce').fillna(0)

        # Sheet 内汇总（按人名 cn 分组）
        sheet_summary = new_df.groupby('cn').agg({
            'list': lambda x: '，'.join([str(i) for i in x if pd.notna(i) and str(i).strip() != '']),
            '点数': 'sum',
            '重量': 'sum',
            '金额': 'sum'
        }).reset_index()

        # 为当前 Sheet 的列加上前缀，防止合并时冲突
        sheet_summary = sheet_summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '重量': f'{sheet_name}_重量',
            '金额': f'{sheet_name}_原始金额'
        })

        # 统计该 Sheet 的总重量用于分摊
        total_w = sheet_summary[f'{sheet_name}_重量'].sum()
        sheet_total_weights[sheet_name] = total_w

        # 合并到主表
        if result_df is None:
            result_df = sheet_summary
        else:
            result_df = result_df.merge(sheet_summary, on='cn', how='outer')

    # --- 计算分摊逻辑 ---
    for sheet_name, total_weight in sheet_total_weights.items():
        w_col = f'{sheet_name}_重量'
        orig_col = f'{sheet_name}_原始金额'
        intl_col = f'{sheet_name}_国际分摊'
        subtotal_col = f'{sheet_name}_单表小计'

        if w_col in result_df.columns:
            # 计算分摊：(个人重量 / 总重量) * 总国际金额
            if total_weight > 0:
                result_df[intl_col] = (result_df[w_col].fillna(0) / total_weight) * total_international_amount
            else:
                result_df[intl_col] = 0.0
            
            # 单表小计 = 原始金额 + 国际分摊
            result_df[subtotal_col] = result_df[orig_col].fillna(0) + result_df[intl_col]

    # --- 全局大汇总列 ---
    list_cols = [c for c in result_df.columns if '_list' in c]
    point_cols = [c for c in result_df.columns if '_点数' in c]
    weight_cols = [c for c in result_df.columns if '_重量' in c]
    orig_cols = [c for c in result_df.columns if '_原始金额' in c]
    intl_cols = [c for c in result_df.columns if '_国际分摊' in c]
    subtotal_cols = [c for c in result_df.columns if '_单表小计' in c]

    result_df['【总List汇总】'] = result_df[list_cols].apply(
        lambda row: '；'.join([str(v) for v in row if pd.notna(v) and str(v).strip() != '']), axis=1
    )
    result_df['【总点数】'] = result_df[point_cols].sum(axis=1)
    result_df['【总重量】'] = result_df[weight_cols].sum(axis=1)
    result_df['【总原始金额】'] = result_df[orig_cols].sum(axis=1)
    result_df['【总国际金额】'] = result_df[intl_cols].sum(axis=1)
    result_df['【最终总支出】'] = result_df[subtotal_cols].sum(axis=1)

    # --- 整理列顺序 ---
    final_cols = ['cn']
    for s in sorted(sheet_total_weights.keys()):
        for suffix in ['_list', '_点数', '_重量', '_原始金额', '_国际分摊', '_单表小计']:
            c_name = f{s}{suffix}
            if c_name in result_df.columns:
                final_cols.append(c_name)
    
    final_cols.extend(['【总List汇总】', '【总点数】', '【总重量】', '【总原始金额】', '【总国际金额】', '【最终总支出】'])
    result_df = result_df[final_cols]

    # 小数点处理
    for c in result_df.columns:
        if '重量' in c:
            result_df[c] = result_df[c].round(2)
        elif '金额' in c or '分摊' in c or '支出' in c or '小计' in c:
            result_df[c] = result_df[c].round(3)

    return result_df

# ================= Streamlit UI =================
st.set_page_config(page_title="国际金额汇总", layout="wide")
st.title("📊 国际金额分摊工具 (修正版)")

st.markdown("""
**注意：** 请确保您的 Excel 每个工作表前 5 列分别是：
1. **cn** (人名) | 2. **list** (内容) | 3. **点数** | 4. **重量** | 5. **金额**
""")

uploaded_file = st.file_uploader("上传 Excel", type=["xlsx"])
total_intl = st.number_input("输入总国际金额", min_value=0.0, step=10.0)

if uploaded_file and total_intl >= 0:
    if st.button("🚀 开始生成汇总"):
        try:
            res = merge_excel_with_international_amount(uploaded_file, total_intl)
            st.success("计算完成！")
            st.dataframe(res, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                res.to_excel(writer, index=False, sheet_name='汇总结果')
            
            st.download_button(
                label="⬇️ 下载 Excel 汇总表",
                data=output.getvalue(),
                file_name="最终分摊汇总表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"处理失败，错误原因：{e}")
