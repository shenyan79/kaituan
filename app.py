import streamlit as st
import pandas as pd
import io

def process_excel_summary_only(uploaded_file):
    # 1. 读取所有工作表
    all_sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)
    
    valid_sheets = {}
    for sheet_name, df in all_sheets_dict.items():
        if not df.empty:
            valid_sheets[sheet_name] = df

    # --- 处理每个 Sheet ---
    result_df = None
    sheet_names_list = []

    for sheet_name, df in valid_sheets.items():
        # 严格按照图示表头顺序映射：A:cn, B:list, C:点数, D:重量, E:金额
        temp_df = pd.DataFrame()
        temp_df['cn'] = df.iloc[:, 0].astype(str)
        temp_df['list'] = df.iloc[:, 1].fillna('')
        temp_df['点数'] = pd.to_numeric(df.iloc[:, 2], errors='coerce').fillna(0)
        temp_df['重量'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
        temp_df['金额'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)

        # 按人名汇总当前 Sheet 内的数据
        summary = temp_df.groupby('cn').agg({
            'list': lambda x: '，'.join([str(i) for i in x if str(i).strip() != '']),
            '点数': 'sum',
            '重量': 'sum',
            '金额': 'sum'
        }).reset_index()

        # 重命名列名以区分 Sheet
        summary = summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '重量': f'{sheet_name}_重量',
            '金额': f'{sheet_name}_金额'
        })
        
        sheet_names_list.append(sheet_name)

        # 合并到主结果表
        if result_df is None:
            result_df = summary
        else:
            result_df = result_df.merge(summary, on='cn', how='outer')

    # --- 全局大汇总计算 ---
    # 获取所有相关的列名
    list_cols = [c for c in result_df.columns if '_list' in c]
    point_cols = [c for c in result_df.columns if '_点数' in c]
    weight_cols = [c for c in result_df.columns if '_重量' in c]
    amount_cols = [c for c in result_df.columns if '_金额' in c]

    # 1. 汇总所有 List 内容
    result_df['【总List汇总】'] = result_df[list_cols].apply(
        lambda row: '；'.join([str(v) for v in row if pd.notna(v) and str(v).strip() != '']), axis=1
    )
    # 2. 数值累加
    result_df['【总点数】'] = result_df[point_cols].sum(axis=1)
    result_df['【总重量】'] = result_df[weight_cols].sum(axis=1)
    result_df['【最终总金额】'] = result_df[amount_cols].sum(axis=1)

    # 数值格式化
    for c in result_df.columns:
        if '重量' in c:
            result_df[c] = result_df[c].round(2)
        elif '金额' in c or '最终' in c:
            result_df[c] = result_df[c].round(3)

    return result_df

# ================= Streamlit UI =================

st.set_page_config(page_title="多Sheet金额重量汇总工具", layout="wide")

st.title("📊 多工作表数据自动汇总工具")
st.markdown(f"**表头要求：** A: `cn` | B: `list` | C: `点数` | D: `重量` | E: `金额`")

uploaded_file = st.file_uploader("直接上传 Excel 文件进行汇总", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 生成汇总报告"):
        with st.spinner("正在提取并计算各工作表数据..."):
            try:
                final_df = process_excel_summary_only(uploaded_file)
                
                st.success("✅ 全表汇总完成")
                
                # 结果预览
                st.dataframe(final_df, use_container_width=True)
                
                # 提供下载
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='汇总结果')
                
                st.download_button(
                    label="⬇️ 下载汇总 Excel",
                    data=output.getvalue(),
                    file_name="多Sheet汇总结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"处理失败，请检查文件格式。错误详情: {e}")
else:
    st.info("💡 请上传需要汇总的 Excel 文件。系统将自动按人名合并所有 Sheet 的重量与金额。")
