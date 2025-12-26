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

        rename_map = {
            cols[0]: '名字',
            cols[1]: 'list',
            cols[2]: '点数',
            cols[3]: '金额'
        }

        if len(cols) >= 5:
            rename_map[cols[4]] = '重量'
        else:
            df['重量'] = pd.NA

        df = df.rename(columns=rename_map)

        # 分类前缀处理
        if '分类' in df.columns:
            df['list'] = df.apply(
                lambda r: f"（{r['分类']}）{r['list']}"
                if pd.notna(r['分类']) and str(r['分类']).strip() != ''
                else r['list'],
                axis=1
            )

        # sheet 内汇总
        sheet_summary = df.groupby('名字').agg({
            'list': lambda x: '，'.join(x.astype(str)),
            '点数': 'sum',
            '金额': 'sum',
            '重量': 'sum'
        }).reset_index()

        sheet_summary = sheet_summary.rename(columns={
            'list': f'{sheet_name}_list',
            '点数': f'{sheet_name}_点数',
            '金额': f'{sheet_name}_金额',
            '重量': f'{sheet_name}_总重量'
        })

        sheet_total_weights[sheet_name] = (
            sheet_summary[f'{sheet_name}_总重量'].sum(skipna=True)
        )

        if result_df is None:
            result_df = sheet_summary
        else:
            result_df = result_df.merge(
                sheet_summary,
                on='名字',
                how='outer'
            )

    # 国际金额计算
    for sheet_name, total_weight in sheet_total_weights.items():
        weight_col = f'{sheet_name}_总重量'
        intl_col = f'{sheet_name}_国际金额'

        if weight_col not in result_df.columns or total_weight == 0:
            result_df[intl_col] = pd.NA
        else:
            result_df[intl_col] = (
                result_df[weight_col] / total_weight * total_international_amount
            )

    # 汇总金额
    amount_cols = [
        c for c in result_df.columns
        if c.endswith('_金额') and not c.endswith('_国际金额')
    ]
    result_df['汇总金额'] = result_df[amount_cols].sum(axis=1, skipna=True)

    # 总国际金额
    intl_cols = [c for c in result_df.columns if c.endswith('_国际金额')]
    result_df['总国际金额'] = result_df[intl_cols].sum(axis=1, skipna=True)

    # 列顺序（国际金额紧跟总重量）
    new_cols = ['名字']
    sheet_names = sorted(sheet_total_weights.keys())

    for s in sheet_names:
        for suffix in ['_list', '_点数', '_金额', '_总重量']:
            col = f'{s}{suffix}'
            if col in result_df.columns:
                new_cols.append(col)

        intl_col = f'{s}_国际金额'
        if intl_col in result_df.columns:
            new_cols.append(intl_col)

    new_cols.extend(['汇总金额', '总国际金额'])
    result_df = result_df[new_cols]

    # 小数控制
    weight_cols = [c for c in result_df.columns if c.endswith('_总重量')]
    result_df[weight_cols] = result_df[weight_cols].round(2)

    money_cols = [
        c for c in result_df.columns
        if c.endswith('_金额') or c.endswith('_国际金额')
    ]
    result_df[money_cols] = result_df[money_cols].round(3)

    return result_df


# ================= Streamlit UI =================

st.set_page_config(page_title="国际金额分摊工具", layout="wide")

st.title("📊 国际金额按重量分摊（多 Sheet）")

uploaded_file = st.file_uploader(
    "上传 gj.xlsx",
    type=["xlsx"]
)

total_international_amount = st.number_input(
    "输入总国际金额",
    min_value=0.0,
    step=100.0
)

if uploaded_file and total_international_amount > 0:
    if st.button("🚀 生成汇总表"):
        with st.spinner("正在计算，请稍候..."):
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
            label="⬇ 下载 Excel",
            data=buffer.getvalue(),
            file_name="国际汇总表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("📌 请上传文件并输入总国际金额")
