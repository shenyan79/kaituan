import streamlit as st
import pandas as pd
import tempfile

st.set_page_config(page_title="重量 & 金额分摊工具", layout="wide")

# =====================================================
# Step 1：原始表 → 重量表
# =====================================================
def step1_generate_weight_excel(input_file):
    all_sheets = pd.read_excel(input_file, sheet_name=None, header=None)
    output_sheets = {}

    for sheet_name, df in all_sheets.items():
        if df.shape[0] < 6 or df.shape[1] < 3:
            continue

        # 制品重量（第 2 行，C 列起）
        weights = pd.to_numeric(df.iloc[1, 2:], errors="coerce")

        # 名字（第 6 行起，B 列）
        names = df.iloc[5:, 1]

        # 数量矩阵
        qty = df.iloc[5:, 2:].fillna(0)
        qty = qty.apply(pd.to_numeric, errors="coerce").fillna(0)

        total_weight = qty.dot(weights)

        result_df = pd.DataFrame({
            "名字": names.values,
            "总重量(g)": total_weight.round(2)  # ⭐ 保留 2 位
        }).dropna(subset=["名字"])

        output_sheets[sheet_name] = result_df

    return output_sheets


# =====================================================
# Step 2：重量表 → 金额分摊表
# =====================================================
def step2_weight_to_amount(weight_excel, total_amount):
    all_sheets = pd.read_excel(weight_excel, sheet_name=None)
    final_df = None

    for sheet_name, df in all_sheets.items():
        if df.empty or "总重量(g)" not in df.columns:
            continue

        sheet_total_weight = df["总重量(g)"].sum()
        if sheet_total_weight == 0:
            continue

        temp = df.copy()
        temp[f"{sheet_name}_重量"] = temp["总重量(g)"].round(2)
        temp[f"{sheet_name}_金额"] = (
            temp["总重量(g)"] / sheet_total_weight * total_amount
        ).round(3)  # ⭐ 金额 3 位小数

        temp = temp[["名字", f"{sheet_name}_重量", f"{sheet_name}_金额"]]

        if final_df is None:
            final_df = temp
        else:
            final_df = final_df.merge(temp, on="名字", how="outer")

    # 汇总金额
    amount_cols = [c for c in final_df.columns if c.endswith("_金额")]
    final_df["汇总金额"] = final_df[amount_cols].sum(axis=1, skipna=True).round(3)

    return final_df


# =====================================================
# 🌈 Streamlit 前端
# =====================================================
st.title("📊 重量 & 金额分摊工具")

tab1, tab2 = st.tabs(["Step 1：生成重量表", "Step 2：重量 → 金额分摊"])


# ==========================
# Step 1 UI
# ==========================
with tab1:
    st.subheader("Step 1：原始 Excel → 重量表")

    uploaded_step1 = st.file_uploader(
        "上传原始 Excel（含制品重量和数量）",
        type=["xlsx"],
        key="step1"
    )

    if uploaded_step1:
        weight_sheets = step1_generate_weight_excel(uploaded_step1)

        if weight_sheets:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    for sheet, df in weight_sheets.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)

                st.success("✅ 重量表生成成功")
                st.download_button(
                    "📥 下载：重量表.xlsx",
                    open(tmp.name, "rb"),
                    file_name="重量表.xlsx"
                )
        else:
            st.warning("未识别到有效的 Sheet")


# ==========================
# Step 2 UI
# ==========================
with tab2:
    st.subheader("Step 2：重量表 → 金额分摊表（国际表）")

    uploaded_step2 = st.file_uploader(
        "上传 Step 1 生成的【重量表.xlsx】",
        type=["xlsx"],
        key="step2"
    )

    total_amount = st.number_input(
        "输入总金额",
        min_value=0.0,
        step=100.0
    )

    if uploaded_step2 and total_amount > 0:
        final_df = step2_weight_to_amount(uploaded_step2, total_amount)

        if final_df is not None:
            st.dataframe(final_df)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                final_df.to_excel(tmp.name, index=False)
                st.download_button(
                    "📥 下载：国际表_重量分摊.xlsx",
                    open(tmp.name, "rb"),
                    file_name="国际表_重量分摊.xlsx"
                )
        else:
            st.warning("未生成有效数据")
