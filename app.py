import streamlit as st
import pandas as pd
import os
from io import BytesIO


def transform_excel(df: pd.DataFrame, original_filename: str):
    # 输出文件名
    file_name_part = os.path.splitext(original_filename)[0]
    output_filename = f"改_{file_name_part}.xlsx"

    # --- 1. 横向扫描分类 (第2行, 索引1) ---
    col_to_category = {}
    last_category = "默认分类"

    for col_idx in range(1, df.shape[1]):
        cat_val = df.iloc[1, col_idx]
        if pd.notna(cat_val) and str(cat_val).strip() not in ["", "分类"]:
            last_category = str(cat_val).strip()
        col_to_category[col_idx] = last_category

    # --- 2. 提取制品名称 (第3行, 索引2) ---
    product_names = {}
    for col_idx in range(1, df.shape[1]):
        name_val = df.iloc[2, col_idx]
        if pd.isna(name_val) or str(name_val).strip() == "":
            break
        product_names[col_idx] = str(name_val).strip()

    # --- 3. 遍历数据行 (从第6行[索引5]开始) ---
    results = []

    for i in range(5, len(df)):
        val_a = df.iloc[i, 0]  # A列：总金额
        val_b = df.iloc[i, 1]  # B列：昵称

        if pd.notna(val_a) and pd.notna(val_b):
            person_name = str(val_b).strip()
            total_money = str(val_a).strip()

            purchased_details = []
            row_total_points = 0

            for col_idx in product_names.keys():
                count = df.iloc[i, col_idx]
                if pd.notna(count) and isinstance(count, (int, float)) and count > 0:
                    category = col_to_category.get(col_idx, "默认分类")
                    item_name = product_names[col_idx]

                    row_total_points += int(count)

                    if category == "默认分类":
                        detail_str = f"{item_name}✖{int(count)}"
                    else:
                        detail_str = f"({category}){item_name}✖{int(count)}"

                    purchased_details.append(detail_str)

            if purchased_details:
                results.append({
                    "名字": person_name,
                    "（分类名称）/种类✖个数": " / ".join(purchased_details),
                    "总点数": row_total_points,
                    "对应的总金额": total_money
                })

    if not results:
        return None, None

    final_df = pd.DataFrame(results)
    final_df = final_df[["名字", "（分类名称）/种类✖个数", "总点数", "对应的总金额"]]

    return final_df, output_filename


# ================= Streamlit UI =================

st.set_page_config(page_title="Excel 汇总转换工具", layout="centered")

st.title("📊 Excel 汇总表 → 清单表转换工具")
st.write("上传 Excel 文件，自动生成整理后的清单表（保持原有逻辑）")

uploaded_file = st.file_uploader(
    "📤 上传 Excel 文件",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, header=None)
        st.success("✅ 文件读取成功")

        if st.button("🚀 开始处理"):
            with st.spinner("处理中，请稍候..."):
                result_df, out_name = transform_excel(df, uploaded_file.name)

            if result_df is None:
                st.error("❌ 未提取到有效数据，请检查 A 列和 B 列内容")
            else:
                st.success("🎉 处理完成！")
                st.dataframe(result_df)

                # 转成 Excel 供下载
                buffer = BytesIO()
                result_df.to_excel(buffer, index=False)
                buffer.seek(0)

                st.download_button(
                    label="⬇️ 下载处理后的 Excel",
                    data=buffer,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
