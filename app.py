import pandas as pd
import os

# ==================================================
# 1️⃣ 汇总表 → 清单表
# ==================================================
def transform_summary_to_list(input_file, output_file):
    df = pd.read_excel(input_file, header=None)

    # --- 1. 横向扫描分类（第2行） ---
    col_to_category = {}
    last_category = "默认分类"
    for col_idx in range(1, df.shape[1]):
        cat_val = df.iloc[1, col_idx]
        if pd.notna(cat_val) and str(cat_val).strip() not in ["", "分类"]:
            last_category = str(cat_val).strip()
        col_to_category[col_idx] = last_category

    # --- 2. 提取制品名称（第3行） ---
    product_names = {}
    for col_idx in range(1, df.shape[1]):
        name_val = df.iloc[2, col_idx]
        if pd.isna(name_val) or str(name_val).strip() == "":
            break
        product_names[col_idx] = str(name_val).strip()

    # --- 3. 遍历数据行（第6行起） ---
    results = []
    for i in range(5, len(df)):
        total_money = df.iloc[i, 0]
        person_name = df.iloc[i, 1]

        if pd.notna(total_money) and pd.notna(person_name):
            purchased_details = []
            total_points = 0

            for col_idx, item_name in product_names.items():
                count = df.iloc[i, col_idx]
                if pd.notna(count) and isinstance(count, (int, float)) and count > 0:
                    category = col_to_category.get(col_idx, "默认分类")
                    total_points += int(count)

                    if category == "默认分类":
                        detail = f"{item_name}✖{int(count)}"
                    else:
                        detail = f"({category}){item_name}✖{int(count)}"

                    purchased_details.append(detail)

            if purchased_details:
                results.append({
                    "名字": str(person_name).strip(),
                    "（分类名称）/种类×个数": " / ".join(purchased_details),
                    "总点数": total_points,
                    "对应的总金额": str(total_money).strip()
                })

    if results:
        pd.DataFrame(results).to_excel(output_file, index=False)


# ==================================================
# 2️⃣ 重量表（计算每个人的制品总重量）
# ==================================================
def generate_weight_table(input_file, output_file):
    all_sheets = pd.read_excel(input_file, sheet_name=None, header=None)
    writer = pd.ExcelWriter(output_file, engine="openpyxl")

    for sheet_name, df in all_sheets.items():
        if df.shape[0] < 6 or df.shape[1] < 3:
            continue

        weights = df.iloc[1, 2:].astype(float)        # 第2行：重量
        names = df.iloc[5:, 1]                        # B列：名字
        qty = df.iloc[5:, 2:].fillna(0)
        qty = qty.apply(pd.to_numeric, errors="coerce").fillna(0)

        total_weight = qty.dot(weights)

        result_df = pd.DataFrame({
            "名字": names.values,
            "总重量(g)": total_weight.values
        })

        result_df = result_df[result_df["名字"].notna()]
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()


# ==================================================
# 3️⃣ 国际表（多 sheet 横向汇总）
# ==================================================
def generate_international_table(input_file, output_file):
    all_sheets = pd.read_excel(input_file, sheet_name=None)
    result_df = None

    for sheet_name, df in all_sheets.items():
        if df.empty:
            continue

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

        result_df = sheet_summary if result_df is None else \
            result_df.merge(sheet_summary, on='名字', how='outer')

    amount_cols = [c for c in result_df.columns if c.endswith('_金额')]
    result_df['汇总金额'] = result_df[amount_cols].sum(axis=1, skipna=True)

    result_df.to_excel(output_file, index=False)


# ==================================================
# 4️⃣ 本地调试入口（Streamlit 用不到）
# ==================================================
if __name__ == "__main__":
    transform_summary_to_list("2.xlsx", "改_2.xlsx")
    generate_weight_table("test.xlsx", "制品总重量统计.xlsx")
    generate_international_table("gj.xlsx", "国际表.xlsx")
