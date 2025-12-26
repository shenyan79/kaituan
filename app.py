import streamlit as st
import tempfile
import os

from app import (
    transform_summary_to_list,
    generate_weight_table,
    generate_international_table
)

st.set_page_config(page_title="制品清单全能转换工具", layout="wide")

# ================= 左侧栏 =================
st.sidebar.title("请选择转换功能")

mode = st.sidebar.radio(
    "转换模式",
    (
        "单页转换：横向区间模式",
        "重量表模式",
        "多Sheet合并汇总模式"
    )
)

st.sidebar.info(f"当前模式：{mode}")

# ================= 主界面 =================
st.title("🛠️ 制品清单全能转换工具")

uploaded_file = st.file_uploader(
    "上传 Excel 文件（.xlsx）",
    type=["xlsx"]
)

if uploaded_file:
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, uploaded_file.name)
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # ===== 根据模式调用不同功能 =====
        if mode == "单页转换：横向区间模式":
            output_path = os.path.join(tmpdir, "清单表.xlsx")
            transform_summary_to_list(input_path, output_path)

        elif mode == "重量表模式":
            output_path = os.path.join(tmpdir, "重量表.xlsx")
            generate_weight_table(input_path, output_path)

        elif mode == "多Sheet合并汇总模式":
            output_path = os.path.join(tmpdir, "国际表.xlsx")
            generate_international_table(input_path, output_path)

        # ===== 下载按钮 =====
        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 下载转换后的 Excel",
                data=f,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
