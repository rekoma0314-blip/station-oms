import streamlit as st
import pandas as pd
from io import BytesIO
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="站点拣货单生成系统", layout="wide")

st.title("⛽ 便利店 & 参股站 拣货单生成系统")
st.caption("版本 v1.0 · 自动读取 Google Sheet 站点仓库表 · 不再需要每次上传")

############################################################
# 1. 连接 Google Sheet - 站点仓库表
############################################################

st.markdown("### 🏭 自动读取：站点仓库对照表（Google Sheet）")

conn = st.connection("gsheets", type=GSheetsConnection)

sheet_url = "https://docs.google.com/spreadsheets/d/1Tqh0eGvawJoVhCVu_dupuYXx9-DuBoz-1nwJl4yQYnk/edit?usp=sharing"

try:
    site = conn.read(spreadsheet=sheet_url, header=0)
    st.success("已成功从 Google Sheet 读取站点仓库表！")
    st.dataframe(site.head())
except Exception as e:
    st.error(f"❌ 无法读取 Google Sheet，请检查链接或权限：{e}")
    st.stop()

############################################################
# 2. 上传官网订单、手工订单、主表
############################################################

st.markdown("### 第一步：上传订单文件（官网 + 手工 + 主表）")

col1, col2 = st.columns(2)

with col1:
    oil_file = st.file_uploader("📄 ① 官网页订单（新编码）", type=["xlsx"])
    manual_file = st.file_uploader("📄 ② 手工订单（参股站旧编码）", type=["xlsx"])

with col2:
    master_file = st.file_uploader("📄 ③ 主表（SKU）", type=["xlsx"])

if not (oil_file and manual_file and master_file):
    st.info("👆 请上传全部文件后继续。")
    st.stop()

############################################################
# 3. 读取文件
############################################################

try:
    oil = pd.read_excel(oil_file)
    manual = pd.read_excel(manual_file)
    master = pd.read_excel(master_file)
except Exception as e:
    st.error(f"❌ Excel 文件读取失败：{e}")
    st.stop()

############################################################
# 4. 字段重命名（适配你的真实列名）
############################################################

# 官网订单字段映射
oil.rename(columns={
    "收货组织编码": "站点编码",
    "收货组织名称": "站点名称",
    "商品编码": "商品编码",
    "商品名称": "商品名称",
    "订货数量": "数量"
}, inplace=True)

# 手工订单字段映射
manual.rename(columns={
    "油站编码": "站点编码",     # 旧编码
    "油站名称": "站点名称",
    "商品编码": "商品编码",
    "订货数量": "数量"
}, inplace=True)

# 主表字段映射
master.rename(columns={
    "商品编码": "商品编码",
    "油站订货目录": "油站订货目录"
}, inplace=True)

############################################################
# 5. 开始生成拣货单
############################################################

st.markdown("### 第二步：生成拣货单")

if st.button("🚀 一键生成今日拣货单", type="primary"):

    # 清洗编码
    oil["站点编码"] = oil["站点编码"].astype(str).str.strip()
    manual["站点编码"] = manual["站点编码"].astype(str).str.strip()
    site["油站编码"] = site["油站编码"].astype(str).str.strip()
    site["便利店新编码"] = site["便利店新编码"].astype(str).str.strip()

    ############################################################
    # 6. 匹配仓库
    ############################################################

    st.write("▶ 站点仓库匹配中…")

    # 官网订单：新编码（收货组织编码） → 便利店新编码 匹配
    oil = oil.merge(
        site[["便利店新编码", "站点名称", "仓库"]],
        left_on="站点编码",
        right_on="便利店新编码",
        how="left",
        suffixes=("", "_站点表")
    )

    # 手工订单：旧编码（油站编码） → 油站编码 匹配
    manual = manual.merge(
        site[["油站编码", "站点名称", "仓库"]],
        left_on="站点编码",
        right_on="油站编码",
        how="left",
        suffixes=("", "_站点表")
    )

    # 合并订单
    oil["来源"] = "官网订单"
    manual["来源"] = "手工订单"

    common_cols = list(set(oil.columns) | set(manual.columns))
    oil = oil.reindex(columns=common_cols)
    manual = manual.reindex(columns=common_cols)

    orders = pd.concat([oil, manual], ignore_index=True)

    ############################################################
    # 7. SKU 校验
    ############################################################

    st.write("▶ SKU 校验中…")

    orders = orders.merge(
        master[["商品编码", "油站订货目录"]],
        on="商品编码",
        how="left"
    )

    abnormal_sku = orders[(orders["油站订货目录"].isna()) | (orders["油站订货目录"] != "油站可订")]

    abnormal_site = orders[orders["仓库"].isna()]

    normal_orders = orders.drop(abnormal_sku.index).drop(abnormal_site.index)

    st.success("校验完成！")

    ############################################################
    # 8. 下载功能
    ############################################################

    def df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.markdown("### 📦 第三步：下载拣货单（按仓拆分）")

    if normal_orders.empty:
        st.warning("⚠ 没有正常订单，请检查文件。")
    else:
        for wh in normal_orders["仓库"].dropna().unique():
            df_wh = normal_orders[normal_orders["仓库"] == wh]
            st.download_button(
                label=f"📦 下载拣货单：{wh}",
                data=df_to_excel(df_wh),
                file_name=f"拣货单_{wh}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")

    st.markdown("### 🧾 异常报告")

    colA, colB = st.columns(2)

    with colA:
        st.write("❌ 异常 SKU（停订或未找到）")
        st.dataframe(abnormal_sku.head())
        st.download_button(
            "下载异常SKU",
            df_to_excel(abnormal_sku),
            "异常SKU.xlsx"
        )

    with colB:
        st.write("❌ 异常站点（未匹配到仓库）")
        st.dataframe(abnormal_site.head())
        st.download_button(
            "下载异常站点",
            df_to_excel(abnormal_site),
            "异常站点.xlsx"
        )

