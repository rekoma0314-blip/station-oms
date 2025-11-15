import streamlit as st
import pandas as pd
from io import BytesIO
from utils.site_utils import get_site_by_code

st.set_page_config(page_title="站点拣货单生成系统（升级版）", layout="wide")

st.title("⛽ 便利店 & 参股站 拣货单生成系统（升级版）")
st.caption("版本 v2.0 · 站点仓库从 Supabase 自动读取 · 支持铺货记录（防止重复铺货）")


# ---------------------------------------------------
# 上传 Excel 文件
# ---------------------------------------------------
st.markdown("### 第一步：上传订单文件")

col1, col2 = st.columns(2)

with col1:
    oil_file = st.file_uploader("📄 ① 官网页订单（新编码）", type=["xlsx"])
    manual_file = st.file_uploader("📄 ② 手工订单（参股站旧编码）", type=["xlsx"])

with col2:
    master_file = st.file_uploader("📄 ③ 主表（SKU）", type=["xlsx"])

if not (oil_file and manual_file and master_file):
    st.info("👆 请上传全部文件后继续")
    st.stop()


# ---------------------------------------------------
# 读取 Excel
# ---------------------------------------------------
try:
    oil = pd.read_excel(oil_file)
    manual = pd.read_excel(manual_file)
    master = pd.read_excel(master_file)
except Exception as e:
    st.error(f"❌ Excel 文件读取失败：{e}")
    st.stop()

st.success("文件读取成功！")


# ---------------------------------------------------
# 字段统一（按你提供的真实字段）
# ---------------------------------------------------

oil.rename(columns={
    "收货组织编码": "站点编码",
    "收货组织名称": "站点名称",
    "商品编码": "商品编码",
    "商品名称": "商品名称",
    "订货数量": "数量"
}, inplace=True)

manual.rename(columns={
    "油站编码": "站点编码",
    "油站名称": "站点名称",
    "商品编码": "商品编码",
    "订货数量": "数量"
}, inplace=True)

master.rename(columns={
    "商品编码": "商品编码",
    "油站订货目录": "油站订货目录"
}, inplace=True)


# ---------------------------------------------------
# 从 Supabase 匹配仓库
# ---------------------------------------------------

def attach_site_info(df):
    df["仓库"] = None
    df["公司归属"] = None
    df["站点名称（数据库）"] = None

    for idx, row in df.iterrows():

        code = str(row["站点编码"]).strip()
        site = get_site_by_code(code)

        if site:
            df.at[idx, "仓库"] = site.get("warehouse")
            df.at[idx, "公司归属"] = site.get("company")
            df.at[idx, "站点名称（数据库）"] = site.get("name")

    return df


st.markdown("### 第二步：生成拣货单")

if st.button("🚀 一键生成今日拣货单", type="primary"):

    st.write("▶ 正在匹配 Supabase 站点仓库信息…")

    oil = attach_site_info(oil)
    manual = attach_site_info(manual)

    oil["来源"] = "官网"
    manual["来源"] = "手工"

    orders = pd.concat([oil, manual], ignore_index=True)

    # ---------------------------------------------------
    # SKU 校验
    # ---------------------------------------------------
    st.write("▶ SKU 校验中…")

    orders = orders.merge(
        master[["商品编码", "油站订货目录"]],
        on="商品编码",
        how="left"
    )

    abnormal_sku = orders[
        (orders["油站订货目录"].isna()) |
        (orders["油站订货目录"] != "油站可订")
    ]

    abnormal_site = orders[orders["仓库"].isna()]

    normal_orders = orders.drop(abnormal_sku.index).drop(abnormal_site.index)

    st.success("校验完成！")

    # ---------------------------------------------------
    # 下载区
    # ---------------------------------------------------
    def df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.markdown("### 📦 第三步：下载拣货单（按仓库拆分）")

    if normal_orders.empty:
        st.warning("⚠ 没有正常订单，请检查文件")
    else:
        for wh in normal_orders["仓库"].dropna().unique():
            sub_df = normal_orders[normal_orders["仓库"] == wh]
            st.download_button(
                label=f"📦 下载拣货单：{wh}",
                data=df_to_excel(sub_df),
                file_name=f"拣货单_{wh}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ---------------------------------------------------
    # 异常报告
    # ---------------------------------------------------
    st.markdown("---")
    st.markdown("### 🧾 异常报告")

    col1, col2 = st.columns(2)

    with col1:
        st.write("❌ 异常 SKU（停订或未找到）")
        st.dataframe(abnormal_sku)
        st.download_button(
            "下载异常 SKU",
            df_to_excel(abnormal_sku),
            "异常SKU.xlsx"
        )

    with col2:
        st.write("❌ 异常站点（Supabase 未找到仓库）")
        st.dataframe(abnormal_site)
        st.download_button(
            "下载异常站点",
            df_to_excel(abnormal_site),
            "异常站点.xlsx"
        )
