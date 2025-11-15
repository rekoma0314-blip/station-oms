import streamlit as st
import pandas as pd
from io import BytesIO

# 引入 Supabase
from utils.supabase_client import supabase
from utils.site_utils import get_all_sites, get_site_by_code


st.set_page_config(page_title="站点拣货单系统（升级版）", layout="wide")

st.title("⛽ 便利店 & 参股站 拣货单生成系统（升级版）")
st.caption("版本 v2.0 · 站点仓库从 Supabase 自动读取 · 支持铺货记录（防止重复铺）")


# =============================================================
# 1. 自动读取 Supabase 站点表
# =============================================================
st.markdown("### 🏭 站点仓库：自动从 Supabase 加载")

with st.spinner("正在读取数据库站点列表…"):
    site = pd.DataFrame(get_all_sites())

if site.empty:
    st.error("❌ Supabase 没有任何站点，请先在数据库 manually 导入数据。")
    st.stop()

st.success(f"已加载 {len(site)} 个站点。")


# =============================================================
# 2. 上传订单文件
# =============================================================
st.markdown("### 第一步：上传订单文件")

col1, col2 = st.columns(2)
with col1:
    oil_file = st.file_uploader("📄 ① 官网页订单（新编码）", type=["xlsx"])
    manual_file = st.file_uploader("📄 ② 手工订单（参股站旧编码）", type=["xlsx"])
with col2:
    master_file = st.file_uploader("📄 ③ 主表（SKU）", type=["xlsx"])

if not (oil_file and manual_file and master_file):
    st.info("👆 请上传全部文件。")
    st.stop()


# =============================================================
# 3. 读取 Excel
# =============================================================
try:
    oil = pd.read_excel(oil_file)
    manual = pd.read_excel(manual_file)
    master = pd.read_excel(master_file)
except Exception as e:
    st.error(f"❌ Excel 读取失败: {e}")
    st.stop()

st.success("Excel 文件读取成功！")


# =============================================================
# 4. 字段重命名
# =============================================================
oil.rename(columns={
    "收货组织编码": "站点编码",
    "商品编码": "商品编码",
    "商品名称": "商品名称",
    "订货数量": "数量",
}, inplace=True)

manual.rename(columns={
    "油站编码": "站点编码",
    "商品编码": "商品编码",
    "订货数量": "数量",
}, inplace=True)

master.rename(columns={
    "商品编码": "商品编码",
    "油站订货目录": "油站订货目录",
}, inplace=True)

# Supabase 站点字段统一：
# site 表结构（你的数据库定义）
# site_code（新）、old_code（旧）、warehouse、name
site.rename(columns={
    "site_code": "新编码",
    "old_code": "旧编码",
    "warehouse": "仓库",
    "name": "站点名称"
}, inplace=True)


# =============================================================
# 5. 合并订单
# =============================================================
st.markdown("### 第二步：生成拣货单")

if st.button("🚀 一键生成今日拣货单", type="primary"):

    oil["站点编码"] = oil["站点编码"].astype(str).str.strip()
    manual["站点编码"] = manual["站点编码"].astype(str).str.strip()

    # 官网订单：新编码匹配
    oil = oil.merge(
        site[["新编码", "仓库", "站点名称"]],
        left_on="站点编码",
        right_on="新编码",
        how="left"
    )

    # 手工订单：旧编码匹配
    manual = manual.merge(
        site[["旧编码", "仓库", "站点名称"]],
        left_on="站点编码",
        right_on="旧编码",
        how="left"
    )

    # 区分来源
    oil["来源"] = "官网订单"
    manual["来源"] = "手工订单"

    orders = pd.concat([oil, manual], ignore_index=True)

    # SKU 校验
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


    # =============================================================
    # 6. 铺货记录写入 Supabase（防止重复铺）
    # =============================================================
    st.markdown("### 🪣 自动记录铺货（防重复）")

    inserted = 0
    skipped = 0

    for _, row in normal_orders.iterrows():
        site_code = str(row["站点编码"])
        sku_code = str(row["商品编码"])

        # 判断是否已经铺过
        exists = (
            supabase.table("activity_records")
            .select("*")
            .eq("site_code", site_code)
            .eq("sku_code", sku_code)
            .execute()
        )

        if exists.data:
            skipped += 1
            continue

        # 新增铺货记录
        supabase.table("activity_records").insert({
            "site_code": site_code,
            "sku_code": sku_code,
            "activity_name": "自动拣货铺货"
        }).execute()

        inserted += 1

    st.success(f"新增铺货记录 {inserted} 条，跳过重复铺货 {skipped} 条。")


    # =============================================================
    # 7. 导出
    # =============================================================

    def df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.markdown("### 📦 下载拣货单（按仓拆分）")

    for wh in normal_orders["仓库"].dropna().unique():
        df_wh = normal_orders[normal_orders["仓库"] == wh]
        st.download_button(
            label=f"📦 下载拣货单：{wh}",
            data=df_to_excel(df_wh),
            file_name=f"拣货单_{wh}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("---")
    colA, colB = st.columns(2)

    with colA:
        st.write("❌ 异常 SKU")
        st.dataframe(abnormal_sku)

    with colB:
        st.write("❌ 异常站点")
        st.dataframe(abnormal_site)
