import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="站点拣货单生成系统", layout="wide")

st.title("⛽ 便利店 & 参股站 拣货单生成系统")
st.caption("版本 v0.1 · by 泡泡自用 · 站点/仓库/主表均可随时更新")

st.markdown("### 第一步：上传 4 份基础数据文件")

st.write("""
请依次上传以下 4 个 Excel 文件（后续都可以随时换新的）：

1. **官网订单（新编码）**：从官网导出的油站订单  
2. **手工订单（旧编码-参股站）**：二级公司/参股站提交的模板  
3. **主表（SKU 可订表）**：含“商品编码”“油站订货目录”等字段  
4. **站点仓库对照表（统一结构表）**：你维护的站点编码 ↔ 仓库 映射
""")

col1, col2 = st.columns(2)

with col1:
    oil_file = st.file_uploader("① 官网页订单（新编码）", type=["xlsx"], key="oil")
    manual_file = st.file_uploader("② 手工订单（旧编码 - 参股站）", type=["xlsx"], key="manual")

with col2:
    master_file = st.file_uploader("③ 主表（SKU 可订）", type=["xlsx"], key="master")
    site_file = st.file_uploader("④ 站点仓库对照表（统一结构)", type=["xlsx"], key="site")

if not (oil_file and manual_file and master_file and site_file):
    st.info("👆 请先把 4 个文件全部上传，右上角可以看到上传状态。")
    st.stop()

# 读取 Excel
try:
    oil = pd.read_excel(oil_file)
    manual = pd.read_excel(manual_file)
    master = pd.read_excel(master_file)
    site = pd.read_excel(site_file)
except Exception as e:
    st.error(f"读取 Excel 失败，请检查文件格式是否为 .xlsx，错误信息：{e}")
    st.stop()

st.success("✅ 文件读取成功，可以继续生成拣货单。")

st.markdown("---")
st.markdown("### 第二步：检查字段并一键生成拣货单")

# 你实际的列名如果不同，这里我们后面可以再改
# 先假定：
# 官网订单：包含 “站点编码”“商品编码”
# 手工订单：包含 “油站编码”“商品编码”
# 主表：包含 “商品编码”“油站订货目录”
# 站点表：包含 “站点编码类型”“站点编码”“站点名称”“仓库”

required_oil_cols = ["站点编码", "商品编码"]
required_manual_cols = ["油站编码", "商品编码"]
required_master_cols = ["商品编码", "油站订货目录"]
required_site_cols = ["站点编码类型", "站点编码", "站点名称", "仓库"]

def check_columns(df, required, name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"❌ {name} 缺少必需列：{missing}")
        return False
    return True

ok = True
ok &= check_columns(oil, required_oil_cols, "官网订单")
ok &= check_columns(manual, required_manual_cols, "手工订单")
ok &= check_columns(master, required_master_cols, "主表（SKU）")
ok &= check_columns(site, required_site_cols, "站点仓库对照表")

if not ok:
    st.stop()

if st.button("🚀 一键生成今日拣货单", type="primary"):

    # ===== 清洗 & 标准化 =====
    st.write("▶ 正在清洗字段...")

    # 统一编码为字符串去空格
    oil["站点编码"] = oil["站点编码"].astype(str).str.strip()
    manual["油站编码"] = manual["油站编码"].astype(str).str.strip()
    oil["商品编码"] = oil["商品编码"].astype(str).str.strip()
    manual["商品编码"] = manual["商品编码"].astype(str).str.strip()
    master["商品编码"] = master["商品编码"].astype(str).str.strip()
    site["站点编码"] = site["站点编码"].astype(str).str.strip()

    # ===== 官网订单：新编码匹配站点表 =====
    st.write("▶ 官网订单：匹配站点仓库（新编码）...")
    site_new = site[site["站点编码类型"] == "新编码"]
    oil = oil.merge(
        site_new[["站点编码", "站点名称", "仓库"]],
        on="站点编码",
        how="left",
        suffixes=("", "_站点表")
    )

    # ===== 手工订单：旧编码匹配站点表 =====
    st.write("▶ 手工订单：匹配站点仓库（旧编码）...")
    site_old = site[site["站点编码类型"] == "旧编码"]

    manual = manual.merge(
        site_old[["站点编码", "站点名称", "仓库"]],
        left_on="油站编码",
        right_on="站点编码",
        how="left",
        suffixes=("", "_站点表")
    )

    # 手工统一站点字段名
    manual["站点编码"] = manual["油站编码"]

    # ===== 合并订单 =====
    st.write("▶ 合并官网订单 + 手工订单...")
    oil["来源"] = "官网订单"
    manual["来源"] = "手工订单"

    # 为了便于后续处理，缺的列我们可以补空（视你真实数据情况）
    common_cols = list(set(oil.columns) | set(manual.columns))
    oil = oil.reindex(columns=common_cols)
    manual = manual.reindex(columns=common_cols)

    orders = pd.concat([oil, manual], ignore_index=True)

    # ===== SKU 校验：是否为油站可订 =====
    st.write("▶ 匹配主表（SKU 可订/停订 信息）...")
    orders = orders.merge(
        master[["商品编码", "油站订货目录"]],
        on="商品编码",
        how="left"
    )

    # 异常 SKU：主表无此编码 或 油站停订
    abnormal_sku = orders[
        (orders["油站订货目录"].isna()) |
        (orders["油站订货目录"] != "油站可订")
    ].copy()

    # 异常站点：没有匹配到仓库
    abnormal_site = orders[orders["仓库"].isna()].copy()

    # 正常订单 = 去掉异常 SKU + 异常站点
    normal_orders = orders.drop(abnormal_sku.index).drop(abnormal_site.index)

    st.success("✅ 订单合并 & 校验完成！")

    # ===== 导出为 Excel，需要先写入内存 BytesIO =====
    def to_excel_bytes(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        return output.getvalue()

    st.markdown("### 第三步：下载拣货单（按仓拆分）")

    if normal_orders.empty:
        st.warning("⚠ 没有任何正常订单，请检查主表、站点表或源数据。")
    else:
        warehouses = normal_orders["仓库"].dropna().unique()
        for wh in warehouses:
            df_wh = normal_orders[normal_orders["仓库"] == wh]
            st.write(f"📦 仓库：**{wh}**，订单行数：{len(df_wh)}")
            excel_bytes = to_excel_bytes(df_wh)
            st.download_button(
                label=f"下载拣货单：{wh}",
                data=excel_bytes,
                file_name=f"拣货单_{wh}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")
    st.markdown("### 第四步：下载异常报告（方便你人工处理）")

    col_a, col_b = st.columns(2)
    with col_a:
        st.write("🧾 异常 SKU 报告")
        if abnormal_sku.empty:
            st.success("没有异常 SKU，太棒了 ✨")
        else:
            excel_bytes = to_excel_bytes(abnormal_sku)
            st.download_button(
                "下载异常 SKU 报告",
                data=excel_bytes,
                file_name="异常SKU报告.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.dataframe(abnormal_sku.head(10))

    with col_b:
        st.write("📍 异常站点 报告（编码未匹配到仓库）")
        if abnormal_site.empty:
            st.success("所有站点都匹配到仓库 ✅")
        else:
            excel_bytes = to_excel_bytes(abnormal_site)
            st.download_button(
                "下载异常站点报告",
                data=excel_bytes,
                file_name="异常站点报告.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.dataframe(abnormal_site.head(10))
