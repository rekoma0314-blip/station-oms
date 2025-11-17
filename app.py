import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="拣货单生成系统", layout="wide")

st.title("⛽ 便利店 & 参股站 拣货单生成系统")
st.caption("版本 v2.1 · 修复字段名 · 支持你当前 Excel 格式")

############################################################
# 1. 上传文件
############################################################
st.markdown("### 📤 第一步：上传订单 & 数据文件")

col1, col2 = st.columns(2)

with col1:
    oil_file = st.file_uploader("① 官网订单（新编码）", type=["xlsx"])
    manual_file = st.file_uploader("② 手工订单（旧编码）", type=["xlsx"])

with col2:
    master_file = st.file_uploader("③ 主表（SKU）", type=["xlsx"])
    site_file = st.file_uploader("④ 站点仓库对照表（code+name+warehouse）", type=["xlsx"])

if not oil_file or not manual_file or not master_file or not site_file:
    st.info("👆 请上传全部 4 个文件后继续")
    st.stop()

############################################################
# 2. 读取
############################################################
try:
    oil = pd.read_excel(oil_file)
    manual = pd.read_excel(manual_file)
    master = pd.read_excel(master_file)
    site = pd.read_excel(site_file)
    st.success("🎉 文件读取成功！")
except Exception as e:
    st.error(f"❌ Excel 读取失败：{e}")
    st.stop()

############################################################
# 3. 字段标准化（替换旧逻辑）
############################################################

# 必须包含：code / name / warehouse
required_cols = ["code", "name", "warehouse"]

for col in required_cols:
    if col not in site.columns:
        st.error(f"❌ 站点表缺少字段：{col}")
        st.stop()

site["code"] = site["code"].astype(str).str.strip()

############################################################
# 4. 订单字段清洗
############################################################

oil.rename(columns={
    "收货组织编码": "code",
    "订货数量": "数量",
    "商品编码": "商品编码",
}, inplace=True)

manual.rename(columns={
    "油站编码": "code",
    "订货数量": "数量",
    "商品编码": "商品编码",
}, inplace=True)

############################################################
# 5. 合并订单
############################################################

oil["来源"] = "官网"
manual["来源"] = "手工"

orders = pd.concat([oil, manual], ignore_index=True)

############################################################
# 6. 匹配站点仓库
############################################################

orders = orders.merge(site[["code", "name", "warehouse"]], on="code", how="left")

unmatched_sites = orders[orders["warehouse"].isna()]

############################################################
# 7. SKU 校验
############################################################

master.rename(columns={
    "商品编码": "商品编码",
    "油站订货目录": "可订",
}, inplace=True)

orders = orders.merge(master[["商品编码", "可订"]], on="商品编码", how="left")

unavailable_sku = orders[orders["可订"] != "油站可订"]

############################################################
# 8. 正常订单
############################################################

valid_orders = orders.drop(unmatched_sites.index).drop(unavailable_sku.index)

############################################################
# 9. 下载按钮
############################################################

def df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

st.markdown("## 📦 第二步：下载拣货单")

if st.button("🚀 一键生成今日拣货单", type="primary"):
    if valid_orders.empty:
        st.warning("⚠ 没有可用订单，请检查源文件")
    else:
        for wh in valid_orders["warehouse"].dropna().unique():
            df = valid_orders[valid_orders["warehouse"] == wh]
            st.download_button(
                label=f"📦 下载：{wh}",
                data=df_to_excel(df),
                file_name=f"拣货单_{wh}.xlsx",
            )

############################################################
# 10. 异常报告
############################################################

st.markdown("---")
st.markdown("### ❌ 异常站点")
st.dataframe(unmatched_sites)

st.markdown("### ❌ 异常 SKU")
st.dataframe(unavailable_sku)

