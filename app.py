import streamlit as st
import pandas as pd
from utils.site_utils import get_site_map
from io import BytesIO


st.set_page_config(page_title="站点拣货单系统", layout="wide")
st.title("🚚 站点拣货单生成系统 v2.0")

# ----------- 上传文件区 -----------

col1, col2 = st.columns(2)

with col1:
    oil_file = st.file_uploader("① 官网订单（新编码）", type=["xlsx"])
    manual_file = st.file_uploader("② 手工订单（参股旧编码）", type=["xlsx"])

with col2:
    sku_file = st.file_uploader("③ 主表（SKU）", type=["xlsx"])


if not (oil_file and manual_file and sku_file):
    st.stop()

oil = pd.read_excel(oil_file)
manual = pd.read_excel(manual_file)
sku = pd.read_excel(sku_file)

st.success("文件读取成功！")

# ----------- 字段适配 -----------

oil.rename(columns={
    "收货组织编码": "站点编码",
    "商品编码": "商品编码",
    "订货数量": "数量"
}, inplace=True)

manual.rename(columns={
    "油站编码": "站点编码",
    "商品编码": "商品编码",
    "订货数量": "数量"
}, inplace=True)

# ----------- 联合订单 -----------

orders = pd.concat([oil, manual], ignore_index=True)

# ----------- 匹配 Supabase 仓库 -----------

site_map = get_site_map()

orders["仓库"] = orders["站点编码"].map(lambda c: site_map.get(c, {}).get("warehouse"))
orders["站点名称"] = orders["站点编码"].map(lambda c: site_map.get(c, {}).get("name"))

abnormal = orders[orders["仓库"].isna()]
orders = orders.dropna(subset=["仓库"])

st.info(f"成功匹配 {len(orders)} 条数据")

# ----------- 下载按钮 -----------

def df_to_excel(df):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return bio.getvalue()

st.subheader("📦 分仓导出拣货单")

for wh in orders["仓库"].unique():
    df = orders[orders["仓库"] == wh]
    st.download_button(
        label=f"下载仓库：{wh}",
        data=df_to_excel(df),
        file_name=f"拣货单_{wh}.xlsx"
    )

if len(abnormal):
    st.warning("❌ 有未匹配站点：")
    st.dataframe(abnormal)


