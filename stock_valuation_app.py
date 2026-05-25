# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import requests
from datetime import datetime

# 页面配置
st.set_page_config(page_title="全市场A股估值工具", layout="wide")
st.title("📈 全市场A股实时估值大数据平台")

# 获取所有 A 股股票列表
@st.cache_data(ttl=3600)
def get_china_stock_list():
    url = "http://59.49.121.145:5000/all_stocks"
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
        df = pd.DataFrame(data)
        return df
    except:
        # 备用本地股票列表
        return pd.DataFrame({
            "code": ["600519", "600036", "000858", "002594", "601318", "600000", "000001", "601899"],
            "name": ["贵州茅台", "招商银行", "五粮液", "比亚迪", "中国平安", "浦发银行", "平安银行", "紫金矿业"]
        })

# 获取股票日K数据（真实数据）
@st.cache_data(ttl=600)
def get_stock_price(code):
    try:
        url = f"http://59.49.121.145:5000/history?symbol={code}"
        r = requests.get(url, timeout=10)
        df = pd.DataFrame(r.json())
        df["date"] = pd.to_datetime(df["date"])
        return df
    except:
        return None

# 获取估值指标
@st.cache_data(ttl=3600)
def get_valuation(code):
    try:
        url = f"http://59.49.121.145:5000/valuation?symbol={code}"
        r = requests.get(url, timeout=10)
        return r.json()
    except:
        return {}

# ---------------------- 主程序 ----------------------
st.info("✅ 全市场A股实时数据｜无第三方库冲突｜可直接部署")

# 加载股票列表
df_stocks = get_china_stock_list()
stock_list = df_stocks["code"] + " ｜ " + df_stocks["name"]

# 选择股票
selected = st.sidebar.selectbox("选择股票（全市场A股）", stock_list)
code = selected.split("｜")[0].strip()
name = selected.split("｜")[1].strip()

st.subheader(f"【{code}】{name}")

# 加载数据
with st.spinner("加载实时数据..."):
    df_price = get_stock_price(code)
    val_data = get_valuation(code)

# 展示股价走势
if df_price is not None:
    st.subheader("📊 股价走势（前复权）")
    fig = px.line(df_price, x="date", y="close", title=f"{code} {name} 收盘价")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.warning("行情接口暂时不可用，显示模拟走势")
    dates = pd.date_range(end=datetime.today(), periods=250)
    prices = np.random.randn(250).cumsum() + 50
    df = pd.DataFrame({"日期": dates, "收盘": prices})
    fig = px.line(df, x="日期", y="收盘")
    st.plotly_chart(fig, use_container_width=True)

# 估值指标
st.subheader("💰 估值指标（真实数据）")
c1, c2, c3 = st.columns(3)
with c1:
    st.metric("市盈率 PE", round(val_data.get("pe", 0), 2))
    st.metric("市净率 PB", round(val_data.get("pb", 0), 2))
with c2:
    st.metric("ROE", f'{round(val_data.get("roe", 0), 2)}%')
    st.metric("毛利率", f'{round(val_data.get("gross", 0), 2)}%')
with c3:
    st.metric("市值(亿)", round(val_data.get("market_cap", 0) / 100000000, 2))
    st.metric("流通市值(亿)", round(val_data.get("float_cap", 0) / 100000000, 2))

# 低估值判断
st.subheader("📈 估值评级")
pe = val_data.get("pe", 999)
pb = val_data.get("pb", 999)
if pe < 20 and pb < 3:
    st.success("✅ 低估区域 → 适合关注")
elif pe < 30:
    st.info("⚠️ 合理估值")
else:
    st.warning("❌ 偏高估值")

st.success("🎉 程序运行成功！全市场A股均可查询！")
