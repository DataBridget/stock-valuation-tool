# -*- coding: utf-8 -*-
import streamlit as st
import akshare as ak
import pandas as pd
import numpy as np
import plotly.express as px
from cachetools import TTLCache

# 缓存设置
cache = TTLCache(maxsize=100, ttl=3600)

@st.cache_data
def get_all_stocks():
    if "all_stocks" not in cache:
        df = ak.stock_info_a_code_name()
        cache["all_stocks"] = df
    return cache["all_stocks"]

@st.cache_data
def get_stock_daily(code):
    df = ak.stock_zh_a_hist(symbol=code, period="daily", adjust="qfq")
    df["日期"] = pd.to_datetime(df["日期"])
    return df

@st.cache_data
def get_stock_financial(code):
    df = ak.stock_financial_analysis_indicator(symbol=code)
    return df

def relative_valuation(df_fin):
    pe = df_fin["市盈率(静)"].dropna()
    pb = df_fin["市净率"].dropna()
    pe_q = np.percentile(pe, [10, 50, 90])
    pb_q = np.percentile(pb, [10, 50, 90])
    return {"PE分位": pe_q, "PB分位": pb_q}

def dcf_valuation(df_fin, growth=0.05, wacc=0.08):
    net_profit = df_fin["净利润(元)"].dropna().iloc[-5:].mean()
    fcf = net_profit * 0.8
    value = 0
    for i in range(1, 4):
        fcf *= (1 + growth)
        value += fcf / ((1 + wacc) ** i)
    terminal_value = fcf * (1 + 0.02) / (wacc - 0.02)
    value += terminal_value / ((1 + wacc) ** 3)
    return round(value / 1e8, 2)

st.set_page_config(page_title="股票估值大数据平台", layout="wide")
st.title("📈 股票估值大数据App")

all_stocks = get_all_stocks()
stock_code = st.sidebar.selectbox("选择股票", all_stocks["代码"] + " - " + all_stocks["名称"])
code = stock_code.split(" - ")[0]

with st.spinner("加载数据中..."):
    df_daily = get_stock_daily(code)
    df_fin = get_stock_financial(code)

st.subheader("股价走势（近1年）")
fig = px.line(df_daily.tail(250), x="日期", y="收盘", title=f"{stock_code} 收盘价")
st.plotly_chart(fig, use_container_width=True)

st.subheader("相对估值（PE/PB分位）")
rel_val = relative_valuation(df_fin)
st.metric("PE 10%/50%/90%分位", f"{rel_val['PE分位'][0]:.1f} / {rel_val['PE分位'][1]:.1f} / {rel_val['PE分位'][2]:.1f}")
st.metric("PB 10%/50%/90%分位", f"{rel_val['PB分位'][0]:.1f} / {rel_val['PB分位'][1]:.1f} / {rel_val['PB分位'][2]:.1f}")

st.subheader("DCF绝对估值（亿元）")
dcf_value = dcf_valuation(df_fin)
current_price = df_daily.iloc[-1]["收盘"]
st.metric("DCF内在价值", dcf_value)
st.metric("当前股价", current_price)
