# -*- coding: utf-8 -*-
"""
全市场A股实时估值大数据平台 v4.0
基于Baostock真实数据，支持PE/DCF/PB多维估值、可视化分析、投资建议
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import io
import warnings
warnings.filterwarnings("ignore")

# 报告生成依赖
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors as rl_colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# ==================== Baostock 数据接口 ====================
import baostock as bs

_bs_logged_in = False

def bs_login():
    global _bs_logged_in
    if not _bs_logged_in:
        lg = bs.login()
        if lg.error_code == "0":
            _bs_logged_in = True
            return True
        return False
    return True

def bs_logout():
    global _bs_logged_in
    if _bs_logged_in:
        bs.logout()
        _bs_logged_in = False

def get_exchange_prefix(code):
    code = str(code).strip().zfill(6)
    if code.startswith(('6', '9', '5')):
        return 'sh'
    elif code.startswith(('0', '3', '2')):
        return 'sz'
    elif code.startswith(('8', '4')):
        return 'bj'
    return 'sh'

@st.cache_data(ttl=3600)
def get_stock_list_baostock():
    if not bs_login():
        return pd.DataFrame()
    rs = bs.query_stock_basic()
    stock_list = []
    while rs.error_code == "0" and rs.next():
        row = rs.get_row_data()
        if row[4] == "1":
            code = row[0][2:].strip().zfill(6)
            stock_list.append({
                '代码': code, '名称': row[1],
                '上市日期': row[2], '交易所': row[0][:2]
            })
    bs_logout()
    return pd.DataFrame(stock_list)

@st.cache_data(ttl=1800)
def get_stock_history_baostock(code, start_date, end_date):
    if not bs_login():
        return pd.DataFrame()
    prefix = get_exchange_prefix(code)
    full_code = f"{prefix}.{code}"
    rs = bs.query_history_k_data_plus(
        full_code,
        "date,open,high,low,close,volume,amount,turn,pctChg,peTTM,pbMRQ",
        start_date=start_date, end_date=end_date,
        frequency="d", adjustflag="2"
    )
    data_list = []
    while rs.error_code == "0" and rs.next():
        data_list.append(rs.get_row_data())
    bs_logout()
    if not data_list:
        return pd.DataFrame()
    df = pd.DataFrame(data_list, columns=rs.fields)
    for col in ['open','high','low','close','volume','amount','turn','pctChg','peTTM','pbMRQ']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    df['date'] = pd.to_datetime(df['date'])
    return df

@st.cache_data(ttl=1800)
def get_stock_profit_baostock(code, year, quarter=4):
    if not bs_login():
        return None
    prefix = get_exchange_prefix(code)
    full_code = f"{prefix}.{code}"
    rs = bs.query_profit_data(code=full_code, year=year, quarter=quarter)
    data_list = []
    while rs.error_code == "0" and rs.next():
        data_list.append(rs.get_row_data())
    bs_logout()
    if not data_list:
        return None
    return pd.DataFrame(data_list, columns=rs.fields)

@st.cache_data(ttl=1800)
def get_stock_growth_baostock(code, year, quarter=4):
    if not bs_login():
        return None
    prefix = get_exchange_prefix(code)
    full_code = f"{prefix}.{code}"
    rs = bs.query_growth_data(code=full_code, year=year, quarter=quarter)
    data_list = []
    while rs.error_code == "0" and rs.next():
        data_list.append(rs.get_row_data())
    bs_logout()
    if not data_list:
        return None
    return pd.DataFrame(data_list, columns=rs.fields)

@st.cache_data(ttl=1800)
def get_stock_balance_baostock(code, year, quarter=4):
    if not bs_login():
        return None
    prefix = get_exchange_prefix(code)
    full_code = f"{prefix}.{code}"
    rs = bs.query_balance_data(code=full_code, year=year, quarter=quarter)
    data_list = []
    while rs.error_code == "0" and rs.next():
        data_list.append(rs.get_row_data())
    bs_logout()
    if not data_list:
        return None
    return pd.DataFrame(data_list, columns=rs.fields)

@st.cache_data(ttl=1800)
def get_stock_dupont_baostock(code, year, quarter=4):
    if not bs_login():
        return None
    prefix = get_exchange_prefix(code)
    full_code = f"{prefix}.{code}"
    rs = bs.query_dupont_data(code=full_code, year=year, quarter=quarter)
    data_list = []
    while rs.error_code == "0" and rs.next():
        data_list.append(rs.get_row_data())
    bs_logout()
    if not data_list:
        return None
    return pd.DataFrame(data_list, columns=rs.fields)

def safe_float(val):
    if not val or str(val).strip() == "":
        return np.nan
    try:
        return float(val)
    except (ValueError, TypeError):
        return np.nan

def get_full_financial_data(code):
    all_data = {}
    for year in range(2020, 2026):
        profit_df = get_stock_profit_baostock(code, year, 4)
        if profit_df is not None and not profit_df.empty:
            row = profit_df.iloc[-1]
            stat_date = str(row.get('statDate', ''))
            if stat_date:
                all_data.setdefault(stat_date, {})
                all_data[stat_date]['roeAvg'] = safe_float(row.get('roeAvg', 0))
                all_data[stat_date]['npMargin'] = safe_float(row.get('npMargin', 0))
                all_data[stat_date]['gpMargin'] = safe_float(row.get('gpMargin', 0))
                all_data[stat_date]['netProfit'] = safe_float(row.get('netProfit', 0))
                all_data[stat_date]['epsTTM'] = safe_float(row.get('epsTTM', 0))
                all_data[stat_date]['MBRevenue'] = safe_float(row.get('MBRevenue', 0))
        growth_df = get_stock_growth_baostock(code, year, 4)
        if growth_df is not None and not growth_df.empty:
            row = growth_df.iloc[-1]
            stat_date = str(row.get('statDate', ''))
            if stat_date:
                all_data.setdefault(stat_date, {})
                all_data[stat_date]['YOYAsset'] = safe_float(row.get('YOYAsset', 0))
                all_data[stat_date]['YOYNP'] = safe_float(row.get('YOYNP', 0))
                all_data[stat_date]['YOYEPSBasic'] = safe_float(row.get('YOYEPSBasic', 0))
                all_data[stat_date]['YOYEquity'] = safe_float(row.get('YOYEquity', 0))
        balance_df = get_stock_balance_baostock(code, year, 4)
        if balance_df is not None and not balance_df.empty:
            row = balance_df.iloc[-1]
            stat_date = str(row.get('statDate', ''))
            if stat_date:
                all_data.setdefault(stat_date, {})
                all_data[stat_date]['currentRatio'] = safe_float(row.get('currentRatio', 0))
                all_data[stat_date]['debtToAssets'] = safe_float(row.get('debtToAssets', 0))
        dupont_df = get_stock_dupont_baostock(code, year, 4)
        if dupont_df is not None and not dupont_df.empty:
            row = dupont_df.iloc[-1]
            stat_date = str(row.get('statDate', ''))
            if stat_date:
                all_data.setdefault(stat_date, {})
                all_data[stat_date]['dupontROE'] = safe_float(row.get('dupontROE', 0))
                all_data[stat_date]['assetStoEquity'] = safe_float(row.get('assetStoEquity', 0))
                all_data[stat_date]['assetTurn'] = safe_float(row.get('assetTurn', 0))
    return all_data

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="A股智能估值分析系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== 自定义CSS ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+SC:wght@400;500;700&display=swap');
    html, body, [class*="css"] { font-family: 'Noto Sans SC', sans-serif; }
    .main-header {
        font-size: 2.8rem; font-weight: 700;
        background: linear-gradient(90deg, #667eea, #764ba2, #f093fb);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        text-align: center; margin-bottom: 0.3rem;
    }
    .sub-header {
        color: #6b7280; text-align: center; font-size: 1.1rem; margin-bottom: 2rem;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 16px; padding: 1.2rem; color: white;
        text-align: center; box-shadow: 0 4px 15px rgba(102,126,234,0.3);
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-3px); }
    .metric-card-green {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        box-shadow: 0 4px 15px rgba(17,153,142,0.3);
    }
    .metric-card-orange {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        box-shadow: 0 4px 15px rgba(240,147,251,0.3);
    }
    .metric-card-blue {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        box-shadow: 0 4px 15px rgba(79,172,254,0.3);
    }
    .valuation-card {
        background: white; border-radius: 16px; padding: 1.5rem;
        border-left: 5px solid #667eea;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
        margin-bottom: 1rem; transition: all 0.3s;
    }
    .valuation-card:hover { box-shadow: 0 4px 20px rgba(0,0,0,0.12); }
    .advice-box {
        border-radius: 16px; padding: 1.5rem; margin: 1rem 0;
        text-align: center; font-size: 1.1rem;
    }
    .advice-buy { background: linear-gradient(135deg, #d4edda, #c3e6cb); border: 2px solid #28a745; }
    .advice-hold { background: linear-gradient(135deg, #fff3cd, #ffeeba); border: 2px solid #ffc107; }
    .advice-sell { background: linear-gradient(135deg, #f8d7da, #f5c6cb); border: 2px solid #dc3545; }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white; border-radius: 10px; border: none;
        padding: 0.7rem 2rem; font-weight: 600;
        box-shadow: 0 4px 15px rgba(102,126,234,0.3);
    }
    .stButton>button:hover {
        box-shadow: 0 6px 20px rgba(102,126,234,0.4);
        transform: translateY(-1px);
    }
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; font-weight: 700 !important; }
    .streamlit-expanderHeader { font-weight: 600 !important; }
</style>
""", unsafe_allow_html=True)

# ==================== 估值计算 ====================

def calculate_valuation(eps, bvps, roe, net_profit, revenue,
                        pe_low, pe_mid, pe_high,
                        growth_rate, discount_rate, terminal_growth):
    pe_vals = {
        '悲观': eps * pe_low if eps > 0 else 0,
        '基准': eps * pe_mid if eps > 0 else 0,
        '乐观': eps * pe_high if eps > 0 else 0
    }
    if net_profit > 0 and revenue > 0:
        fcf = net_profit * 0.7
        growth_rates = [growth_rate] * 5 + [growth_rate * 0.5] * 5
        pv_fcf = 0
        current_fcf = fcf
        for year in range(1, 11):
            g = growth_rates[min(year-1, len(growth_rates)-1)]
            current_fcf *= (1 + g)
            pv_fcf += current_fcf / ((1 + discount_rate) ** year)
        terminal = current_fcf * (1 + terminal_growth) / (discount_rate - terminal_growth)
        pv_terminal = terminal / ((1 + discount_rate) ** 10)
        total_shares = max(revenue / (eps * 10000), 1) if eps > 0 else 1e8
        dcf_value = (pv_fcf + pv_terminal) / total_shares
    else:
        dcf_value = 0
    pb_value = bvps * 2.5 if bvps > 0 else 0
    fair_price = pe_vals['基准'] * 0.4 + dcf_value * 0.3 + pb_value * 0.3
    return {'pe': pe_vals, 'dcf': dcf_value, 'pb': pb_value, 'fair': fair_price}

def calculate_financial_health(fin_data):
    """计算财务健康评分 (0-100)"""
    if not fin_data:
        return 50, {}
    dates = sorted(fin_data.keys(), reverse=True)
    latest = fin_data.get(dates[0], {})
    scores = {}
    # ROE评分
    roe = latest.get('roeAvg', 0)
    scores['ROE'] = min(roe * 2, 30) if roe > 0 else 0
    # 毛利率评分
    gp = latest.get('gpMargin', 0)
    scores['毛利率'] = min(gp * 0.5, 20) if gp > 0 else 0
    # 净利率评分
    np = latest.get('npMargin', 0)
    scores['净利率'] = min(np * 0.8, 15) if np > 0 else 0
    # 成长性评分
    yoy_np = latest.get('YOYNP', 0)
    scores['成长性'] = min(max(yoy_np, 0) * 0.5, 15) if yoy_np > 0 else 0
    # 偿债能力评分
    debt = latest.get('debtToAssets', 100)
    scores['偿债能力'] = max(20 - debt * 0.2, 0)
    total = sum(scores.values())
    return min(total, 100), scores

def get_investment_advice(current_price, fair_price, pe_ttm, pb_mrq, health_score):
    """生成投资建议"""
    upside = ((fair_price / current_price - 1) * 100) if current_price > 0 else 0
    reasons = []
    if upside > 30:
        rating = "强烈买入"
        color = "buy"
        reasons.append(f"股价低于合理估值 {abs(upside):.1f}%，存在显著安全边际")
    elif upside > 10:
        rating = "买入"
        color = "buy"
        reasons.append(f"股价低于合理估值 {abs(upside):.1f}%，具备投资价值")
    elif upside > -10:
        rating = "持有"
        color = "hold"
        reasons.append("股价处于合理估值区间，建议观望")
    elif upside > -30:
        rating = "减持"
        color = "sell"
        reasons.append(f"股价高于合理估值 {abs(upside):.1f}%，估值偏高")
    else:
        rating = "卖出"
        color = "sell"
        reasons.append(f"股价大幅高于合理估值 {abs(upside):.1f}%，存在回调风险")
    if pe_ttm > 0:
        if pe_ttm < 15:
            reasons.append(f"PE({pe_ttm:.1f})处于历史低位，估值吸引力强")
        elif pe_ttm > 50:
            reasons.append(f"PE({pe_ttm:.1f})偏高，需警惕估值泡沫")
    if pb_mrq > 0:
        if pb_mrq < 1.5:
            reasons.append(f"PB({pb_mrq:.1f})较低，资产价值被低估")
        elif pb_mrq > 5:
            reasons.append(f"PB({pb_mrq:.1f})较高，资产溢价明显")
    if health_score >= 80:
        reasons.append(f"财务健康评分{health_score:.0f}分，基本面优秀")
    elif health_score < 50:
        reasons.append(f"财务健康评分{health_score:.0f}分，需关注财务风险")
    return rating, color, upside, reasons

# ==================== 可视化 ====================

def create_kline_chart(df_history, fair_price=None):
    if df_history.empty or len(df_history) < 20:
        return None
    fig = make_subplots(
        rows=3, cols=1, shared_xaxes=True,
        vertical_spacing=0.02,
        row_heights=[0.55, 0.25, 0.20],
        subplot_titles=('', '', '')
    )
    # K线
    fig.add_trace(go.Candlestick(
        x=df_history['date'], open=df_history['open'], high=df_history['high'],
        low=df_history['low'], close=df_history['close'],
        name='K线', increasing_line_color='#ff4757', decreasing_line_color='#2ed573',
        increasing_fillcolor='#ff4757', decreasing_fillcolor='#2ed573'
    ), row=1, col=1)
    # MA
    ma5 = df_history['close'].rolling(5).mean()
    ma20 = df_history['close'].rolling(20).mean()
    ma60 = df_history['close'].rolling(60).mean()
    fig.add_trace(go.Scatter(x=df_history['date'], y=ma5, name='MA5',
                              line=dict(color='#ffa502', width=1.2)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_history['date'], y=ma20, name='MA20',
                              line=dict(color='#3742fa', width=1.2)), row=1, col=1)
    fig.add_trace(go.Scatter(x=df_history['date'], y=ma60, name='MA60',
                              line=dict(color='#ff6348', width=1.2)), row=1, col=1)
    # 估值线
    if fair_price and fair_price > 0:
        fig.add_hline(y=fair_price, line_dash="dash", line_color="#e84393",
                      annotation_text=f"合理估值 ¥{fair_price:.2f}",
                      annotation_position="top right", row=1, col=1)
    # 成交量
    vol_colors = ['#ff4757' if c >= o else '#2ed573'
                  for c, o in zip(df_history['close'], df_history['open'])]
    fig.add_trace(go.Bar(x=df_history['date'], y=df_history['volume'],
                         marker_color=vol_colors, name='成交量', showlegend=False), row=2, col=1)
    # MACD
    ema12 = df_history['close'].ewm(span=12).mean()
    ema26 = df_history['close'].ewm(span=26).mean()
    dif = ema12 - ema26
    dea = dif.ewm(span=9).mean()
    macd = (dif - dea) * 2
    fig.add_trace(go.Bar(x=df_history['date'], y=macd,
                         marker_color=['#ff4757' if v >= 0 else '#2ed573' for v in macd],
                         name='MACD', showlegend=False), row=3, col=1)
    fig.add_trace(go.Scatter(x=df_history['date'], y=dif, name='DIF',
                              line=dict(color='#3742fa', width=1)), row=3, col=1)
    fig.add_trace(go.Scatter(x=df_history['date'], y=dea, name='DEA',
                              line=dict(color='#ffa502', width=1)), row=3, col=1)
    fig.update_layout(
        title=dict(text='📈 K线走势 + 成交量 + MACD', font=dict(size=18, color='#2f3542')),
        xaxis_rangeslider_visible=False, height=700,
        template='plotly_white', hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    fig.update_xaxes(showgrid=True, gridwidth=0.5, gridcolor='rgba(0,0,0,0.05)')
    fig.update_yaxes(showgrid=True, gridwidth=0.5, gridcolor='rgba(0,0,0,0.05)')
    return fig

def create_pe_bands_chart(df_history, eps):
    """PE估值带图"""
    if df_history.empty or eps <= 0:
        return None
    df = df_history.copy()
    df['pe_low'] = eps * 10
    df['pe_mid'] = eps * 20
    df['pe_high'] = eps * 35
    df['pe_extreme'] = eps * 50
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df['date'], y=df['pe_extreme'], name='极端高估(50x)',
                              line=dict(color='rgba(255,71,87,0.3)', width=0), fill='tonexty',
                              fillcolor='rgba(255,71,87,0.15)'))
    fig.add_trace(go.Scatter(x=df['date'], y=df['pe_high'], name='高估区间(35x)',
                              line=dict(color='rgba(255,99,72,0.4)', width=0), fill='tonexty',
                              fillcolor='rgba(255,99,72,0.15)'))
    fig.add_trace(go.Scatter(x=df['date'], y=df['pe_mid'], name='合理区间(20x)',
                              line=dict(color='rgba(55,66,250,0.5)', width=0), fill='tonexty',
                              fillcolor='rgba(55,66,250,0.1)'))
    fig.add_trace(go.Scatter(x=df['date'], y=df['pe_low'], name='低估区间(10x)',
                              line=dict(color='rgba(46,213,115,0.5)', width=0), fill='tonexty',
                              fillcolor='rgba(46,213,115,0.15)'))
    fig.add_trace(go.Scatter(x=df['date'], y=df['close'], name='收盘价',
                              line=dict(color='#2f3542', width=2)))
    fig.update_layout(
        title='📊 PE估值带分析', height=400, template='plotly_white',
        hovermode='x unified',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1)
    )
    return fig

def create_valuation_waterfall(current_price, pe_vals, dcf_val, pb_val, fair_price):
    """估值对比瀑布图"""
    fig = go.Figure(go.Waterfall(
        name="估值对比",
        orientation="v",
        x=['当前股价', 'PE悲观', 'PE基准', 'PE乐观', 'DCF估值', 'PB估值', '综合估值'],
        y=[current_price,
           pe_vals['悲观'] - current_price,
           pe_vals['基准'] - pe_vals['悲观'],
           pe_vals['乐观'] - pe_vals['基准'],
           dcf_val - pe_vals['乐观'],
           pb_val - dcf_val,
           fair_price - pb_val],
        measure=['absolute', 'relative', 'relative', 'relative', 'relative', 'relative', 'total'],
        text=[f'¥{current_price:.2f}', f'¥{pe_vals["悲观"]:.2f}', f'¥{pe_vals["基准"]:.2f}',
              f'¥{pe_vals["乐观"]:.2f}', f'¥{dcf_val:.2f}', f'¥{pb_val:.2f}', f'¥{fair_price:.2f}'],
        textposition='outside',
        connector={'line': {'color': 'rgb(63, 63, 63)'}},
        decreasing={'marker': {'color': '#2ed573'}},
        increasing={'marker': {'color': '#ff4757'}},
        totals={'marker': {'color': '#3742fa'}}
    ))
    fig.update_layout(title='💰 估值方法对比瀑布图', height=420, template='plotly_white')
    return fig

def create_radar_chart(fin_data):
    """财务能力雷达图"""
    if not fin_data:
        return None
    dates = sorted(fin_data.keys(), reverse=True)
    latest = fin_data.get(dates[0], {})
    prev = fin_data.get(dates[1] if len(dates) > 1 else dates[0], {})
    categories = ['ROE', '毛利率', '净利率', '资产增速', '利润增速', '偿债能力']
    # 归一化到0-100
    def norm(val, max_val=100):
        return min(max(val, 0), max_val) / max_val * 100 if val and val == val else 30
    curr_vals = [
        norm(latest.get('roeAvg', 0), 30),
        norm(latest.get('gpMargin', 0), 60),
        norm(latest.get('npMargin', 0), 40),
        norm(latest.get('YOYAsset', 0), 50),
        norm(latest.get('YOYNP', 0), 50),
        max(0, 100 - latest.get('debtToAssets', 100) * 0.8)
    ]
    prev_vals = [
        norm(prev.get('roeAvg', 0), 30),
        norm(prev.get('gpMargin', 0), 60),
        norm(prev.get('npMargin', 0), 40),
        norm(prev.get('YOYAsset', 0), 50),
        norm(prev.get('YOYNP', 0), 50),
        max(0, 100 - prev.get('debtToAssets', 100) * 0.8)
    ]
    fig = go.Figure()
    fig.add_trace(go.Scatterpolar(
        r=curr_vals + [curr_vals[0]], theta=categories + [categories[0]],
        fill='toself', name='最新年度', fillcolor='rgba(102,126,234,0.25)',
        line=dict(color='#667eea', width=2)
    ))
    fig.add_trace(go.Scatterpolar(
        r=prev_vals + [prev_vals[0]], theta=categories + [categories[0]],
        fill='toself', name='上一年度', fillcolor='rgba(240,147,251,0.15)',
        line=dict(color='#f093fb', width=2, dash='dash')
    ))
    fig.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
        title='🎯 财务能力雷达图', height=420, template='plotly_white',
        legend=dict(orientation='h', yanchor='bottom', y=-0.15, xanchor='center', x=0.5)
    )
    return fig

def create_health_gauge(score):
    """财务健康评分仪表盘"""
    color = '#2ed573' if score >= 80 else '#ffa502' if score >= 60 else '#ff4757'
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=score,
        number={'suffix': '分', 'font': {'size': 36, 'color': color}},
        title={'text': "财务健康评分", 'font': {'size': 16}},
        gauge={
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': color, 'thickness': 0.7},
            'bgcolor': 'white',
            'steps': [
                {'range': [0, 40], 'color': '#ff6b6b20'},
                {'range': [40, 60], 'color': '#feca5720'},
                {'range': [60, 80], 'color': '#48dbfb20'},
                {'range': [80, 100], 'color': '#1dd1a120'}
            ],
            'threshold': {'line': {'color': color, 'width': 3}, 'thickness': 0.8, 'value': score}
        }
    ))
    fig.update_layout(height=300, template='plotly_white')
    return fig

def create_financial_trend(fin_data):
    if not fin_data:
        return None
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=['💰 净利润趋势', '📈 ROE趋势', '📊 毛利率趋势', '📉 净利率趋势'],
        vertical_spacing=0.12, horizontal_spacing=0.08
    )
    metrics = [
        ('netProfit', '净利润(亿)', (1,1), '#667eea'),
        ('roeAvg', 'ROE(%)', (1,2), '#f5576c'),
        ('gpMargin', '毛利率(%)', (2,1), '#11998e'),
        ('npMargin', '净利率(%)', (2,2), '#4facfe')
    ]
    dates = sorted(fin_data.keys())
    date_labels = [d[:4] for d in dates]
    for key, label, (r, c), color in metrics:
        values = [fin_data.get(d, {}).get(key, np.nan) for d in dates]
        if key == 'netProfit':
            values = [v/1e8 if v and v == v else np.nan for v in values]
        fig.add_trace(go.Scatter(
            x=date_labels, y=values, mode='lines+markers',
            name=label, line=dict(width=3, color=color),
            marker=dict(size=10, color=color, line=dict(width=2, color='white')),
            fill='tozeroy', fillcolor=color.replace(')', ',0.1)').replace('rgb', 'rgba')
        ), row=r, col=c)
    fig.update_layout(height=500, template='plotly_white', showlegend=False)
    fig.update_xaxes(showgrid=True, gridwidth=0.5, gridcolor='rgba(0,0,0,0.05)')
    fig.update_yaxes(showgrid=True, gridwidth=0.5, gridcolor='rgba(0,0,0,0.05)')
    return fig

def create_growth_chart(fin_data):
    """成长性分析图"""
    if not fin_data:
        return None
    dates = sorted(fin_data.keys())
    date_labels = [d[:4] for d in dates]
    fig = go.Figure()
    metrics = [
        ('YOYNP', '净利润增速(%)', '#667eea'),
        ('YOYEPSBasic', 'EPS增速(%)', '#f5576c'),
        ('YOYAsset', '资产增速(%)', '#11998e'),
        ('YOYEquity', '净资产增速(%)', '#4facfe')
    ]
    for key, label, color in metrics:
        values = [fin_data.get(d, {}).get(key, np.nan) for d in dates]
        fig.add_trace(go.Bar(name=label, x=date_labels, y=values, marker_color=color))
    fig.update_layout(
        title='🚀 成长性分析', barmode='group', height=400,
        template='plotly_white', legend=dict(orientation='h', yanchor='bottom', y=1.02)
    )
    return fig

# ==================== 报告生成 ====================

def generate_word_report(code, name, price, valuation, fin_data, history_df, advice, health_score, health_details):
    doc = Document()
    title = doc.add_heading(f'{name} ({code}) 投资价值深度分析报告', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f'报告生成日期: {datetime.now().strftime("%Y年%m月%d日")}').alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    # 基本信息
    doc.add_heading('一、基本信息', level=1)
    doc.add_paragraph(f'当前股价: ¥{price:.2f}')
    if not history_df.empty:
        doc.add_paragraph(f'最新PE-TTM: {history_df["peTTM"].iloc[-1]:.2f}')
        doc.add_paragraph(f'最新PB-MRQ: {history_df["pbMRQ"].iloc[-1]:.2f}')
    doc.add_paragraph()
    # 投资建议
    doc.add_heading('二、投资建议', level=1)
    p = doc.add_paragraph()
    p.add_run(f'投资评级: {advice[0]}').bold = True
    p.add_run(f'  (溢价空间: {advice[2]:+.1f}%)')
    doc.add_heading('评级理由:', level=2)
    for reason in advice[3]:
        doc.add_paragraph(reason, style='List Bullet')
    doc.add_paragraph()
    # 估值分析
    doc.add_heading('三、估值分析', level=1)
    pe = valuation['pe']
    doc.add_paragraph(f'综合合理估值: ¥{valuation["fair"]:.2f}')
    doc.add_paragraph(f'PE悲观估值: ¥{pe["悲观"]:.2f} (PE={valuation.get("pe_low",10)}x)')
    doc.add_paragraph(f'PE基准估值: ¥{pe["基准"]:.2f} (PE={valuation.get("pe_mid",20)}x)')
    doc.add_paragraph(f'PE乐观估值: ¥{pe["乐观"]:.2f} (PE={valuation.get("pe_high",35)}x)')
    doc.add_paragraph(f'DCF估值: ¥{valuation["dcf"]:.2f}')
    doc.add_paragraph(f'PB估值: ¥{valuation["pb"]:.2f}')
    doc.add_paragraph()
    # 财务健康
    doc.add_heading('四、财务健康评分', level=1)
    doc.add_paragraph(f'综合评分: {health_score:.0f}/100')
    for k, v in health_details.items():
        doc.add_paragraph(f'{k}: {v:.1f}分', style='List Bullet')
    doc.add_paragraph()
    # 财务摘要
    if fin_data:
        doc.add_heading('五、最新财务指标', level=1)
        dates = sorted(fin_data.keys(), reverse=True)
        latest = fin_data[dates[0]]
        doc.add_paragraph(f'EPS(TTM): ¥{latest.get("epsTTM",0):.2f}')
        doc.add_paragraph(f'ROE: {latest.get("roeAvg",0):.2f}%')
        doc.add_paragraph(f'毛利率: {latest.get("gpMargin",0):.2f}%')
        doc.add_paragraph(f'净利率: {latest.get("npMargin",0):.2f}%')
        doc.add_paragraph(f'资产负债率: {latest.get("debtToAssets",0):.2f}%')
    doc.add_paragraph()
    doc.add_paragraph('免责声明: 本报告基于公开财务数据生成，仅供参考研究使用，不构成任何投资建议。投资有风险，入市需谨慎。')
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

def generate_pdf_report(code, name, price, valuation, advice, health_score):
    pdf_io = io.BytesIO()
    doc = SimpleDocTemplate(pdf_io, pagesize=A4, topMargin=50, bottomMargin=50)
    styles = getSampleStyleSheet()
    story = []
    ts = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=22,
                        textColor=rl_colors.HexColor('#667eea'), alignment=1, spaceAfter=15)
    story.append(Paragraph(f'{name} ({code})', ts))
    story.append(Paragraph('投资价值深度分析报告', ts))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f'<b>报告日期:</b> {datetime.now().strftime("%Y-%m-%d")}', styles['Normal']))
    story.append(Paragraph(f'<b>当前股价:</b> ¥{price:.2f}', styles['Normal']))
    story.append(Spacer(1, 15))
    # 投资建议
    story.append(Paragraph('投资建议', styles['Heading2']))
    rating_color = '#28a745' if advice[1] == 'buy' else '#ffc107' if advice[1] == 'hold' else '#dc3545'
    story.append(Paragraph(f'<b>投资评级:</b> <font color="{rating_color}">{advice[0]}</font>', styles['Normal']))
    story.append(Paragraph(f'<b>溢价空间:</b> {advice[2]:+.1f}%', styles['Normal']))
    story.append(Spacer(1, 10))
    # 估值表
    story.append(Paragraph('估值详情', styles['Heading2']))
    pe = valuation['pe']
    td = [['估值方法', '悲观', '基准', '乐观', '综合']]
    td.append(['PE估值', f'¥{pe["悲观"]:.2f}', f'¥{pe["基准"]:.2f}', f'¥{pe["乐观"]:.2f}', '-'])
    td.append(['DCF估值', '-', f'¥{valuation["dcf"]:.2f}', '-', '-'])
    td.append(['PB估值', '-', f'¥{valuation["pb"]:.2f}', '-', '-'])
    td.append(['综合估值', '-', '-', '-', f'¥{valuation["fair"]:.2f}'])
    table = Table(td)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), rl_colors.HexColor('#667eea')),
        ('TEXTCOLOR', (0, 0), (-1, 0), rl_colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, rl_colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor('#f8fafc')]),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
    ]))
    story.append(table)
    story.append(Spacer(1, 15))
    story.append(Paragraph(f'<b>财务健康评分:</b> {health_score:.0f}/100', styles['Normal']))
    story.append(Spacer(1, 20))
    story.append(Paragraph('免责声明: 本报告基于公开数据生成，仅供参考，不构成投资建议。', styles['Normal']))
    doc.build(story)
    pdf_io.seek(0)
    return pdf_io

# ==================== 主应用 ====================

def main():
    # 侧边栏
    with st.sidebar:
        st.markdown("## 🔍 股票搜索")
        search_term = st.text_input("输入代码或名称", placeholder="如: 600519 或 贵州茅台")
        selected_stock = None
        if search_term:
            stock_list = get_stock_list_baostock()
            if not stock_list.empty:
                matched = stock_list[
                    stock_list['代码'].str.contains(search_term, na=False) |
                    stock_list['名称'].str.contains(search_term, na=False)
                ]
                if not matched.empty:
                    selected_stock = matched.iloc[0]
                    st.success(f"✅ {selected_stock['名称']} ({selected_stock['代码']})")
        st.markdown("---")
        st.markdown("## ⚙️ 估值参数设置")
        with st.expander("📊 PE市盈率参数", expanded=True):
            pe_low = st.slider("悲观PE", 5, 30, 10, help="保守估值对应的PE倍数")
            pe_mid = st.slider("基准PE", 10, 50, 20, help="合理估值对应的PE倍数")
            pe_high = st.slider("乐观PE", 15, 80, 35, help="乐观估值对应的PE倍数")
        with st.expander("💰 DCF现金流折现参数", expanded=False):
            growth = st.slider("预期增长率(%)", -10, 50, 15, help="未来10年净利润预期年化增长率") / 100
            discount = st.slider("折现率(%)", 5, 20, 10, help="投资者要求的年化回报率") / 100
            terminal = st.slider("永续增长率(%)", 0, 8, 3, help="长期稳定增长率") / 100
        with st.expander("📈 其他参数", expanded=False):
            pb_multiple = st.slider("PB倍数", 1.0, 5.0, 2.5, step=0.1, help="市净率估值倍数")
        st.markdown("---")
        st.markdown("<small>数据源自 Baostock 开源证券数据平台</small>", unsafe_allow_html=True)

    # 主页面头部
    st.markdown('<div class="main-header">📊 A股智能估值分析系统</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">基于真实财务数据 · 多维估值模型 · 智能投资建议</div>', unsafe_allow_html=True)

    if selected_stock is None:
        # 欢迎页
        stock_list = get_stock_list_baostock()
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div style="font-size:0.9rem;opacity:0.9">A股上市公司</div>
                <div style="font-size:2rem;font-weight:700">{len(stock_list):,}</div>
            </div>""", unsafe_allow_html=True)
        with col2:
            st.markdown(f"""
            <div class="metric-card metric-card-green">
                <div style="font-size:0.9rem;opacity:0.9">估值模型</div>
                <div style="font-size:2rem;font-weight:700">PE+DCF+PB</div>
            </div>""", unsafe_allow_html=True)
        with col3:
            st.markdown(f"""
            <div class="metric-card metric-card-orange">
                <div style="font-size:0.9rem;opacity:0.9">可视化图表</div>
                <div style="font-size:2rem;font-weight:700">8+</div>
            </div>""", unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
            <div class="metric-card metric-card-blue">
                <div style="font-size:0.9rem;opacity:0.9">数据源</div>
                <div style="font-size:2rem;font-weight:700">Baostock</div>
            </div>""", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### 🚀 使用指南")
        st.info("在左侧搜索框输入股票代码(如600519)或名称(如贵州茅台)，系统将自动获取真实K线数据、财务报表，并生成多维估值分析和投资建议。")
        if not stock_list.empty:
            st.markdown("### 🔥 热门股票")
            hot_stocks = stock_list[stock_list['代码'].isin(['600519','600036','000858','002594','601318','600900','000001','601012'])]
            st.dataframe(hot_stocks if not hot_stocks.empty else stock_list.head(20),
                         use_container_width=True, hide_index=True, height=300)
        return

    # 获取数据
    symbol = str(selected_stock['代码']).strip().zfill(6)
    stock_name = selected_stock['名称']
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=500)).strftime('%Y-%m-%d')

    with st.spinner(f'🔄 正在获取 {stock_name}({symbol}) 的实时数据...'):
        history_df = get_stock_history_baostock(symbol, start_date, end_date)
        fin_data = get_full_financial_data(symbol)

    # 提取价格数据
    if not history_df.empty:
        current_price = float(history_df['close'].iloc[-1])
        latest_pe = float(history_df['peTTM'].iloc[-1]) if pd.notna(history_df['peTTM'].iloc[-1]) else 0
        latest_pb = float(history_df['pbMRQ'].iloc[-1]) if pd.notna(history_df['pbMRQ'].iloc[-1]) else 0
        pct_chg = float(history_df['pctChg'].iloc[-1]) if pd.notna(history_df['pctChg'].iloc[-1]) else 0
        high_52w = history_df['high'].max() if 'high' in history_df.columns else 0
        low_52w = history_df['low'].min() if 'low' in history_df.columns else 0
    else:
        current_price = latest_pe = latest_pb = pct_chg = high_52w = low_52w = 0

    # 提取财务数据
    dates = sorted(fin_data.keys(), reverse=True)
    latest_fin = fin_data.get(dates[0], {}) if dates else {}
    eps = latest_fin.get('epsTTM', 0)
    roe = latest_fin.get('roeAvg', 0)
    net_profit = latest_fin.get('netProfit', 0)
    revenue = latest_fin.get('MBRevenue', 0)
    gp_margin = latest_fin.get('gpMargin', 0)
    np_margin = latest_fin.get('npMargin', 0)
    debt_ratio = latest_fin.get('debtToAssets', 0)
    bvps = current_price / latest_pb if latest_pb > 0 else 0

    # 计算估值
    valuation = calculate_valuation(
        eps, bvps, roe, net_profit, revenue,
        pe_low, pe_mid, pe_high, growth, discount, terminal
    )
    valuation['pe_low'] = pe_low
    valuation['pe_mid'] = pe_mid
    valuation['pe_high'] = pe_high
    fair_price = valuation['fair']

    # 财务健康评分
    health_score, health_details = calculate_financial_health(fin_data)

    # 投资建议
    advice = get_investment_advice(current_price, fair_price, latest_pe, latest_pb, health_score)

    # ========== 股票标题与核心指标 ==========
    st.markdown(f"""
    <div style="text-align:center; margin: 1rem 0 2rem 0;">
        <h1 style="margin:0; color:#2f3542;">【{symbol}】{stock_name}</h1>
        <p style="color:#6b7280; margin:0.3rem 0;">{selected_stock.get('交易所','').upper()}交易所 | 上市日期: {selected_stock.get('上市日期','')}</p>
    </div>
    """, unsafe_allow_html=True)

    # 核心指标卡片
    cols = st.columns(6)
    metrics_data = [
        ("当前股价", f"¥{current_price:.2f}", f"{pct_chg:+.2f}%", pct_chg >= 0),
        ("市盈率PE", f"{latest_pe:.1f}x" if latest_pe > 0 else "N/A", "TTM", None),
        ("市净率PB", f"{latest_pb:.1f}x" if latest_pb > 0 else "N/A", "MRQ", None),
        ("52周最高", f"¥{high_52w:.2f}", "", None),
        ("52周最低", f"¥{low_52w:.2f}", "", None),
        ("财务评分", f"{health_score:.0f}", "满分100", None),
    ]
    colors_list = ['#667eea', '#11998e', '#f5576c', '#4facfe', '#ffa502', '#e84393']
    for col, (label, value, delta, is_up), color in zip(cols, metrics_data, colors_list):
        with col:
            if is_up is not None:
                delta_color = "normal" if is_up else "inverse"
            else:
                delta_color = "off"
            st.metric(label, value, delta, delta_color=delta_color)

    # ========== 投资建议 ==========
    st.markdown("---")
    st.markdown("### 🎯 智能投资建议")
    advice_class = f"advice-{advice[1]}"
    st.markdown(f"""
    <div class="advice-box {advice_class}">
        <h2 style="margin:0 0 0.5rem 0;">投资评级: <b>{advice[0]}</b></h2>
        <p style="margin:0.3rem 0;">合理估值: <b>¥{fair_price:.2f}</b> | 当前股价: <b>¥{current_price:.2f}</b> | 溢价空间: <b>{advice[2]:+.1f}%</b></p>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("**评级理由:**")
    for reason in advice[3]:
        st.markdown(f"- {reason}")

    # ========== 估值分析 ==========
    st.markdown("---")
    st.markdown("### 💰 多维估值分析")
    col_v1, col_v2 = st.columns([3, 2])
    with col_v1:
        pe_vals = valuation['pe']
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"""
            <div class="valuation-card" style="border-left-color:#2ed573;">
                <div style="color:#2ed573; font-size:0.9rem; font-weight:600;">🟢 悲观估值</div>
                <div style="font-size:1.8rem; font-weight:700; margin:0.3rem 0;">¥{pe_vals['悲观']:.2f}</div>
                <div style="color:#6b7280; font-size:0.8rem;">PE = {pe_low}x</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            st.markdown(f"""
            <div class="valuation-card" style="border-left-color:#3742fa;">
                <div style="color:#3742fa; font-size:0.9rem; font-weight:600;">🔵 基准估值</div>
                <div style="font-size:1.8rem; font-weight:700; margin:0.3rem 0;">¥{pe_vals['基准']:.2f}</div>
                <div style="color:#6b7280; font-size:0.8rem;">PE = {pe_mid}x</div>
            </div>""", unsafe_allow_html=True)
        with c3:
            st.markdown(f"""
            <div class="valuation-card" style="border-left-color:#ff4757;">
                <div style="color:#ff4757; font-size:0.9rem; font-weight:600;">🔴 乐观估值</div>
                <div style="font-size:1.8rem; font-weight:700; margin:0.3rem 0;">¥{pe_vals['乐观']:.2f}</div>
                <div style="color:#6b7280; font-size:0.8rem;">PE = {pe_high}x</div>
            </div>""", unsafe_allow_html=True)
        # 瀑布图
        fig_waterfall = create_valuation_waterfall(current_price, pe_vals, valuation['dcf'], valuation['pb'], fair_price)
        if fig_waterfall:
            st.plotly_chart(fig_waterfall, use_container_width=True, key="waterfall")
    with col_v2:
        # 估值仪表盘
        upside = advice[2]
        gauge_color = '#2ed573' if upside > 10 else '#ffa502' if upside > -10 else '#ff4757'
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=fair_price,
            delta={'reference': current_price, 'valueformat': '.2f'},
            number={'prefix': '¥', 'font': {'size': 28}},
            title={'text': f"合理估值 (溢价 {upside:+.1f}%)", 'font': {'size': 14}},
            gauge={
                'axis': {'range': [0, max(current_price, fair_price) * 1.5]},
                'bar': {'color': gauge_color, 'thickness': 0.65},
                'steps': [
                    {'range': [0, current_price * 0.7], 'color': '#2ed57320'},
                    {'range': [current_price * 0.7, current_price * 1.3], 'color': '#ffa50220'},
                    {'range': [current_price * 1.3, max(current_price, fair_price) * 1.5], 'color': '#ff475720'}
                ],
                'threshold': {'line': {'color': '#e84393', 'width': 3}, 'thickness': 0.8, 'value': current_price}
            }
        ))
        fig_gauge.update_layout(height=320, template='plotly_white')
        st.plotly_chart(fig_gauge, use_container_width=True, key="gauge")
        # DCF/PB详情
        st.markdown(f"""
        <div class="valuation-card" style="border-left-color:#e84393;">
            <div style="color:#e84393; font-size:0.85rem; font-weight:600;">💰 DCF现金流折现估值</div>
            <div style="font-size:1.4rem; font-weight:700; margin:0.2rem 0;">¥{valuation['dcf']:.2f}</div>
            <div style="color:#6b7280; font-size:0.75rem;">增长率{growth*100:.0f}% | 折现率{discount*100:.0f}%</div>
        </div>""", unsafe_allow_html=True)
        st.markdown(f"""
        <div class="valuation-card" style="border-left-color:#8e44ad;">
            <div style="color:#8e44ad; font-size:0.85rem; font-weight:600;">📊 PB市净率估值</div>
            <div style="font-size:1.4rem; font-weight:700; margin:0.2rem 0;">¥{valuation['pb']:.2f}</div>
            <div style="color:#6b7280; font-size:0.75rem;">BVPS ¥{bvps:.2f} x {pb_multiple}x</div>
        </div>""", unsafe_allow_html=True)

    # ========== 图表区域 ==========
    st.markdown("---")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown("### 📈 K线走势 + 技术指标")
        fig_kline = create_kline_chart(history_df, fair_price)
        if fig_kline:
            st.plotly_chart(fig_kline, use_container_width=True, key="kline")
        else:
            st.warning("暂无K线数据")
    with col_c2:
        st.markdown("### 📊 PE估值带")
        fig_pe = create_pe_bands_chart(history_df, eps)
        if fig_pe:
            st.plotly_chart(fig_pe, use_container_width=True, key="pe_bands")
        else:
            st.info("暂无EPS数据，无法生成PE估值带")

    # ========== 财务分析 ==========
    st.markdown("---")
    st.markdown("### 📋 财务深度分析")
    col_f1, col_f2 = st.columns([3, 2])
    with col_f1:
        fig_trend = create_financial_trend(fin_data)
        if fig_trend:
            st.plotly_chart(fig_trend, use_container_width=True, key="trend")
        else:
            st.info("暂无财务趋势数据")
    with col_f2:
        # 财务健康评分仪表盘
        fig_health = create_health_gauge(health_score)
        st.plotly_chart(fig_health, use_container_width=True, key="health")
        # 评分详情
        st.markdown("**评分详情:**")
        for k, v in health_details.items():
            bar_color = '#2ed573' if v >= 20 else '#ffa502' if v >= 10 else '#ff4757'
            pct = min(v / 30 * 100, 100) if k == 'ROE' else min(v / 20 * 100, 100)
            st.markdown(f"""
            <div style="margin:0.3rem 0;">
                <div style="display:flex; justify-content:space-between; font-size:0.85rem;">
                    <span>{k}</span><span>{v:.1f}分</span>
                </div>
                <div style="background:#f1f2f6; border-radius:4px; height:8px;">
                    <div style="background:{bar_color}; width:{pct}%; height:8px; border-radius:4px; transition:width 0.5s;"></div>
                </div>
            </div>""", unsafe_allow_html=True)

    # 成长性 + 雷达图
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown("### 🚀 成长性分析")
        fig_growth = create_growth_chart(fin_data)
        if fig_growth:
            st.plotly_chart(fig_growth, use_container_width=True, key="growth")
        else:
            st.info("暂无成长性数据")
    with col_g2:
        st.markdown("### 🎯 财务能力雷达图")
        fig_radar = create_radar_chart(fin_data)
        if fig_radar:
            st.plotly_chart(fig_radar, use_container_width=True, key="radar")
        else:
            st.info("暂无雷达图数据")

    # ========== 财务摘要卡片 ==========
    st.markdown("---")
    st.markdown("### 📊 最新财务指标")
    fcols = st.columns(6)
    fin_metrics = [
        ("EPS(TTM)", f"¥{eps:.2f}", '#667eea'),
        ("ROE", f"{roe:.2f}%", '#f5576c'),
        ("毛利率", f"{gp_margin:.2f}%", '#11998e'),
        ("净利率", f"{np_margin:.2f}%", '#4facfe'),
        ("资产负债率", f"{debt_ratio:.2f}%", '#ffa502'),
        ("净利润", f"{net_profit/1e8:.1f}亿" if net_profit > 0 else "N/A", '#e84393'),
    ]
    for col, (label, value, color) in zip(fcols, fin_metrics):
        with col:
            st.markdown(f"""
            <div style="background:linear-gradient(135deg, {color}20, {color}08); border-radius:12px; padding:1rem; text-align:center; border:1px solid {color}30;">
                <div style="color:{color}; font-size:0.8rem; font-weight:600;">{label}</div>
                <div style="color:#2f3542; font-size:1.4rem; font-weight:700; margin-top:0.3rem;">{value}</div>
            </div>""", unsafe_allow_html=True)

    # ========== 报告下载 ==========
    st.markdown("---")
    st.markdown("### 📥 下载专业分析报告")
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        word_file = generate_word_report(symbol, stock_name, current_price, valuation, fin_data, history_df, advice, health_score, health_details)
        st.download_button(
            label="📄 下载 Word 报告", data=word_file,
            file_name=f"{stock_name}_{symbol}_估值分析报告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    with col_dl2:
        pdf_file = generate_pdf_report(symbol, stock_name, current_price, valuation, advice, health_score)
        st.download_button(
            label="📕 下载 PDF 报告", data=pdf_file,
            file_name=f"{stock_name}_{symbol}_估值分析报告.pdf",
            mime="application/pdf", use_container_width=True
        )

    # 底部免责声明
    st.markdown("---")
    st.caption("⚠️ 免责声明：本系统基于Baostock公开财务数据进行估值分析，所有结果仅供参考研究使用，不构成任何投资建议。股市有风险，投资需谨慎。")

if __name__ == "__main__":
    main()
