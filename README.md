A股智能估值分析系统 v4.0
基于Baostock真实数据的A股股票智能估值分析工具，部署在Streamlit Cloud。

核心功能
数据层（真实数据）
Baostock API - 免费开源证券数据平台
全A股上市公司实时行情（沪深北交易所）
K线数据含PE-TTM、PB-MRQ、成交量、换手率
财务报表：盈利能力、成长能力、偿债能力、杜邦分析
估值模型
PE市盈率法 - 悲观/基准/乐观三档估值
DCF现金流折现 - 10年自由现金流折现模型
PB市净率法 - 基于每股净资产的估值
综合估值 - 加权平均得出合理股价
智能投资建议
根据估值溢价空间自动评级：强烈买入 / 买入 / 持有 / 减持 / 卖出
结合PE、PB、财务健康度给出具体投资建议
详细评级理由说明
可视化图表（8+种）
图表	说明
K线走势	含MA5/MA20/MA60均线 + 成交量 + MACD指标
PE估值带	动态显示低估/合理/高估区间
估值瀑布图	对比当前价与各估值方法的差异
估值仪表盘	直观显示合理估值与当前价的关系
财务趋势图	净利润、ROE、毛利率、净利率5年趋势
成长性分析	净利润/EPS/资产/净资产增速对比
财务雷达图	6维财务能力对比（今年vs去年）
健康评分仪表盘	0-100分财务健康度量化评分
财务健康评分
综合评分（0-100分）
5个维度：ROE、毛利率、净利率、成长性、偿债能力
可视化进度条展示各维度得分
报告下载
Word报告 - 含投资建议、估值分析、财务摘要的完整文档
PDF报告 - 专业排版的投资价值分析报告
在线访问
Streamlit Cloud 部署地址

本地运行

Bash

pip install -r requirements.txt
streamlit run stock_valuation_app.py
部署到Streamlit Cloud
将代码推送到GitHub仓库 DataBridget/stock-valuation-tool
在 Streamlit Cloud 连接仓库
主文件设置为 stock_valuation_app.py
自动部署完成
技术栈
前端框架: Streamlit
数据源: Baostock
可视化: Plotly（交互式图表）
报告生成: python-docx, reportlab
数据处理: pandas, numpy
免责声明
本系统基于Baostock公开财务数据进行估值分析，所有结果仅供参考研究使用，不构成任何投资建议。股市有风险，投资需谨慎。

License
MIT
