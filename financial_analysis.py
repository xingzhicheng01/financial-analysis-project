# -*- coding: utf-8 -*-
"""
白酒行业双龙头财务分析脚本 v2
贵州茅台(600519) vs 五粮液(000858)

数据来源：东方财富(EM) + 新浪财经
分析维度：偿债能力 / 营运能力 / 盈利能力 / 发展能力 / 杜邦分析
"""

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import akshare as ak
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import warnings
import os
warnings.filterwarnings('ignore')

# ===================== 配置 =====================
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

COMPANIES = {'600519': '贵州茅台', '000858': '五粮液'}

def em_symbol(code):
    return f"SH{code}" if code.startswith('6') else f"SZ{code}"

OUTPUT_DIR = r'C:\Users\Xzc\financial_report'
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ===================== 格式化工具 =====================
def pct(v, digits=1):
    """已是百分数的值直接显示"""
    if pd.isna(v) or np.isinf(v):
        return 'N/A'
    return f"{v:.{digits}f}%"

def ratio(v, digits=2):
    if pd.isna(v) or np.isinf(v):
        return 'N/A'
    return f"{v:.{digits}f}"

def yi(v, digits=2):
    """转亿元"""
    if pd.isna(v) or np.isinf(v):
        return 'N/A'
    return f"{v / 1e8:.{digits}f}"

def sdiv(a, b):
    return a / b if b != 0 else np.nan

# ===================== 获取数据 =====================
print("=" * 70)
print("  贵州茅台(600519) vs 五粮液(000858) 财务分析")
print("=" * 70)
print()
print("正在获取财务数据...")

all_data = {}

for code, name in COMPANIES.items():
    print(f"  获取 {name}({code})...")
    try:
        sym = em_symbol(code)

        # Sina 财务指标（已有计算好的比率）
        indicator = ak.stock_financial_analysis_indicator(symbol=code, start_year="2020")

        # EM 报表（宽表格式：每列是一个科目）
        balance = ak.stock_balance_sheet_by_report_em(symbol=sym)
        income = ak.stock_profit_sheet_by_report_em(symbol=sym)
        cashflow = ak.stock_cash_flow_sheet_by_report_em(symbol=sym)

        # 统一处理 REPORT_DATE 为 datetime
        for df in [balance, income, cashflow]:
            if 'REPORT_DATE' in df.columns:
                df['REPORT_DATE'] = pd.to_datetime(df['REPORT_DATE'])

        all_data[code] = {
            'name': name, 'indicator': indicator,
            'balance': balance, 'income': income, 'cashflow': cashflow,
        }
        print(f"     OK")
    except Exception as e:
        print(f"     FAIL: {e}")
        all_data[code] = None

# ===================== 指标提取 =====================
# 新浪财务指标列名映射
SINA_INDICATOR_MAP = {
    '流动比率': ['流动比率'],
    '速动比率': ['速动比率'],
    '资产负债率': ['资产负债率'],
    '产权比率': ['产权比率'],
    '毛利率': ['毛利率'],
    '净利率': ['净利率'],
    'ROE': ['净资产收益率', '净资产收益率(%)'],
    'ROA': ['总资产报酬率', '总资产净利润率'],
    '应收账款周转率': ['应收账款周转率(次)', '应收账款周转率'],
    '存货周转率': ['存货周转率(次)', '存货周转率'],
    '总资产周转率': ['总资产周转率(次)', '总资产周转率'],
    '每股收益': ['摊薄每股收益(元)'],
    '每股净资产': ['每股净资产_调整前(元)'],
}
# 注意：新浪数据中的比率已经是百分数（如 54.17 表示 54.17%），不需要再乘100
# 部分指标需要判断是比率还是次数的单位

# EM 报表列名映射（宽表格式，直接从列取值）
EM_BALANCE_MAP = {
    '资产总计': 'TOTAL_ASSETS',
    '负债合计': 'TOTAL_LIABILITIES',
    '所有者权益合计': 'TOTAL_EQUITY',
    '归母权益': 'TOTAL_PARENT_EQUITY',
    '流动资产合计': 'TOTAL_CURRENT_ASSETS',
    '流动负债合计': 'TOTAL_CURRENT_LIAB',
    '存货': 'INVENTORY',
    '应收账款': 'ACCOUNTS_RECE',
    '货币资金': 'MONETARYFUNDS',
    '固定资产': 'FIXED_ASSET',
    '无形资产': 'INTANGIBLE_ASSET',
    '非流动资产合计': 'TOTAL_NONCURRENT_ASSETS',
    '非流动负债合计': 'TOTAL_NONCURRENT_LIAB',
    '短期借款': 'SHORT_LOAN',
    '预付款项': 'PREPAYMENT',
    '应付账款': 'ACCOUNTS_PAYABLE',
    '合同负债': 'CONTRACT_LIAB',
}

EM_INCOME_MAP = {
    '营业总收入': 'TOTAL_OPERATE_INCOME',
    '营业收入': 'OPERATE_INCOME',
    '营业成本': 'OPERATE_COST',
    '营业利润': 'OPERATE_PROFIT',
    '利润总额': 'TOTAL_PROFIT',
    '净利润': 'NETPROFIT',
    '归母净利润': 'PARENT_NETPROFIT',
    '销售费用': 'SALE_EXPENSE',
    '管理费用': 'MANAGE_EXPENSE',
    '研发费用': 'RD_EXPENSE',
    '财务费用': 'FINANCE_EXPENSE',
    '税金及附加': 'OPERATE_TAX_ADD',
}

EM_CASHFLOW_MAP = {
    '经营活动现金流净额': 'NETCASH_OPERATE',
    '投资活动现金流净额': 'NETCASH_INVEST',
    '筹资活动现金流净额': 'NETCASH_FINANCE',
}

# ===================== 提取数据 =====================
print()
print("正在提取财务数据...")

def get_latest_row(df, date_col='REPORT_DATE'):
    """获取最新报告期的行"""
    if df is None or df.empty:
        return None
    df_sorted = df.sort_values(date_col, ascending=False)
    return df_sorted.iloc[0]

def get_latest_val(df, col, date_col='REPORT_DATE'):
    """获取最新报告期某个科目的值"""
    if df is None or col not in df.columns:
        return np.nan
    row = get_latest_row(df, date_col)
    if row is None:
        return np.nan
    return pd.to_numeric(row[col], errors='coerce')

def get_sina_indicator(ind_df, targets):
    """从新浪指标表提取最新值"""
    if ind_df is None:
        return np.nan
    for target in targets:
        for col in ind_df.columns:
            if target in str(col):
                vals = pd.to_numeric(ind_df[col], errors='coerce').dropna()
                if len(vals) > 0:
                    return vals.iloc[0]
    return np.nan

# 构建对比数据
results = {}

for code, data in all_data.items():
    if data is None:
        continue
    name = data['name']
    ind = data['indicator']
    bs = data['balance']
    inc = data['income']
    cf = data['cashflow']

    r = {}

    # --- 从 EM 报表提取（宽表，直接取列） ---
    for item_name, col_name in EM_BALANCE_MAP.items():
        r[f'BS_{item_name}'] = get_latest_val(bs, col_name)

    for item_name, col_name in EM_INCOME_MAP.items():
        r[f'INC_{item_name}'] = get_latest_val(inc, col_name)

    for item_name, col_name in EM_CASHFLOW_MAP.items():
        r[f'CF_{item_name}'] = get_latest_val(cf, col_name)

    # --- 从新浪指标表提取 ---
    for metric, targets in SINA_INDICATOR_MAP.items():
        val = get_sina_indicator(ind, targets)
        r[f'SINA_{metric}'] = val if pd.notna(val) else np.nan

    # --- 报表日期 ---
    latest_bs = get_latest_row(bs)
    latest_inc = get_latest_row(inc)
    r['report_date_bs'] = latest_bs['REPORT_DATE'] if latest_bs is not None else 'N/A'
    r['report_date_inc'] = latest_inc['REPORT_DATE'] if latest_inc is not None else 'N/A'

    # --- 计算比率（使用 EM 数据手工计算，用于验证） ---
    营收 = r.get('INC_营业总收入', np.nan)
    营业成本 = r.get('INC_营业成本', np.nan)
    净利润 = r.get('INC_净利润', np.nan)
    总资产 = r.get('BS_资产总计', np.nan)
    总负债 = r.get('BS_负债合计', np.nan)
    净资产 = r.get('BS_所有者权益合计', np.nan)
    流动资产 = r.get('BS_流动资产合计', np.nan)
    流动负债 = r.get('BS_流动负债合计', np.nan)
    存货 = r.get('BS_存货', np.nan)
    应收账款 = r.get('BS_应收账款', np.nan)

    # 偿债能力
    r['CAL_流动比率'] = sdiv(流动资产, 流动负债)
    速动资产 = 流动资产 - 存货 if not pd.isna(流动资产) and not pd.isna(存货) else np.nan
    r['CAL_速动比率'] = sdiv(速动资产, 流动负债)
    r['CAL_资产负债率'] = sdiv(总负债, 总资产)
    r['CAL_权益乘数'] = sdiv(总资产, 净资产)

    # 盈利能力
    毛利 = 营收 - 营业成本 if not pd.isna(营收) and not pd.isna(营业成本) else np.nan
    r['CAL_毛利率'] = sdiv(毛利, 营收)
    r['CAL_净利率'] = sdiv(净利润, 营收)
    r['CAL_ROE'] = sdiv(净利润, 净资产)
    r['CAL_ROA'] = sdiv(净利润, 总资产)

    # 营运能力
    r['CAL_应收账款周转率'] = sdiv(营收, 应收账款)
    r['CAL_存货周转率'] = sdiv(营业成本, 存货)
    r['CAL_总资产周转率'] = sdiv(营收, 总资产)

    # 杜邦验证
    cal_净利率 = r['CAL_净利率']
    cal_总资产周转率 = r['CAL_总资产周转率']
    cal_权益乘数 = r['CAL_权益乘数']
    r['CAL_DUPONT_ROE'] = cal_净利率 * cal_总资产周转率 * cal_权益乘数

    results[code] = r

    # 打印提取结果
    print(f"\n  {'='*50}")
    print(f"  {name} 最新财报数据（报表日期: {r['report_date_inc']}）")
    print(f"  {'='*50}")

    def print_item(label, val, unit='亿'):
        if pd.isna(val):
            print(f"    {label}: N/A")
        elif unit == '亿':
            print(f"    {label}: {val/1e8:.2f} 亿元")
        else:
            print(f"    {label}: {val:.4f}")

    print_item('营业总收入', 营收)
    print_item('营业成本', 营业成本)
    print_item('净利润', 净利润)
    print_item('资产总计', 总资产)
    print_item('负债合计', 总负债)
    print_item('所有者权益', 净资产)
    print_item('流动资产', 流动资产)
    print_item('流动负债', 流动负债)
    print_item('存货', 存货)
    print_item('应收账款', 应收账款)

    print()
    print(f"  ─── 偿债能力 ───")
    print(f"    流动比率:       {ratio(r['CAL_流动比率'])}")
    print(f"    速动比率:       {ratio(r['CAL_速动比率'])}")
    print(f"    资产负债率:     {pct(r['CAL_资产负债率']*100)}")
    print(f"    权益乘数:       {ratio(r['CAL_权益乘数'])}")
    print(f"    新浪流动比率:   {r.get('SINA_流动比率', 'N/A')}")
    print(f"    新浪速动比率:   {r.get('SINA_速动比率', 'N/A')}")

    print(f"  ─── 盈利能力 ───")
    print(f"    毛利率:         {pct(r['CAL_毛利率']*100)}")
    print(f"    净利率:         {pct(r['CAL_净利率']*100)}")
    print(f"    ROE:            {pct(r['CAL_ROE']*100)}")
    print(f"    ROA:            {pct(r['CAL_ROA']*100)}")
    print(f"    新浪净利率:     {r.get('SINA_净利率', 'N/A')}")
    print(f"    新浪ROE:        {r.get('SINA_ROE', 'N/A')}")

    print(f"  ─── 营运能力 ───")
    print(f"    应收账款周转率: {ratio(r['CAL_应收账款周转率'])} 次")
    print(f"    存货周转率:     {ratio(r['CAL_存货周转率'])} 次")
    print(f"    总资产周转率:   {ratio(r['CAL_总资产周转率'])} 次")

    print(f"  ─── 杜邦分析 ───")
    print(f"    净利率:         {pct(r['CAL_净利率']*100)}")
    print(f"    总资产周转率:   {ratio(r['CAL_总资产周转率'])}")
    print(f"    权益乘数:       {ratio(r['CAL_权益乘数'])}")
    print(f"    ROE(杜邦):      {pct(r['CAL_DUPONT_ROE']*100)}")
    print(f"    ROE(直接):      {pct(r['CAL_ROE']*100)}")


# ===================== 对比报告 =====================
print()
print()
print("=" * 70)
print("  财报对比分析报告")
print("=" * 70)

def compare_report(title, items, key1, key2, fmt_func, note_map=None):
    """输出对比表格"""
    print()
    print(f"  {'─'*55}")
    print(f"  {title}")
    print(f"  {'─'*55}")
    header = f"  {'指标':<18s} {'贵州茅台':>12s} {'五粮液':>12s}  {'说明':<20s}"
    print(header)
    print(f"  {'-'*65}")

    for item in items:
        if isinstance(item, tuple):
            label, key = item[0], item[1]
        else:
            label, key = item, item
        v1 = results.get('600519', {}).get(key1.format(key), np.nan)
        v2 = results.get('000858', {}).get(key2.format(key), np.nan)
        note = ''
        if note_map and key in note_map:
            note = note_map[key]
        s1 = fmt_func(v1)
        s2 = fmt_func(v2)
        print(f"  {label:<18s} {s1:>12s} {s2:>12s}  {note}")

# 偿债能力
DEBT_NOTES = {
    'CAL_流动比率': '>2 优秀（白酒通常很高）',
    'CAL_速动比率': '>1 安全',
    'CAL_资产负债率': '越低越稳健',
    'CAL_权益乘数': '财务杠杆倍数',
}
compare_report('一、偿债能力分析', [
    ('流动比率', 'CAL_流动比率'),
    ('速动比率', 'CAL_速动比率'),
    ('资产负债率', 'CAL_资产负债率'),
    ('权益乘数', 'CAL_权益乘数'),
], '{}', '{}', lambda v: ratio(v), DEBT_NOTES)

# 盈利能力
PROFIT_NOTES = {
    'CAL_毛利率': '越高品牌越强',
    'CAL_净利率': '成本管控能力',
    'CAL_ROE': '股东回报 >15%=优',
    'CAL_ROA': '资产使用效率',
}
compare_report('二、盈利能力分析', [
    ('毛利率', 'CAL_毛利率'),
    ('净利率', 'CAL_净利率'),
    ('ROE', 'CAL_ROE'),
    ('ROA', 'CAL_ROA'),
], '{}', '{}', lambda v: pct(v * 100) if not pd.isna(v) else 'N/A', PROFIT_NOTES)

# 营运能力
OP_NOTES = {
    'CAL_应收账款周转率': '越高回款越快',
    'CAL_存货周转率': '越高变现越快',
    'CAL_总资产周转率': '越高效率越高',
}
compare_report('三、营运能力分析', [
    ('应收账款周转率(次)', 'CAL_应收账款周转率'),
    ('存货周转率(次)', 'CAL_存货周转率'),
    ('总资产周转率(次)', 'CAL_总资产周转率'),
], '{}', '{}', lambda v: ratio(v), OP_NOTES)

# 经营规模
print()
print(f"  {'─'*70}")
print(f"  四、经营规模对比")
print(f"  {'─'*70}")
print(f"  {'项目':<18s} {'贵州茅台':>14s} {'五粮液':>14s} {'茅台/五粮液':>12s}")
print(f"  {'-'*60}")

scale_items = [
    ('营业总收入', 'INC_营业总收入'),
    ('营业成本', 'INC_营业成本'),
    ('净利润', 'INC_净利润'),
    ('资产总计', 'BS_资产总计'),
    ('负债合计', 'BS_负债合计'),
    ('归母权益', 'BS_归母权益'),
    ('流动资产', 'BS_流动资产合计'),
    ('货币资金', 'BS_货币资金'),
]

for label, key in scale_items:
    v1 = results.get('600519', {}).get(key, np.nan)
    v2 = results.get('000858', {}).get(key, np.nan)
    s1 = yi(v1) if not pd.isna(v1) else 'N/A'
    s2 = yi(v2) if not pd.isna(v2) else 'N/A'
    ratio_val = ''
    if not pd.isna(v1) and not pd.isna(v2) and v2 != 0:
        ratio_val = f"{v1/v2:.2f}x"
    print(f"  {label:<18s} {s1:>14s} {s2:>14s} {ratio_val:>12s}")

# 杜邦分析
print()
print(f"  {'─'*70}")
print(f"  五、杜邦分析 (ROE = 净利率 x 总资产周转率 x 权益乘数)")
print(f"  {'─'*70}")

for code, data in all_data.items():
    if data is None:
        continue
    name = data['name']
    r = results.get(code, {})

    净利率 = r.get('CAL_净利率', np.nan)
    周转率 = r.get('CAL_总资产周转率', np.nan)
    乘数 = r.get('CAL_权益乘数', np.nan)
    roe_dupont = r.get('CAL_DUPONT_ROE', np.nan)
    roe_direct = r.get('CAL_ROE', np.nan)

    print(f"\n  {name}:")
    print(f"    净利率       = {pct(净利率 * 100)}")
    print(f"    总资产周转率  = {ratio(周转率)}")
    print(f"    权益乘数     = {ratio(乘数)}")
    print(f"    ─────────────────────")
    print(f"    ROE (杜邦)   = {pct(roe_dupont * 100)}")
    print(f"    ROE (直接)   = {pct(roe_direct * 100)}")

    # ROE 驱动力判断
    if not pd.isna(净利率):
        if 净利率 > 0.30:
            print(f"    >> 核心驱动：超高净利率 ({pct(净利率*100)})，高端白酒盈利模式")
        elif 净利率 > 0.15:
            print(f"    >> 核心驱动：较高净利率 + 资产周转")
        else:
            print(f"    >> 核心驱动：资产周转 + 杠杆")


# ===================== 图表 =====================
print()
print("正在生成图表...")

# 图1：四维对比
fig, axes = plt.subplots(2, 2, figsize=(14, 12))
width = 0.35

# 偿债能力
ax = axes[0, 0]
debt_cats = ['流动比率', '速动比率', '权益乘数']
mt_vals = [results.get('600519', {}).get('CAL_流动比率', 0) or 0,
           results.get('600519', {}).get('CAL_速动比率', 0) or 0,
           results.get('600519', {}).get('CAL_权益乘数', 0) or 0]
wly_vals = [results.get('000858', {}).get('CAL_流动比率', 0) or 0,
            results.get('000858', {}).get('CAL_速动比率', 0) or 0,
            results.get('000858', {}).get('CAL_权益乘数', 0) or 0]
x = np.arange(len(debt_cats))
ax.bar(x - width/2, mt_vals, width, label='贵州茅台', color='#c0392b', alpha=0.85)
ax.bar(x + width/2, wly_vals, width, label='五粮液', color='#2980b9', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(debt_cats, fontsize=11)
ax.set_title('偿债能力', fontsize=14, fontweight='bold')
ax.legend(fontsize=10)
ax.grid(axis='y', alpha=0.3)

# 盈利能力
ax = axes[0, 1]
profit_cats = ['毛利率(%)', '净利率(%)', 'ROE(%)', 'ROA(%)']
mt_profit = [results.get('600519', {}).get(f'CAL_{k}', 0) * 100 or 0 for k in ['毛利率', '净利率', 'ROE', 'ROA']]
wly_profit = [results.get('000858', {}).get(f'CAL_{k}', 0) * 100 or 0 for k in ['毛利率', '净利率', 'ROE', 'ROA']]
x = np.arange(len(profit_cats))
ax.bar(x - width/2, mt_profit, width, label='贵州茅台', color='#c0392b', alpha=0.85)
ax.bar(x + width/2, wly_profit, width, label='五粮液', color='#2980b9', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(profit_cats, fontsize=11)
ax.set_title('盈利能力 (%)', fontsize=14, fontweight='bold')
ax.legend(fontsize=10)
ax.grid(axis='y', alpha=0.3)

# 经营规模（亿元）
ax = axes[1, 0]
scale_cats = ['营业总收入', '净利润', '总资产']
mt_scale = [
    (results.get('600519', {}).get('INC_营业总收入', 0) or 0) / 1e8,
    (results.get('600519', {}).get('INC_净利润', 0) or 0) / 1e8,
    (results.get('600519', {}).get('BS_资产总计', 0) or 0) / 1e8,
]
wly_scale = [
    (results.get('000858', {}).get('INC_营业总收入', 0) or 0) / 1e8,
    (results.get('000858', {}).get('INC_净利润', 0) or 0) / 1e8,
    (results.get('000858', {}).get('BS_资产总计', 0) or 0) / 1e8,
]
x = np.arange(len(scale_cats))
ax.bar(x - width/2, mt_scale, width, label='贵州茅台', color='#c0392b', alpha=0.85)
ax.bar(x + width/2, wly_scale, width, label='五粮液', color='#2980b9', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(scale_cats, fontsize=11)
ax.set_title('经营规模 (亿元)', fontsize=14, fontweight='bold')
ax.legend(fontsize=10)
ax.grid(axis='y', alpha=0.3)

# 杜邦拆解
ax = axes[1, 1]
dupont_cats = ['净利率(%)', '总资产周转率', '权益乘数']
mt_dupont = [
    (results.get('600519', {}).get('CAL_净利率', 0) or 0) * 100,
    results.get('600519', {}).get('CAL_总资产周转率', 0) or 0,
    results.get('600519', {}).get('CAL_权益乘数', 0) or 0,
]
wly_dupont = [
    (results.get('000858', {}).get('CAL_净利率', 0) or 0) * 100,
    results.get('000858', {}).get('CAL_总资产周转率', 0) or 0,
    results.get('000858', {}).get('CAL_权益乘数', 0) or 0,
]
x = np.arange(len(dupont_cats))
ax.bar(x - width/2, mt_dupont, width, label='贵州茅台', color='#c0392b', alpha=0.85)
ax.bar(x + width/2, wly_dupont, width, label='五粮液', color='#2980b9', alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(dupont_cats, fontsize=11)
ax.set_title('杜邦分析拆解', fontsize=14, fontweight='bold')
ax.legend(fontsize=10)
ax.grid(axis='y', alpha=0.3)

plt.suptitle('贵州茅台 vs 五粮液 — 财务对比分析', fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
path1 = os.path.join(OUTPUT_DIR, '财务对比分析.png')
plt.savefig(path1, dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print(f"  OK: {path1}")

# 图2：ROE 驱动瀑布堆积
fig, ax = plt.subplots(figsize=(12, 6))

mt_净利 = (results.get('600519', {}).get('CAL_净利率', 0) or 0) * 100
mt_周转 = results.get('600519', {}).get('CAL_总资产周转率', 0) or 0
mt_乘数 = results.get('600519', {}).get('CAL_权益乘数', 0) or 0
mt_roe = (results.get('600519', {}).get('CAL_ROE', 0) or 0) * 100

wly_净利 = (results.get('000858', {}).get('CAL_净利率', 0) or 0) * 100
wly_周转 = results.get('000858', {}).get('CAL_总资产周转率', 0) or 0
wly_乘数 = results.get('000858', {}).get('CAL_权益乘数', 0) or 0
wly_roe = (results.get('000858', {}).get('CAL_ROE', 0) or 0) * 100

categories = ['净利率\n(%)', '总资产\n周转率', '权益\n乘数', 'ROE\n(%)']
mt_bars = [mt_净利, mt_周转, mt_乘数, mt_roe]
wly_bars = [wly_净利, wly_周转, wly_乘数, wly_roe]

x = np.arange(len(categories))
width = 0.35
bars1 = ax.bar(x - width/2, mt_bars, width, label='贵州茅台', color='#c0392b', alpha=0.85, edgecolor='white', linewidth=0.5)
bars2 = ax.bar(x + width/2, wly_bars, width, label='五粮液', color='#2980b9', alpha=0.85, edgecolor='white', linewidth=0.5)

# 标注数值
for bar in bars1:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 0.5, f'{h:.1f}', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#c0392b')
for bar in bars2:
    h = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., h + 0.5, f'{h:.1f}', ha='center', va='bottom', fontsize=10, fontweight='bold', color='#2980b9')

ax.set_xticks(x)
ax.set_xticklabels(categories, fontsize=12)
ax.set_title('ROE 驱动因素对比', fontsize=16, fontweight='bold')
ax.legend(fontsize=12, loc='upper right')
ax.grid(axis='y', alpha=0.2)
ax.set_ylabel('数值', fontsize=11)
plt.tight_layout()
path2 = os.path.join(OUTPUT_DIR, '杜邦分析对比.png')
plt.savefig(path2, dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print(f"  OK: {path2}")


# ===================== Excel 报告 =====================
print()
print("正在生成 Excel 报告...")

wb = Workbook()
ws = wb.active
ws.title = "财务分析报告"

# 样式
title_font = Font(name='Microsoft YaHei', size=14, bold=True, color='FFFFFF')
title_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
header_font = Font(name='Microsoft YaHei', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='34495E', end_color='34495E', fill_type='solid')
section_font = Font(name='Microsoft YaHei', size=12, bold=True, color='C0392B')
normal_font = Font(name='Microsoft YaHei', size=10)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# 标题
ws.merge_cells('A1:G1')
ws['A1'] = '贵州茅台(600519) vs 五粮液(000858) — 财务分析报告'
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
ws.row_dimensions[1].height = 35

def write_section(ws, start_row, title, headers, data_rows):
    ws.merge_cells(f'A{start_row}:G{start_row}')
    ws[f'A{start_row}'] = title
    ws[f'A{start_row}'].font = section_font
    ws.row_dimensions[start_row].height = 25

    start_row += 1
    for j, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=j, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin_border
    ws.row_dimensions[start_row].height = 22

    start_row += 1
    for i, row in enumerate(data_rows):
        for j, val in enumerate(row, 1):
            cell = ws.cell(row=start_row + i, column=j, value=str(val) if val is not None else 'N/A')
            cell.font = normal_font
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center' if j > 1 else 'left')

    return start_row + len(data_rows) + 1

row = 3
cal_prefix = 'CAL_'

# 偿债能力
row = write_section(ws, row, '一、偿债能力分析',
    ['指标', '贵州茅台', '五粮液', '说明'],
    [
        ['流动比率', ratio(results['600519'].get(f'{cal_prefix}流动比率')),
         ratio(results['000858'].get(f'{cal_prefix}流动比率')), '流动资产/流动负债，>2为优'],
        ['速动比率', ratio(results['600519'].get(f'{cal_prefix}速动比率')),
         ratio(results['000858'].get(f'{cal_prefix}速动比率')), '(流动资产-存货)/流动负债，>1为优'],
        ['资产负债率', pct((results['600519'].get(f'{cal_prefix}资产负债率') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}资产负债率') or 0) * 100), '越低越稳健'],
        ['权益乘数', ratio(results['600519'].get(f'{cal_prefix}权益乘数')),
         ratio(results['000858'].get(f'{cal_prefix}权益乘数')), '总资产/净资产，杠杆倍数'],
    ])

# 盈利能力
row = write_section(ws, row, '二、盈利能力分析',
    ['指标', '贵州茅台', '五粮液', '说明'],
    [
        ['毛利率', pct((results['600519'].get(f'{cal_prefix}毛利率') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}毛利率') or 0) * 100), '(收入-成本)/收入'],
        ['净利率', pct((results['600519'].get(f'{cal_prefix}净利率') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}净利率') or 0) * 100), '净利润/收入'],
        ['ROE', pct((results['600519'].get(f'{cal_prefix}ROE') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}ROE') or 0) * 100), '净利润/净资产，>15%为优'],
        ['ROA', pct((results['600519'].get(f'{cal_prefix}ROA') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}ROA') or 0) * 100), '净利润/总资产'],
    ])

# 营运能力
row = write_section(ws, row, '三、营运能力分析',
    ['指标', '贵州茅台', '五粮液', '说明'],
    [
        ['应收账款周转率(次)', ratio(results['600519'].get(f'{cal_prefix}应收账款周转率')),
         ratio(results['000858'].get(f'{cal_prefix}应收账款周转率')), '收入/应收账款，越高越好'],
        ['存货周转率(次)', ratio(results['600519'].get(f'{cal_prefix}存货周转率')),
         ratio(results['000858'].get(f'{cal_prefix}存货周转率')), '成本/存货，越高越好'],
        ['总资产周转率(次)', ratio(results['600519'].get(f'{cal_prefix}总资产周转率')),
         ratio(results['000858'].get(f'{cal_prefix}总资产周转率')), '收入/总资产，越高越好'],
    ])

# 经营规模
row = write_section(ws, row, '四、经营规模对比（亿元）',
    ['项目', '贵州茅台', '五粮液', '茅台/五粮液'],
    [
        ['营业总收入',
         yi(results['600519'].get('INC_营业总收入')),
         yi(results['000858'].get('INC_营业总收入')),
         f"{results['600519'].get('INC_营业总收入', 0) / results['000858'].get('INC_营业总收入', 1):.2f}x" if results['000858'].get('INC_营业总收入', 0) else 'N/A'],
        ['净利润',
         yi(results['600519'].get('INC_净利润')),
         yi(results['000858'].get('INC_净利润')),
         f"{results['600519'].get('INC_净利润', 0) / results['000858'].get('INC_净利润', 1):.2f}x" if results['000858'].get('INC_净利润', 0) else 'N/A'],
        ['资产总计',
         yi(results['600519'].get('BS_资产总计')),
         yi(results['000858'].get('BS_资产总计')),
         f"{results['600519'].get('BS_资产总计', 0) / results['000858'].get('BS_资产总计', 1):.2f}x" if results['000858'].get('BS_资产总计', 0) else 'N/A'],
        ['所有者权益',
         yi(results['600519'].get('BS_所有者权益合计')),
         yi(results['000858'].get('BS_所有者权益合计')),
         f"{results['600519'].get('BS_所有者权益合计', 0) / results['000858'].get('BS_所有者权益合计', 1):.2f}x" if results['000858'].get('BS_所有者权益合计', 0) else 'N/A'],
    ])

# 杜邦分析
row = write_section(ws, row, '五、杜邦分析',
    ['指标', '贵州茅台', '五粮液', '说明'],
    [
        ['净利率', pct((results['600519'].get(f'{cal_prefix}净利率') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}净利率') or 0) * 100), 'ROE 第一因子'],
        ['总资产周转率', ratio(results['600519'].get(f'{cal_prefix}总资产周转率')),
         ratio(results['000858'].get(f'{cal_prefix}总资产周转率')), 'ROE 第二因子'],
        ['权益乘数', ratio(results['600519'].get(f'{cal_prefix}权益乘数')),
         ratio(results['000858'].get(f'{cal_prefix}权益乘数')), 'ROE 第三因子'],
        ['ROE(杜邦计算)', pct((results['600519'].get(f'{cal_prefix}DUPONT_ROE') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}DUPONT_ROE') or 0) * 100), '三因子乘积'],
        ['ROE(直接计算)', pct((results['600519'].get(f'{cal_prefix}ROE') or 0) * 100),
         pct((results['000858'].get(f'{cal_prefix}ROE') or 0) * 100), '净利润/净资产'],
    ])

# 设置列宽
for col, width in [('A', 22), ('B', 16), ('C', 16), ('D', 14), ('E', 18), ('F', 18), ('G', 30)]:
    ws.column_dimensions[col].width = width

excel_path = os.path.join(OUTPUT_DIR, '财务分析报告.xlsx')
wb.save(excel_path)
print(f"  OK: {excel_path}")


# ===================== 总结 =====================
print()
print("=" * 70)
print("  分析总结")
print("=" * 70)

mt_roe = (results.get('600519', {}).get('CAL_ROE', 0) or 0) * 100
wly_roe = (results.get('000858', {}).get('CAL_ROE', 0) or 0) * 100
mt_net = (results.get('600519', {}).get('CAL_净利率', 0) or 0) * 100
wly_net = (results.get('000858', {}).get('CAL_净利率', 0) or 0) * 100
mt_gross = (results.get('600519', {}).get('CAL_毛利率', 0) or 0) * 100
wly_gross = (results.get('000858', {}).get('CAL_毛利率', 0) or 0) * 100
mt_debt = (results.get('600519', {}).get('CAL_资产负债率', 0) or 0) * 100
wly_debt = (results.get('000858', {}).get('CAL_资产负债率', 0) or 0) * 100
mt_turn = results.get('600519', {}).get('CAL_总资产周转率', 0) or 0
wly_turn = results.get('000858', {}).get('CAL_总资产周转率', 0) or 0

print(f"""
  盈利能力对比:
    茅台 毛利率={pct(mt_gross)}  净利率={pct(mt_net)}  ROE={pct(mt_roe)}
    五粮液 毛利率={pct(wly_gross)}  净利率={pct(wly_net)}  ROE={pct(wly_roe)}

  偿债能力:
    茅台 资产负债率={pct(mt_debt)}  五粮液 资产负债率={pct(wly_debt)}
    (白酒行业 <30% 属于非常稳健水平)

  资产效率:
    茅台 总资产周转率={ratio(mt_turn)}  五粮液 总资产周转率={ratio(wly_turn)}

  ROE 驱动模式:
    茅台 ROE 主要驱动力: 超高净利率（品牌溢价）
    五粮液 ROE 主要驱动力: 高净利率 + 资产效率
""")

print(f"\n输出文件:")
print(f"  {path1}")
print(f"  {path2}")
print(f"  {excel_path}")
print()
print("分析完成!")
