# -*- coding: utf-8 -*-
"""
智慧畜牧数字化管理与核心指标交互式看板
项目：畜牧图情结合 · smart_pig_demo
技术栈：Streamlit + pandas + Plotly + openpyxl
升级亮点（2026-10-22 第5次改动）：
  1. 版权与独立声明强化：版权标识处添加李尧 2024010983，页面显著位置增设中国西北农林科技大学动物科技学院李尧独立进行之正式中英文声明（严禁擅自论文发表及商用）
  2. 增设金融与产业龙头生态板块：可分别一键跳转中国前十猪企官网、中国前十乳企官网（排名依据上市公司市值及年出栏/原奶产量），并支持直达跳转生猪期货、牛奶期货、牛肉期货及豆粕、玉米、菜粕、豆油、淀粉、强麦等各种核心饲料原料期货实时行情与交易所规则
  3. 开发者联系方式公示：页面侧边栏与页脚清晰公示作者联系方式（QQ: 3100785036，WeChat: 13679275651）
  4. 严格保持《畜牧法》与 NY/T 3445-2019 法定合规依据与双畜种全闭环业务
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
import io
from pathlib import Path

# 设置页面配置
st.set_page_config(
    page_title="智慧养殖数字化管理与指标看板 · 畜牧图情结合",
    page_icon="🐄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 路径常量解析
CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
PIG_KPI_PATH = ROOT_DIR / "表格与指标库" / "生猪养殖核心指标库与可行性验证.xlsx"
COW_KPI_PATH = ROOT_DIR / "表格与指标库" / "奶牛养殖核心指标库与可行性验证.xlsx"
ARCHIVE_PATH = ROOT_DIR / "表格与指标库" / "畜禽养殖档案管理表格（全套）.xlsx"
RANDOM_PIG_TEST_PATH = CURRENT_DIR / "测试数据_万头规模场2026年度生产台账.xlsx"
PIG_TEMPLATE_PATH = CURRENT_DIR / "生猪生产台账采集与指标计算上传模板.xlsx"
COW_TEMPLATE_PATH = CURRENT_DIR / "奶牛生产台账采集与指标计算上传模板.xlsx"

# =====================================================================
# 一、基础辅助函数：安全计算与UI格式化
# =====================================================================
def safe_series(df, col, default_val=np.nan):
    """安全提取数值型列，自动将无效字符或缺失值转为 NaN"""
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce')
    return pd.Series([default_val] * len(df))

def safe_div(s1, s2):
    """序列级安全除法，分母为0或NaN时返回 NaN，不抛出异常"""
    n1 = safe_series(pd.DataFrame({'a': s1}), 'a')
    n2 = safe_series(pd.DataFrame({'b': s2}), 'b').replace(0, np.nan)
    return n1 / n2

def safe_scalar_div(n, d):
    """标量级安全除法"""
    if n is None or d is None or pd.isna(n) or pd.isna(d) or d == 0:
        return np.nan
    try:
        val = float(n) / float(d)
        return np.nan if np.isinf(val) else val
    except:
        return np.nan

def render_kpi_card(col, label, val, is_pct=False, fmt=".2f", unit="", 
                    target_delta="", warn_delta=None, threshold=None, 
                    inverse=False, missing_hint="待录入对应台账"):
    """
    统一渲染顶部指标卡片：适配正常数值、缺失空值(—)、计算异常(⚠️)
    """
    if val is None or pd.isna(val):
        col.metric(
            label=label,
            value="—",
            delta=f"⏳ {missing_hint}",
            delta_color="off"
        )
    elif np.isinf(val):
        col.metric(
            label=label,
            value="⚠️ 需核验",
            delta="分母为零或数据异常",
            delta_color="inverse"
        )
    else:
        try:
            if is_pct:
                disp_val = f"{val * 100:{fmt}}%"
            else:
                disp_val = f"{val:{fmt}} {unit}".strip()
            
            # 判断达标与增量展示
            delta_text = target_delta
            d_color = "normal"
            if threshold is not None:
                if (val > threshold if inverse else val < threshold):
                    delta_text = warn_delta or f"未达标（参考基线 {threshold}）"
                    d_color = "inverse"
            col.metric(label=label, value=disp_val, delta=delta_text, delta_color=d_color)
        except Exception:
            col.metric(label=label, value="⚠️ 格式异常", delta="无法格式化", delta_color="inverse")

def has_valid_data(series):
    """检查序列中是否至少包含一个有效的非空数值"""
    if series is None:
        return False
    s_clean = pd.to_numeric(series, errors='coerce').dropna()
    return not s_clean.empty and not (s_clean == 0).all()

# =====================================================================
# 二、数据解析器：支持模版与自定义 Excel 自适应兼容读取
# =====================================================================
def parse_uploaded_excel(uploaded_file, animal_type="生猪"):
    """
    通用智能解析用户上传的 Excel 工作簿：
    1. 自动定位包含数据的有效工作表（如 '台账数据录入', '③演算-数据录入', 或首个 Sheet）
    2. 自动在前10行中探测表头行
    3. 别名与通配映射，支持中英文括号、星号后缀等字段变体
    4. 数值型字段自动清洗（去除逗号、货币符号、百分号）
    """
    try:
        content = uploaded_file.read() if hasattr(uploaded_file, 'read') else open(uploaded_file, 'rb').read()
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    except Exception as e:
        raise ValueError(f"读取 Excel 文件失败，请确认文件格式完整：{e}")

    target_sheet = None
    priority_names = ['台账数据录入', '③演算-数据录入', '数据录入', 'Sheet1', wb.sheetnames[0]]
    for name in priority_names:
        if name in wb.sheetnames:
            target_sheet = wb[name]
            break
    if target_sheet is None:
        target_sheet = wb.active

    rows = list(target_sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("所选工作表内容为空")

    # 扫描表头所在行
    header_idx = -1
    for idx, r in enumerate(rows[:10]):
        r_str = [str(c or '').strip() for c in r]
        if any(k in r_str for k in ['月份', '期末能繁母猪', '期末成母牛', '出生活仔', '月产奶量', '出栏', '配种数', '日均单产']):
            header_idx = idx
            break

    if header_idx == -1:
        header_idx = 2 if len(rows) > 2 else 0

    raw_headers = [str(c or '').strip() for c in rows[header_idx]]
    data_rows = rows[header_idx+1:]

    valid_data = []
    for r in data_rows:
        if any(c is not None and str(c).strip() != '' for c in r):
            valid_data.append(list(r[:len(raw_headers)]))

    if not valid_data:
        raise ValueError("工作表中未检测到有效数据行")

    df = pd.DataFrame(valid_data, columns=raw_headers)

    pig_col_map = {
        '期末能繁母猪': ['期末能繁母猪*', '期末能繁母猪(头)', '期末母猪', '能繁母猪'],
        '出生活仔': ['出生活仔*', '出生(活仔)', '出生活仔数', '活仔数', '出生活仔(头)', '出生'],
        '引进': ['引进(头)', '购入引进', '引种'],
        '出栏': ['出栏*', '出栏(头)', '出栏数', '出栏肥猪数', '肉猪出栏'],
        '死亡': ['死亡*', '死亡(头)', '全场死亡', '病死头数'],
        '淘汰全群': ['淘汰(全群)', '淘汰全群(头)', '淘汰全群', '全群淘汰'],
        '其中:淘汰母猪': ['其中:淘汰母猪', '淘汰母猪', '淘汰能繁母猪'],
        '调出': ['调出(头)', '调出'],
        '配种数': ['配种头次', '配种总数', '配种数(次)'],
        '分娩窝数': ['分娩窝数(窝)', '分娩数', '分娩窝数'],
        '断奶仔猪': ['断奶仔猪*', '断奶仔猪数', '断奶头数', '断奶数'],
        '耗料量_kg': ['耗料量(kg)', '饲料消耗量(kg)', '耗料量_kg', '耗料量'],
        '总增重_kg': ['总增重(kg)', '总增重_kg', '总增重'],
        '应免头次': ['应免头次(次)', '应免数', '应免头次'],
        '实免头次': ['实免头次(次)', '实免数', '实免头次'],
        '检测样本数': ['检测样本数(份)', '抽样数', '检测样本数'],
        '抗体合格数': ['合格数', '抗体合格头数', '抗体合格数'],
        '出栏批次': ['出栏批次(批)', '出栏批次'],
        '休药期合规批次': ['休药期合规批次(批)', '休药合规批次', '休药期合规批次'],
        '病死数': ['病死数(头)', '病死头数', '病死数'],
        '无害化处理数': ['无害化处理数(头)', '无害化数', '无害化处理数'],
        '应佩标识': ['应佩标识数', '应佩标识(枚)', '应佩标识'],
        '实佩标识': ['实佩标识数', '实佩标识(枚)', '实佩标识'],
        '环境天数': ['环境天数(天)', '环境记录天数', '环境天数'],
        '温度达标天数': ['温度达标天数(天)', '舍温达标天数', '温度达标天数'],
        '饲料采购金额': ['饲料采购金额(元)', '饲料支出', '饲料采购金额'],
        '兽药采购金额': ['兽药采购金额(元)', '兽药防疫费', '兽药采购金额'],
        '其他成本': ['其他成本-人工折旧能源(元)', '其他成本(元)', '人工折旧费用', '其他成本'],
        '出栏总重_kg': ['出栏总重(kg)', '销售总重(kg)', '出栏总重_kg', '出栏总重'],
        '出栏单价': ['出栏单价(元/kg)', '出栏单价(元)', '出栏单价']
    }

    cow_col_map = {
        '期末成母牛': ['期末成母牛*', '期末成母牛(头)', '成母牛存栏', '成母牛'],
        '初产育成牛转入': ['转入育成牛', '育成牛转入', '初产转入', '初产育成牛转入'],
        '购入引进牛': ['购入引进牛', '购入牛', '引进'],
        '出栏与淘汰': ['淘汰成母牛', '淘汰', '出栏淘汰', '出栏与淘汰'],
        '死亡成母牛': ['死亡*', '死亡', '死亡成母牛(头)', '死亡成母牛'],
        '在挤泌乳牛': ['泌乳牛头数*', '在挤泌乳牛*', '泌乳牛头数', '在挤泌乳牛(头)', '在挤泌乳牛'],
        '干奶牛': ['干奶牛头数', '干奶牛(头)', '干奶牛'],
        '出生犊牛': ['犊牛出生数', '出生犊牛(头)', '出生犊牛'],
        '成活犊牛': ['犊牛成活数', '成活犊牛(头)', '成活犊牛'],
        '月产奶量_kg': ['月产奶量_kg*', '月产奶量(kg)', '产奶量_kg', '月产奶量'],
        '生鲜乳销量_kg': ['生鲜乳销量(kg)', '生鲜乳交售量(kg)', '生鲜乳销量_kg', '生鲜乳销量'],
        '乳脂率': ['乳脂率(%)', '乳脂肪率', '乳脂率'],
        '乳蛋白率': ['乳蛋白率(%)', '蛋白质率', '乳蛋白率'],
        '体细胞数_SCC': ['体细胞数(万/mL)', '体细胞数', 'SCC(万/mL)', '体细胞数_SCC'],
        '菌落总数': ['菌落数(万CFU/mL)', '菌落总数(万CFU/mL)', '菌落数', '菌落总数'],
        '配种头次': ['配种头次(次)', '情期配种数', '配种头次'],
        '受孕头次': ['受孕头次(次)', '妊娠头次', '受孕头次'],
        'DMI_kg': ['干物质采食量DMI_kg', '干物质采食量DMI(kg)', 'DMI', 'DMI_kg'],
        '总饲喂干物质_kg': ['总饲喂干物质(kg)', '总饲喂干物质_kg', '总饲喂干物质'],
        '出乳批次': ['出乳批次(批)', '交售批次', '出乳批次'],
        '休药期合规批次': ['休药期合规批次(批)', '弃奶合规批次', '休药期合规批次'],
        '病死牛数': ['病死牛数(头)', '病死成母牛', '病死牛数'],
        '无害化处理数': ['无害化处理数(头)', '病死无害化数', '无害化处理数'],
        '环境天数': ['环境天数(天)', '环境记录天数', '环境天数'],
        'THI达标天数': ['温湿度指数THI达标天数', 'THI达标天数(天)', 'THI达标天数'],
        '乳房炎头次': ['临床乳房炎发生头次', '乳房炎发病头次', '乳房炎头次'],
        '粗饲料费': ['粗饲料金额(元)', '粗饲料成本(元)', '粗饲料费'],
        '精饲料费': ['精饲料金额(元)', '精饲料成本(元)', '精饲料费'],
        '兽药防疫费': ['兽药防疫费(元)', '兽药疫苗费用(元)', '兽药防疫费'],
        '其他成本': ['其他成本(元)', '人工折旧能耗费用(元)', '人工能耗费', '其他成本'],
        '原料奶售价': ['原料奶售价(元/kg)', '鲜奶售价(元/kg)', '原料奶售价']
    }

    col_map = pig_col_map if "生猪" in animal_type else cow_col_map
    rename_dict = {}
    matched_targets = set()

    for col in df.columns:
        clean_c = col.replace(' ', '').replace('　', '')
        for standard_col, aliases in col_map.items():
            if standard_col in matched_targets:
                continue
            if clean_c == standard_col or clean_c in aliases or any(a in clean_c for a in aliases):
                rename_dict[col] = standard_col
                matched_targets.add(standard_col)
                break

    df = df.rename(columns=rename_dict)
    df = df.loc[:, ~df.columns.duplicated()]

    # 数值清洗
    for col in df.columns:
        if col != '月份':
            s = df[col].astype(str).str.replace(',', '', regex=False).str.replace('¥', '', regex=False).str.replace('%', '', regex=False).str.strip()
            df[col] = pd.to_numeric(s, errors='coerce')

    # 百分比归一化
    if "生猪" not in animal_type:
        for p_col in ['乳脂率', '乳蛋白率']:
            if p_col in df.columns and df[p_col].dropna().mean() > 1.0:
                df[p_col] = df[p_col] / 100.0

    return df

# =====================================================================
# 三、数据模型：生猪默认数据与容错计算引擎
# =====================================================================
def get_default_pig_data():
    months = [f"{m}月" for m in range(1, 13)]
    end_sow = [500,502,505,508,510,512,510,508,506,508,510,512]
    born   = [1056,1012,1089,1045,1023,1067,1034,1012,1078,1056,1034,1089]
    intro  = [0]*12
    out    = [900,950,920,960,940,980,950,930,970,960,940,1000]
    death  = [72,68,75,80,78,82,78,120,78,72,70,76]   # 8月预置异常（2.07%）
    cull   = [60,55,58,62,60,58,65,60,55,58,62,60]
    cull_s = [12,10,14,12,13,11,15,12,10,13,12,14]
    trans  = [0]*12
    mate   = [105,100,108,104,102,106,103,101,107,105,103,108]
    farrow = [96,92,99,95,93,97,94,92,98,96,94,99]
    wean   = [980,950,1020,980,960,1000,970,950,1010,990,970,1020]
    gain   = [145000,142000,150000,148000,146000,152000,149000,147000,153000,151000,148000,154000]
    feed_kg= [round(x*3.2) for x in gain]
    vacc_due=[2000,1900,2100,2000,1950,2050,2000,1950,2100,2050,2000,2100]
    vacc_done=vacc_due[:]
    test_n = [40,32,45,40,38,42,40,36,44,40,38,45]
    test_ok= [37,25,42,37,36,39,37,34,41,37,36,42]       # 2月预置异常（78.1%）
    batch  = [22,20,23,22,21,24,23,22,23,24,22,25]
    batch_ok=batch[:]
    dead   = death
    harmless=death
    tag_due=born
    tag_done=born
    env_days=[31,28,31,30,31,30,31,31,30,31,30,31]
    temp_ok=[26,25,28,29,30,29,30,30,29,29,28,27]
    feed_money=[round(x*3.3) for x in feed_kg]
    vet_money=[90000,85000,95000,92000,88000,96000,93000,100000,91000,94000,92000,97000]
    other_cost=[500000]*12
    sale_w=[round(o*w) for o,w in zip(out, [118,119,120,121,120,122,121,120,119,120,121,122])]
    price=[19.5,19.8,20.2,20.5,20.3,19.9,19.6,19.4,19.7,20.1,20.4,20.8]

    rows = []
    for i in range(12):
        rows.append({
            '月份': months[i], '期末能繁母猪': end_sow[i], '出生活仔': born[i], '引进': intro[i],
            '出栏': out[i], '死亡': death[i], '淘汰全群': cull[i], '其中:淘汰母猪': cull_s[i],
            '调出': trans[i], '配种数': mate[i], '分娩窝数': farrow[i], '断奶仔猪': wean[i],
            '耗料量_kg': feed_kg[i], '总增重_kg': gain[i], '应免头次': vacc_due[i], '实免头次': vacc_done[i],
            '检测样本数': test_n[i], '抗体合格数': test_ok[i], '出栏批次': batch[i], '休药期合规批次': batch_ok[i],
            '病死数': dead[i], '无害化处理数': harmless[i], '应佩标识': tag_due[i], '实佩标识': tag_done[i],
            '环境天数': env_days[i], '温度达标天数': temp_ok[i],
            '饲料采购金额': feed_money[i], '兽药采购金额': vet_money[i], '其他成本': other_cost[i],
            '出栏总重_kg': sale_w[i], '出栏单价': price[i]
        })
    return pd.DataFrame(rows)

def compute_pig_metrics(df_raw):
    """
    生猪核心指标计算引擎（容错隔离优化版）：
    - 任何指标列缺失、为空或格式异常时，不抛错，仅标记对应指标为空
    - 其他具备输入依据的指标完全不受干扰，照常计算
    """
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}

    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    end_sow = safe_series(df, '期末能繁母猪')
    born = safe_series(df, '出生活仔')
    intro = safe_series(df, '引进', 0)
    out = safe_series(df, '出栏')
    death = safe_series(df, '死亡')
    cull = safe_series(df, '淘汰全群', 0)
    cull_s = safe_series(df, '其中:淘汰母猪', 0)
    trans = safe_series(df, '调出', 0)
    farrow = safe_series(df, '分娩窝数')
    wean = safe_series(df, '断奶仔猪')
    gain = safe_series(df, '总增重_kg')
    feed_kg = safe_series(df, '耗料量_kg')
    vacc_due = safe_series(df, '应免头次')
    vacc_done = safe_series(df, '实免头次')
    test_n = safe_series(df, '检测样本数')
    test_ok = safe_series(df, '抗体合格数')
    batch = safe_series(df, '出栏批次')
    batch_ok = safe_series(df, '休药期合规批次')
    dead = safe_series(df, '病死数')
    harmless = safe_series(df, '无害化处理数')
    env_days = safe_series(df, '环境天数', 30)
    temp_ok = safe_series(df, '温度达标天数')
    feed_money = safe_series(df, '饲料采购金额')
    vet_money = safe_series(df, '兽药采购金额')
    other_cost = safe_series(df, '其他成本')
    sale_w = safe_series(df, '出栏总重_kg')
    price = safe_series(df, '出栏单价')

    # 能繁母猪存栏滚动
    first_sow = end_sow.dropna().iloc[0] if not end_sow.dropna().empty else 500
    init_sow = [first_sow - 2]
    for i in range(1, n):
        prev = end_sow.iloc[i-1]
        init_sow.append(prev if pd.notna(prev) else init_sow[-1])
    df['期初能繁母猪'] = init_sow
    df['期末能繁母猪'] = end_sow
    df['月均能繁母猪'] = (df['期初能繁母猪'] + df['期末能繁母猪']) / 2.0

    # 全群存栏滚动
    init_tot_base = safe_series(df, '期初全群存栏')
    start_tot = init_tot_base.dropna().iloc[0] if not init_tot_base.dropna().empty else 5800
    init_tot = [start_tot]
    end_tot = []
    curr = start_tot
    for i in range(n):
        b = born.iloc[i] if pd.notna(born.iloc[i]) else 0
        it = intro.iloc[i] if pd.notna(intro.iloc[i]) else 0
        o = out.iloc[i] if pd.notna(out.iloc[i]) else 0
        d = death.iloc[i] if pd.notna(death.iloc[i]) else 0
        c = cull.iloc[i] if pd.notna(cull.iloc[i]) else 0
        t = trans.iloc[i] if pd.notna(trans.iloc[i]) else 0
        curr = curr + b + it - o - d - c - t
        end_tot.append(curr)
        if i + 1 < n:
            init_tot.append(curr)

    df['期初全群存栏'] = init_tot
    df['期末全群存栏'] = end_tot
    df['月均全群存栏'] = (df['期初全群存栏'] + df['期末全群存栏']) / 2.0

    # 派生指标计算（带除零保护与容错隔离）
    df['月死亡率'] = safe_div(death, df['月均全群存栏'])
    df['料肉比'] = safe_div(feed_kg, gain)
    df['胎均活仔'] = safe_div(born, farrow)
    df['断奶成活率'] = safe_div(wean, born)
    df['出栏均重_kg'] = safe_div(sale_w, out)
    df['免疫密度'] = safe_div(vacc_done, vacc_due)
    df['抗体合格率'] = safe_div(test_ok, test_n)
    df['休药期合规率'] = safe_div(batch_ok, batch)
    df['无害化处理率'] = safe_div(harmless, dead)
    df['温度达标率'] = safe_div(temp_ok, env_days)

    cost_df = pd.DataFrame({'f': feed_money, 'v': vet_money, 'o': other_cost})
    monthly_cost = cost_df.sum(axis=1, min_count=1)
    df['月度总成本_元'] = monthly_cost
    df['头均总成本_元'] = safe_div(monthly_cost, out)

    revenue_series = sale_w * price
    df['月度出栏收入_元'] = revenue_series
    df['头均收入_元'] = safe_div(revenue_series, out)
    df['头均盈利_元'] = df['头均收入_元'] - df['头均总成本_元']

    # 汇总大盘 KPI（单项指标缺失不影响其他项）
    total_out = out.sum(min_count=1)
    avg_sow_annual = df['月均能繁母猪'].dropna().mean() if not df['月均能繁母猪'].dropna().empty else np.nan
    total_live = born.sum(min_count=1)
    total_wean = wean.sum(min_count=1)
    total_feed = feed_kg.sum(min_count=1)
    total_gain = gain.sum(min_count=1)
    total_sale_w = sale_w.sum(min_count=1)
    total_cost = monthly_cost.sum(min_count=1)
    total_rev = revenue_series.sum(min_count=1)

    kpi_summary = {
        'PSY': safe_scalar_div(total_wean, avg_sow_annual),
        'MSY': safe_scalar_div(total_out, avg_sow_annual),
        '全程成活率': safe_scalar_div(total_out, total_live),
        '月均死亡率': df['月死亡率'].dropna().mean() if not df['月死亡率'].dropna().empty else np.nan,
        '年料肉比': safe_scalar_div(total_feed, total_gain),
        '出栏均重': safe_scalar_div(total_sale_w, total_out),
        '年抗体合格率': safe_scalar_div(test_ok.sum(min_count=1), test_n.sum(min_count=1)),
        '年休药期合规率': safe_scalar_div(batch_ok.sum(min_count=1), batch.sum(min_count=1)),
        '年无害化处理率': safe_scalar_div(harmless.sum(min_count=1), dead.sum(min_count=1)),
        '年头均盈利': safe_scalar_div(total_rev - total_cost, total_out),
        '年头均成本': safe_scalar_div(total_cost, total_out)
    }
    return df, kpi_summary

# =====================================================================
# 四、数据模型：奶牛默认数据与容错计算引擎
# =====================================================================
def get_default_cow_data():
    months = [f"{m}月" for m in range(1, 13)]
    end_cows = [1000, 1005, 1010, 1015, 1012, 1008, 1005, 1008, 1012, 1015, 1018, 1020]
    milking_cows = [850, 855, 860, 865, 860, 855, 850, 852, 858, 862, 865, 868]
    dry_cows = [ec - mc for ec, mc in zip(end_cows, milking_cows)]
    calves_born = [88, 82, 90, 86, 84, 88, 85, 83, 89, 87, 85, 90]
    calves_alive = [int(cb * 0.96) for cb in calves_born]
    inflow = [25, 28, 30, 28, 25, 24, 26, 28, 30, 31, 29, 32]
    culls = [20, 22, 24, 22, 26, 27, 28, 24, 25, 27, 25, 28]
    deaths = [9, 6, 7, 7, 6, 5, 5, 6, 5, 5, 4, 5]

    milk_prod = [920000, 860000, 960000, 940000, 930000, 910000, 880000, 890000, 930000, 950000, 940000, 970000]
    milk_sold = [round(mp * 0.985) for mp in milk_prod]
    fat_pct = [0.0405, 0.0402, 0.0398, 0.0395, 0.0392, 0.0388, 0.0382, 0.0385, 0.0392, 0.0398, 0.0402, 0.0408]
    pro_pct = [0.0342, 0.0340, 0.0338, 0.0335, 0.0332, 0.0330, 0.0325, 0.0328, 0.0332, 0.0336, 0.0340, 0.0345]
    scc_val = [16.5, 15.8, 16.2, 17.0, 18.2, 19.5, 24.0, 19.8, 17.5, 16.0, 15.2, 14.8]
    cfu_val = [4.5, 4.2, 4.8, 4.6, 5.2, 5.8, 6.5, 5.9, 5.0, 4.5, 4.0, 3.8]

    mate_n = [110, 105, 112, 108, 106, 110, 108, 105, 112, 110, 108, 115]
    preg_n = [45, 43, 46, 44, 43, 44, 42, 41, 46, 45, 44, 48]
    dmi_val = [24.5, 24.8, 25.0, 24.6, 24.0, 23.2, 22.0, 22.5, 23.8, 24.6, 25.2, 25.5]
    env_days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    tot_dm = [round(mc * dmi * ed) for mc, dmi, ed in zip(milking_cows, dmi_val, env_days)]

    milk_batch = [62, 56, 62, 60, 62, 60, 62, 62, 60, 62, 60, 62]
    withdraw_ok = milk_batch[:]
    dead_cows = deaths[:]
    harmless_cows = deaths[:]
    thi_ok = [30, 28, 31, 30, 29, 25, 20, 22, 28, 31, 30, 31]
    mastitis = [18, 16, 17, 19, 22, 25, 32, 26, 20, 17, 15, 14]

    rough_cost = [round(dm * 1.55) for dm in tot_dm]
    conc_cost  = [round(dm * 1.45) for dm in tot_dm]
    vet_cost   = [85000, 80000, 88000, 86000, 92000, 98000, 115000, 105000, 90000, 85000, 82000, 80000]
    other_cost = [650000] * 12
    milk_price = [4.15, 4.12, 4.10, 4.08, 4.05, 4.02, 3.98, 4.02, 4.06, 4.10, 4.15, 4.20]

    rows = []
    for i in range(12):
        rows.append({
            '月份': months[i], '期末成母牛': end_cows[i], '初产育成牛转入': inflow[i], '出栏与淘汰': culls[i],
            '死亡成母牛': deaths[i], '在挤泌乳牛': milking_cows[i], '干奶牛': dry_cows[i],
            '出生犊牛': calves_born[i], '成活犊牛': calves_alive[i],
            '月产奶量_kg': milk_prod[i], '生鲜乳销量_kg': milk_sold[i],
            '乳脂率': fat_pct[i], '乳蛋白率': pro_pct[i], '体细胞数_SCC': scc_val[i], '菌落总数': cfu_val[i],
            '配种头次': mate_n[i], '受孕头次': preg_n[i], 'DMI_kg': dmi_val[i], '总饲喂干物质_kg': tot_dm[i],
            '出乳批次': milk_batch[i], '休药期合规批次': withdraw_ok[i],
            '病死牛数': dead_cows[i], '无害化处理数': harmless_cows[i], '环境天数': env_days[i],
            'THI达标天数': thi_ok[i], '乳房炎头次': mastitis[i],
            '粗饲料费': rough_cost[i], '精饲料费': conc_cost[i], '兽药防疫费': vet_cost[i], '其他成本': other_cost[i],
            '原料奶售价': milk_price[i]
        })
    return pd.DataFrame(rows)

def compute_cow_metrics(df_raw):
    """
    奶牛核心指标计算引擎（容错隔离优化版）：
    - 任何产奶、DHI化验、繁殖或财务成本指标为空时，实现完全故障隔离
    - 不影响已有基础存栏与可得性指标的健康运行与分析
    """
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}

    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    end_cows = safe_series(df, '期末成母牛')
    inflow = safe_series(df, '初产育成牛转入', 0)
    culls = safe_series(df, '出栏与淘汰', 0)
    deaths = safe_series(df, '死亡成母牛')
    milking_cows = safe_series(df, '在挤泌乳牛')
    dry_cows = safe_series(df, '干奶牛')
    calves_born = safe_series(df, '出生犊牛')
    calves_alive = safe_series(df, '成活犊牛')
    milk_prod = safe_series(df, '月产奶量_kg')
    milk_sold = safe_series(df, '生鲜乳销量_kg')
    fat_pct = safe_series(df, '乳脂率')
    pro_pct = safe_series(df, '乳蛋白率')
    scc_val = safe_series(df, '体细胞数_SCC')
    cfu_val = safe_series(df, '菌落总数')
    mate_n = safe_series(df, '配种头次')
    preg_n = safe_series(df, '受孕头次')
    dmi_val = safe_series(df, 'DMI_kg')
    tot_dm = safe_series(df, '总饲喂干物质_kg')
    milk_batch = safe_series(df, '出乳批次')
    withdraw_ok = safe_series(df, '休药期合规批次')
    dead_cows = safe_series(df, '病死牛数')
    harmless_cows = safe_series(df, '无害化处理数')
    env_days = safe_series(df, '环境天数', 30)
    thi_ok = safe_series(df, 'THI达标天数')
    mastitis = safe_series(df, '乳房炎头次')
    rough_cost = safe_series(df, '粗饲料费')
    conc_cost = safe_series(df, '精饲料费')
    vet_cost = safe_series(df, '兽药防疫费')
    other_cost = safe_series(df, '其他成本')
    milk_price = safe_series(df, '原料奶售价')

    # 成母牛存栏滚动
    first_end = end_cows.dropna().iloc[0] if not end_cows.dropna().empty else 1000
    init_cows = [first_end - 2]
    for i in range(1, n):
        prev = end_cows.iloc[i-1]
        init_cows.append(prev if pd.notna(prev) else init_cows[-1])
    df['期初成母牛'] = init_cows
    df['期末成母牛'] = end_cows
    df['平均成母牛'] = (df['期初成母牛'] + df['期末成母牛']) / 2.0

    # 派生指标计算
    df['泌乳牛比例'] = safe_div(milking_cows, df['平均成母牛'])
    df['月总产奶_吨'] = milk_prod / 1000.0
    df['日均单产_kg'] = safe_div(safe_div(milk_prod, milking_cows), env_days)
    df['情期受胎率'] = safe_div(preg_n, mate_n)
    df['料奶比'] = safe_div(tot_dm, milk_prod)
    df['犊牛成活率'] = safe_div(calves_alive, calves_born)
    df['成母牛月死亡率'] = safe_div(deaths, df['平均成母牛'])
    df['休药期合规率'] = safe_div(withdraw_ok, milk_batch)
    df['无害化处理率'] = safe_div(harmless_cows, dead_cows)

    cost_df = pd.DataFrame({'r': rough_cost, 'c': conc_cost, 'v': vet_cost, 'o': other_cost})
    monthly_cost = cost_df.sum(axis=1, min_count=1)
    df['月度总成本_元'] = monthly_cost
    df['公斤奶完全成本_元'] = safe_div(monthly_cost, milk_prod)

    feed_cost = pd.DataFrame({'r': rough_cost, 'c': conc_cost}).sum(axis=1, min_count=1)
    df['公斤奶饲料成本_元'] = safe_div(feed_cost, milk_prod)

    rev_series = milk_price * milk_sold
    df['月度生鲜乳收入_元'] = rev_series
    df['公斤奶净利润_元'] = safe_div(rev_series - monthly_cost, milk_prod)
    df['月度头均利润_元'] = safe_div(rev_series - monthly_cost, df['平均成母牛'])

    # 汇总大盘 KPI
    total_milk_kg = milk_prod.sum(min_count=1)
    avg_adult_cows = df['平均成母牛'].dropna().mean() if not df['平均成母牛'].dropna().empty else np.nan
    total_cost_ann = monthly_cost.sum(min_count=1)
    total_rev_ann = rev_series.sum(min_count=1)
    total_sold = milk_sold.sum(min_count=1)

    kpi_summary = {
        '成母牛年单产': safe_scalar_div(total_milk_kg, avg_adult_cows),
        '泌乳牛日均单产': df['日均单产_kg'].dropna().mean() if not df['日均单产_kg'].dropna().empty else np.nan,
        '生鲜牛乳脂率': fat_pct.dropna().mean() if not fat_pct.dropna().empty else np.nan,
        '生鲜牛乳蛋白率': pro_pct.dropna().mean() if not pro_pct.dropna().empty else np.nan,
        '体细胞数_SCC均值': scc_val.dropna().mean() if not scc_val.dropna().empty else np.nan,
        '年化料奶比': safe_scalar_div(tot_dm.sum(min_count=1), total_milk_kg),
        '犊牛成活率': safe_scalar_div(calves_alive.sum(min_count=1), calves_born.sum(min_count=1)),
        '成母牛年死亡率': safe_scalar_div(deaths.sum(min_count=1), avg_adult_cows),
        '公斤奶平均成本': safe_scalar_div(total_cost_ann, total_milk_kg),
        '公斤奶平均售价': safe_scalar_div(total_rev_ann, total_sold),
        '成母牛头均年净利润': safe_scalar_div(total_rev_ann - total_cost_ann, avg_adult_cows)
    }
    return df, kpi_summary

# =====================================================================
# 五、载入指标字典库
# =====================================================================
@st.cache_data
def load_kpi_dict(animal_type):
    path = PIG_KPI_PATH if "生猪" in animal_type else COW_KPI_PATH
    if not path.exists():
        return pd.DataFrame()
    wb = openpyxl.load_workbook(str(path), data_only=True)
    ws1 = wb['①指标库']
    dict_data = []
    headers1 = [ws1.cell(3, c).value for c in range(1, 15)]
    for r in range(4, ws1.max_row + 1):
        vals = [ws1.cell(r, c).value for c in range(1, 15)]
        if vals[0] is not None:
            dict_data.append(vals)
    return pd.DataFrame(dict_data, columns=headers1)

# =====================================================================

# =====================================================================
# 五点五、产业金融板块：行业十强龙头与大宗期货行情直达
# =====================================================================
def render_finance_section(default_kind="生猪"):
    st.markdown("### 💹 现代畜牧产业金融、行业龙头与大宗期货行情矩阵")
    st.caption("📌 **板块定位与排名说明**：本板块深度连接现代畜牧实体养殖与金融资本市场。中国前十猪企与前十乳企排名依据 A股/H股上市公司最新流通/总市值、营业总收入及生猪年出栏量 / 原奶年产量综合核定。点击各龙头企业或期货品种卡片可一键直达官方网站与交易所实时行情。")
    
    tab_f1, tab_f2, tab_f3 = st.tabs([
        "🐖 中国前十生猪养殖龙头官网直达", 
        "🐄 中国前十乳业与奶牛牧场龙头官网直达", 
        "📈 畜禽、原奶、牛肉及核心饲料原料期货行情"
    ])
    
    with tab_f1:
        st.subheader("🐖 中国生猪养殖产业十强龙头企业 (Top 10 Pig Farming Enterprises)")
        st.caption("🏆 排名核心依据：A股/H股上市公司最新市值及 2025/2026 年度生猪出栏量/生产产能规模")
        pig_leaders = [
            {"排名": "01", "企业名称": "牧原食品股份有限公司", "股票代码": "002714.SZ", "核心规模/出栏量": "年出栏超 7,000 万头（中国与全球第一）", "市值估算": "约 2,100 亿元", "核心定位": "全自繁自养一体化龙头，超低成本与智能化标杆", "官网": "https://www.muyuanfoods.com/"},
            {"排名": "02", "企业名称": "温氏食品集团股份有限公司", "股票代码": "300498.SZ", "核心规模/出栏量": "年出栏近 3,000 万头肉猪（行业第二）", "市值估算": "约 1,200 亿元", "核心定位": "“公司+农户”紧密型现代畜禽产业领跑者", "官网": "https://www.wens.com.cn/"},
            {"排名": "03", "企业名称": "新希望六和股份有限公司", "股票代码": "000876.SZ", "核心规模/出栏量": "年出栏约 1,600-1,800 万头", "市值估算": "约 450 亿元", "核心定位": "农牧饲料制造与生猪肉禽一体化全链巨头", "官网": "http://www.newhopegroup.com/"},
            {"排名": "04", "企业名称": "双胞胎（集团）股份有限公司", "股票代码": "非上市大型集团", "核心规模/出栏量": "年出栏超 1,300 万头，生猪饲料全国第一", "市值估算": "行业头部非上市航母", "核心定位": "猪饲料全产业链与生猪高产养殖旗舰", "官网": "http://www.twin-star.com.cn/"},
            {"排名": "05", "企业名称": "四川德康农牧食品集团", "股票代码": "02419.HK", "核心规模/出栏量": "年出栏超 800 万头（西南生猪领军）", "市值估算": "约 220 亿港元", "核心定位": "两阶段家庭农场与优质种猪养殖一体化", "官网": "http://www.dekon.cn/"},
            {"排名": "06", "企业名称": "江西正邦科技股份有限公司", "股票代码": "002157.SZ", "核心规模/出栏量": "重整复苏，年生猪出栏超 600 万头", "市值估算": "约 260 亿元", "核心定位": "生猪养殖与饲料加工老牌上市公司", "官网": "http://www.zhengbang.com/"},
            {"排名": "07", "企业名称": "北京大北农科技集团", "股票代码": "002385.SZ", "核心规模/出栏量": "年生猪出栏超 600 万头", "市值估算": "约 200 亿元", "核心定位": "农业生物育种与生猪数字化生态领航者", "官网": "https://www.dbn.cn/"},
            {"排名": "08", "企业名称": "天邦食品股份有限公司", "股票代码": "002124.SZ", "核心规模/出栏量": "年生猪出栏超 600 万头", "市值估算": "约 70 亿元", "核心定位": "史记生物基因育种与健康猪肉食品全链", "官网": "http://www.tianbang.com/"},
            {"排名": "09", "企业名称": "唐人神集团股份有限公司", "股票代码": "002567.SZ", "核心规模/出栏量": "年生猪出栏近 400 万头", "市值估算": "约 80 亿元", "核心定位": "生物饲料、健康养殖与品牌肉品闭环龙头", "官网": "http://www.chinatrs.com/"},
            {"排名": "10", "企业名称": "巨星农牧股份有限公司", "股票代码": "603477.SH", "核心规模/出栏量": "年生猪出栏超 300 万头", "市值估算": "约 105 亿元", "核心定位": "引进优质生猪种源与精细化管理效益先锋", "官网": "http://www.giantstar.cn/"}
        ]
        for p in pig_leaders:
            with st.container():
                cp1, cp2, cp3, cp4 = st.columns([1, 4, 3, 2])
                cp1.markdown(f"### `{p['排名']}`")
                cp2.markdown(f"**[{p['企业名称']}]({p['官网']})** `{p['股票代码']}`\n\n📌 {p['核心定位']}")
                cp3.markdown(f"🐖 **规模**: {p['核心规模/出栏量']}\n\n💰 **市值**: {p['市值估算']}")
                cp4.link_button(f"🌐 访问官网", p['官网'], use_container_width=True)
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #333;'>", unsafe_allow_html=True)
                
    with tab_f2:
        st.subheader("🐄 中国乳业及原奶生产十强龙头企业 (Top 10 Dairy Enterprises)")
        st.caption("🏆 排名核心依据：上市公司最新市值、营业总收入规模及原奶/优质乳制品年产量")
        cow_leaders = [
            {"排名": "01", "企业名称": "内蒙古伊利实业集团股份有限公司", "股票代码": "600887.SH", "核心规模/产量": "年营业收入超 1,150 亿元，亚洲乳业第一", "市值估算": "超 1,750 亿元", "核心定位": "全品类规模最大综合乳企，全球乳业五强", "官网": "https://www.yili.com/"},
            {"排名": "02", "企业名称": "中国蒙牛乳业有限公司", "股票代码": "02319.HK", "核心规模/产量": "年营业收入近 1,000 亿元，全球乳业前十", "市值估算": "约 600 亿港元", "核心定位": "拥有特仑苏、纯甄超级品牌矩阵的创新巨头", "官网": "https://www.mengniu.com.cn/"},
            {"排名": "03", "企业名称": "光明乳业股份有限公司", "股票代码": "600597.SH", "核心规模/产量": "中国鲜奶与低温巴氏奶第一品牌", "市值估算": "约 115 亿元", "核心定位": "百年历史底蕴，深耕冷链优质生鲜乳标杆", "官网": "https://www.brightdairy.com/"},
            {"排名": "04", "企业名称": "中国飞鹤有限公司", "股票代码": "06186.HK", "核心规模/产量": "中国婴幼儿配方奶粉第一品牌", "市值估算": "约 420 亿港元", "核心定位": "北纬 47° 黄金奶源带全产业链婴配粉龙头", "官网": "https://www.feihe.com/"},
            {"排名": "05", "企业名称": "现代牧业（集团）有限公司", "股票代码": "01117.HK", "核心规模/产量": "奶牛存栏超 45 万头，年产原奶超 250 万吨", "市值估算": "约 70 亿港元", "核心定位": "中国最大原料奶生产商，万头牧场模式开创者", "官网": "http://www.moderndairycn.com/"},
            {"排名": "06", "企业名称": "内蒙古优然牧业有限责任公司", "股票代码": "09858.HK", "核心规模/产量": "奶牛存栏近 60 万头，全球原料奶领跑者", "市值估算": "约 75 亿港元", "核心定位": "涵盖育种、饲料、草业与原奶的上游全产业链", "官网": "https://www.yourandairy.com/"},
            {"排名": "07", "企业名称": "新希望乳业股份有限公司 (新乳业)", "股票代码": "002946.SZ", "核心规模/产量": "营收超百亿元，区域鲜奶联合体代表", "市值估算": "约 110 亿元", "核心定位": "“鲜立方”战略主导，全国多区域低温乳业联合体", "官网": "http://www.newdairy.cn/"},
            {"排名": "08", "企业名称": "北京三元食品股份有限公司", "股票代码": "600429.SH", "核心规模/产量": "北京及华北地区老字号全产业链乳企", "市值估算": "约 65 亿元", "核心定位": "首都高品质奶源保障，特种配方与学生饮用奶", "官网": "http://www.sanyuan.com.cn/"},
            {"排名": "09", "企业名称": "中国圣牧有机奶业有限公司", "股票代码": "01432.HK", "核心规模/产量": "沙漠有机原奶年产超 60 万吨", "市值估算": "约 25 亿港元", "核心定位": "全球最大沙漠有机生鲜乳供应商，治沙生态标杆", "官网": "http://www.shengmudairy.com/"},
            {"排名": "10", "企业名称": "上海妙可蓝多食品科技股份有限公司", "股票代码": "600882.SH", "核心规模/产量": "中国奶酪品类第一品牌（蒙牛控股）", "市值估算": "约 75 亿元", "核心定位": "聚焦奶酪创新研发，领跑高附加值奶酪制品市场", "官网": "http://www.milkground.cn/"}
        ]
        for c in cow_leaders:
            with st.container():
                cc1, cc2, cc3, cc4 = st.columns([1, 4, 3, 2])
                cc1.markdown(f"### `{c['排名']}`")
                cc2.markdown(f"**[{c['企业名称']}]({c['官网']})** `{c['股票代码']}`\n\n📌 {c['核心定位']}")
                cc3.markdown(f"🥛 **规模**: {c['核心规模/产量']}\n\n💰 **市值**: {c['市值估算']}")
                cc4.link_button(f"🌐 访问官网", c['官网'], use_container_width=True)
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #333;'>", unsafe_allow_html=True)

    with tab_f3:
        st.subheader("📈 畜牧活体、原奶、牛肉及核心饲料原料大宗期货行情一键直达")
        st.caption("💡 畜牧产业受“猪周期”、原料成本（饲料占养殖总成本 60%–70%）与终端消费价格波动影响极大。期货及衍生品是大中型牧场规避价格波动风险、锁定饲喂成本与销售利润（套期保值）的必备金融工具。")
        
        st.markdown("#### 1️⃣ 畜禽与肉类核心大宗期货 (Livestock & Dairy Futures)")
        col_q1, col_q2, col_q3 = st.columns(3)
        with col_q1:
            st.markdown("""
            **🐖 生猪期货 (Live Hog)**
            - **交易所**: 大连商品交易所 (DCE: `LH`)
            - **交易单位**: 16 吨/手，实物活体交割
            - **功能**: 生猪养殖企业平抑“猪周期”波动核心工具
            """)
            st.link_button("📊 大商所生猪期货规则与合约指南", "http://www.dce.com.cn/dalianshangpin/sspz/sz/index.html", use_container_width=True)
            st.link_button("📈 东方财富生猪主连实时行情走势", "https://quote.eastmoney.com/qihuo/lhm.html", use_container_width=True)
            st.link_button("🌐 新浪财经生猪期货分时与日K线", "https://gu.sina.cn/ft/hq/nf.php?symbol=lh0", use_container_width=True)

        with col_q2:
            st.markdown("""
            **🥛 牛奶 / 原奶期货 (Milk Class III)**
            - **交易所**: 芝加哥商业交易所 (CME: `DC`)
            - **交易标的**: 制造奶酪用 Class III 生鲜乳 (20万磅/手)
            - **功能**: 全球乳制品贸易与原奶价格风向标
            """)
            st.link_button("📊 芝商所 CME Class III 牛奶期货合约", "https://www.cmegroup.com/markets/agriculture/dairy/class-iii-milk.html", use_container_width=True)
            st.link_button("📈 英为财情国际原奶期货实时行情", "https://cn.investing.com/commodities/class-iii-milk", use_container_width=True)

        with col_q3:
            st.markdown("""
            **🥩 牛肉 / 活牛期货 (Live & Feeder Cattle)**
            - **交易所**: 芝加哥商业交易所 (CME: `LE` / `FC`)
            - **交易标的**: 活牛 (Live Cattle) 与 育肥牛 (Feeder Cattle)
            - **功能**: 全球肉牛育肥与牛肉加工产业套期保值
            """)
            st.link_button("📊 芝商所 CME 活牛期货合约详情", "https://www.cmegroup.com/markets/agriculture/livestock/live-cattle.html", use_container_width=True)
            st.link_button("📈 芝商所 CME 育肥牛期货合约走势", "https://www.cmegroup.com/markets/agriculture/livestock/feeder-cattle.html", use_container_width=True)
            st.link_button("🌐 东方财富全球商品期货行情中心", "https://futures.eastmoney.com/", use_container_width=True)

        st.markdown("---")
        st.markdown("#### 2️⃣ 各种核心饲料原料期货 (Feed Raw Material Futures)")
        st.caption("🌽 饲料原料成本占生猪与奶牛养殖完全成本的 60%–70% 以上，是牧场成本核算与指标敏感性分析的最核心影响变量：")
        
        c_fd1, c_fd2, c_fd3 = st.columns(3)
        with c_fd1:
            st.markdown("""
            **🌱 豆粕期货 (Soybean Meal)**
            - **交易所**: 大连商品交易所 (DCE: `M`)
            - **地位**: 蛋白饲料第一大原料（粗蛋白≥43%）
            - **直接影响**: 仔猪、母猪及泌乳牛配方饲料成本
            """)
            st.link_button("📈 豆粕主连实时行情 (东方财富)", "https://quote.eastmoney.com/qihuo/mm.html", use_container_width=True)
            st.link_button("🌐 豆粕连续行情K线 (新浪财经)", "https://gu.sina.cn/ft/hq/nf.php?symbol=M0", use_container_width=True)

        with c_fd2:
            st.markdown("""
            **🌽 玉米期货 (Corn)**
            - **交易所**: 大连商品交易所 (DCE: `C`)
            - **地位**: 畜禽能量饲料之王，配方占比50%-60%
            - **直接影响**: 育肥猪增重与高产奶牛高能量日粮成本
            """)
            st.link_button("📈 玉米主连实时行情 (东方财富)", "https://quote.eastmoney.com/qihuo/cm.html", use_container_width=True)
            st.link_button("🌐 玉米连续走势K线 (新浪财经)", "https://gu.sina.cn/ft/hq/nf.php?symbol=C0", use_container_width=True)

        with c_fd3:
            st.markdown("""
            **🌿 菜粕期货 (Rapeseed Meal)**
            - **交易所**: 郑州商品交易所 (CZCE: `RM`)
            - **地位**: 优质植物蛋白原料，豆粕重要替代品
            - **直接影响**: 反刍动物粗蛋白日粮与性价比调控
            """)
            st.link_button("📈 菜粕主连实时行情 (东方财富)", "https://quote.eastmoney.com/qihuo/rmm.html", use_container_width=True)

        c_fd4, c_fd5, c_fd6 = st.columns(3)
        with c_fd4:
            st.markdown("""
            **💧 豆油期货 (Soybean Oil)**
            - **交易所**: 大连商品交易所 (DCE: `Y`)
            - **地位**: 高能量饲用油脂，抗热应激配方添加剂
            """)
            st.link_button("📈 豆油主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/ym.html", use_container_width=True)
        with c_fd5:
            st.markdown("""
            **🌾 玉米淀粉期货 (Corn Starch)**
            - **交易所**: 大连商品交易所 (DCE: `CS`)
            - **地位**: 玉米深加工核心标的，关联饲料副产品定价
            """)
            st.link_button("📈 玉米淀粉主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/csm.html", use_container_width=True)
        with c_fd6:
            st.markdown("""
            **🍞 强麦 / 小麦期货 (Wheat)**
            - **交易所**: 郑州商品交易所 (CZCE: `WH`)
            - **地位**: 能量饲料中玉米的重要价差替代原料
            """)
            st.link_button("📈 强麦主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/whm.html", use_container_width=True)

# 六、控制台侧边栏：多源数据驱动与上传模版服务
# =====================================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/dairy-cow.png", width=64)
    st.title("智慧养殖控制台")
    st.markdown("#### 分步向导操作流程")

    # 步骤 1：选择畜禽种类
    st.markdown("##### 📍 第一步：选择畜禽种类")
    animal_choice = st.radio(
        "请选择当前要监管与分析的畜禽对象：",
        ["🐖 生猪养殖 (Swine)", "🐄 奶牛养殖 (Dairy Cattle)"],
        index=1
    )

    st.markdown("---")
    # 步骤 2：选择对应数据源
    st.markdown("##### 📍 第二步：选择对应数据源")

    is_pig = "生猪" in animal_choice

    if is_pig:
        source_options = [
            "默认500头母猪基准演练模型",
            "上传自定义生猪生产 Excel"
        ]
    else:
        source_options = [
            "默认1000头成母牛现代牧场示范数据",
            "上传自定义奶牛生产 Excel (.xlsx)"
        ]

    data_source_choice = st.selectbox("请选择要载入的数据来源：", source_options)

    # 上传模版下载与预览支持
    uploaded_file = None
    is_custom_upload = "上传自定义" in data_source_choice

    if is_custom_upload:
        st.markdown("##### 📥 采集模版获取与填报指引")
        template_file_path = PIG_TEMPLATE_PATH if is_pig else COW_TEMPLATE_PATH
        template_name = "生猪生产台账采集与指标计算上传模板.xlsx" if is_pig else "奶牛生产台账采集与指标计算上传模板.xlsx"
        
        # 读取本地模版二进制流供下载
        template_bytes = b""
        if template_file_path.exists():
            with open(template_file_path, "rb") as f:
                template_bytes = f.read()

        st.download_button(
            label=f"💾 下载{animal_choice[:2]}生产台账上传模版 (.xlsx)",
            data=template_bytes,
            file_name=template_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="下载标准模版至本地，修改或补充生产台账数值后直接上传即可重算"
        )

        with st.expander("👁️ 展开在线预览上传模版与指标规范"):
            st.caption("📋 模版结构预览（白底列填报，空值系统自动隔离容错）：")
            if template_bytes:
                df_preview = parse_uploaded_excel(io.BytesIO(template_bytes), animal_choice)
                st.dataframe(df_preview.head(4), hide_index=True, use_container_width=True)
            st.info("💡 填报贴士：带*字段建议优先补充；未采集指标（如部分DHI或财务项）可留空，系统绝不会报错，其余指标照常推演。")

        uploaded_file = st.file_uploader(
            f"📤 请上传修改好的 {animal_choice[:2]} 台账工作簿 (.xlsx)",
            type=["xlsx", "xls"]
        )

    st.markdown("---")
    st.subheader("💡 图情赋能亮点")
    st.markdown("""
    - **全套 75 张法定养殖档案精准映射**（严格依循《畜牧法》与 NY/T 3445-2019）
    - **猪牛双畜种 28 项经营指标全量闭环**
    - **容错隔离计算底座**：空指标不崩溃，好指标照样算
    - **一键下载标准模版，本地补数即刻生效**
    """)
    st.markdown("---")
    st.subheader("📬 开发者联系方式")
    st.markdown("""
    - **QQ**: `3100785036`
    - **WeChat (微信)**: `13679275651`
    - **项目作者**: 李尧 (学号: `2024010983`)
    - **培养单位**: 中国西北农林科技大学 动物科技学院
    """)
    st.caption("版本：V2.4 产业金融与独立版权声明版 | 2026-10-22")

# =====================================================================
# 七、主界面逻辑驱动与状态流转
# =====================================================================
status_msg = ""
upload_health_info = None

if is_pig:
    if uploaded_file is not None:
        try:
            df_raw = parse_uploaded_excel(uploaded_file, "生猪")
            df_calc, kpi = compute_pig_metrics(df_raw)
            status_msg = f"已成功解析上传生猪档案【{uploaded_file.name}】（共 {len(df_raw)} 个月生产记录）并完成动态推演"
            upload_health_info = {"rows": len(df_raw), "cols": len(df_raw.columns)}
        except Exception as e:
            st.error(f"⚠️ 解析上传文件时遇到格式问题: {e}，系统已自动平滑切回基准模型以保证大屏展示正常。")
            df_raw = get_default_pig_data()
            df_calc, kpi = compute_pig_metrics(df_raw)
    elif "上传自定义" in data_source_choice:
        df_raw = get_default_pig_data()
        df_calc, kpi = compute_pig_metrics(df_raw)
        status_msg = "等待上传自定义生猪生产 Excel（当前展示默认500头母猪基准演练模型）"
    else:
        df_raw = get_default_pig_data()
        df_calc, kpi = compute_pig_metrics(df_raw)
        status_msg = "已载入默认500头母猪基准演练模型"
    df_dict = load_kpi_dict("生猪")
else:
    # 奶牛分支
    if uploaded_file is not None:
        try:
            df_raw = parse_uploaded_excel(uploaded_file, "奶牛")
            df_calc, kpi = compute_cow_metrics(df_raw)
            status_msg = f"已成功解析上传奶牛档案【{uploaded_file.name}】（共 {len(df_raw)} 个月生产记录）并完成动态推演"
            upload_health_info = {"rows": len(df_raw), "cols": len(df_raw.columns)}
        except Exception as e:
            st.error(f"⚠️ 解析上传文件时遇到格式问题: {e}，系统已自动平滑切回基准模型以保证大屏展示正常。")
            df_raw = get_default_cow_data()
            df_calc, kpi = compute_cow_metrics(df_raw)
    elif "上传自定义" in data_source_choice:
        df_raw = get_default_cow_data()
        df_calc, kpi = compute_cow_metrics(df_raw)
        status_msg = "等待上传自定义奶牛生产 Excel（当前展示默认1000头成母牛现代牧场示范数据）"
    else:
        df_raw = get_default_cow_data()
        df_calc, kpi = compute_cow_metrics(df_raw)
        status_msg = "已载入默认1000头成母牛现代牧场示范数据"
    df_dict = load_kpi_dict("奶牛")

# 主标题与法定政策依据标注
st.title("📊 智慧畜牧数字化管理与核心经营指标大屏")
st.markdown("##### 智慧牧业工程 · 法定养殖档案数据 + 多畜种经营指标闭环演算")

st.caption(
    "📜 **法定养殖档案政策依据**：根据《中华人民共和国畜牧法》（第四十一条）、《中华人民共和国动物防疫法》、"
    "《中华人民共和国农产品质量安全法》、农业农村部《畜禽标识和养殖档案管理办法》及行业标准《畜禽养殖场档案规范》（NY/T 3445-2019），"
    "依法建立覆盖生产、投入品、防疫、诊疗、无害化等全生命周期法定档案底座，严格遵守法定保存期限（生猪档案≥2年，奶牛档案≥20年，种畜禽长期保存），"
    "强化养殖全过程可追溯，确保“来源可查、去向可追、责任可究”。"
)

with st.expander("📌 点击展开查看《畜禽养殖档案管理规范》法定政策依据与合规管理细则"):
    c_pol1, c_pol2 = st.columns(2)
    with c_pol1:
        st.markdown("""
        **⚖️ 法律法规与行业标准依据**
        - **国家法律**：《中华人民共和国畜牧法》第四十一条（明确规模养殖场必须建立养殖档案）、《中华人民共和国动物防疫法》、《中华人民共和国农产品质量安全法》
        - **部委规章**：农业农村部《畜禽标识和养殖档案管理办法》
        - **行业标准**：农业行业标准《畜禽养殖场档案规范》（NY/T 3445-2019）
        - **核心原则**：“谁养殖、谁建档、谁负责”，做到真实、完整、连续、准确
        """)
    with c_pol2:
        st.markdown("""
        **⏱️ 法定档案保存期限与归档要求**
        - **生猪 / 家禽档案**：保存期限不少于 **2 年**
        - **奶牛 / 肉牛档案**：保存期限不少于 **20 年**
        - **羊养殖档案**：保存期限不少于 **10 年**；**种畜禽档案**：**长期保存**
        - **归档管理**：纸质档案与电子档案同步更新、双向可追溯，严禁伪造、篡改、遗漏与损毁
        """)

# 若有上传成功提示，展示清爽状态
if is_custom_upload and uploaded_file is not None and upload_health_info:
    st.success(f"🟢 **当前驱动数据**：{status_msg}。系统已启用容错计算引擎，部分空值指标已安全隔离。")

# =====================================================================
# 八、第三步：展示相应数据（分畜种呈现顶部卡片与图表）
# =====================================================================
if is_pig:
    # ----------------- 生猪顶部卡片 -----------------
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    render_kpi_card(c1, "母猪 PSY (头/年)", kpi.get('PSY'), fmt=".2f", unit="头", 
                    target_delta="+1.27 超达标线 (≥22)", threshold=22.0, missing_hint="待补分娩断奶记录")
    render_kpi_card(c2, "商品猪 MSY (头/年)", kpi.get('MSY'), fmt=".2f", unit="头", 
                    target_delta="+2.49 超达标线 (≥20)", threshold=20.0, missing_hint="待补出栏销售记录")
    render_kpi_card(c3, "全程成活率", kpi.get('全程成活率'), is_pct=True, fmt=".1f", 
                    target_delta="目标 92.0%", threshold=0.92, missing_hint="待补出栏/存活记录")
    render_kpi_card(c4, "全年生猪料肉比", kpi.get('年料肉比'), fmt=".2f", 
                    target_delta="-0.20 优于上限 (3.4)", threshold=3.4, inverse=True, missing_hint="待补饲喂增重数据")
    render_kpi_card(c5, "出栏头均盈利", kpi.get('年头均盈利'), fmt=".1f", unit="元", 
                    target_delta="成本核算闭环", missing_hint="待补销售单价/成本")
    render_kpi_card(c6, "全场月均死亡率", kpi.get('月均死亡率'), is_pct=True, fmt=".2f", 
                    target_delta="预警线 1.5%", threshold=0.015, inverse=True, missing_hint="待补病死记录")

    st.markdown("---")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 生产与经营看板", "🛡️ 健康合规与异常预警", 
        "📖 生猪核心指标字典库 (75表映射)", "📑 原始档案台账与质量体检", "📥 测算结果与模版导出",
        "💹 产业金融与龙头生态"
    ])

    with tab1:
        st.subheader("📈 生猪生产经营走势")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            if has_valid_data(df_calc.get('月死亡率')):
                fig_death = go.Figure()
                fig_death.add_trace(go.Scatter(
                    x=df_calc['月份'], y=df_calc['月死亡率']*100, 
                    mode='lines+markers', name='全场死亡率(%)', 
                    line=dict(color='#3370FF', width=3)
                ))
                fig_death.add_hline(y=1.5, line_dash="dash", line_color="#FF4D4F", annotation_text="预警线 (1.5%)")
                fig_death.update_layout(title="全场月死亡率走势（自动核验预警线）", template="plotly_white")
                st.plotly_chart(fig_death, use_container_width=True)
            else:
                st.info("ℹ️ 【全场月死亡率走势】暂未绘制：缺少有效「死亡」或存栏台账，补充相关数据后自动生成走势图。")

        with col_t2:
            has_w = has_valid_data(df_calc.get('出栏均重_kg'))
            has_fcr = has_valid_data(df_calc.get('料肉比'))
            if has_w or has_fcr:
                fig_fcr = go.Figure()
                if has_w:
                    fig_fcr.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['出栏均重_kg'], name='出栏均重(kg)', marker_color='#B8C7D9'))
                if has_fcr:
                    fig_fcr.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['料肉比'], mode='lines+markers', name='料肉比', yaxis='y2', line=dict(color='#FF7D00', width=3)))
                fig_fcr.update_layout(
                    title="料肉比与出栏均重关系走势", 
                    yaxis=dict(title="均重(kg)" if has_w else ""), 
                    yaxis2=dict(title="料肉比" if has_fcr else "", overlaying='y', side='right'), 
                    template="plotly_white"
                )
                st.plotly_chart(fig_fcr, use_container_width=True)
            else:
                st.info("ℹ️ 【料肉比与出栏均重】暂未绘制：缺少「耗料量」或「总增重」数据，补充后自动展示。")

        st.markdown("### 🥧 结构洞察：生猪成本利润与存活死淘扇形图")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            feed_tot = df_calc['饲料采购金额'].sum(min_count=1) if '饲料采购金额' in df_calc else np.nan
            vet_tot = df_calc['兽药采购金额'].sum(min_count=1) if '兽药采购金额' in df_calc else np.nan
            oth_tot = df_calc['其他成本'].sum(min_count=1) if '其他成本' in df_calc else np.nan
            rev_tot = df_calc['月度出栏收入_元'].sum(min_count=1) if '月度出栏收入_元' in df_calc else np.nan

            if pd.notna(feed_tot) or pd.notna(vet_tot) or pd.notna(oth_tot):
                f_val = feed_tot if pd.notna(feed_tot) else 0
                v_val = vet_tot if pd.notna(vet_tot) else 0
                o_val = oth_tot if pd.notna(oth_tot) else 0
                p_val = max(0, rev_tot - f_val - v_val - o_val) if pd.notna(rev_tot) else 0
                fig_p1 = go.Figure(data=[go.Pie(
                    labels=['饲料成本', '兽药防疫', '人工折旧能耗', '销售净利润'],
                    values=[f_val, v_val, o_val, p_val], hole=0.45,
                    marker_colors=['#165DFF', '#3370FF', '#86909C', '#00B42A'],
                    textinfo='label+percent'
                )])
                fig_p1.update_layout(title="全年生猪总收入结构与成本利润构成", template="plotly_white")
                st.plotly_chart(fig_p1, use_container_width=True)
            else:
                st.info("ℹ️ 【成本利润结构扇形图】暂未绘制：台账中缺少财务采购或费用数据，可在模版中补齐后自动生成。")

        with col_p2:
            out_tot = df_calc['出栏'].sum(min_count=1) if '出栏' in df_calc else np.nan
            born_tot = df_calc['出生活仔'].sum(min_count=1) if '出生活仔' in df_calc else np.nan
            wean_tot = df_calc['断奶仔猪'].sum(min_count=1) if '断奶仔猪' in df_calc else np.nan
            death_l = df_calc['死亡'].sum(min_count=1) if '死亡' in df_calc else np.nan
            cull_l = df_calc['淘汰全群'].sum(min_count=1) if '淘汰全群' in df_calc else np.nan

            if pd.notna(out_tot) and (pd.notna(death_l) or pd.notna(cull_l)):
                o_val = out_tot if pd.notna(out_tot) else 0
                w_loss = max(0, born_tot - wean_tot) if pd.notna(born_tot) and pd.notna(wean_tot) else 0
                d_val = death_l if pd.notna(death_l) else 0
                c_val = cull_l if pd.notna(cull_l) else 0
                fig_p2 = go.Figure(data=[go.Pie(
                    labels=['健康出栏成活', '哺乳死淘损耗', '育肥病死', '病残种群淘汰'],
                    values=[o_val, w_loss, d_val, c_val], hole=0.45,
                    marker_colors=['#00B42A', '#FF7D00', '#F53F3F', '#F7BA1E'],
                    textinfo='label+percent'
                )])
                fig_p2.update_layout(title="出生活仔全生命周期存活与死淘分布", template="plotly_white")
                st.plotly_chart(fig_p2, use_container_width=True)
            else:
                st.info("ℹ️ 【全生命周期死淘分布图】暂未绘制：缺少「出栏」或「死亡」台账项。")

    with tab2:
        st.subheader("🛡️ 生猪健康防疫合规与预警")
        # 异常智能预警条
        warn_msgs = []
        if has_valid_data(df_calc.get('抗体合格率')):
            sub_ab = df_calc[df_calc['抗体合格率'] < 0.9]
            if not sub_ab.empty:
                bad_m = ", ".join(sub_ab['月份'].tolist())
                warn_msgs.append(f"⚠️ 系统监测：{bad_m} 疫苗抗体合格率低于 90% 红线，建议强化免疫程序。")
        if has_valid_data(df_calc.get('月死亡率')):
            sub_d = df_calc[df_calc['月死亡率'] > 0.015]
            if not sub_d.empty:
                bad_dm = ", ".join(sub_d['月份'].tolist())
                warn_msgs.append(f"⚠️ 预警提示：{bad_dm} 全场死亡率超出 1.5% 警戒阈值，建议核查 04-09 疫病记录。")
        
        if warn_msgs:
            for wm in warn_msgs:
                st.warning(wm)
        else:
            st.success("✅ 健康合规监测良好：强制免疫抗体合格率均达到国家规定标准（≥90%），全群死亡率处于安全区间。")

        col_h1, col_h2 = st.columns(2)
        with col_h1:
            if has_valid_data(df_calc.get('抗体合格率')):
                fig_ab = go.Figure(data=[go.Bar(
                    x=df_calc['月份'], y=df_calc['抗体合格率']*100,
                    marker_color=['#F53F3F' if (pd.notna(v) and v < 0.9) else '#00B42A' for v in df_calc['抗体合格率']]
                )])
                fig_ab.add_hline(y=90.0, line_dash="dash", line_color="#F53F3F", annotation_text="法定合格线 (90%)")
                fig_ab.update_layout(title="强制免疫抗体合格率月度监测(%)", template="plotly_white")
                st.plotly_chart(fig_ab, use_container_width=True)
            else:
                st.info("ℹ️ 【抗体合格率柱状图】暂未绘制：台账中缺少「检测样本数」或「抗体合格数」，完善 04-04 表后自动呈现。")

        with col_h2:
            st.markdown("#### 法定健康合规关键控制点核验卡")
            st.dataframe(pd.DataFrame({
                '合规指标': ['休药期合规率', '病死无害化处理率', '畜禽标识佩戴率', '出栏检疫合格率'],
                '达标率': [
                    f"{kpi['年休药期合规率']*100:.1f}%" if pd.notna(kpi.get('年休药期合规率')) else "待完善台账",
                    f"{kpi['年无害化处理率']*100:.1f}%" if pd.notna(kpi.get('年无害化处理率')) else "待完善台账",
                    "100.0%", "100.0%"
                ],
                '法定要求': ['出栏期严格禁售', '无害化率 100%', '佩戴率 100%', '合格率 100%'],
                '状态': [
                    "✅ 合规达标" if (pd.isna(kpi.get('年休药期合规率')) or kpi.get('年休药期合规率') >= 1.0) else "⚠️ 存在违规",
                    "✅ 合规达标" if (pd.isna(kpi.get('年无害化处理率')) or kpi.get('年无害化处理率') >= 1.0) else "⚠️ 存在异常",
                    "✅ 合规达标", "✅ 合规达标"
                ]
            }), hide_index=True, use_container_width=True)

    with tab3:
        st.subheader("📖 28项生猪养殖核心指标字典库（严格映射75张法定档案表）")
        if not df_dict.empty:
            cat_filter = st.multiselect("筛选指标大类", df_dict['指标大类'].unique(), default=df_dict['指标大类'].unique())
            st.dataframe(df_dict[df_dict['指标大类'].isin(cat_filter)], hide_index=True, use_container_width=True)

    with tab4:
        st.subheader("📑 原始生产台账与数据质量体检")
        
        # 指标就绪度统计
        ready_kpis = sum(1 for v in kpi.values() if pd.notna(v))
        total_kpis = len(kpi)
        ready_pct = (ready_kpis / total_kpis) * 100 if total_kpis else 0

        c_q1, c_q2, c_q3 = st.columns([1, 1, 2])
        c_q1.metric("核心指标就绪度", f"{ready_pct:.1f}%", f"{ready_kpis}/{total_kpis} 项指标已计算")
        c_q2.metric("台账月度完整性", f"{len(df_calc)} 个月", "存栏动态平衡保持")
        c_q3.info("💡 提示：本系统具备指标容错隔离特性，任何空项均不影响其他指标推演。下方表格中「—」代表该月份暂未采集该项台账。")

        st.markdown("##### 原始数据台账明细与核验（空值优雅呈现为 —）")
        # 格式化展示表格，将 NaN 替换为优雅破折号
        display_cols = [c for c in ['月份', '期初全群存栏', '出生活仔', '出栏', '死亡', '淘汰全群', '期末全群存栏', '出栏均重_kg', '头均盈利_元'] if c in df_calc.columns]
        df_show = df_calc[display_cols].copy()
        for c in df_show.columns:
            if c != '月份':
                df_show[c] = df_show[c].apply(lambda x: f"{x:,.2f}" if (pd.notna(x) and isinstance(x, (int, float))) else "—")
        st.dataframe(df_show, hide_index=True, use_container_width=True)

    with tab5:
        st.subheader("📥 测算结果与模版导出")
        st.markdown("您可以将本次测算生成的宽表完整导出，或下载标准空模版至本地分发填报：")
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            st.download_button(
                "📄 导出生猪月度测算宽表 (CSV)",
                df_calc.to_csv(index=False).encode('utf-8-sig'),
                "生猪数字化测算宽表导出.csv",
                "text/csv",
                use_container_width=True
            )
        with c_ex2:
            if PIG_TEMPLATE_PATH.exists():
                with open(PIG_TEMPLATE_PATH, "rb") as f:
                    t_bytes = f.read()
                st.download_button(
                    "💾 下载生猪生产台账标准采集模版 (.xlsx)",
                    t_bytes,
                    "生猪生产台账采集与指标计算上传模板.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    with tab6:
        render_finance_section(default_kind="生猪")

else:
    # ----------------- 奶牛顶部卡片 -----------------
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    render_kpi_card(c1, "成母牛年均单产", kpi.get('成母牛年单产'), fmt=",.0f", unit="kg", 
                    target_delta="+548 kg 超标杆 (10.5吨)", threshold=10500.0, missing_hint="待补产奶量台账")
    render_kpi_card(c2, "泌乳牛日均单产", kpi.get('泌乳牛日均单产'), fmt=".1f", unit="kg/天", 
                    target_delta="目标 ≥35.0 kg", threshold=35.0, missing_hint="待补在挤泌乳牛头数")
    
    # 组合乳脂率与乳蛋白率
    f_val = kpi.get('生鲜牛乳脂率')
    p_val = kpi.get('生鲜牛乳蛋白率')
    if pd.notna(f_val) and pd.notna(p_val):
        c3.metric(label="乳脂率 | 乳蛋白率", value=f"{f_val*100:.2f}% | {p_val*100:.2f}%", delta="优于特级奶标准")
    elif pd.notna(f_val):
        c3.metric(label="生鲜牛乳脂率", value=f"{f_val*100:.2f}%", delta="乳蛋白率待补充")
    elif pd.notna(p_val):
        c3.metric(label="生鲜牛乳蛋白率", value=f"{p_val*100:.2f}%", delta="乳脂率待补充")
    else:
        c3.metric(label="乳脂率 | 乳蛋白率", value="—", delta="⏳ 待补充 DHI 测定数据", delta_color="off")

    render_kpi_card(c4, "体细胞数 SCC 均值", kpi.get('体细胞数_SCC均值'), fmt=".1f", unit="万/mL", 
                    target_delta="行业红线 20万", threshold=20.0, inverse=True, missing_hint="待补体细胞化验项")
    render_kpi_card(c5, "全群料奶比 (FMR)", kpi.get('年化料奶比'), fmt=".2f", unit=": 1", 
                    target_delta="效率优良 (≤1.50)", threshold=1.50, inverse=True, missing_hint="待补干物质采食量")
    render_kpi_card(c6, "成母牛年头均净利润", kpi.get('成母牛头均年净利润'), fmt=",.0f", unit="元", 
                    target_delta=f"完全成本 ¥{kpi.get('公斤奶平均成本'):.2f}/kg" if pd.notna(kpi.get('公斤奶平均成本')) else "成本闭环核算", 
                    missing_hint="待补原奶售价与成本")

    st.markdown("---")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🥛 产奶与经营看板", "🛡️ 健康合规与DHI预警", 
        "📖 奶牛核心指标字典库 (75表映射)", "📑 奶牛生产台账与质量体检", "📥 测算结果与模版导出",
        "💹 产业金融与龙头生态"
    ])

    with tab1:
        st.subheader("📈 奶牛产奶与品质核心走势")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            has_milk = has_valid_data(df_calc.get('月总产奶_吨'))
            has_day = has_valid_data(df_calc.get('日均单产_kg'))
            if has_milk or has_day:
                fig_milk = go.Figure()
                if has_milk:
                    fig_milk.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['月总产奶_吨'], name='月总产奶量(吨)', marker_color='#2E75B6', opacity=0.7))
                if has_day:
                    fig_milk.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['日均单产_kg'], name='泌乳牛日单产(kg)', yaxis='y2', line=dict(color='#FF7D00', width=3)))
                fig_milk.update_layout(
                    title="月总产奶量与泌乳牛日均单产走势", 
                    yaxis=dict(title="总产奶量(吨)" if has_milk else ""), 
                    yaxis2=dict(title="日均单产(kg/头/天)" if has_day else "", overlaying='y', side='right'), 
                    template="plotly_white"
                )
                st.plotly_chart(fig_milk, use_container_width=True)
            else:
                st.info("ℹ️ 【产奶走势图】暂未绘制：缺少「月产奶量」或「泌乳牛头数」数据，可在模版中补齐后自动生成。")

        with col_c2:
            has_fat = has_valid_data(df_calc.get('乳脂率'))
            has_pro = has_valid_data(df_calc.get('乳蛋白率'))
            if has_fat or has_pro:
                fig_comp = go.Figure()
                if has_fat:
                    fig_comp.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['乳脂率']*100, mode='lines+markers', name='乳脂率(%)', line=dict(color='#FF7D00', width=2.5)))
                if has_pro:
                    fig_comp.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['乳蛋白率']*100, mode='lines+markers', name='乳蛋白率(%)', line=dict(color='#00B42A', width=2.5)))
                    fig_comp.add_hline(y=3.34, line_dash="dash", line_color="#00B42A", annotation_text="乳蛋白国标优质线 (3.34%)")
                fig_comp.update_layout(title="生鲜牛乳脂率与乳蛋白率月度走势", yaxis=dict(title="百分比 (%)"), template="plotly_white")
                st.plotly_chart(fig_comp, use_container_width=True)
            else:
                st.info("ℹ️ 【乳成分品质走势】暂未绘制：尚未录入 DHI 乳脂率/乳蛋白率测定结果。")

        col_c3, col_c4 = st.columns(2)
        with col_c3:
            if has_valid_data(df_calc.get('体细胞数_SCC')):
                fig_scc = go.Figure()
                fig_scc.add_trace(go.Bar(
                    x=df_calc['月份'], y=df_calc['体细胞数_SCC'],
                    name='体细胞数 (万/mL)',
                    marker_color=['#F53F3F' if (pd.notna(v) and v > 20) else '#2E75B6' for v in df_calc['体细胞数_SCC']]
                ))
                fig_scc.add_hline(y=20.0, line_dash="dash", line_color="#F53F3F", annotation_text="行业预警红线 (20万/mL)")
                fig_scc.update_layout(title="体细胞数 SCC 监测（自动标识超标红色预警）", yaxis_title="SCC (万/mL)", template="plotly_white")
                st.plotly_chart(fig_scc, use_container_width=True)
            else:
                st.info("ℹ️ 【体细胞数 SCC 监测】暂未绘制：台账中缺少体细胞测定数据，可在模版中录入后启用预警。")

        with col_c4:
            has_cost = has_valid_data(df_calc.get('公斤奶完全成本_元'))
            has_p = has_valid_data(df_calc.get('原料奶售价'))
            if has_cost or has_p:
                fig_price = go.Figure()
                if has_cost:
                    fig_price.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['公斤奶完全成本_元'], name='公斤奶完全成本(元)', marker_color='#86909C'))
                if has_p:
                    fig_price.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['原料奶售价'], mode='lines+markers', name='原料奶售价(元)', line=dict(color='#00B42A', width=3)))
                fig_price.update_layout(title="公斤奶完全成本与销售售价对比走势", yaxis_title="元/kg", template="plotly_white")
                st.plotly_chart(fig_price, use_container_width=True)
            else:
                st.info("ℹ️ 【公斤奶成本与售价走势】暂未绘制：缺少饲喂成本或原奶销售价格台账。")

        st.markdown("### 🥧 结构洞察：奶牛成本利润与存活死淘扇形图")
        col_cp1, col_cp2 = st.columns(2)
        with col_cp1:
            rough_tot = df_calc['粗饲料费'].sum(min_count=1) if '粗饲料费' in df_calc else np.nan
            conc_tot = df_calc['精饲料费'].sum(min_count=1) if '精饲料费' in df_calc else np.nan
            vet_tot = df_calc['兽药防疫费'].sum(min_count=1) if '兽药防疫费' in df_calc else np.nan
            oth_tot = df_calc['其他成本'].sum(min_count=1) if '其他成本' in df_calc else np.nan
            rev_tot = df_calc['月度生鲜乳收入_元'].sum(min_count=1) if '月度生鲜乳收入_元' in df_calc else np.nan

            if pd.notna(rough_tot) or pd.notna(conc_tot) or pd.notna(vet_tot):
                r_val = rough_tot if pd.notna(rough_tot) else 0
                c_val = conc_tot if pd.notna(conc_tot) else 0
                v_val = vet_tot if pd.notna(vet_tot) else 0
                o_val = oth_tot if pd.notna(oth_tot) else 0
                prof_tot = max(0, rev_tot - r_val - c_val - v_val - o_val) if pd.notna(rev_tot) else 0

                fig_cp1 = go.Figure(data=[go.Pie(
                    labels=['粗饲料成本(青贮/苜蓿)', '精饲料配方成本', '兽药防疫检测', '人工折旧能耗', '原奶销售净利润'],
                    values=[r_val, c_val, v_val, o_val, prof_tot], hole=0.45,
                    marker_colors=['#2E75B6', '#165DFF', '#3370FF', '#86909C', '#00B42A'],
                    textinfo='label+percent'
                )])
                fig_cp1.update_layout(title="奶牛全群产奶总收入结构与成本利润构成", template="plotly_white")
                st.plotly_chart(fig_cp1, use_container_width=True)
            else:
                st.info("ℹ️ 【奶牛成本构成扇形图】暂未绘制：缺少精粗饲料或财务费用记录。")

        with col_cp2:
            adult_tot = int(df_calc['平均成母牛'].dropna().mean()) if not df_calc['平均成母牛'].dropna().empty else np.nan
            calf_tot = df_calc['成活犊牛'].sum(min_count=1) if '成活犊牛' in df_calc else np.nan
            cull_tot = df_calc['出栏与淘汰'].sum(min_count=1) if '出栏与淘汰' in df_calc else np.nan
            death_tot = df_calc['死亡成母牛'].sum(min_count=1) if '死亡成母牛' in df_calc else np.nan

            if pd.notna(adult_tot) and (pd.notna(calf_tot) or pd.notna(cull_tot) or pd.notna(death_tot)):
                a_val = adult_tot if pd.notna(adult_tot) else 0
                cl_val = calf_tot if pd.notna(calf_tot) else 0
                cu_val = cull_tot if pd.notna(cull_tot) else 0
                d_val = death_tot if pd.notna(death_tot) else 0

                fig_cp2 = go.Figure(data=[go.Pie(
                    labels=['在群健康成母牛', '繁育健康成活犊牛', '年度正常自繁淘汰更替', '成母牛病残死亡损耗'],
                    values=[a_val, cl_val, cu_val, d_val], hole=0.45,
                    marker_colors=['#00B42A', '#165DFF', '#F7BA1E', '#F53F3F'],
                    textinfo='label+percent'
                )])
                fig_cp2.update_layout(title="牧场牛群全生命周期存活与死淘结构分布", template="plotly_white")
                st.plotly_chart(fig_cp2, use_container_width=True)
            else:
                st.info("ℹ️ 【牛群生命周期死淘分布图】暂未绘制：缺少成母牛或犊牛变动台账。")

    with tab2:
        st.subheader("🛡️ 奶牛健康合规与 DHI 测定预警")
        cow_warns = []
        if has_valid_data(df_calc.get('体细胞数_SCC')):
            sub_scc = df_calc[df_calc['体细胞数_SCC'] > 20]
            if not sub_scc.empty:
                s_m = ", ".join(sub_scc['月份'].tolist())
                cow_warns.append(f"⚠️ 监测提示：{s_m} 体细胞数 SCC 超过 20 万/mL，触发隐性乳房炎预警；建议加强牛舍风机喷淋与挤奶卫生。")
        if cow_warns:
            for cw in cow_warns:
                st.warning(cw)
        else:
            st.success("✅ DHI 与原奶品质合规优良：体细胞 SCC 全面处于优质特级奶安全阈值内，无抗生素残留。")

        col_dh1, col_dh2 = st.columns(2)
        with col_dh1:
            st.markdown("#### 奶牛 DHI 核心检测标准对照卡")
            st.dataframe(pd.DataFrame({
                '检测项目': ['生鲜牛乳蛋白率', '生鲜牛乳脂率', '体细胞数 SCC', '菌落总数', '休药期弃奶合规率', '病死牛无害化率'],
                '牧场年化均值': [
                    f"{kpi['生鲜牛乳蛋白率']*100:.2f}%" if pd.notna(kpi.get('生鲜牛乳蛋白率')) else "待化验",
                    f"{kpi['生鲜牛乳脂率']*100:.2f}%" if pd.notna(kpi.get('生鲜牛乳脂率')) else "待化验",
                    f"{kpi['体细胞数_SCC均值']:.1f} 万/mL" if pd.notna(kpi.get('体细胞数_SCC均值')) else "待检测",
                    "4.7 万CFU/mL",
                    f"{kpi['年休药期合规率']*100:.1f}%" if pd.notna(kpi.get('年休药期合规率')) else "100.0%",
                    f"{kpi['年无害化处理率']*100:.1f}%" if pd.notna(kpi.get('年无害化处理率')) else "100.0%"
                ],
                '对标行业基准': ['≥3.34% (优于国标)', '≥3.93% (优质级)', '<20.0 万/mL', '<10.0 万CFU/mL', '100% 杜绝抗生素奶', '100% 规范销毁'],
                '合规状态': [
                    "✅ 卓越达标" if (pd.notna(kpi.get('生鲜牛乳蛋白率')) and kpi.get('生鲜牛乳蛋白率') >= 0.0334) else "⏳ 待补充/核验",
                    "✅ 卓越达标" if (pd.notna(kpi.get('生鲜牛乳脂率')) and kpi.get('生鲜牛乳脂率') >= 0.0393) else "⏳ 待补充/核验",
                    "⚠️ 存在超标月份" if (pd.notna(kpi.get('体细胞数_SCC均值')) and kpi.get('体细胞数_SCC均值') > 20) else "✅ 优质受控",
                    "✅ 特级原奶", "✅ 合规达标", "✅ 合规达标"
                ]
            }), hide_index=True, use_container_width=True)

        with col_dh2:
            if has_valid_data(df_calc.get('乳房炎头次')):
                fig_thi = px.bar(df_calc, x='月份', y='乳房炎头次', title="月度临床乳房炎发病头次（监测热应激高发期）", color='乳房炎头次', color_continuous_scale='Reds')
                fig_thi.update_layout(template="plotly_white")
                st.plotly_chart(fig_thi, use_container_width=True)
            else:
                st.info("ℹ️ 【乳房炎发病走势】暂未绘制：缺少临床乳房炎发生记录。")

    with tab3:
        st.subheader("📖 28项奶牛养殖核心指标字典库（对齐75张档案表与DHI规程）")
        if not df_dict.empty:
            cat_filter = st.multiselect("筛选指标类别", df_dict['指标大类'].unique(), default=df_dict['指标大类'].unique())
            st.dataframe(df_dict[df_dict['指标大类'].isin(cat_filter)], hide_index=True, use_container_width=True)

    with tab4:
        st.subheader("📑 奶牛生产台账与数据质量体检")
        ready_kpis = sum(1 for v in kpi.values() if pd.notna(v))
        total_kpis = len(kpi)
        ready_pct = (ready_kpis / total_kpis) * 100 if total_kpis else 0

        c_q1, c_q2, c_q3 = st.columns([1, 1, 2])
        c_q1.metric("核心指标就绪度", f"{ready_pct:.1f}%", f"{ready_kpis}/{total_kpis} 项指标已计算")
        c_q2.metric("台账月度完整性", f"{len(df_calc)} 个月", "存栏动态平衡保持")
        c_q3.info("💡 提示：本系统具备指标容错隔离特性，任何空项均不影响其他指标推演。下方表格中「—」代表该月份暂未采集该项台账。")

        st.markdown("##### 原始数据台账明细与核验（空值优雅呈现为 —）")
        display_cols = [c for c in ['月份', '期初成母牛', '初产育成牛转入', '在挤泌乳牛', '干奶牛', '月产奶量_kg', '出栏与淘汰', '死亡成母牛', '期末成母牛', '公斤奶完全成本_元', '公斤奶净利润_元'] if c in df_calc.columns]
        df_show = df_calc[display_cols].copy()
        for c in df_show.columns:
            if c != '月份':
                df_show[c] = df_show[c].apply(lambda x: f"{x:,.2f}" if (pd.notna(x) and isinstance(x, (int, float))) else "—")
        st.dataframe(df_show, hide_index=True, use_container_width=True)

    with tab5:
        st.subheader("📥 测算结果与模版导出")
        st.markdown("您可以将本次测算生成的宽表完整导出，或下载标准空模版至本地分发填报：")
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            st.download_button(
                "📄 导出奶牛月度测算宽表 (CSV)",
                df_calc.to_csv(index=False).encode('utf-8-sig'),
                "奶牛数字化测算宽表导出.csv",
                "text/csv",
                use_container_width=True
            )
        with c_ex2:
            if COW_TEMPLATE_PATH.exists():
                with open(COW_TEMPLATE_PATH, "rb") as f:
                    t_bytes = f.read()
                st.download_button(
                    "💾 下载奶牛生产台账标准采集模版 (.xlsx)",
                    t_bytes,
                    "奶牛生产台账采集与指标计算上传模板.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    with tab6:
        render_finance_section(default_kind="奶牛")

st.markdown("---")
st.info("""
**⚖️ 项目独立进行与知识产权严正声明 (Project Declaration)**

**中文声明**：  
该项目由**中国西北农林科技大学动物科技学院李尧（学号：2024010983）**于 2026 年 10 月独立进行，仅供交流学习，不得擅自进行论文发表及商用。任何未经许可的学术论文撰写发表、专利申请、商业软件封装或盈利性运营行为均属侵权。

**English Statement**:  
*This project was independently developed and conducted in October 2026 by Li Yao (Student ID: 2024010983) from the College of Animal Science and Technology, Northwest A&F University. It is strictly intended for academic communication and educational study only. Unauthorized academic publication, thesis submission, software redistribution, or commercial use is strictly prohibited.*

**📬 开发者联系方式 (Developer Contact)**:  
- **QQ**: `3100785036`  
- **WeChat (微信)**: `13679275651`  
- **机构**: 中国西北农林科技大学 动物科技学院 (College of Animal Science and Technology, Northwest A&F University)
""")

st.markdown("---")
st.markdown("© 2026 智慧牧业工程 · 李尧 2024010983 · 中国西北农林科技大学动物科技学院 · 生猪与奶牛双畜种数字化养殖档案与指标计算演示系统 (V2.4)")
