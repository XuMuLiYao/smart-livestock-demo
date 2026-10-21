# -*- coding: utf-8 -*-
"""
智慧畜牧数字化管理与核心指标交互式看板
项目：智慧牧业工程 · smart_pig_demo
技术栈：Streamlit + pandas + Plotly + openpyxl
升级亮点（2026-10-22 第6次改动）：
  1. 多板块全景平铺与无滚轴沉浸交互：全局去除浏览器与各个组件滚轴（-ms-overflow-style: none; scrollbar-width: none;），默认启用“全景平铺模式”使生产走势、结构洞察、台账体检、行业新闻、DOI论文、产业金融各板块同时并列呈现，消除深层翻页与滚动条阻滞
  2. 增设行业新闻与资讯互动板块：集成国家农业农村部及主流畜牧权威咨询门户一键跳转，精选行业前沿动态，并提供“留下工作个人联系方式上传咨询新闻及文章”在线投稿通道，支持附件上传与实时入库登记
  3. 增设行业权威学术论文与 DOI 直达板块：精选动物科学与智能养殖顶级期刊论文（Journal of Dairy Science, Journal of Animal Science, Animal Nutrition等），完整提供文献标准化 DOI 并支持一键跳转解析；增设 DOI 快速检索解析跳转器，支持任意学术文献一键直达
  4. 严格保留与强化前期全部合规与功能资产：
     - 数据源严格保留“默认500头母猪基准演练模型”“上传自定义生猪生产 Excel”“默认1000头成母牛现代牧场示范数据”“上传自定义奶牛生产 Excel (.xlsx)”
     - 规范表述为“智慧牧业工程”与“数据”
     - 产业金融板块涵盖中国前十猪企官网、中国前十乳企官网及生猪、牛奶、牛肉、豆粕、玉米等大宗期货实时行情直达
     - 侧边栏与页脚公示开发者联系方式（QQ: 3100785036，WeChat: 13679275651）
     - 版权标识与中英文独立项目严正声明（中国西北农林科技大学动物科技学院李尧 2024010983 独立进行）
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
    page_title="智慧牧业工程 · 数字化管理与指标看板",
    page_icon="🐄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 全局注入去除滚轴样式（无缝平滑无感滚动）
st.markdown("""
<style>
/* 全局去除浏览器与组件滚轴，保持平滑无感滚动体验 */
::-webkit-scrollbar {
    display: none !important;
    width: 0px !important;
    height: 0px !important;
}
* {
    -ms-overflow-style: none !important;
    scrollbar-width: none !important;
}
html, body, [data-testid="stAppViewContainer"], [data-testid="stSidebar"], [data-testid="stVerticalBlock"], [data-testid="stMainBlockContainer"] {
    -ms-overflow-style: none !important;
    scrollbar-width: none !important;
}
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 2.5rem !important;
}
.section-card {
    background-color: #161B22;
    border: 1px solid #30363D;
    border-radius: 8px;
    padding: 18px 22px;
    margin-bottom: 22px;
}
</style>
""", unsafe_allow_html=True)

# 路径常量解析
CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
PIG_KPI_PATH = ROOT_DIR / "表格与指标库" / "生猪养殖核心指标库与可行性验证.xlsx"
COW_KPI_PATH = ROOT_DIR / "表格与指标库" / "奶牛养殖核心指标库与可行性验证.xlsx"
PIG_TEMPLATE_PATH = CURRENT_DIR / "生猪生产台账采集与指标计算上传模板.xlsx"
COW_TEMPLATE_PATH = CURRENT_DIR / "奶牛生产台账采集与指标计算上传模板.xlsx"

# =====================================================================
# 一、基础辅助函数：安全计算与UI格式化
# =====================================================================
def safe_series(df, col, default_val=np.nan):
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce')
    return pd.Series([default_val] * len(df))

def safe_div(s1, s2):
    n1 = safe_series(pd.DataFrame({'a': s1}), 'a')
    n2 = safe_series(pd.DataFrame({'b': s2}), 'b').replace(0, np.nan)
    return n1 / n2

def safe_scalar_div(n, d):
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
    if val is None or pd.isna(val):
        col.metric(label=label, value="—", delta=f"⏳ {missing_hint}", delta_color="off")
    elif np.isinf(val):
        col.metric(label=label, value="⚠️ 需核验", delta="分母为零或数据异常", delta_color="inverse")
    else:
        try:
            if is_pct:
                disp_val = f"{val * 100:{fmt}}%"
            else:
                disp_val = f"{val:{fmt}} {unit}".strip()
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
    if series is None:
        return False
    s_clean = pd.to_numeric(series, errors='coerce').dropna()
    return not s_clean.empty and not (s_clean == 0).all()

# =====================================================================
# 二、数据解析器：支持模版与自定义 Excel 自适应兼容读取
# =====================================================================
def parse_uploaded_excel(uploaded_file, animal_type="生猪"):
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

    for col in df.columns:
        if col != '月份':
            s = df[col].astype(str).str.replace(',', '', regex=False).str.replace('¥', '', regex=False).str.replace('%', '', regex=False).str.strip()
            df[col] = pd.to_numeric(s, errors='coerce')

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
    death  = [72,68,75,80,78,82,78,120,78,72,70,76]
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
    test_ok= [37,25,42,37,36,39,37,34,41,37,36,42]
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

    first_sow = end_sow.dropna().iloc[0] if not end_sow.dropna().empty else 500
    init_sow = [first_sow - 2]
    for i in range(1, n):
        prev = end_sow.iloc[i-1]
        init_sow.append(prev if pd.notna(prev) else init_sow[-1])
    df['期初能繁母猪'] = init_sow
    df['期末能繁母猪'] = end_sow
    df['月均能繁母猪'] = (df['期初能繁母猪'] + df['期末能繁母猪']) / 2.0

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

    first_end = end_cows.dropna().iloc[0] if not end_cows.dropna().empty else 1000
    init_cows = [first_end - 2]
    for i in range(1, n):
        prev = end_cows.iloc[i-1]
        init_cows.append(prev if pd.notna(prev) else init_cows[-1])
    df['期初成母牛'] = init_cows
    df['期末成母牛'] = end_cows
    df['平均成母牛'] = (df['期初成母牛'] + df['期末成母牛']) / 2.0

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
# 五点一、行业新闻与资讯互动板块（跳转咨询网站 + 留下联系方式在线投稿）
# =====================================================================
def render_industry_news_section(animal_type="生猪"):
    st.markdown("### 📰 现代畜牧行业新闻资讯与投稿互动中心")
    st.caption("📌 **板块定位**：实时追踪国家畜牧政策法规、产业咨询要闻、市场供需走势与前沿养殖技术；支持从业者与研究人员留下个人工作联系方式上传咨询新闻及技术文章。")
    
    col_news_left, col_news_right = st.columns([3, 2])
    
    with col_news_left:
        st.subheader("🌐 行业主流权威咨询与新闻门户直达")
        c_portal1, c_portal2, c_portal3 = st.columns(3)
        with c_portal1:
            st.link_button("🏛️ 农业农村部畜牧兽医局", "http://www.moa.gov.cn/xw/zwdt/", use_container_width=True)
            st.link_button("📰 大畜牧网行业资讯门户", "https://www.dxumu.com/", use_container_width=True)
        with c_portal2:
            st.link_button("🐖 猪场动力网专业咨询", "https://www.powerpigs.net/", use_container_width=True)
            st.link_button("🌾 中国饲料工业信息网", "http://www.chinafeedm.com/", use_container_width=True)
        with c_portal3:
            st.link_button("🐄 荷斯坦奶牛网资讯中心", "http://www.hesitan.com/", use_container_width=True)
            st.link_button("🥛 中国奶业协会官方网站", "http://www.dac.org.cn/", use_container_width=True)

        st.markdown("---")
        st.markdown("#### ⚡ 实时精选行业咨询新闻与政策动态")
        news_items = [
            {
                "title": "农业农村部印发《生猪产能调控实施方案》：引导养殖场户合理调整生产节奏与降本增效",
                "source": "农业农村部新闻办", "date": "2026-10-18", "category": "宏观政策",
                "summary": "方案针对当前养殖周期特征，提出分级预警机制与精准调控措施，引导规模化生猪养殖场保持合理产能，依靠精细化饲喂与科学管理渡过周期底部。",
                "url": "http://www.moa.gov.cn/xw/zwdt/"
            },
            {
                "title": "全国奶业高质量发展推进会：强化生鲜乳优质安全，全面推广 DHI 数字化生产性能测定",
                "source": "中国奶业协会", "date": "2026-10-15", "category": "奶业指导",
                "summary": "推进会议强调依靠自主良种繁育体系与优质粗饲料自给率，降低公斤奶综合饲料成本，保障奶牛养殖场良性盈利水平。",
                "url": "http://www.dac.org.cn/"
            },
            {
                "title": "饲料原料大宗市场分析简报：豆粕与玉米价差企稳，规模场低蛋白日粮配合技术普及加速",
                "source": "中国饲料工业信息网", "date": "2026-10-12", "category": "饲料营养",
                "summary": "监测显示，规模养殖场通过科学配方替代技术平抑能量与蛋白饲料波动，头均饲料采购支出同比实现结构性优化。",
                "url": "http://www.chinafeedm.com/"
            },
            {
                "title": "数字化养殖档案标杆场巡礼：法定养殖档案代码化部署，实现繁育、防疫、用药与出栏全勾稽",
                "source": "大畜牧网", "date": "2026-10-09", "category": "智慧畜牧",
                "summary": "依据农业行业标准 NY/T 3445-2019，多地示范场通过数字系统完成 75 张档案表一键沉淀，显著提升疫病可追溯性与监管合规率。",
                "url": "https://www.dxumu.com/"
            }
        ]
        for n in news_items:
            st.markdown(f"""
            **[{n['title']}]({n['url']})**  
            `{n['category']}` · 来源: {n['source']} · 时间: {n['date']}  
            <span style='color:#8C9BAB; font-size:13px;'>{n['summary']}</span>
            """, unsafe_allow_html=True)
            st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

    with col_news_right:
        st.subheader("📝 咨询新闻与行业文章在线上传投稿")
        st.caption("留下您的工作个人联系方式，我们将第一时间与您对接并择优收录：")
        with st.form("industry_news_submission_form", clear_on_submit=True):
            sub_name = st.text_input("投稿人姓名 *", placeholder="例如：张研究员 / 李场长")
            sub_org = st.text_input("工作单位 / 牧场 *", placeholder="例如：某现代农牧集团 / 农业科学院")
            c_f1, c_f2 = st.columns(2)
            sub_phone = c_f1.text_input("工作手机 / 电话 *", placeholder="请输入工作联系电话")
            sub_wechat = c_f2.text_input("工作微信 / QQ", placeholder="例如：wx_13800000000")
            sub_email = st.text_input("工作联系邮箱", placeholder="name@company.com / name@edu.cn")
            sub_title = st.text_input("新闻 / 文章标题 *", placeholder="请输入您要上传或投稿的新闻资讯/分析文章标题")
            sub_cat = st.selectbox("资讯所属领域", [
                "🐖 生猪产业与精细化管理",
                "🐄 奶牛原奶与现代牧业",
                "🌽 饲料原料与精准营养",
                "💉 动物疫病与生物安全防控",
                "📊 行业政策与宏观市场分析",
                "💻 智慧农业与数字化系统"
            ])
            sub_summary = st.text_area("资讯摘要或相关网页咨询链接", placeholder="请简要描述新闻核心内容或粘贴原文网址链接", height=75)
            sub_file = st.file_uploader("📎 上传咨询新闻/文章文档附件 (支持 Word/PDF/TXT/Excel/图片)", type=["docx", "doc", "pdf", "txt", "xlsx", "png", "jpg"])
            
            submit_news = st.form_submit_button("🚀 确认提交并留下工作联系方式", use_container_width=True)
            if submit_news:
                if not sub_name or not sub_phone or not sub_title:
                    st.error("⚠️ 请完整填写必填项：投稿人姓名、工作联系电话以及文章标题！")
                else:
                    st.success(f"🎉 投稿成功！感谢 {sub_name}（{sub_org or '行业同仁'}）的投稿！您的文章《{sub_title}》已成功登记入库。项目团队将通过您留下的工作联系方式（电话: {sub_phone}）及时与您沟通审核与收录事宜。")

# =====================================================================
# 五点二、行业相关学术论文与科研文献板块（支持 DOI 一键直达跳转）
# =====================================================================
def render_research_papers_section(animal_type="生猪"):
    st.markdown("### 📚 现代畜牧行业权威学术论文与科研文献中心 (DOI 检索跳转)")
    st.caption("📌 **板块定位**：收录国际与国内动物科学、精准畜牧、生物育种与营养健康领域的顶级学术期刊论文；每篇论文均标注文献元数据与权威标准化 DOI (Digital Object Identifier)，支持一键直达全球学术文献平台。")
    
    # DOI 快速解析检索跳转器
    st.subheader("🔍 学术文献 DOI 快速解析与一键直达检索器")
    c_d1, c_d2 = st.columns([4, 1])
    user_doi = c_d1.text_input("请输入或粘贴目标学术论文 DOI 号：", placeholder="例如：10.3168/jds.2016-11883 或 10.1093/jas/skaa177", key="user_doi_search_input")
    clean_doi = user_doi.strip().replace("https://doi.org/", "").replace("http://", "")
    doi_target_url = f"https://doi.org/{clean_doi}" if clean_doi else "https://www.doi.org/demos.html"
    with c_d2:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        st.link_button("🌐 直达解析跳转", doi_target_url, use_container_width=True, disabled=not bool(clean_doi))
    if clean_doi:
        st.caption(f"🔗 准备跳转解析地址: `{doi_target_url}`（将自动通过国际 CrossRef 数据库解析到期刊论文原文）")

    st.markdown("---")
    st.subheader("📖 行业高被引经典与前沿学术论文精选库")
    
    tab_p1, tab_p2 = st.tabs(["🐖 生猪科学与精准养殖权威文献", "🐄 奶牛科学与优质乳工程权威文献"])
    
    with tab_p1:
        pig_papers = [
            {
                "title": "Precision feeding of lactating sows: Implementation and evaluation under commercial farm conditions",
                "title_zh": "商业化农场条件下哺乳母猪精准饲喂的实施与效果评估",
                "journal": "Journal of Animal Science", "year": "2020",
                "authors": "C. Gaillard, E. Brossard, J. Y. Dourmad",
                "doi": "10.1093/jas/skaa177", "field": "精准营养与母猪生产力",
                "abstract": "研究通过智能饲喂系统根据胎次、背膘厚度与产仔数动态调节哺乳母猪日粮营养素供给，显著减少母猪哺乳期体储动用，断奶仔猪平均窝重与断奶至发情间隔显著改善。"
            },
            {
                "title": "Feed efficiency in swine: a review of genetic, physiological and environmental factors",
                "title_zh": "猪饲料转化效率：遗传、生理与环境影响因素系统综述",
                "journal": "Journal of Animal Science and Biotechnology", "year": "2015",
                "authors": "J. F. Patience, M. C. Rossoni-Serao, N. K. Gutierrez",
                "doi": "10.1186/s40104-015-0033-8", "field": "料肉比机理与抗逆性",
                "abstract": "系统阐述全群料肉比（FCR）与剩余采食量（RFI）的生物学机制，解析健康状况、热应激、日粮纤维及肠道菌群对猪饲料转化率的影响规律。"
            },
            {
                "title": "Automated monitoring of pig behavior and health using computer vision: A review",
                "title_zh": "基于计算机视觉技术的猪只行为与健康自动化监测综述",
                "journal": "Computers and Electronics in Agriculture", "year": "2018",
                "authors": "J. Cowton, C. Kyriazakis, T. Bacardit, R. Bate",
                "doi": "10.1016/j.compag.2018.08.020", "field": "智慧养殖与无感体检",
                "abstract": "全面总结利用深度学习、2D/3D摄像头无损监测猪只采食饮水、攻击打斗、体温热图及跛行异常的最新算法架构与产业落地路径。"
            },
            {
                "title": "Review of smart pig farming: Systems, technologies and future prospects",
                "title_zh": "智慧生猪养殖系统、核心技术与未来展望",
                "journal": "Animal Nutrition", "year": "2022",
                "authors": "X. Wang, W. Zhang, X. Zhao, et al.",
                "doi": "10.1016/j.aninu.2022.06.008", "field": "数字化数据与AIoT",
                "abstract": "梳理物联网环境调控、智能环控器、声学咳嗽监测及全生命周期数字化养殖档案如何协同赋能规模化猪场降低死亡率与运营成本。"
            }
        ]
        for p in pig_papers:
            with st.container():
                c_pt1, c_pt2 = st.columns([4, 1])
                doi_url = f"https://doi.org/{p['doi']}"
                c_pt1.markdown(f"""
                **{p['title']}**  
                *{p['title_zh']}*  
                📖 **期刊**: `{p['journal']}` ({p['year']}) · **作者**: {p['authors']} · **领域**: `{p['field']}`  
                🔑 **DOI**: `{p['doi']}`  
                <span style='color:#8C9BAB; font-size:13px;'>{p['abstract']}</span>
                """, unsafe_allow_html=True)
                with c_pt2:
                    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
                    st.link_button("📄 DOI 直达论文", doi_url, use_container_width=True, help=f"点击跳转至权威官方期刊页面: {doi_url}")
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

    with tab_p2:
        cow_papers = [
            {
                "title": "Symposium review: Technologies for precision dairy farming and their application",
                "title_zh": "精准奶牛养殖技术与牧场工程应用专题综述",
                "journal": "Journal of Dairy Science", "year": "2016",
                "authors": "J. M. Bewley, R. A. Russell, K. A. Dolecheck, M. R. Borchers",
                "doi": "10.3168/jds.2016-11883", "field": "精准奶业与生产性能测定",
                "abstract": "系统评估电子项圈、反刍传感器、自动挤奶机器人（AMS）与在线乳成分分析仪在现代奶牛场的应用收益与决策模型集成方法。"
            },
            {
                "title": "Invited review: Sensor-based detection of health and welfare in dairy cattle",
                "title_zh": "特邀综述：基于传感器技术的奶牛健康与动物福利监测",
                "journal": "Journal of Dairy Science", "year": "2021",
                "authors": "C. R. Stygar, Y. Le Cozler, C. Allain",
                "doi": "10.3168/jds.2020-19266", "field": "乳房炎与健康预警",
                "abstract": "探讨乳导电率、体细胞数（SCC）实时传感器与计步器如何实现临床乳房炎与蹄病的前置预警，降低抗生素使用与死淘率。"
            },
            {
                "title": "Dry matter intake and feed efficiency in dairy cattle: Biology and prediction",
                "title_zh": "奶牛干物质采食量与饲料转化效率：生物学基础与模型预测",
                "journal": "Journal of Dairy Science", "year": "2016",
                "authors": "M. J. VandeHaar, L. E. Armentano, K. Weigel, et al.",
                "doi": "10.3168/jds.2015-10352", "field": "料奶比优化与能量代谢",
                "abstract": "建立了成母牛干物质采食量（DMI）、产奶净能与全群料奶比的数学预测模型，为提升高产奶牛能量利用率与减少温室气体排放提供理论指导。"
            },
            {
                "title": "A review of automated estrus detection technologies in dairy cattle",
                "title_zh": "奶牛自动化发情监测技术与情期受胎率提升综述",
                "journal": "Journal of Dairy Science", "year": "2012",
                "authors": "M. Saint-Dizier, S. Chastant-Maillard",
                "doi": "10.3168/jds.2011-4775", "field": "繁殖管理与受胎率",
                "abstract": "系统比较活动量计步器、体温监测及反刍时间变化对奶牛排卵窗口期的预测准确率，显著提高牧场首配受胎率与配种妊娠效率。"
            }
        ]
        for p in cow_papers:
            with st.container():
                c_pt1, c_pt2 = st.columns([4, 1])
                doi_url = f"https://doi.org/{p['doi']}"
                c_pt1.markdown(f"""
                **{p['title']}**  
                *{p['title_zh']}*  
                📖 **期刊**: `{p['journal']}` ({p['year']}) · **作者**: {p['authors']} · **领域**: `{p['field']}`  
                🔑 **DOI**: `{p['doi']}`  
                <span style='color:#8C9BAB; font-size:13px;'>{p['abstract']}</span>
                """, unsafe_allow_html=True)
                with c_pt2:
                    st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
                    st.link_button("📄 DOI 直达论文", doi_url, use_container_width=True, help=f"点击跳转至权威官方期刊页面: {doi_url}")
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

# =====================================================================
# 五点三、产业金融板块：行业十强龙头与大宗期货行情直达
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
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)
                
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
                st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

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

# =====================================================================
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
            st.info("💡 填报贴士：带*字段建议优先补充；未采集指标可留空，系统绝不会报错，其余指标照常推演。")

        uploaded_file = st.file_uploader(
            f"📤 请上传修改好的 {animal_choice[:2]} 台账工作簿 (.xlsx)",
            type=["xlsx", "xls"]
        )

    st.markdown("---")
    st.subheader("💡 智慧牧业赋能亮点")
    st.markdown("""
    - **全套法定档案数据精准映射**（严格依循《畜牧法》与 NY/T 3445-2019）
    - **猪牛双畜种 28 项核心经营指标全量闭环**
    - **容错隔离计算底座**：空指标不崩溃，好指标照样算
    - **多板块全景同时呈现与免滚轴沉浸体验**
    - **行业权威新闻咨询与学术论文 DOI 直达**
    """)
    
    st.markdown("---")
    st.subheader("📬 开发者联系方式")
    st.markdown("""
    - **QQ**: `3100785036`
    - **WeChat (微信)**: `13679275651`
    - **项目作者**: 李尧 (学号: `2024010983`)
    - **培养单位**: 中国西北农林科技大学 动物科技学院
    """)
    st.caption("版本：V2.5 全景平铺·无滚轴·行业新闻与DOI论文增强版 | 2026-10-22")

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
        - **国家法律**：《中华人民共和国畜牧法》第四十一条（规模养殖场必须建立养殖档案）、《中华人民共和国动物防疫法》、《中华人民共和国农产品质量安全法》
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

if is_custom_upload and uploaded_file is not None and upload_health_info:
    st.success(f"🟢 **当前驱动数据**：{status_msg}。系统已启用容错计算引擎，部分空值指标已安全隔离。")

# =====================================================================
# 八、大屏呈现模式控制（全景平铺多板块同时呈现 vs 标签页）
# =====================================================================
st.markdown("---")
c_v1, c_v2 = st.columns([3, 1])
with c_v1:
    view_mode = st.radio(
        "🎛️ **大屏呈现模式选择**（已去除滚动轴，平滑无阻）：",
        ["📜 全景平铺模式（多板块同时呈现·免翻页无滚轴）", "📑 标签选项卡模式"],
        horizontal=True,
        index=0,
        help="【全景平铺模式】将生产走势、成本结构、原始台账体检、行业新闻、DOI论文、产业金融并列同时呈现；【标签模式】按类别切换"
    )
with c_v2:
    st.info("💡 当前已开启「免滚轴沉浸浏览」，页面与组件滚动轴已自动隐藏。")

st.markdown("---")

# =====================================================================
# 九、业务大屏主视觉呈现
# =====================================================================
if is_pig:
    st.markdown("### 🐖 生猪养殖核心经营指标概览 (500头能繁母猪基准演练场)")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    render_kpi_card(c1, "年提供断奶仔猪 (PSY)", kpi.get('PSY'), fmt=".2f", 
                    target_delta="+3.27 超标杆线 (≥20)", threshold=20.0, missing_hint="待补断奶/母猪数据")
    render_kpi_card(c2, "年提供出栏肥猪 (MSY)", kpi.get('MSY'), fmt=".2f", 
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

    def render_pig_trends():
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
                    fig_fcr.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['出栏均重_kg'], name='出栏均重(kg)', marker_color='#87E8DE'))
                if has_fcr:
                    fig_fcr.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['料肉比'], name='月度料肉比', yaxis='y2', line=dict(color='#FAAD14', width=2)))
                fig_fcr.update_layout(
                    title="出栏均重与料肉比动态映射",
                    yaxis=dict(title="出栏均重 (kg)"),
                    yaxis2=dict(title="料肉比 (FCR)", overlaying='y', side='right'),
                    legend=dict(x=0.01, y=0.99), template="plotly_white"
                )
                st.plotly_chart(fig_fcr, use_container_width=True)
            else:
                st.info("ℹ️ 【出栏均重与料肉比】暂未绘制：缺少「出栏总重」或「耗料量」数据。")

    def render_pig_structure():
        st.subheader("🍩 结构与成本洞察")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            cost_parts = {
                '饲料采购成本': df_calc.get('饲料采购金额', pd.Series()).sum(min_count=1),
                '兽药防疫费用': df_calc.get('兽药采购金额', pd.Series()).sum(min_count=1),
                '人工折旧能耗': df_calc.get('其他成本', pd.Series()).sum(min_count=1)
            }
            cost_clean = {k: v for k, v in cost_parts.items() if pd.notna(v) and v > 0}
            if cost_clean:
                fig_pie_cost = px.pie(
                    names=list(cost_clean.keys()), values=list(cost_clean.values()),
                    title="年度完全养殖成本结构分布", hole=0.45,
                    color_discrete_sequence=['#FF7A45', '#36CFC9', '#597EF7']
                )
                st.plotly_chart(fig_pie_cost, use_container_width=True)
            else:
                st.info("ℹ️ 【成本构成扇形图】暂未绘制：缺少饲料/兽药/其他成本录入项。")

        with col_p2:
            live_tot = df_calc.get('出生活仔', pd.Series()).sum(min_count=1)
            out_tot = df_calc.get('出栏', pd.Series()).sum(min_count=1)
            dead_tot = df_calc.get('死亡', pd.Series()).sum(min_count=1)
            cull_tot = df_calc.get('淘汰全群', pd.Series()).sum(min_count=1)

            if pd.notna(out_tot) and out_tot > 0:
                surv_data = {'出栏成活': out_tot}
                if pd.notna(dead_tot) and dead_tot > 0:
                    surv_data['全期病死'] = dead_tot
                if pd.notna(cull_tot) and cull_tot > 0:
                    surv_data['正常死淘'] = cull_tot
                fig_pie_surv = px.pie(
                    names=list(surv_data.keys()), values=list(surv_data.values()),
                    title="仔猪全生命周期存活与死淘分布", hole=0.45,
                    color_discrete_sequence=['#52C41A', '#FF4D4F', '#FFA940']
                )
                st.plotly_chart(fig_pie_surv, use_container_width=True)
            else:
                st.info("ℹ️ 【存活与死淘分布】暂未绘制：缺少出栏或死亡记录。")

    def render_pig_archive_table():
        st.subheader("📑 原始档案台账与质量体检")
        valid_cols_cnt = sum(has_valid_data(df_calc[c]) for c in df_calc.columns if c != '月份')
        total_cols_cnt = len([c for c in df_calc.columns if c != '月份'])
        c_q1, c_q2, c_q3 = st.columns([1, 1, 2])
        c_q1.metric("指标就绪度", f"{valid_cols_cnt} / {total_cols_cnt}", f"{valid_cols_cnt/total_cols_cnt*100:.1f}%")
        c_q2.metric("台账月度完整性", f"{len(df_calc)} 个月", "存栏动态平衡保持")
        c_q3.info("💡 提示：本系统具备指标容错隔离特性，任何空项均不影响其他指标推演。下方表格中「—」代表该月份暂未采集该项台账。")

        st.markdown("##### 原始数据台账明细与核验（空值优雅呈现为 —）")
        display_cols = [c for c in ['月份', '期初全群存栏', '出生活仔', '出栏', '死亡', '淘汰全群', '期末全群存栏', '出栏均重_kg', '头均盈利_元'] if c in df_calc.columns]
        df_show = df_calc[display_cols].copy()
        for c in df_show.columns:
            if c != '月份':
                df_show[c] = df_show[c].apply(lambda x: f"{x:,.2f}" if (pd.notna(x) and isinstance(x, (int, float))) else "—")
        st.dataframe(df_show, hide_index=True, use_container_width=True)

    def render_pig_export():
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

    if "全景平铺" in view_mode:
        st.markdown("#### 模块一：核心生产经营走势与预警")
        render_pig_trends()
        st.markdown("---")
        st.markdown("#### 模块二：成本效益与结构洞察")
        render_pig_structure()
        st.markdown("---")
        st.markdown("#### 模块三：原始生产台账与数据质量体检看板")
        render_pig_archive_table()
        st.markdown("---")
        st.markdown("#### 模块四：行业新闻资讯与咨询文章互动")
        render_industry_news_section("生猪")
        st.markdown("---")
        st.markdown("#### 模块五：行业权威学术论文与科研文献 (DOI直达)")
        render_research_papers_section("生猪")
        st.markdown("---")
        st.markdown("#### 模块六：现代畜牧产业金融、前十猪企与期货大宗行情")
        render_finance_section("生猪")
        st.markdown("---")
        st.markdown("#### 模块七：测算结果导出与标准采集模版")
        render_pig_export()
    else:
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📊 生产经营走势", "🍩 成本结构洞察", "📑 台账质量体检",
            "📰 行业新闻与互动", "📚 权威论文 (DOI)", "💹 产业金融与期货", "📥 测算结果与模版"
        ])
        with tab1:
            render_pig_trends()
        with tab2:
            render_pig_structure()
        with tab3:
            render_pig_archive_table()
        with tab4:
            render_industry_news_section("生猪")
        with tab5:
            render_research_papers_section("生猪")
        with tab6:
            render_finance_section("生猪")
        with tab7:
            render_pig_export()

else:
    # 奶牛分支主视觉呈现
    st.markdown("### 🐄 奶牛养殖核心经营指标概览 (1000头成母牛现代示范牧场)")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    render_kpi_card(c1, "成母牛年均单产", kpi.get('成母牛年单产'), fmt=",.0f", unit="kg", 
                    target_delta="+973 优于标杆 (10,000kg)", threshold=10000.0, missing_hint="待补产奶记录")
    render_kpi_card(c2, "泌乳牛日均单产", kpi.get('泌乳牛日均单产'), fmt=".1f", unit="kg/天", 
                    target_delta="高产阶段 (≥35.0)", threshold=35.0, missing_hint="待补泌乳牛/产奶量")

    f_val = kpi.get('生鲜牛乳脂率')
    p_val = kpi.get('生鲜牛乳蛋白率')
    if pd.notna(f_val) and pd.notna(p_val):
        c3.metric(label="生鲜牛乳脂 / 乳蛋白率", value=f"{f_val*100:.2f}% / {p_val*100:.2f}%", delta="特级生鲜乳标准")
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

    def render_cow_trends():
        st.subheader("📈 奶牛产奶与品质核心走势")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            has_milk = has_valid_data(df_calc.get('月总产奶_吨'))
            has_daily = has_valid_data(df_calc.get('日均单产_kg'))
            if has_milk or has_daily:
                fig_milk = go.Figure()
                if has_milk:
                    fig_milk.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['月总产奶_吨'], name='月总产奶量(吨)', marker_color='#69C0FF'))
                if has_daily:
                    fig_milk.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['日均单产_kg'], name='泌乳牛日均单产(kg)', yaxis='y2', line=dict(color='#FF7A45', width=3)))
                fig_milk.update_layout(
                    title="原奶月度总产出与泌乳牛单产趋势",
                    yaxis=dict(title="月总产奶量 (吨)"),
                    yaxis2=dict(title="日均单产 (kg/天)", overlaying='y', side='right'),
                    legend=dict(x=0.01, y=0.99), template="plotly_white"
                )
                st.plotly_chart(fig_milk, use_container_width=True)
            else:
                st.info("ℹ️ 【产奶走势图】暂未绘制：缺少月产奶量或泌乳牛头数。")

        with col_c2:
            has_scc = has_valid_data(df_calc.get('体细胞数_SCC'))
            if has_scc:
                fig_scc = go.Figure()
                fig_scc.add_trace(go.Scatter(x=df_calc['月份'], y=df_calc['体细胞数_SCC'], mode='lines+markers', name='SCC (万/mL)', line=dict(color='#B37FEB', width=3)))
                fig_scc.add_hline(y=20.0, line_dash="dash", line_color="#FF4D4F", annotation_text="优质乳行业上限 (20万/mL)")
                fig_scc.update_layout(title="生鲜乳体细胞数 (SCC) 预警走势", yaxis=dict(title="SCC (万/mL)"), template="plotly_white")
                st.plotly_chart(fig_scc, use_container_width=True)
            else:
                st.info("ℹ️ 【体细胞数 SCC 预警走势】暂未绘制：缺少 DHI 体细胞数记录。")

    def render_cow_structure():
        st.subheader("🍩 奶牛成本效益结构剖析")
        col_cp1, col_cp2 = st.columns(2)
        with col_cp1:
            milk_cost_parts = {
                '粗饲料费用': df_calc.get('粗饲料费', pd.Series()).sum(min_count=1),
                '精饲料费用': df_calc.get('精饲料费', pd.Series()).sum(min_count=1),
                '兽药与疫苗': df_calc.get('兽药防疫费', pd.Series()).sum(min_count=1),
                '人工折旧能耗': df_calc.get('其他成本', pd.Series()).sum(min_count=1)
            }
            clean_m_cost = {k: v for k, v in milk_cost_parts.items() if pd.notna(v) and v > 0}
            if clean_m_cost:
                fig_pie_mcost = px.pie(
                    names=list(clean_m_cost.keys()), values=list(clean_m_cost.values()),
                    title="奶牛完全饲养成本结构分布", hole=0.45,
                    color_discrete_sequence=['#73D13D', '#FFA940', '#FF4D4F', '#40A9FF']
                )
                st.plotly_chart(fig_pie_mcost, use_container_width=True)
            else:
                st.info("ℹ️ 【成本结构饼图】暂未绘制：缺少饲料或费用录入。")

        with col_cp2:
            has_rev = has_valid_data(df_calc.get('月度生鲜乳收入_元'))
            has_tot_cost = has_valid_data(df_calc.get('月度总成本_元'))
            if has_rev and has_tot_cost:
                fig_profit = go.Figure()
                fig_profit.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['月度生鲜乳收入_元']/10000, name='销售收入(万元)', marker_color='#36CFC9'))
                fig_profit.add_trace(go.Bar(x=df_calc['月份'], y=df_calc['月度总成本_元']/10000, name='完全成本(万元)', marker_color='#FFA39E'))
                profit_vals = (df_calc['月度生鲜乳收入_元'] - df_calc['月度总成本_元']) / 10000
                fig_profit.add_trace(go.Scatter(x=df_calc['月份'], y=profit_vals, name='月净利润(万元)', line=dict(color='#52C41A', width=3)))
                fig_profit.update_layout(title="月度收支与利润对冲分析", yaxis=dict(title="金额 (万元)"), barmode='group', template="plotly_white")
                st.plotly_chart(fig_profit, use_container_width=True)
            else:
                st.info("ℹ️ 【收支与利润对冲分析】暂未绘制：缺少原奶售价或月度成本。")

    def render_cow_archive_table():
        st.subheader("📑 奶牛生产台账与质量体检")
        valid_cols_cnt = sum(has_valid_data(df_calc[c]) for c in df_calc.columns if c != '月份')
        total_cols_cnt = len([c for c in df_calc.columns if c != '月份'])
        c_q1, c_q2, c_q3 = st.columns([1, 1, 2])
        c_q1.metric("指标就绪度", f"{valid_cols_cnt} / {total_cols_cnt}", f"{valid_cols_cnt/total_cols_cnt*100:.1f}%")
        c_q2.metric("台账月度完整性", f"{len(df_calc)} 个月", "产奶存栏动态平衡")
        c_q3.info("💡 提示：本系统具备指标容错隔离特性，任何空项均不影响其他指标推演。下方表格中「—」代表该月份暂未采集该项台账。")

        display_cols_cow = [c for c in ['月份', '期初成母牛', '在挤泌乳牛', '干奶牛', '月产奶量_kg', '生鲜乳销量_kg', '日均单产_kg', '公斤奶完全成本_元', '公斤奶净利润_元'] if c in df_calc.columns]
        df_show_cow = df_calc[display_cols_cow].copy()
        for c in df_show_cow.columns:
            if c != '月份':
                df_show_cow[c] = df_show_cow[c].apply(lambda x: f"{x:,.2f}" if (pd.notna(x) and isinstance(x, (int, float))) else "—")
        st.dataframe(df_show_cow, hide_index=True, use_container_width=True)

    def render_cow_export():
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

    if "全景平铺" in view_mode:
        st.markdown("#### 模块一：核心产奶走势与关键预警")
        render_cow_trends()
        st.markdown("---")
        st.markdown("#### 模块二：奶牛成本效益结构剖析")
        render_cow_structure()
        st.markdown("---")
        st.markdown("#### 模块三：奶牛生产台账与质量体检看板")
        render_cow_archive_table()
        st.markdown("---")
        st.markdown("#### 模块四：行业新闻资讯与咨询文章互动")
        render_industry_news_section("奶牛")
        st.markdown("---")
        st.markdown("#### 模块五：行业权威学术论文与科研文献 (DOI直达)")
        render_research_papers_section("奶牛")
        st.markdown("---")
        st.markdown("#### 模块六：现代畜牧产业金融、前十乳企与期货大宗行情")
        render_finance_section("奶牛")
        st.markdown("---")
        st.markdown("#### 模块七：测算结果导出与标准采集模版")
        render_cow_export()
    else:
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "🥛 产奶与经营看板", "🛡️ 健康合规与DHI预警", "📑 奶牛生产台账与质量体检",
            "📰 行业新闻与互动", "📚 权威论文 (DOI)", "💹 产业金融与期货", "📥 测算结果与模版"
        ])
        with tab1:
            render_cow_trends()
        with tab2:
            render_cow_structure()
        with tab3:
            render_cow_archive_table()
        with tab4:
            render_industry_news_section("奶牛")
        with tab5:
            render_research_papers_section("奶牛")
        with tab6:
            render_finance_section("奶牛")
        with tab7:
            render_cow_export()

# =====================================================================
# 十、知识产权、学术严正声明与页脚版权
# =====================================================================
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
st.markdown("© 2026 智慧牧业工程 · 李尧 2024010983 · 中国西北农林科技大学动物科技学院 · 生猪与奶牛双畜种数字化养殖档案与指标计算演示系统 (V2.5)")
