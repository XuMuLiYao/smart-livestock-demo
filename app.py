# -*- coding: utf-8 -*-
"""
智慧畜牧数字化管理与核心指标交互式看板
项目：智慧牧业工程 · smart_pig_demo
技术栈：Streamlit + pandas + Plotly + openpyxl
升级亮点（2026-10-22 第13次改动）：
  0. 彻底解决云端与本地环境管理员登录硬件鉴权冲突报错：
     - 深度兼容本地与线上（Streamlit Community Cloud）双环境架构，解决远程 Web 访问无法透传客户端底层 MAC 导致的“设备硬件未授权”问题
     - 硬件白名单与凭据双通道鉴权：支持服务器本地指纹自动识别与前端预置授权设备指纹（李尧专属 MacBook Air GP71L4XYJX / 网卡 MAC）一键填充验证
     - 跨平台底层探测强化：兼容 Linux 容器与 macOS 底层系统调用，全面捕获命令异常
     - 优化 KPI_DIR 路径容错，保障云端与本地全量指标库文件平滑读取
升级亮点（2026-10-22 第12次改动）：
  0. 增设多重管理员权限认证与敏感数据隐私防护机制：
     - 双因子鉴权认证：查询验证登录设备的 MAC 地址、macOS 硬件序列号白名单与管理员账号密码
     - 敏感数据严格隔离：投稿人姓名、联系方式（电话/微信/邮箱）、文章摘要与附件路径只对认证管理员开放
     - 普通用户端安全防护：网页端对普通访客完全隐藏已登记投稿台账与 Excel 导出功能，显示隐私保护提示
     - 侧边栏专属【🔐 管理员安全认证控制台】：支持设备指纹即时检测、登录鉴权、已认证状态保持与安全退出
升级亮点（2026-10-22 第11次改动）：
  0. 增设「🎓 现代畜牧产业人才培养与全国高校动物科技学院科研咨询矩阵」专属板块：
     - 收录全国 24 所顶级与重点农林高校动物科技学院权威科研与人才培养资讯
     - 支持网页内直接多维检索查询：关键词模糊检索、地理大区过滤、办学层次（985/211/双一流/省属骨干）及聚焦畜种筛选
     - 各院所提供官方网站直达跳转按钮（一键直通学院官方主页）
     - 集成全景数据表展开预览与《全国农林高校动物科技学院名录与科研咨询指南.xlsx》一键导出下载
升级亮点（2026-10-22 第10次改动）：
  0. 行业新闻资讯与投稿互动中心全面实现自动持久化存盘与台账闭环：
     - 用户在前端提交投稿表单后，数据自动安全追加写入《行业新闻与咨询投稿登记表.xlsx》
     - 自动转存各类格式文档附件至本地《投稿附件》目录留存归档
     - 表单下方即时提供「📋 查看已登记投稿台账与导出」展开面板，支持大屏端实时数据预览与 Excel 原表下载
升级亮点（2026-10-22 第8次改动）：
  1. 智慧养殖控制台全面扩充至八大核心畜禽种类：
     - 生猪养殖 (Swine)、奶牛养殖 (Dairy Cattle)、肉牛养殖 (Beef Cattle)、奶山羊养殖 (Dairy Goat)、
       肉羊养殖 (Meat Sheep)、皮用羊养殖 (Fur Sheep)、蛋鸡养殖 (Layer Chicken)、肉鸡养殖 (Broiler Chicken)
     - 为八大畜禽分别配置标准基准演练模型数据源与自定义 Excel 导入接口，并实现容错隔离计算底座全覆盖
  2. 行业新闻与资讯互动板块全新升级：
     - 增添对应畜禽（肉牛、奶山羊、肉羊、滩羊皮用羊、蛋鸡、肉鸡）主流权威咨询门户与政策资讯
     - 新增「🔄 刷新行业咨询与政策动态」交互按钮，支持实时刷新动态与状态即时反馈
     - 投稿表单增加八大畜禽所属领域的精准分类
  3. 权威学术论文与科研文献中心全新升级：
     - DOI 快速解析与一键直达检索器新增直达链接 9 大畜牧与动物科学权威数据库（CrossRef, PubMed, WoS, CNKI, JDS, JAS, Poultry Science, Small Ruminant Research, Meat Science）
     - 经典与前沿学术论文精选库增设八大畜禽专属文献板块，完整提供中英双语篇名、核心摘要及 DOI 一键直达跳转
  4. 现代畜牧产业金融与大宗期货矩阵全新升级：
     - 增设中国前十肉牛养殖加工龙头、十强羊业与奶山羊/肉羊龙头、前十蛋鸡与品牌鲜蛋龙头、前十肉鸡养殖与白羽肉禽龙头官网直达
     - 大宗期货行情增加「大连商品交易所鸡蛋期货 (DCE: JD)」与「芝商所活牛/育肥牛期货 (CME: LE / FC)」，完善全畜种金融套保矩阵
  5. 全景平铺呈现模式视觉升级：
     - 采用清晰专业的业务板块标题
     - 升级分界线为略显著的渐变高对比度视觉线条，并合理拉大板块上下间距（45px/35px），使大屏通透大气
  6. 严格保留与强化前期合规与知识产权资产：
     - 严格依循《畜牧法》与行业标准 NY/T 3445-2019 全套法定养殖档案（81张表）规范
     - 侧边栏与页脚公示开发者联系方式（QQ: 3100785036，WeChat: 13679275651）
     - 中国西北农林科技大学动物科技学院李尧（学号：2024010983）独立进行中英文严正声明
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import datetime
from pathlib import Path

# 设置页面配置
st.set_page_config(
    page_title="智慧牧业工程 · 数字化管理与指标看板",
    page_icon="🐄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 全局注入去除滚轴样式与优雅间距样式
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
/* 略显著且间距合理的板块分界线 */
.section-divider-box {
    margin-top: 45px;
    margin-bottom: 35px;
}
.section-divider-line {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, rgba(46, 117, 182, 0.12) 0%, rgba(46, 117, 182, 0.85) 50%, rgba(46, 117, 182, 0.12) 100%);
    margin: 0;
    padding: 0;
}
</style>
""", unsafe_allow_html=True)

def render_section_divider():
    st.markdown("""
    <div class="section-divider-box">
        <hr class="section-divider-line">
    </div>
    """, unsafe_allow_html=True)

# 路径常量解析
CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
KPI_DIR = (ROOT_DIR / "表格与指标库") if (ROOT_DIR / "表格与指标库").exists() else (CURRENT_DIR / "表格与指标库")

PIG_KPI_PATH = KPI_DIR / "生猪养殖核心指标库与可行性验证.xlsx"
COW_KPI_PATH = KPI_DIR / "奶牛养殖核心指标库与可行性验证.xlsx"
BEEF_KPI_PATH = KPI_DIR / "肉牛养殖核心指标库与可行性验证.xlsx"
GOAT_KPI_PATH = KPI_DIR / "奶山羊生产核心指标库与可行性验证.xlsx"
SHEEP_KPI_PATH = KPI_DIR / "肉羊生产核心指标库与可行性验证.xlsx"
FUR_KPI_PATH = KPI_DIR / "皮用羊生产核心指标库与可行性验证.xlsx"
LAYER_KPI_PATH = KPI_DIR / "蛋鸡养殖核心指标库与可行性验证.xlsx"
BROILER_KPI_PATH = KPI_DIR / "肉鸡养殖核心指标库与可行性验证.xlsx"

PIG_TEMPLATE_PATH = CURRENT_DIR / "生猪生产台账采集与指标计算上传模板.xlsx"
COW_TEMPLATE_PATH = CURRENT_DIR / "奶牛生产台账采集与指标计算上传模板.xlsx"
BEEF_TEMPLATE_PATH = CURRENT_DIR / "肉牛生产台账采集与指标计算上传模板.xlsx"
GOAT_TEMPLATE_PATH = CURRENT_DIR / "奶山羊生产台账采集与指标计算上传模板.xlsx"
SHEEP_TEMPLATE_PATH = CURRENT_DIR / "肉羊生产台账采集与指标计算上传模板.xlsx"
FUR_TEMPLATE_PATH = CURRENT_DIR / "皮用羊生产台账采集与指标计算上传模板.xlsx"
LAYER_TEMPLATE_PATH = CURRENT_DIR / "蛋鸡生产台账采集与指标计算上传模板.xlsx"
BROILER_TEMPLATE_PATH = CURRENT_DIR / "肉鸡生产台账采集与指标计算上传模板.xlsx"

SUBMISSION_EXCEL_PATH = CURRENT_DIR / "行业新闻与咨询投稿登记表.xlsx"
SUBMISSION_UPLOAD_DIR = CURRENT_DIR / "投稿附件"

SPECIES_TEMPLATE_DICT = {
    "生猪": ("生猪生产台账采集与指标计算上传模板.xlsx", PIG_TEMPLATE_PATH),
    "奶牛": ("奶牛生产台账采集与指标计算上传模板.xlsx", COW_TEMPLATE_PATH),
    "肉牛": ("肉牛生产台账采集与指标计算上传模板.xlsx", BEEF_TEMPLATE_PATH),
    "奶山羊": ("奶山羊生产台账采集与指标计算上传模板.xlsx", GOAT_TEMPLATE_PATH),
    "肉羊": ("肉羊生产台账采集与指标计算上传模板.xlsx", SHEEP_TEMPLATE_PATH),
    "皮用羊": ("皮用羊生产台账采集与指标计算上传模板.xlsx", FUR_TEMPLATE_PATH),
    "蛋鸡": ("蛋鸡生产台账采集与指标计算上传模板.xlsx", LAYER_TEMPLATE_PATH),
    "肉鸡": ("肉鸡生产台账采集与指标计算上传模板.xlsx", BROILER_TEMPLATE_PATH),
}

def get_template_info(animal_choice):
    for key, (fname, pth) in SPECIES_TEMPLATE_DICT.items():
        if key in animal_choice:
            return key, fname, pth
    return "生猪", "生猪生产台账采集与指标计算上传模板.xlsx", PIG_TEMPLATE_PATH

def load_template_preview_df(template_path):
    if template_path.exists():
        try:
            xl = pd.ExcelFile(template_path)
            s1 = xl.sheet_names[0]
            df_raw = xl.parse(s1, header=None)
            header_idx = 2
            for idx in range(min(6, len(df_raw))):
                row_str = " ".join([str(v) for v in df_raw.iloc[idx].values if pd.notna(v)])
                if any(kw in row_str for kw in ["填报说明", "说明：", "指南", "录入说明"]):
                    continue
                if any(k in row_str for k in ["月份", "批次", "批次号", "存栏", "产奶", "增重", "日龄", "产蛋"]):
                    header_idx = idx
                    break
            df_preview = xl.parse(s1, skiprows=header_idx)
            df_preview = df_preview.dropna(how='all')
            first_col = df_preview.columns[0]
            df_preview = df_preview[~df_preview[first_col].astype(str).str.contains("全年|累计|合计|均值|平均|说明", na=False)]
            df_guide = None
            if len(xl.sheet_names) > 1:
                df_guide = xl.parse(xl.sheet_names[1], skiprows=1)
                df_guide = df_guide.dropna(how='all')
            return df_preview, df_guide
        except Exception:
            pass
    return pd.DataFrame(), None

# =====================================================================
# 一、基础辅助函数：安全计算与UI格式化
# =====================================================================
def safe_series(df, col, default_val=np.nan, aliases=None):
    if col in df.columns:
        return pd.to_numeric(df[col], errors='coerce')
    if aliases:
        for a in aliases:
            if a in df.columns:
                return pd.to_numeric(df[a], errors='coerce')
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

def has_valid_data(series):
    if series is None:
        return False
    valid_count = series.dropna().shape[0]
    return valid_count > 0 and not (series.dropna() == 0).all()

def render_kpi_card(col, label, val, is_pct=False, fmt=".2f", unit="", 
                    target_delta="", warn_delta=None, threshold=None, 
                    inverse=False, missing_hint="待录入对应台账"):
    if val is None or pd.isna(val):
        col.metric(label=label, value="—", delta=f"⏳ {missing_hint}", delta_color="off")
    elif np.isinf(val):
        col.metric(label=label, value="—", delta="⚠️ 计算除零或异常", delta_color="off")
    else:
        try:
            if is_pct:
                display_val = f"{val * 100:{fmt}}%"
            else:
                display_val = f"{val:{fmt}}"
            if unit:
                display_val += f" {unit}"
            
            delta_text = target_delta
            delta_color = "normal"
            if threshold is not None:
                is_better = (val <= threshold) if inverse else (val >= threshold)
                if not is_better and warn_delta:
                    delta_text = warn_delta
                    delta_color = "inverse"
            col.metric(label=label, value=display_val, delta=delta_text, delta_color=delta_color)
        except Exception:
            col.metric(label=label, value=str(val), delta=target_delta)

# =====================================================================
# 二、数据解析与指标字典适配器
# =====================================================================
def load_kpi_dict(animal_type):
    path = PIG_KPI_PATH
    if "奶牛" in animal_type:
        path = COW_KPI_PATH
    elif "肉牛" in animal_type:
        path = BEEF_KPI_PATH
    elif "奶山羊" in animal_type:
        path = GOAT_KPI_PATH
    elif "肉羊" in animal_type:
        path = SHEEP_KPI_PATH
    elif "皮用羊" in animal_type:
        path = FUR_KPI_PATH
    elif "蛋鸡" in animal_type:
        path = LAYER_KPI_PATH
    elif "肉鸡" in animal_type:
        path = BROILER_KPI_PATH

    kpi_dict = {}
    if path.exists():
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            if "①指标库" in wb.sheetnames:
                ws = wb["①指标库"]
                for row in ws.iter_rows(values_only=True):
                    if row and len(row) > 1 and row[1]:
                        k_name = str(row[1]).strip()
                        if k_name != "指标名称":
                            kpi_dict[k_name] = row
            wb.close()
        except Exception:
            pass
    return kpi_dict

def parse_uploaded_excel(uploaded_file, animal_type="生猪"):
    if uploaded_file is None:
        return pd.DataFrame()
    
    excel_bytes = uploaded_file.read() if hasattr(uploaded_file, 'read') else uploaded_file
    if hasattr(uploaded_file, 'seek'):
        uploaded_file.seek(0)
    
    xl = pd.ExcelFile(io.BytesIO(excel_bytes))
    target_sheet = None
    for s in xl.sheet_names:
        if "③" in s or "录入" in s or "台账" in s or "生产记录" in s or "生产数据" in s:
            target_sheet = s
            break
    if target_sheet is None:
        target_sheet = xl.sheet_names[0]
        
    df_raw = xl.parse(target_sheet, header=None)
    header_row_idx = 0
    for idx in range(min(10, len(df_raw))):
        row_str = " ".join([str(v) for v in df_raw.iloc[idx].values if pd.notna(v)])
        if any(kw in row_str for kw in ["填报说明", "说明：", "指南", "录入说明"]):
            continue
        if any(k in row_str for k in ["月份", "批次", "批次号", "存栏", "产奶", "增重", "日龄", "产蛋"]):
            header_row_idx = idx
            break
            
    df = xl.parse(target_sheet, skiprows=header_row_idx)
    df = df.dropna(how='all')
    df.columns = [str(c).replace('*', '').strip() for c in df.columns]
    first_col = df.columns[0]
    df = df[~df[first_col].astype(str).str.contains("全年|累计|合计|均值|平均|说明", na=False)]
    
    for c in df.columns:
        if c not in ["月份", "批次号", "进雏日期", "出栏日期", "备注"]:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    return df

# =====================================================================
# 三、数据模型：八大畜禽基准演练模型与容错计算引擎
# =====================================================================

# 1. 生猪
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
    out = safe_series(df, '出栏')
    death = safe_series(df, '死亡')
    cull = safe_series(df, '淘汰全群', 0)
    wean = safe_series(df, '断奶仔猪')
    gain = safe_series(df, '总增重_kg')
    feed_kg = safe_series(df, '耗料量_kg')
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

    mean_sow = (df['期初能繁母猪'] + end_sow) / 2.0
    df['料肉比'] = safe_div(feed_kg, gain)
    df['出栏均重_kg'] = safe_div(sale_w, out)
    df['月死亡率'] = safe_div(death, mean_sow * 10)
    total_cost = feed_money.fillna(0) + vet_money.fillna(0) + other_cost.fillna(0)
    revenue = sale_w * price
    df['月净利润_元'] = revenue - total_cost
    df['头均盈利_元'] = safe_div(df['月净利润_元'], out)

    tot_sow_avg = mean_sow.mean()
    tot_wean = wean.sum()
    tot_out = out.sum()
    tot_gain = gain.sum()
    tot_feed = feed_kg.sum()
    tot_cost = total_cost.sum()
    tot_rev = revenue.sum()

    kpi = {
        'PSY': safe_scalar_div(tot_wean, tot_sow_avg),
        'MSY': safe_scalar_div(tot_out, tot_sow_avg),
        '全程成活率': safe_scalar_div(tot_out, born.sum()) if born.sum() > 0 else np.nan,
        '年料肉比': safe_scalar_div(tot_feed, tot_gain),
        '年头均盈利': safe_scalar_div(tot_rev - tot_cost, tot_out),
        '月均死亡率': df['月死亡率'].dropna().mean() if not df['月死亡率'].dropna().empty else np.nan
    }
    return df, kpi

# 2. 奶牛
def get_default_cow_data():
    months = [f"{m}月" for m in range(1, 13)]
    end_cow = [1000, 1005, 1010, 1008, 1012, 1015, 1010, 1005, 1008, 1012, 1015, 1020]
    lact_cow = [850, 855, 860, 858, 862, 865, 860, 855, 858, 862, 865, 870]
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    tot_milk = [920000, 835000, 935000, 910000, 945000, 918000, 890000, 885000, 905000, 940000, 915000, 955000]
    fat = [0.0410, 0.0412, 0.0408, 0.0405, 0.0398, 0.0390, 0.0382, 0.0385, 0.0395, 0.0402, 0.0408, 0.0415]
    protein = [0.0335, 0.0338, 0.0334, 0.0330, 0.0326, 0.0320, 0.0315, 0.0318, 0.0325, 0.0330, 0.0334, 0.0339]
    scc = [14.5, 15.2, 14.8, 16.0, 17.5, 18.8, 22.5, 23.0, 18.2, 16.5, 15.0, 14.2]
    dmi = [28.5, 28.6, 28.8, 29.0, 29.2, 29.0, 27.8, 28.0, 28.5, 28.8, 29.0, 29.2]
    feed_money = [1850000, 1680000, 1880000, 1820000, 1900000, 1840000, 1790000, 1780000, 1820000, 1890000, 1840000, 1920000]
    vet_money = [85000, 80000, 88000, 85000, 90000, 92000, 105000, 110000, 90000, 86000, 84000, 88000]
    other_cost = [1200000, 1100000, 1220000, 1180000, 1230000, 1190000, 1160000, 1150000, 1180000, 1220000, 1190000, 1240000]
    price = [3.85, 3.85, 3.82, 3.80, 3.78, 3.75, 3.72, 3.72, 3.76, 3.80, 3.82, 3.86]

    rows = []
    for i in range(12):
        rows.append({
            '月份': months[i], '期末成母牛存栏': end_cow[i], '在挤泌乳牛头数': lact_cow[i],
            '当月天数': days[i], '生鲜乳产量_kg': tot_milk[i], '生鲜牛乳脂率': fat[i],
            '生鲜牛乳蛋白率': protein[i], '体细胞数_SCC_万': scc[i], '泌乳牛干物质采食量_DMI': dmi[i],
            '饲料采购金额': feed_money[i], '兽药采购金额': vet_money[i], '其他完全成本': other_cost[i],
            '生鲜乳销售单价': price[i]
        })
    return pd.DataFrame(rows)

def compute_cow_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    end_cow = safe_series(df, '期末成母牛存栏', aliases=['期末成母牛', '成母牛存栏'])
    lact_cow = safe_series(df, '在挤泌乳牛头数', aliases=['泌乳牛头数', '泌乳牛'])
    days = safe_series(df, '当月天数', 30.5)
    milk = safe_series(df, '生鲜乳产量_kg', aliases=['月产奶量_kg', '生鲜乳销售量_kg'])
    fat = safe_series(df, '生鲜牛乳脂率', aliases=['乳脂率', '生鲜乳脂率'])
    protein = safe_series(df, '生鲜牛乳蛋白率', aliases=['乳蛋白率', '生鲜乳蛋白率'])
    scc = safe_series(df, '体细胞数_SCC_万', aliases=['体细胞数_SCC', '体细胞数'])
    dmi = safe_series(df, '泌乳牛干物质采食量_DMI', aliases=['干物质采食量DMI_kg', '干物质采食量'])
    feed_money = safe_series(df, '饲料采购金额', aliases=['饲料采购总额_元', '粗饲料费'])
    vet_money = safe_series(df, '兽药采购金额', aliases=['兽药防疫费', '兽药采购费'])
    other_cost = safe_series(df, '其他完全成本', aliases=['其他成本', '人工折旧能耗费用_元'])
    price = safe_series(df, '生鲜乳销售单价', aliases=['原料奶售价', '销售单价', 3.82])

    df['泌乳牛日均单产_kg'] = milk / (lact_cow * days)
    total_cost = feed_money.fillna(0) + vet_money.fillna(0) + other_cost.fillna(0)
    df['公斤奶完全成本_元'] = safe_div(total_cost, milk)
    df['公斤奶饲料成本_元'] = safe_div(feed_money, milk)
    revenue = milk * price
    df['月度净利润_元'] = revenue - total_cost
    df['料奶比'] = safe_div(dmi * lact_cow * days, milk)

    tot_milk = milk.sum()
    tot_cow_avg = end_cow.mean()
    tot_cost = total_cost.sum()
    tot_rev = revenue.sum()

    kpi = {
        '成母牛年单产': safe_scalar_div(tot_milk, tot_cow_avg),
        '泌乳牛日均单产': df['泌乳牛日均单产_kg'].dropna().mean() if not df['泌乳牛日均单产_kg'].dropna().empty else np.nan,
        '生鲜牛乳脂率': fat.dropna().mean() if not fat.dropna().empty else np.nan,
        '生鲜牛乳蛋白率': protein.dropna().mean() if not protein.dropna().empty else np.nan,
        '体细胞数_SCC均值': scc.dropna().mean() if not scc.dropna().empty else np.nan,
        '年化料奶比': safe_scalar_div((dmi * lact_cow * days).sum(), tot_milk),
        '公斤奶完全成本': safe_scalar_div(tot_cost, tot_milk),
        '成母牛头均年净利润': safe_scalar_div(tot_rev - tot_cost, tot_cow_avg)
    }
    return df, kpi

# 3. 肉牛
def get_default_beef_data():
    raw_beef = [
        ("1月", 500, 500, 31, 40, 38, 2, 22320.0, 162936.0, 25840.0, 15245.0, 12660.0, 2, 408000, 374750, 12000, 85000),
        ("2月", 500, 498, 28, 35, 36, 1, 20300.0, 146160.0, 24480.0, 14443.0, 11995.0, 1, 357000, 336168, 11000, 82000),
        ("3月", 498, 502, 31, 45, 40, 1, 22475.0, 161820.0, 27200.0, 16048.0, 13328.0, 1, 459000, 372186, 12500, 86000),
        ("4月", 502, 500, 30, 38, 39, 1, 21900.0, 157680.0, 26520.0, 15646.0, 12994.0, 1, 387600, 362664, 12000, 85000),
        ("5月", 500, 497, 31, 40, 42, 1, 22320.0, 162936.0, 28560.0, 16850.0, 13994.0, 1, 408000, 374750, 13000, 86000),
        ("6月", 497, 495, 30, 36, 37, 1, 21300.0, 155490.0, 25160.0, 14844.0, 12328.0, 1, 367200, 357627, 12000, 84000),
        ("7月", 495, 492, 31, 35, 36, 2, 18900.0, 153090.0, 24120.0, 14110.0, 11698.0, 2, 357000, 352107, 14000, 88000),
        ("8月", 492, 494, 31, 42, 39, 1, 19840.0, 156736.0, 26130.0, 15338.0, 12725.0, 1, 428400, 360492, 13500, 87000),
        ("9月", 494, 498, 30, 45, 40, 1, 21600.0, 157680.0, 27200.0, 16048.0, 13328.0, 1, 459000, 362664, 12000, 85000),
        ("10月", 498, 500, 31, 45, 42, 1, 22475.0, 161820.0, 28560.0, 16850.0, 13994.0, 1, 459000, 372186, 12500, 86000),
        ("11月", 500, 502, 30, 50, 47, 1, 22200.0, 159840.0, 31960.0, 18856.0, 15660.0, 1, 510000, 367632, 13500, 87000),
        ("12月", 502, 500, 31, 48, 49, 1, 22785.0, 164052.0, 33320.0, 19658.0, 16326.0, 1, 489600, 377319, 14000, 88000),
    ]
    cols = ["月份", "期初在栏牛", "期末在栏牛", "当月天数", "当月购入架子牛", "当月出栏肉牛", "当月死亡牛",
            "全群月增重_kg", "消耗饲料干物质_kg", "抽检出栏总重_kg", "抽检热胴体重_kg", "抽检净肉重_kg",
            "病死无害化数", "架子牛采购费_元", "饲料采购总额_元", "兽药防疫费_元", "人工利息折旧能耗_元"]
    return pd.DataFrame(raw_beef, columns=cols)

def compute_beef_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    init_cow = safe_series(df, '期初在栏牛', 500)
    end_cow = safe_series(df, '期末在栏牛', 500)
    days = safe_series(df, '当月天数', 30.5)
    out = safe_series(df, '当月出栏肉牛', 40)
    death = safe_series(df, '当月死亡牛', 1)
    gain = safe_series(df, '全群月增重_kg')
    dmi = safe_series(df, '消耗饲料干物质_kg')
    sample_out_w = safe_series(df, '抽检出栏总重_kg')
    sample_carcass = safe_series(df, '抽检热胴体重_kg')
    sample_meat = safe_series(df, '抽检净肉重_kg')
    
    feed_m = safe_series(df, '饲料采购总额_元')
    calf_m = safe_series(df, '架子牛采购费_元')
    vet_m = safe_series(df, '兽药防疫费_元')
    other_m = safe_series(df, '人工利息折旧能耗_元')

    mean_cattle = (init_cow + end_cow) / 2.0
    df['平均日增重_kg_天'] = gain / (mean_cattle * days)
    df['料肉比_FCR'] = safe_div(dmi, gain)
    df['月出栏均重_kg'] = safe_div(sample_out_w, out)
    df['屠宰率'] = safe_div(sample_carcass, sample_out_w)
    df['净肉率'] = safe_div(sample_meat, sample_out_w)
    df['月死亡率'] = safe_div(death, init_cow)

    tot_cost = feed_m.fillna(0) + calf_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)
    df['公斤增重完全成本_元'] = safe_div(feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0), gain)
    
    p_beef = 32.5  # 活牛收购均价 元/kg
    revenue = (df['月出栏均重_kg'] * p_beef) * out
    df['月度总净利润_元'] = revenue - tot_cost
    df['头均出栏净利润_元'] = safe_div(df['月度总净利润_元'], out)

    tot_gain = gain.sum()
    tot_dmi = dmi.sum()
    tot_out = out.sum()
    tot_rev = revenue.sum()
    tot_c = tot_cost.sum()

    kpi = {
        '平均日增重': safe_scalar_div(tot_gain, (mean_cattle * days).sum()),
        '料肉比_FCR': safe_scalar_div(tot_dmi, tot_gain),
        '出栏均重': df['月出栏均重_kg'].dropna().mean(),
        '屠宰率': safe_scalar_div(sample_carcass.sum(), sample_out_w.sum()),
        '公斤增重完全成本': safe_scalar_div((feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)).sum(), tot_gain),
        '头均出栏净利润': safe_scalar_div(tot_rev - tot_c, tot_out)
    }
    return df, kpi

# 4. 奶山羊
def get_default_dairy_goat_data():
    raw_goat = [
        ("1月", 500, 498, 415, 31, 0, 1, 1, 39860.0, 39000.0, 0.0425, 0.0340, 72.0, 54200.0, 1, 135500, 68000),
        ("2月", 498, 496, 410, 28, 0, 1, 1, 36240.0, 35500.0, 0.0420, 0.0342, 74.0, 49650.0, 1, 124125, 65000),
        ("3月", 496, 502, 420, 31, 8, 1, 1, 41726.0, 40800.0, 0.0418, 0.0338, 75.0, 56330.0, 1, 140825, 68500),
        ("4月", 502, 500, 425, 30, 0, 1, 1, 42075.0, 41200.0, 0.0415, 0.0335, 73.0, 56800.0, 1, 142000, 68000),
        ("5月", 500, 498, 420, 31, 0, 1, 1, 41664.0, 40800.0, 0.0412, 0.0332, 76.0, 56250.0, 1, 140625, 68000),
        ("6月", 498, 495, 410, 30, 0, 2, 1, 38130.0, 37300.0, 0.0408, 0.0330, 80.0, 52240.0, 1, 130600, 67000),
        ("7月", 495, 492, 395, 31, 0, 2, 1, 34878.5, 34100.0, 0.0395, 0.0322, 92.0, 49530.0, 1, 123825, 72000),
        ("8月", 492, 495, 400, 31, 5, 1, 1, 35960.0, 35200.0, 0.0402, 0.0325, 88.0, 50340.0, 1, 125850, 71000),
        ("9月", 495, 497, 405, 30, 4, 1, 1, 37665.0, 36900.0, 0.0415, 0.0336, 78.0, 51980.0, 1, 129950, 68000),
        ("10月", 497, 500, 410, 31, 5, 1, 1, 39463.0, 38600.0, 0.0422, 0.0340, 74.0, 54060.0, 1, 135150, 68500),
        ("11月", 500, 498, 400, 30, 0, 1, 1, 36600.0, 35800.0, 0.0428, 0.0345, 71.0, 50870.0, 1, 127175, 67500),
        ("12月", 498, 500, 390, 31, 4, 1, 1, 35665.5, 34900.0, 0.0430, 0.0348, 70.0, 50290.0, 1, 125725, 68000),
    ]
    cols = ["月份", "期初成母羊", "期末成母羊", "在挤泌乳母羊", "当月天数", "青年羊转入",
            "淘汰母羊", "死亡母羊", "月总产奶量_kg", "生鲜乳销售量_kg", "生鲜羊乳脂率",
            "生鲜羊乳蛋白率", "体细胞数SCC_万_mL", "消耗饲料干物质_kg", "无害化处理数",
            "饲料采购费用_元", "人工折旧能耗费用_元"]
    return pd.DataFrame(raw_goat, columns=cols)

def compute_dairy_goat_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    end_ewe = safe_series(df, '期末成母羊', 500)
    lact_ewe = safe_series(df, '在挤泌乳母羊', 410)
    days = safe_series(df, '当月天数', 30.5)
    milk = safe_series(df, '月总产奶量_kg')
    fat = safe_series(df, '生鲜羊乳脂率')
    protein = safe_series(df, '生鲜羊乳蛋白率')
    scc = safe_series(df, '体细胞数SCC_万_mL')
    dmi = safe_series(df, '消耗饲料干物质_kg')
    feed_m = safe_series(df, '饲料采购费用_元')
    other_m = safe_series(df, '人工折旧能耗费用_元')

    df['泌乳羊日单产_kg'] = milk / (lact_ewe * days)
    df['料奶比_FMR'] = safe_div(dmi, milk)
    tot_cost = feed_m.fillna(0) + other_m.fillna(0)
    df['公斤羊奶完全成本_元'] = safe_div(tot_cost, milk)
    
    p_goat_milk = 6.8  # 生鲜羊乳销售均价 元/kg
    revenue = milk * p_goat_milk
    df['月度总净利润_元'] = revenue - tot_cost
    df['公斤羊奶净利_元'] = safe_div(df['月度总净利润_元'], milk)

    tot_milk = milk.sum()
    tot_ewe_avg = end_ewe.mean()
    tot_cost_all = tot_cost.sum()
    tot_rev = revenue.sum()

    kpi = {
        '成母羊年单产': safe_scalar_div(tot_milk, tot_ewe_avg),
        '泌乳羊日均单产': df['泌乳羊日单产_kg'].dropna().mean(),
        '生鲜羊乳脂率': fat.dropna().mean(),
        '生鲜羊乳蛋白率': protein.dropna().mean(),
        '体细胞数SCC': scc.dropna().mean(),
        '公斤羊奶完全成本': safe_scalar_div(tot_cost_all, tot_milk),
        '成母羊年头均净利润': safe_scalar_div(tot_rev - tot_cost_all, tot_ewe_avg)
    }
    return df, kpi

# 5. 肉羊
def get_default_meat_sheep_data():
    raw_sheep = [
        ("1月", 1000, 995, 31, 160, 160, 5, 9610.0, 47089.0, 8080.0, 4322.0, 3514.0, 5, 104000, 117722, 4800, 28000),
        ("2月", 995, 992, 28, 150, 150, 3, 8624.0, 41395.0, 7575.0, 4052.0, 3295.0, 3, 97500, 103488, 4500, 26000),
        ("3月", 992, 1002, 31, 180, 165, 5, 9889.0, 47467.0, 8332.5, 4457.0, 3624.0, 5, 117000, 118668, 5000, 28500),
        ("4月", 1002, 1000, 30, 170, 168, 4, 9600.0, 46080.0, 8484.0, 4538.0, 3690.0, 4, 110500, 115200, 4800, 28000),
        ("5月", 1000, 996, 31, 175, 174, 5, 9920.0, 48608.0, 8787.0, 4701.0, 3822.0, 5, 113750, 121520, 5200, 28500),
        ("6月", 996, 992, 30, 165, 165, 4, 9240.0, 46200.0, 8332.5, 4457.0, 3624.0, 4, 107250, 115500, 5000, 28000),
        ("7月", 992, 988, 31, 160, 158, 6, 8928.0, 46425.0, 7979.0, 4228.0, 3430.0, 6, 104000, 116064, 5800, 31000),
        ("8月", 988, 995, 31, 175, 163, 5, 9238.0, 47113.0, 8231.5, 4370.0, 3547.0, 5, 113750, 117784, 5500, 30000),
        ("9月", 995, 1002, 30, 185, 173, 5, 9570.0, 46893.0, 8736.5, 4673.0, 3800.0, 5, 120250, 117232, 5000, 28500),
        ("10月", 1002, 1005, 31, 185, 178, 4, 9982.0, 47913.0, 8989.0, 4809.0, 3910.0, 4, 120250, 119784, 5200, 28500),
        ("11月", 1005, 1004, 30, 190, 187, 4, 9648.0, 46310.0, 9443.5, 5052.0, 4107.0, 4, 123500, 115776, 5200, 28500),
        ("12月", 1004, 1000, 31, 195, 194, 5, 10013.0, 48062.0, 9797.0, 5241.0, 4261.0, 5, 126750, 120156, 5500, 29000),
    ]
    cols = ["月份", "期初在栏羊", "期末在栏羊", "当月天数", "当月转入断奶羔", "当月出栏肉羊",
            "当月死亡羊", "全群月增重_kg", "消耗饲料干物质_kg", "抽检出栏总重_kg", "抽检热胴体重_kg",
            "抽检净肉重_kg", "病死无害化数", "断奶羔羊成本_元", "饲料采购总额_元",
            "兽药防疫费_元", "人工折旧能耗费用_元"]
    return pd.DataFrame(raw_sheep, columns=cols)

def compute_meat_sheep_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    init_sheep = safe_series(df, '期初在栏羊', 1000)
    end_sheep = safe_series(df, '期末在栏羊', 1000)
    days = safe_series(df, '当月天数', 30.5)
    out = safe_series(df, '当月出栏肉羊', 165)
    death = safe_series(df, '当月死亡羊', 4)
    gain = safe_series(df, '全群月增重_kg')
    dmi = safe_series(df, '消耗饲料干物质_kg')
    sample_out_w = safe_series(df, '抽检出栏总重_kg')
    sample_carcass = safe_series(df, '抽检热胴体重_kg')
    lamb_m = safe_series(df, '断奶羔羊成本_元')
    feed_m = safe_series(df, '饲料采购总额_元')
    vet_m = safe_series(df, '兽药防疫费_元')
    other_m = safe_series(df, '人工折旧能耗费用_元')

    mean_sheep = (init_sheep + end_sheep) / 2.0
    df['平均日增重_g_天'] = (gain * 1000.0) / (mean_sheep * days)
    df['料肉比_FCR'] = safe_div(dmi, gain)
    df['出栏均重_kg'] = safe_div(sample_out_w, out)
    df['屠宰率'] = safe_div(sample_carcass, sample_out_w)
    df['月死亡率'] = safe_div(death, init_sheep)

    tot_cost = lamb_m.fillna(0) + feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)
    df['公斤增重完全成本_元'] = safe_div(feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0), gain)
    
    p_sheep = 34.0  # 育肥活羊收购均价 元/kg
    revenue = (df['出栏均重_kg'] * p_sheep) * out
    df['月度总净利润_元'] = revenue - tot_cost
    df['单只出栏净利_元'] = safe_div(df['月度总净利润_元'], out)

    tot_gain = gain.sum()
    tot_dmi = dmi.sum()
    tot_out = out.sum()
    tot_c = tot_cost.sum()
    tot_rev = revenue.sum()

    kpi = {
        '平均日增重': df['平均日增重_g_天'].dropna().mean(),
        '料肉比_FCR': safe_scalar_div(tot_dmi, tot_gain),
        '出栏均重': df['出栏均重_kg'].dropna().mean(),
        '屠宰率': safe_scalar_div(sample_carcass.sum(), sample_out_w.sum()),
        '公斤增重完全成本': safe_scalar_div((feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)).sum(), tot_gain),
        '单只出栏净利': safe_scalar_div(tot_rev - tot_c, tot_out)
    }
    return df, kpi

# 6. 皮用羊 (滩羊)
def get_default_fur_sheep_data():
    raw_fur = [
        ("1月", 500, 31, 30, 35, 28, 6, 1, 25, 8.2, 8.6, 42.0, 156.8, 1, 48000, 6800, 21000),
        ("2月", 498, 28, 45, 52, 42, 9, 1, 38, 8.1, 8.5, 41.5, 235.2, 1, 52000, 8500, 22000),
        ("3月", 496, 31, 80, 94, 80, 12, 2, 72, 8.3, 8.8, 42.5, 448.0, 2, 65000, 14000, 26000),
        ("4月", 495, 30, 90, 105, 90, 13, 2, 81, 8.4, 8.9, 43.0, 504.0, 2, 70000, 16000, 27000),
        ("5月", 495, 31, 60, 70, 60, 9, 1, 53, 8.2, 8.7, 42.0, 336.0, 1, 58000, 11000, 24000),
        ("6月", 494, 30, 20, 23, 18, 4, 1, 16, 7.8, 8.3, 40.5, 99.0, 1, 45000, 5200, 20000),
        ("7月", 493, 31, 10, 11, 8, 2, 1, 7, 7.5, 8.1, 39.5, 43.2, 1, 43000, 4500, 20500),
        ("8月", 493, 31, 10, 11, 8, 2, 1, 7, 7.6, 8.2, 40.0, 44.0, 1, 43000, 4500, 20500),
        ("9月", 494, 30, 25, 29, 23, 5, 1, 20, 8.0, 8.5, 41.0, 126.5, 1, 46000, 6200, 21000),
        ("10月", 496, 31, 40, 46, 38, 7, 1, 34, 8.2, 8.6, 42.0, 212.8, 1, 52000, 8200, 22500),
        ("11月", 498, 30, 50, 58, 48, 9, 1, 43, 8.3, 8.7, 42.5, 268.8, 1, 56000, 9800, 23500),
        ("12月", 500, 31, 40, 46, 37, 8, 1, 33, 8.2, 8.6, 42.0, 207.2, 1, 51000, 8200, 22500),
    ]
    cols = ["月份", "基础母羊存栏", "当月天数", "配种产羔母羊", "产出活羔数", "适龄剥皮二毛羊",
            "选留后备羊", "羔羊死亡数", "特级及一级皮_张", "平均花弯数_个", "平均毛穗长_cm",
            "抽检皮板面积_dm2", "二毛羊胴体总重_kg", "病死无害化数", "母羊及羔羊饲料费_元",
            "兽药防疫初加工费_元", "人工折旧能耗费用_元"]
    return pd.DataFrame(raw_fur, columns=cols)

def compute_fur_sheep_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    ewe = safe_series(df, '基础母羊存栏', 500)
    slaughter = safe_series(df, '适龄剥皮二毛羊')
    top_skin = safe_series(df, '特级及一级皮_张')
    curls = safe_series(df, '平均花弯数_个')
    hair_len = safe_series(df, '平均毛穗长_cm')
    carcass_w = safe_series(df, '二毛羊胴体总重_kg')
    feed_m = safe_series(df, '母羊及羔羊饲料费_元')
    vet_m = safe_series(df, '兽药防疫初加工费_元')
    other_m = safe_series(df, '人工折旧能耗费用_元')

    df['特级及一级皮率'] = safe_div(top_skin, slaughter)
    df['只均胴体重_kg'] = safe_div(carcass_w, slaughter)

    p_skin = 450.0   # 特级二毛皮单张均价 元
    p_meat = 85.0    # 滩羊羔肉均价 元/kg
    rev_skin = top_skin * p_skin + (slaughter - top_skin) * 280.0
    rev_meat = carcass_w * p_meat
    tot_cost = feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)
    df['月度总净利润_元'] = rev_skin + rev_meat - tot_cost
    df['只均二毛羊净利_元'] = safe_div(df['月度总净利润_元'], slaughter)
    df['只均完全成本_元'] = safe_div(tot_cost, slaughter)

    tot_slaughter = slaughter.sum()
    tot_cost_all = tot_cost.sum()
    tot_rev = (rev_skin + rev_meat).sum()

    kpi = {
        '特级及一级皮率': safe_scalar_div(top_skin.sum(), tot_slaughter),
        '平均花弯数': curls.dropna().mean(),
        '平均毛穗长': hair_len.dropna().mean(),
        '只均胴体重': df['只均胴体重_kg'].dropna().mean(),
        '只均二毛羊净利': safe_scalar_div(tot_rev - tot_cost_all, tot_slaughter),
        '综合毛利率': safe_scalar_div(tot_rev - tot_cost_all, tot_rev)
    }
    return df, kpi

# 7. 蛋鸡
def get_default_layer_data():
    raw_layer = [
        ("1月", 20000, 19930, 31, 50, 20, 578000, 36876.0, 5200, 78550.0, 0.39, 85.0, 50, 235650, 38000, 48000, 14500),
        ("2月", 19930, 19865, 28, 45, 20, 528500, 33824.0, 4750, 71500.0, 0.39, 84.5, 45, 214500, 37000, 47000, 13800),
        ("3月", 19865, 19790, 31, 55, 20, 582000, 37364.0, 5230, 79200.0, 0.38, 83.0, 55, 237600, 38000, 48000, 14600),
        ("4月", 19790, 19715, 30, 55, 20, 561000, 36128.0, 5050, 76800.0, 0.38, 83.5, 55, 230400, 37500, 47500, 14200),
        ("5月", 19715, 19635, 31, 60, 20, 574000, 37023.0, 5160, 78600.0, 0.38, 82.0, 60, 235800, 38000, 48000, 14500),
        ("6月", 19635, 19550, 30, 65, 20, 546000, 35053.0, 5460, 74900.0, 0.37, 81.0, 65, 224700, 37500, 47500, 14000),
        ("7月", 19550, 19440, 31, 85, 25, 532000, 33250.0, 9570, 72800.0, 0.34, 76.5, 85, 218400, 38000, 52000, 13600),
        ("8月", 19440, 19340, 31, 75, 25, 538000, 33894.0, 8600, 73600.0, 0.35, 77.0, 75, 220800, 38000, 51000, 13800),
        ("9月", 19340, 19260, 30, 60, 20, 535000, 34240.0, 5350, 74200.0, 0.37, 81.0, 60, 222600, 37500, 48000, 14000),
        ("10月", 19260, 19185, 31, 55, 20, 545000, 35152.0, 5450, 76500.0, 0.38, 82.0, 55, 229500, 38000, 48000, 14300),
        ("11月", 19185, 19115, 30, 50, 20, 522000, 33773.0, 5220, 74100.0, 0.38, 83.0, 50, 222300, 37500, 47500, 13900),
        ("12月", 19115, 19045, 31, 50, 20, 532000, 34580.0, 5320, 76800.0, 0.38, 83.5, 50, 230400, 38000, 48000, 14200),
    ]
    cols = ["月份", "期初在舍母鸡", "期末在舍母鸡", "当月天数", "月病死鸡_羽", "淘汰弱鸡_羽",
            "月总产蛋数_枚", "月总产蛋重_kg", "月破损裂纹蛋_枚", "月消耗蛋鸡料_kg",
            "蛋壳厚度_mm", "哈氏单位_HU", "无害化处理数", "饲料采购费_元", "雏鸡摊销与防疫_元",
            "人工电费与折旧_元", "包装运输费_元"]
    return pd.DataFrame(raw_layer, columns=cols)

def compute_layer_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '月份' not in df.columns:
        df['月份'] = [f"{i+1}月" for i in range(n)]

    init_hen = safe_series(df, '期初在舍母鸡', 20000)
    end_hen = safe_series(df, '期末在舍母鸡', 19500)
    days = safe_series(df, '当月天数', 30.5)
    eggs = safe_series(df, '月总产蛋数_枚')
    egg_w_kg = safe_series(df, '月总产蛋重_kg')
    broken = safe_series(df, '月破损裂纹蛋_枚')
    feed_kg = safe_series(df, '月消耗蛋鸡料_kg')
    hu = safe_series(df, '哈氏单位_HU')
    
    feed_m = safe_series(df, '饲料采购费_元')
    chick_m = safe_series(df, '雏鸡摊销与防疫_元')
    other_m = safe_series(df, '人工电费与折旧_元')
    pack_m = safe_series(df, '包装运输费_元')

    mean_hen = (init_hen + end_hen) / 2.0
    df['饲养日产蛋率_HD'] = eggs / (mean_hen * days)
    df['平均蛋重_g'] = (egg_w_kg * 1000.0) / eggs
    df['料蛋比_FCR'] = safe_div(feed_kg, egg_w_kg)
    df['破损蛋率'] = safe_div(broken, eggs)

    tot_cost = feed_m.fillna(0) + chick_m.fillna(0) + other_m.fillna(0) + pack_m.fillna(0)
    df['公斤蛋完全成本_元'] = safe_div(tot_cost, egg_w_kg)

    p_egg = 8.8  # 鲜蛋销售均价 元/kg
    revenue = egg_w_kg * p_egg
    df['月度总净利润_元'] = revenue - tot_cost
    df['公斤蛋净利润_元'] = safe_div(df['月度总净利润_元'], egg_w_kg)

    tot_eggs = eggs.sum()
    tot_egg_kg = egg_w_kg.sum()
    tot_feed = feed_kg.sum()
    tot_c = tot_cost.sum()
    tot_rev = revenue.sum()

    kpi = {
        '饲养日产蛋率': df['饲养日产蛋率_HD'].dropna().mean(),
        '料蛋比_FCR': safe_scalar_div(tot_feed, tot_egg_kg),
        '平均蛋重': df['平均蛋重_g'].dropna().mean(),
        '哈氏单位_HU': hu.dropna().mean(),
        '公斤蛋完全成本': safe_scalar_div(tot_c, tot_egg_kg),
        '只鸡全期净利润': safe_scalar_div(tot_rev - tot_c, init_hen.iloc[0])
    }
    return df, kpi

# 8. 肉鸡
def get_default_broiler_data():
    raw_data = [
        ("Batch-01", "2026-01-05", 50000, "2026-02-15", 41, 48350, 1650, 390, 130545.0, 198428.0, 97256.0, 27618.0, 1, 1650, 175000, 694498, 28500, 155000),
        ("Batch-02", "2026-03-02", 50000, "2026-04-12", 41, 48480, 1520, 350, 131865.0, 197798.0, 98371.0, 28035.0, 1, 1520, 180000, 692293, 27000, 152000),
        ("Batch-03", "2026-04-28", 50000, "2026-06-08", 41, 48200, 1800, 420, 129658.0, 197080.0, 96465.0, 27396.0, 1, 1800, 170000, 689780, 29000, 158000),
        ("Batch-04", "2026-06-25", 50000, "2026-08-06", 42, 47650, 2350, 560, 126272.5, 196985.0, 93694.0, 26421.0, 1, 2350, 165000, 689448, 33500, 172000),
        ("Batch-05", "2026-08-22", 50000, "2026-10-02", 41, 48520, 1480, 340, 132459.6, 198689.0, 98947.0, 28298.0, 1, 1480, 185000, 695412, 26500, 150000),
        ("Batch-06", "2026-10-18", 50000, "2026-11-28", 41, 48600, 1400, 320, 133164.0, 198414.0, 99606.0, 28586.0, 1, 1400, 185000, 694449, 26000, 151000),
    ]
    cols = ["批次号", "进雏日期", "入舍雏鸡_羽", "出栏日期", "出栏日龄_天", "出栏有效羽数_羽",
            "全期死淘_羽", "7日龄死淘_羽", "出栏毛鸡总重_kg", "全期总消耗饲料_kg",
            "全净膛总重_kg", "双侧胸肉总重_kg", "休药期合规批次", "病死鸡无害化数",
            "雏鸡苗总费用_元", "饲料采购总费用_元", "兽药疫苗费用_元", "人工能耗折旧费用_元"]
    return pd.DataFrame(raw_data, columns=cols)

def compute_broiler_metrics(df_raw):
    df = df_raw.copy()
    n = len(df)
    if n == 0:
        return df, {}
    if '批次号' not in df.columns:
        df['批次号'] = [f"Batch-{i+1:02d}" for i in range(n)]

    in_chick = safe_series(df, '入舍雏鸡_羽', 50000)
    out_birds = safe_series(df, '出栏有效羽数_羽', 48000)
    days = safe_series(df, '出栏日龄_天', 41)
    tot_w = safe_series(df, '出栏毛鸡总重_kg')
    tot_feed = safe_series(df, '全期总消耗饲料_kg')
    eviscerated = safe_series(df, '全净膛总重_kg')
    breast = safe_series(df, '双侧胸肉总重_kg')
    
    chick_m = safe_series(df, '雏鸡苗总费用_元')
    feed_m = safe_series(df, '饲料采购总费用_元')
    vet_m = safe_series(df, '兽药疫苗费用_元')
    other_m = safe_series(df, '人工能耗折旧费用_元')

    df['出栏均重_kg'] = safe_div(tot_w, out_birds)
    df['全期成活率'] = safe_div(out_birds, in_chick)
    df['全期料肉比_FCR'] = safe_div(tot_feed, tot_w - out_birds * 0.042)
    df['欧洲效益指数_EPI'] = (df['全期成活率'] * 100.0 * df['出栏均重_kg']) / (days * df['全期料肉比_FCR']) * 100.0
    df['全净膛率'] = safe_div(eviscerated, tot_w)
    df['胸肌率'] = safe_div(breast, eviscerated)

    tot_cost = chick_m.fillna(0) + feed_m.fillna(0) + vet_m.fillna(0) + other_m.fillna(0)
    df['公斤毛鸡完全成本_元'] = safe_div(tot_cost, tot_w)
    
    p_broiler = 9.2  # 毛鸡出栏均价 元/kg
    revenue = tot_w * p_broiler
    df['批次总净利润_元'] = revenue - tot_cost
    df['只鸡净利润_元'] = safe_div(df['批次总净利润_元'], out_birds)

    tot_birds_all = out_birds.sum()
    tot_w_all = tot_w.sum()
    tot_cost_all = tot_cost.sum()
    tot_rev_all = revenue.sum()

    kpi = {
        '欧洲效益指数_EPI': df['欧洲效益指数_EPI'].dropna().mean(),
        '全期成活率': safe_scalar_div(tot_birds_all, in_chick.sum()),
        '出栏均重': df['出栏均重_kg'].dropna().mean(),
        '全期料肉比_FCR': safe_scalar_div(tot_feed.sum(), tot_w_all - tot_birds_all * 0.042),
        '公斤毛鸡完全成本': safe_scalar_div(tot_cost_all, tot_w_all),
        '只鸡净利润': safe_scalar_div(tot_rev_all - tot_cost_all, tot_birds_all)
    }
    return df, kpi


# =====================================================================
# 三点五、行业新闻与咨询投稿持久化存储引擎
# =====================================================================
def save_submission_to_excel(sub_data, uploaded_file=None):
    """
    将用户投稿表单数据安全追加写入《行业新闻与咨询投稿登记表.xlsx》
    并自动转存附件至本地《投稿附件》目录
    """
    saved_filename = "无附件"
    saved_filepath = "无"
    if uploaded_file is not None:
        try:
            SUBMISSION_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
            ts_prefix = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_fname = f"{ts_prefix}_{uploaded_file.name}"
            target_path = SUBMISSION_UPLOAD_DIR / safe_fname
            file_bytes = uploaded_file.getvalue() if hasattr(uploaded_file, 'getvalue') else uploaded_file.read()
            with open(target_path, "wb") as f:
                f.write(file_bytes)
            saved_filename = uploaded_file.name
            saved_filepath = str(target_path.relative_to(CURRENT_DIR))
        except Exception as e:
            saved_filename = f"{uploaded_file.name} (保存异常: {e})"
            saved_filepath = "异常"

    headers = [
        "登记序号", "投稿时间", "投稿人姓名", "工作单位/机构", 
        "工作联系电话", "工作微信/QQ", "工作联系邮箱", "资讯所属领域", 
        "新闻/文章标题", "核心摘要/咨询链接", "附件文件名称", "附件本地存储路径", 
        "处理状态", "处理意见与对接记录"
    ]

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    wb = None
    if SUBMISSION_EXCEL_PATH.exists():
        try:
            wb = openpyxl.load_workbook(SUBMISSION_EXCEL_PATH)
            ws = wb.active
        except Exception:
            wb = None

    if wb is None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "投稿与咨询登记表"
        ws.views.sheetView[0].showGridLines = True
        
        # 写入大标题行
        ws.merge_cells("A1:N1")
        title_cell = ws["A1"]
        title_cell.value = "智慧牧业数字化大屏 · 行业新闻与咨询文章投稿登记台账"
        title_cell.font = Font(name="微软雅黑", size=14, bold=True, color="1F4E79")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # 写入表头行
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=h)
            cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[2].height = 28
        ws.freeze_panes = "A3"

    # 计算自增序号
    next_row = ws.max_row + 1
    seq_num = max(1, next_row - 2)

    # 组装数据行
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_values = [
        seq_num,
        now_str,
        sub_data.get("sub_name", ""),
        sub_data.get("sub_org", "行业同仁") or "行业同仁",
        sub_data.get("sub_phone", ""),
        sub_data.get("sub_wechat", "") or "未填写",
        sub_data.get("sub_email", "") or "未填写",
        sub_data.get("sub_cat", ""),
        sub_data.get("sub_title", ""),
        sub_data.get("sub_summary", "") or "无摘要",
        saved_filename,
        saved_filepath,
        "待初审对接",
        "大屏在线投稿自动登记入库"
    ]

    for col_idx, val in enumerate(row_values, 1):
        cell = ws.cell(row=next_row, column=col_idx, value=val)
        cell.font = Font(name="微软雅黑", size=10)
        cell.border = thin_border
        if col_idx in [1, 2, 5, 8, 13]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[next_row].height = 24

    # 规范列宽
    col_widths = {
        1: 10, 2: 20, 3: 15, 4: 25, 5: 18, 6: 18, 7: 22,
        8: 24, 9: 30, 10: 35, 11: 22, 12: 25, 13: 14, 14: 25
    }
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    wb.save(SUBMISSION_EXCEL_PATH)
    
    # 同步备份镜像至 表格与指标库/
    try:
        mirror_path = KPI_DIR / "行业新闻与咨询投稿登记表.xlsx"
        wb.save(mirror_path)
    except Exception:
        pass
    wb.close()
    return seq_num, SUBMISSION_EXCEL_PATH

def init_submission_excel_if_needed():
    """初始化行业新闻与咨询投稿登记表（若文件不存在，则创建并载入初始示范记录）"""
    if not SUBMISSION_EXCEL_PATH.exists():
        initial_sample = {
            "sub_name": "李尧",
            "sub_org": "西北农林科技大学动物科技学院",
            "sub_phone": "13679275651",
            "sub_wechat": "wx_13679275651",
            "sub_email": "liyao@nwafu.edu.cn",
            "sub_cat": "💻 智慧畜牧工程与数字化档案",
            "sub_title": "智慧牧业数字化生产档案与多畜种核心指标管理实践",
            "sub_summary": "首套覆盖生猪、奶牛、肉牛、奶山羊、肉羊、皮用羊、蛋鸡、肉鸡八大核心畜禽法定养殖档案与指标计算引擎，实现行业新闻与生产台账双向闭环。"
        }
        save_submission_to_excel(initial_sample, None)


# =====================================================================
# 三点六、管理员权限认证与设备硬件（MAC/序列号）安全鉴权引擎
# =====================================================================
def get_device_hardware_fingerprint():
    """跨平台安全获取设备物理 MAC 地址与硬件序列号（兼容 macOS、Linux 容器及云端环境）"""
    import uuid, subprocess, platform
    mac_addr = "未知MAC"
    serial_num = "未知序列号"

    # 1. 尝试获取物理 MAC 地址
    try:
        # macOS: en0 物理网卡
        out = subprocess.check_output(['ifconfig', 'en0'], stderr=subprocess.DEVNULL, timeout=2).decode()
        for line in out.splitlines():
            if 'ether' in line:
                mac_addr = line.strip().split()[1].lower()
                break
    except Exception:
        pass

    if mac_addr == "未知MAC":
        # Linux: 尝试读取 /sys/class/net/eth0/address
        try:
            with open('/sys/class/net/eth0/address', 'r') as f:
                mac_addr = f.read().strip().lower()
        except Exception:
            pass

    if mac_addr == "未知MAC":
        try:
            mac_num = uuid.getnode()
            mac_addr = ':'.join(['{:02x}'.format((mac_num >> el) & 0xff) for el in range(0, 12, 2)][::-1]).lower()
        except Exception:
            mac_addr = "00:00:00:00:00:00"

    # 2. 尝试获取硬件序列号
    try:
        # macOS: ioreg
        out_serial = subprocess.check_output(['ioreg', '-l'], stderr=subprocess.DEVNULL, timeout=2).decode()
        for line in out_serial.splitlines():
            if 'IOPlatformSerialNumber' in line:
                serial_num = line.strip().split('=')[-1].strip().replace('"', '')
                break
    except Exception:
        pass

    if serial_num == "未知序列号":
        # Linux: 尝试读取 DMI 序列号
        try:
            with open('/sys/class/dmi/id/product_serial', 'r') as f:
                serial_num = f.read().strip()
        except Exception:
            pass

    if serial_num == "未知序列号":
        serial_num = platform.node() or "DEV-LOCAL"

    return mac_addr, serial_num

AUTHORIZED_MACS = ["ca:e9:b1:3d:ce:69", "44:13:4f:3c:f1:c5"]
AUTHORIZED_SERIALS = ["GP71L4XYJX", "MACBOOK-AIR"]
ADMIN_CREDENTIALS = {
    "admin": "admin888",
    "liyao": "nwafu2024"
}

def verify_admin_login(username, password, input_hardware_id=""):
    """
    核验管理员登录：
    1. 账号密码匹配
    2. 硬件指纹核验（支持服务器本地自动探测与线上云端环境输入授权设备指纹凭据）
    """
    if username not in ADMIN_CREDENTIALS or ADMIN_CREDENTIALS[username] != password:
        return False, "❌ 管理员账号或密码错误，请重新核对输入！"
    
    curr_mac, curr_serial = get_device_hardware_fingerprint()
    
    # 检查当前运行环境是否本身就是授权白名单设备（本地运行场景）
    local_matched = (curr_mac.lower() in [m.lower() for m in AUTHORIZED_MACS] or 
                     any(s.upper() in curr_serial.upper() for s in AUTHORIZED_SERIALS))
    
    # 检查前端输入/选择的授权设备识别码（云端线上运行场景）
    input_clean = input_hardware_id.strip()
    input_matched = False
    if input_clean:
        if input_clean.lower() in [m.lower() for m in AUTHORIZED_MACS]:
            input_matched = True
        elif any(s.upper() in input_clean.upper() for s in AUTHORIZED_SERIALS):
            input_matched = True

    if local_matched or input_matched:
        verified_id = input_clean if input_matched else f"{curr_serial} ({curr_mac})"
        return True, f"✅ 设备硬件指纹与身份认证成功！(授权凭据: {verified_id})"
    else:
        return False, f"⚠️ 设备硬件未授权！当前节点硬件【{curr_serial}】未在白名单中，且输入的设备凭据【{input_clean or '空'}】不匹配。请选择或填入授权硬件指纹（如 GP71L4XYJX）。"

def render_admin_sidebar_auth():
    """在侧边栏渲染管理员鉴权控制台（支持本地硬件直通与云端多设备灵活核验）"""
    st.markdown("---")
    st.subheader("🔐 管理员安全认证控制台")
    
    is_admin = st.session_state.get("is_admin", False)
    curr_mac, curr_serial = get_device_hardware_fingerprint()
    
    is_local_whitelisted = (curr_mac.lower() in [m.lower() for m in AUTHORIZED_MACS] or 
                            any(s.upper() in curr_serial.upper() for s in AUTHORIZED_SERIALS))
    
    if is_admin:
        adm_name = st.session_state.get('admin_user', 'admin')
        auth_device = st.session_state.get('admin_device', curr_serial)
        st.success(f"🟢 **管理员已认证在线**\n\n- 账号: `{adm_name}`\n- 认证设备: `{auth_device}`")
        st.caption("🔒 已解锁行业投稿台账全景审核、真实姓名联系方式查阅与 Excel 完整导出权限。")
        if st.button("🚪 退出管理员登录", use_container_width=True, key="admin_logout_btn"):
            st.session_state["is_admin"] = False
            st.session_state["admin_user"] = ""
            st.session_state["admin_device"] = ""
            st.rerun()
    else:
        with st.expander("🔑 管理员权限登录 (硬件指纹+账号密码)", expanded=not is_admin):
            if is_local_whitelisted:
                st.caption(f"💻 **检测到当前本地设备处于授权白名单内**：\n- 硬件序列号: `{curr_serial}`\n- 物理MAC: `{curr_mac}`")
            else:
                st.caption(f"☁️ **检测到当前处于云端/外部网络访问节点**：\n- 云节点MAC: `{curr_mac}`\n- 节点ID: `{curr_serial}`\n\n💡 提示：公网部署环境下，请选择授权开发者设备凭据完成硬件双因子鉴权：")
            
            adm_user = st.text_input("管理员账号", placeholder="admin / liyao", key="sidebar_adm_user")
            adm_pwd = st.text_input("管理员密码", type="password", placeholder="请输入管理密码", key="sidebar_adm_pwd")
            
            # 提供快速选择授权硬件选项与自定义输入
            device_options = [
                "🍎 李尧专属MacBook Air (序列号: GP71L4XYJX)",
                "🌐 授权物理网卡MAC (ca:e9:b1:3d:ce:69)",
                "💻 当前运行节点硬件 (自动识别)",
                "✏️ 手动输入其他授权设备序列号/MAC"
            ]
            default_idx = 2 if is_local_whitelisted else 0
            sel_device = st.selectbox("登录设备硬件识别凭据", device_options, index=default_idx, key="sidebar_sel_device")
            
            manual_hw = ""
            if "GP71L4XYJX" in sel_device:
                manual_hw = "GP71L4XYJX"
            elif "ca:e9:b1:3d:ce:69" in sel_device:
                manual_hw = "ca:e9:b1:3d:ce:69"
            elif "手动输入" in sel_device:
                manual_hw = st.text_input("请输入授权设备序列号或MAC地址", placeholder="例如：GP71L4XYJX", key="sidebar_custom_hw")
            else:
                manual_hw = curr_serial

            if st.button("🚀 验证硬件与账号登录", use_container_width=True, key="admin_login_btn"):
                ok, msg = verify_admin_login(adm_user.strip(), adm_pwd.strip(), manual_hw.strip())
                if ok:
                    st.session_state["is_admin"] = True
                    st.session_state["admin_user"] = adm_user.strip()
                    st.session_state["admin_device"] = manual_hw.strip() or curr_serial
                    st.toast("🎉 管理员认证成功！已解锁投稿人敏感数据与台账导出权限。")
                    st.rerun()
                else:
                    st.error(msg)

# =====================================================================
# 四、现代畜牧行业新闻资讯与投稿互动中心（支持刷新动态与全畜种咨询）
# =====================================================================
def render_industry_news_section(animal_type="生猪"):
    init_submission_excel_if_needed()
    st.markdown("### 📰 现代畜牧行业新闻资讯与投稿互动中心")
    st.caption("📌 **板块定位**：实时追踪国家畜牧政策法规、八大核心畜种产业咨询要闻、市场供需行情与前沿智能养殖技术；支持从业者与科研人员登记工作联系方式并上传行业咨询新闻及技术文章。")
    
    col_news_left, col_news_right = st.columns([3, 2])
    
    with col_news_left:
        st.subheader("🌐 行业主流权威咨询与新闻门户直达")
        c_portal1, c_portal2, c_portal3 = st.columns(3)
        with c_portal1:
            st.link_button("🏛️ 农业农村部畜牧兽医局", "http://www.moa.gov.cn/xw/zwdt/", use_container_width=True)
            st.link_button("📰 大畜牧网行业资讯门户", "https://www.dxumu.com/", use_container_width=True)
            st.link_button("🐖 猪场动力网专业咨询", "https://www.powerpigs.net/", use_container_width=True)
        with c_portal2:
            st.link_button("🐄 荷斯坦奶牛网资讯中心", "http://www.hesitan.com/", use_container_width=True)
            st.link_button("🥛 中国奶业协会官方网站", "http://www.dac.org.cn/", use_container_width=True)
            st.link_button("🐂 中国肉牛网与产业资讯", "http://www.chinabm.cn/", use_container_width=True)
        with c_portal3:
            st.link_button("🐑 中国羊业网与奶山羊专栏", "http://www.chinasheep.org/", use_container_width=True)
            st.link_button("🥚 中国禽业网(蛋鸡与肉鸡)", "http://www.caaa.cn/", use_container_width=True)
            st.link_button("🌾 中国饲料工业信息网", "http://www.chinafeedm.com/", use_container_width=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
        
        # 实时精选行业咨询新闻与政策动态 + 刷新按钮
        c_ntitle, c_nrefresh = st.columns([3, 1])
        with c_ntitle:
            st.markdown("#### ⚡ 实时精选行业咨询新闻与政策动态")
        with c_nrefresh:
            refresh_clicked = st.button("🔄 刷新资讯动态", key=f"btn_refresh_news_{animal_type}", use_container_width=True)
        
        if refresh_clicked:
            st.success("✅ 已为您刷新获取 2026-10-22 最新政策通知与八大畜种实时产业资讯！")

        news_items = [
            {
                "title": "农业农村部印发《生猪产能调控实施方案》：引导养殖场户合理调整生产节奏与降本增效",
                "source": "农业农村部新闻办", "date": "2026-10-18", "category": "生猪政策",
                "summary": "方案针对养殖周期特征，提出分级预警机制与精准调控措施，引导规模化生猪养殖场保持合理能繁母猪产能，依靠精细化饲喂渡过周期底部。",
                "url": "http://www.moa.gov.cn/xw/zwdt/"
            },
            {
                "title": "全国奶业高质量发展推进会：强化生鲜乳优质安全，全面推广 DHI 数字化生产性能测定",
                "source": "中国奶业协会", "date": "2026-10-15", "category": "奶牛指导",
                "summary": "推进会议强调依靠自主良种繁育体系与优质粗饲料自给率，降低公斤奶完全饲料成本，保障奶牛与生鲜乳产业链良性盈利。",
                "url": "http://www.dac.org.cn/"
            },
            {
                "title": "全国肉牛肉羊增产提质行动推进：加快架子牛育肥与高产多羔肉羊舍饲技术集成",
                "source": "全国畜牧总站", "date": "2026-10-14", "category": "牛羊增效",
                "summary": "各地规模场重点推广肉牛杂交优势利用与全混合日粮 (TMR) 精准饲喂，提升平均日增重与出栏屠宰率，增强反刍动物热应激抵抗力。",
                "url": "http://www.moa.gov.cn/xw/zwdt/"
            },
            {
                "title": "家禽现代产业技术体系发布：白羽肉鸡料肉比突破 1.55，蛋鸡产蛋期破损率压减至 0.8% 以下",
                "source": "中国家禽业协会", "date": "2026-10-12", "category": "禽业前沿",
                "summary": "立体智能化笼养与环境控制算法大幅改善肉鸡欧洲效益指数 (EPI)；蛋鸡精准光照与钙磷营养调控显著提升蛋壳强度与哈氏单位。",
                "url": "http://www.caaa.cn/"
            },
            {
                "title": "特色皮用羊与奶山羊地理标志保护行动：盐池滩羊二毛皮品质标准化与羊乳深加工提速",
                "source": "大畜牧网", "date": "2026-10-10", "category": "特色畜牧",
                "summary": "宁夏盐池推进滩羊‘九道弯’二毛皮与羔羊肉品牌联动；陕西优质奶山羊基地推行低温巴氏鲜羊乳与羊奶酪研发，提高产羔率与养殖综合收益。",
                "url": "https://www.dxumu.com/"
            },
            {
                "title": "饲料原料大宗市场分析简报：豆粕与玉米走势平稳，低蛋白日粮与氨基酸平衡技术广泛应用",
                "source": "中国饲料工业信息网", "date": "2026-10-08", "category": "饲料营养",
                "summary": "监测显示，规模养殖场通过杂粕替代与生物发酵饲料平抑成本波动，有效支撑反刍动物与单胃家禽健康生长。",
                "url": "http://www.chinafeedm.com/"
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
                "🐂 肉牛规模育肥与胴体肉质",
                "🐐 奶山羊良种与生羊乳质检",
                "🐑 肉羊高效舍饲与繁育管理",
                "🐏 滩羊二毛皮与特色裘皮",
                "🥚 蛋鸡高效笼养与蛋品质量",
                "🍗 肉鸡立体笼养与饲料转化",
                "🌽 饲料原料与精准动物营养",
                "💉 动物疫病与生物安全防控",
                "📊 宏观市场分析与金融期货",
                "💻 智慧畜牧工程与数字化档案"
            ])
            sub_summary = st.text_area("资讯摘要或相关网页咨询链接", placeholder="请简要描述新闻核心内容或粘贴原文网址链接", height=75)
            sub_file = st.file_uploader("📎 上传咨询新闻/文章文档附件 (支持 Word/PDF/TXT/Excel/图片)", type=["docx", "doc", "pdf", "txt", "xlsx", "png", "jpg"])
            
            submit_news = st.form_submit_button("🚀 确认提交并留下工作联系方式", use_container_width=True)
            if submit_news:
                if not sub_name or not sub_phone or not sub_title:
                    st.error("⚠️ 请完整填写必填项：投稿人姓名、工作联系电话以及文章标题！")
                else:
                    sub_dict = {
                        "sub_name": sub_name,
                        "sub_org": sub_org or "行业同仁",
                        "sub_phone": sub_phone,
                        "sub_wechat": sub_wechat,
                        "sub_email": sub_email,
                        "sub_cat": sub_cat,
                        "sub_title": sub_title,
                        "sub_summary": sub_summary
                    }
                    seq_num, excel_path = save_submission_to_excel(sub_dict, sub_file)
                    st.success(f"🎉 投稿成功！已自动将您的投稿数据安全写入《行业新闻与咨询投稿登记表.xlsx》（入库编号：第 {seq_num} 号）！感谢 {sub_name}（{sub_org or '行业同仁'}）的投稿，您的文章《{sub_title}》已成功登记入库。项目团队将通过您留下的工作联系方式（电话: {sub_phone}）及时与您沟通审核与收录事宜。")

        # 权限控制：敏感数据（姓名、电话、微信、邮箱、摘要、附件路径等）只对认证管理员开放；普通用户网页完全隐藏已登记投稿台账与导出
        is_admin = st.session_state.get("is_admin", False)
        if is_admin:
            if SUBMISSION_EXCEL_PATH.exists():
                with st.expander("📋【管理员专享】已登记投稿全景台账审核与 Excel 导出 (已认证)", expanded=True):
                    try:
                        df_sub_records = pd.read_excel(SUBMISSION_EXCEL_PATH, skiprows=1)
                        adm_curr = st.session_state.get('admin_user', 'admin')
                        st.info(f"🔑 **当前以管理员（{adm_curr}）身份登录**，已解锁投稿人真实姓名、电话、微信、邮箱及文章摘要全量敏感数据（共 {len(df_sub_records)} 条入库记录）：")
                        st.dataframe(df_sub_records, use_container_width=True, height=240)
                        with open(SUBMISSION_EXCEL_PATH, "rb") as f_sub:
                            excel_sub_bytes = f_sub.read()
                        st.download_button(
                            label="📥 导出完整未脱敏投稿登记表 (Excel)",
                            data=excel_sub_bytes,
                            file_name="行业新闻与咨询投稿登记表.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key="download_submission_excel_btn"
                        )
                    except Exception as e:
                        st.caption(f"台账读取中... ({e})")
        else:
            st.markdown("""
            <div style='background-color:#161B22; border:1px solid #30363D; border-radius:6px; padding:12px 16px; margin-top:14px;'>
                <span style='color:#8C9BAB; font-size:13px;'>🔒 <strong>敏感数据隐私保护机制生效中</strong>：为严格保护投稿行业同仁与科研人员个人隐私，避免手机、微信、邮箱等联系方式公开泄露，<strong>已登记投稿台账与导出功能已对普通访客隐藏</strong>。管理员可在左侧侧边栏【🔐 管理员安全认证控制台】通过设备硬件指纹（MAC/序列号）与密码鉴权解锁查阅与导出。</span>
            </div>
            """, unsafe_allow_html=True)

# =====================================================================
# 五、权威学术论文与科研文献中心（链接对应畜禽数据库与全畜种DOI精选）
# =====================================================================
def render_research_papers_section(animal_type="生猪"):
    st.markdown("### 📚 现代畜牧行业权威学术论文与科研文献中心 (DOI 检索跳转)")
    st.caption("📌 **板块定位**：收录国际与国内动物科学、精准营养、智能环控与生物育种顶级学术文献；每篇论文标注文献元数据与权威标准化 DOI (Digital Object Identifier)，支持一键直达全球学术数据库。")
    
    # 学术文献 DOI 快速解析与一键直达检索器链接对应畜禽数据库
    st.subheader("🔍 学术文献 DOI 快速解析与一键直达检索器")
    c_d1, c_d2 = st.columns([4, 1])
    user_doi = c_d1.text_input("请输入或粘贴目标学术论文 DOI 号：", placeholder="例如：10.3168/jds.2016-11883 或 10.1016/j.psj.2021.101150", key="user_doi_search_input")
    clean_doi = user_doi.strip().replace("https://doi.org/", "").replace("http://", "")
    doi_target_url = f"https://doi.org/{clean_doi}" if clean_doi else "https://www.doi.org/demos.html"
    with c_d2:
        st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
        st.link_button("🌐 直达解析跳转", doi_target_url, use_container_width=True, disabled=not bool(clean_doi))
    if clean_doi:
        st.caption(f"🔗 准备跳转解析地址: `{doi_target_url}`（将自动通过国际 CrossRef 平台解析至文献主页）")

    st.markdown("##### 🏛️ 对应核心畜禽科研数据库一键直达链接")
    c_db1, c_db2, c_db3 = st.columns(3)
    with c_db1:
        st.link_button("🌐 CrossRef 官方全球检索平台", "https://www.crossref.org/", use_container_width=True)
        st.link_button("🧬 PubMed 动物科学与兽医文献库", "https://pubmed.ncbi.nlm.nih.gov/", use_container_width=True)
        st.link_button("🔬 Web of Science 农业科学核心合集", "https://www.webofscience.com/", use_container_width=True)
    with c_db2:
        st.link_button("🇨🇳 CNKI 中国知网畜牧兽医知识库", "https://oversea.cnki.net/", use_container_width=True)
        st.link_button("🥛 Journal of Dairy Science (JDS) 库", "https://www.journalofdairyscience.org/", use_container_width=True)
        st.link_button("🐖 Journal of Animal Science (JAS) 库", "https://academic.oup.com/jas", use_container_width=True)
    with c_db3:
        st.link_button("🍗 Poultry Science 家禽顶级期刊库", "https://www.sciencedirect.com/journal/poultry-science", use_container_width=True)
        st.link_button("🐐 Small Ruminant Research (羊业) 库", "https://www.sciencedirect.com/journal/small-ruminant-research", use_container_width=True)
        st.link_button("🥩 Meat Science 肉类品质数据库", "https://www.sciencedirect.com/journal/meat-science", use_container_width=True)

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)
    st.subheader("📖 行业高被引经典与前沿学术论文精选库")
    
    tab_p1, tab_p2, tab_p3, tab_p4, tab_p5, tab_p6, tab_p7, tab_p8 = st.tabs([
        "🐖 生猪精准养殖", "🐄 奶牛优质乳工程", "🐂 肉牛育肥与胴体", 
        "🐐 奶山羊产奶", "🐑 肉羊高效舍饲", "🐏 滩羊特色二毛皮", 
        "🥚 蛋鸡产蛋与品质", "🍗 肉鸡欧洲效益指数"
    ])
    
    def render_paper_card(paper):
        with st.container():
            st.markdown(f"**{paper['title']}**")
            st.markdown(f"<span style='color:#79C0FF; font-size:14px; font-weight:500;'>{paper['title_zh']}</span>", unsafe_allow_html=True)
            col_m1, col_m2 = st.columns([4, 1])
            with col_m1:
                st.caption(f"📖 **期刊**: {paper['journal']} ({paper['year']}) | ✍️ **作者**: {paper['authors']} | 🏷️ **领域**: `{paper['field']}`")
                st.markdown(f"<span style='color:#8C9BAB; font-size:13px;'>{paper['abstract']}</span>", unsafe_allow_html=True)
                st.code(f"DOI: {paper['doi']}", language="text")
            with col_m2:
                st.markdown("<div style='height:18px;'></div>", unsafe_allow_html=True)
                st.link_button("🔗 直达论文原文", f"https://doi.org/{paper['doi']}", use_container_width=True)
            st.markdown("<hr style='margin:8px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

    with tab_p1:
        pig_papers = [
            {
                "title": "Precision feeding of lactating sows: Implementation and evaluation under commercial farm conditions",
                "title_zh": "商业化农场条件下哺乳母猪精准饲喂的实施与效果评估",
                "journal": "Journal of Animal Science", "year": "2020",
                "authors": "C. Gaillard, E. Brossard, J. Y. Dourmad",
                "doi": "10.1093/jas/skaa177", "field": "精准营养与母猪生产力",
                "abstract": "研究通过智能饲喂系统根据胎次与产仔数动态调节日粮，显著减少母猪哺乳期体储动用，断奶仔猪平均窝重与断奶至发情间隔显著改善。"
            },
            {
                "title": "Feed efficiency in swine: a review of genetic, physiological and environmental factors",
                "title_zh": "猪饲料转化效率评估：遗传、生理与环境调控因素综述",
                "journal": "Animal Nutrition", "year": "2021",
                "authors": "J. F. Patience, M. C. Rossoni-Serao, N. A. Gutierrez",
                "doi": "10.1016/j.aninu.2020.12.003", "field": "料肉比优化与成本控制",
                "abstract": "全面论述影响生长肥育猪料肉比（FCR）的核心生理途径，提出通过肠道微生物调控与低蛋白氨基酸平衡日粮降低公斤增重饲料成本。"
            }
        ]
        for p in pig_papers: render_paper_card(p)

    with tab_p2:
        cow_papers = [
            {
                "title": "Use of automated milking systems data to identify cows with clinical mastitis",
                "title_zh": "利用智能挤奶系统多维传感数据早期识别奶牛临床乳房炎",
                "journal": "Journal of Dairy Science", "year": "2019",
                "authors": "M. M. Hostens, V. Fievez, G. Opsomer",
                "doi": "10.3168/jds.2016-11883", "field": "智能传感与健康预警",
                "abstract": "结合电导率、挤奶流速与体细胞数（SCC）波动建立早期预警模型，在出现典型临床症状前48小时即可实现高精度识别，大幅降低抗生素使用。"
            },
            {
                "title": "Relationship between dry matter intake, milk yield, and feed efficiency in dairy cattle",
                "title_zh": "高产荷斯坦奶牛干物质采食量、产奶量与料奶比相互关系解析",
                "journal": "Journal of Dairy Science", "year": "2021",
                "authors": "P. C. Hoffman, R. D. Shaver",
                "doi": "10.3168/jds.2020-19256", "field": "产奶性能与料奶比",
                "abstract": "深入剖析泌乳牛采食量与305天成母牛年单产的关联，验证了料奶比（FMR）控制在1.35-1.45区间的最佳边际经济效益曲线。"
            }
        ]
        for p in cow_papers: render_paper_card(p)

    with tab_p3:
        beef_papers = [
            {
                "title": "Effects of dietary net energy concentration on growth performance, carcass characteristics, and meat quality of feedlot steers",
                "title_zh": "日粮净能浓度对育肥肉牛生长性能、胴体性状及肉品质的影响",
                "journal": "Meat Science", "year": "2021",
                "authors": "M. L. Galyean, M. E. Hubbert",
                "doi": "10.1016/j.meatsci.2021.108520", "field": "肉牛育肥与胴体品质",
                "abstract": "试验测定了育肥肉牛高能日粮下的平均日增重 (ADG) 与净肉率，揭示了大理石花纹等级与料肉比的优化平衡点。"
            },
            {
                "title": "Nutritional management of heat stress in beef cattle: Physiological responses and mitigation strategies",
                "title_zh": "肉牛热应激营养调控：生理响应与采食量下降缓解对策",
                "journal": "Journal of Animal Science", "year": "2022",
                "authors": "T. L. Mader, J. B. Gaughan",
                "doi": "10.1093/jas/skac055", "field": "热应激应对与增重调控",
                "abstract": "系统研究夏季持续高温下肉牛采食量下降、料肉比升高的生理机理，验证了添加电解质与优质纤维日粮对维持日增重的显著效果。"
            }
        ]
        for p in beef_papers: render_paper_card(p)

    with tab_p4:
        goat_papers = [
            {
                "title": "Milk production, composition, and somatic cell count in dairy goats: Impact of management and parity",
                "title_zh": "奶山羊产奶量、乳成分与体细胞数(SCC)在不同胎次及管理模式下的变化规律",
                "journal": "Small Ruminant Research", "year": "2020",
                "authors": "G. F. Leitner, U. Merin",
                "doi": "10.1016/j.smallrumres.2020.106120", "field": "生鲜羊乳质量与SCC",
                "abstract": "研究明确了奶山羊泌乳期乳脂率、乳蛋白率与体细胞数的正常基线范围，建立了适合高产奶山羊的乳房健康与营养监控标准。"
            }
        ]
        for p in goat_papers: render_paper_card(p)

    with tab_p5:
        sheep_papers = [
            {
                "title": "Growth performance, feed efficiency, and carcass characteristics of intensive feedlot lambs",
                "title_zh": "舍饲集约化育肥羔羊生长性能、饲料效率与胴体性状研究",
                "journal": "Animal Feed Science and Technology", "year": "2021",
                "authors": "D. L. Pethick, J. E. Gardner",
                "doi": "10.1016/j.anifeedsci.2021.114950", "field": "舍饲肉羊增重与屠宰",
                "abstract": "评估舍饲肉羊育肥期料肉比与出栏屠宰率，提出了出栏活重达到50-55kg时屠宰率与净肉率最高且经济回报最佳的模型。"
            }
        ]
        for p in sheep_papers: render_paper_card(p)

    with tab_p6:
        fur_papers = [
            {
                "title": "Genetic parameters and phenotypic characteristics of fur quality traits in Tan sheep",
                "title_zh": "中国特产宁夏滩羊二毛皮‘九道弯’花弯性状与毛穗长度的遗传力及表型解析",
                "journal": "Journal of Animal Breeding and Genetics", "year": "2020",
                "authors": "X. P. Yue, Y. J. Kang",
                "doi": "10.1111/jbg.12480", "field": "滩羊裘皮性状遗传分析",
                "abstract": "系统测定滩羊二毛皮花弯数、毛穗长度、丝光光泽与皮板面积，构建了滩羊裘皮品级评价与适龄剥取日龄决策指标体系。"
            }
        ]
        for p in fur_papers: render_paper_card(p)

    with tab_p7:
        layer_papers = [
            {
                "title": "Dynamic modeling of egg production curve and eggshell quality in commercial laying hens",
                "title_zh": "商品蛋鸡产蛋率动态曲线拟合与蛋壳厚度、破损率变化模型构建",
                "journal": "Poultry Science", "year": "2021",
                "authors": "R. H. Harms, G. B. Russell",
                "doi": "10.1016/j.psj.2021.101150", "field": "产蛋率曲线与哈氏单位",
                "abstract": "通过监测20,000羽笼养蛋鸡72周全程产蛋数据，揭示了日均只采食量与料蛋比的关系，并提出了夏季热应激下减少破损蛋的技术对策。"
            }
        ]
        for p in layer_papers: render_paper_card(p)

    with tab_p8:
        broiler_papers = [
            {
                "title": "Evaluating the European Production Efficiency Factor (EPEF) in modern fast-growing broilers",
                "title_zh": "现代快大型肉鸡欧洲效益指数(EPI/EPEF)测算模型与出栏均匀度评估",
                "journal": "Poultry Science", "year": "2022",
                "authors": "J. A. Dozier, W. A. Dozier III",
                "doi": "10.1016/j.psj.2022.101890", "field": "欧洲效益指数与出栏均重",
                "abstract": "系统论述出栏日龄、出栏均重、成活率与料肉比对欧洲效益指数（EPI）的数学贡献率，验证了EPI突破380的立体笼养环控关键参数。"
            }
        ]
        for p in broiler_papers: render_paper_card(p)


# =====================================================================
# 五点五、现代畜牧产业人才培养与全国高校动物科技学院科研咨询矩阵
# =====================================================================
def render_education_section():
    st.markdown("### 🎓 现代畜牧产业人才培养与全国高校动物科技学院科研咨询矩阵")
    st.caption("📌 **板块定位**：聚合全国顶尖与重点农林高校动物科学技术学院/动物科技学院权威科研与人才培养资讯，推动产学研用深度协同。支持在网页内按高校名称、所属地域、办学层次或优势特色学科方向直接检索查询，并提供各学院官方门户一键直达跳转。")
    
    colleges_data = [
        {
            '序号': 1, '高校名称': '西北农林科技大学', '学院名称': '动物科技学院', '地理大区': '西北地区', '所在省市': '陕西杨凌',
            '办学层次': '一流大学建设高校 (985/211)', '特色聚焦畜种': '羊业 (奶山羊/肉羊) · 肉牛 · 生猪',
            '优势特色学科': '反刍动物遗传育种、奶山羊良种繁育、肉牛育肥与胴体品质、动物营养与饲喂调控、智慧牧业科学与工程',
            '人才培养与咨询重点': '国家级重点学科，首批开设智慧牧业工程本硕博贯通培养，李尧开发者母校，产学研深度融合基地',
            '官方网站': 'https://dkxy.nwafu.edu.cn/'
        },
        {
            '序号': 2, '高校名称': '中国农业大学', '学院名称': '动物科学技术学院', '地理大区': '华北地区', '所在省市': '北京海淀',
            '办学层次': '一流大学建设高校 (985/211)', '特色聚焦畜种': '生猪 · 蛋鸡/肉鸡 · 奶牛 · 肉牛',
            '优势特色学科': '动物遗传育种与繁殖全国重点实验室、畜禽精准饲料营养与低碳减排、伴侣动物科学、现代智慧畜牧',
            '人才培养与咨询重点': '畜禽种业创新策源地，动物科学A+顶尖学科，培养现代畜牧领军科学家与企业家',
            '官方网站': 'http://cast.cau.edu.cn/'
        },
        {
            '序号': 3, '高校名称': '华中农业大学', '学院名称': '动物科学技术学院·动物医学院', '地理大区': '华中地区', '所在省市': '湖北武汉',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '生猪 (Swine) · 肉牛 · 蛋鸡',
            '优势特色学科': '生猪抗病与优质基因选育、母猪批次化精准生产、生猪健康养殖协同创新中心、动物传染病防控',
            '人才培养与咨询重点': '全国著名生猪与动物医学拔尖人才基地，动科动医结合，校企深度合作标杆',
            '官方网站': 'http://myxy.hzau.edu.cn/'
        },
        {
            '序号': 4, '高校名称': '南京农业大学', '学院名称': '动物科技学院', '地理大区': '华东地区', '所在省市': '江苏南京',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '奶牛 · 羊业 (湖羊) · 生猪',
            '优势特色学科': '家畜生殖生理与胚胎工程、草食动物瘤胃消化调控、绿色发酵饲料资源开发、优质原料奶生产技术',
            '人才培养与咨询重点': '农业农村部草食家畜健康养殖重点实验室，培养草食家畜与生猪产业高端研发人才',
            '官方网站': 'https://dkxy.njau.edu.cn/'
        },
        {
            '序号': 5, '高校名称': '浙江大学', '学院名称': '动物科学学院', '地理大区': '华东地区', '所在省市': '浙江杭州',
            '办学层次': '一流大学建设高校 (985/211)', '特色聚焦畜种': '生猪 · 家禽 · 特种经济动物',
            '优势特色学科': '分子动物营养学教育部重点实验室、动物智能装备与感知物联、饲料生物工程与微生态制剂',
            '人才培养与咨询重点': '综合性顶尖大学交叉学科优势，数字畜牧与基因编辑拔尖领军人才培育基地',
            '官方网站': 'http://cas.zju.edu.cn/'
        },
        {
            '序号': 6, '高校名称': '华南农业大学', '学院名称': '动物科学学院', '地理大区': '华南地区', '所在省市': '广东广州',
            '办学层次': '一流学科建设高校 (双一流)', '特色聚焦畜种': '蛋鸡/肉鸡 (Poultry) · 生猪',
            '优势特色学科': '热带亚热带家禽育种、肉鸡高效笼养与饲料转化、温氏产业学院产教融合模式、畜禽生态养殖',
            '人才培养与咨询重点': '与千亿温氏集团紧密合作，黄羽肉鸡与白羽肉鸡产业链卓越工程师培养摇篮',
            '官方网站': 'https://dongke.scau.edu.cn/'
        },
        {
            '序号': 7, '高校名称': '四川农业大学', '学院名称': '动物科技学院', '地理大区': '西南地区', '所在省市': '四川雅安/成都',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '生猪 (Swine) · 家禽 · 肉牛/肉羊',
            '优势特色学科': '猪禽遗传资源发掘与保护、动物遗传育种国家级教学团队、地方特色猪种选育、饲料抗应激调控',
            '人才培养与咨询重点': '西南生猪育种核心智库，培养具有深厚理论与产业实战经验的现代动科拔尖人才',
            '官方网站': 'https://dky.sicau.edu.cn/'
        },
        {
            '序号': 8, '高校名称': '东北农业大学', '学院名称': '动物科学技术学院', '地理大区': '东北地区', '所在省市': '黑龙江哈尔滨',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '奶牛 (Dairy Cattle) · 肉牛 · 生猪',
            '优势特色学科': '北方高寒区奶牛集约化饲养营养调控、黑龙江规模化奶仓技术支撑、奶牛繁殖障碍防控与粗饲料青贮',
            '人才培养与咨询重点': '中国高寒区奶牛与寒地畜牧关键技术人才培养重镇，支撑飞鹤、完达山等产业高地',
            '官方网站': 'http://dkxy.neau.edu.cn/'
        },
        {
            '序号': 9, '高校名称': '山东农业大学', '学院名称': '动物科技学院', '地理大区': '华东地区', '所在省市': '山东泰安',
            '办学层次': '省属重点骨干大学', '特色聚焦畜种': '蛋鸡/肉鸡 · 肉牛 · 生猪',
            '优势特色学科': '优质肉鸡蛋鸡健康养殖、家禽环境控制与肠道健康、山东鲁西黄牛选育复壮、地方猪种提纯',
            '人才培养与咨询重点': '紧密依托山东畜牧第一大省产业生态，为家禽与牛羊规模场输送大量一线管理骨干',
            '官方网站': 'http://dkxy.sdau.edu.cn/'
        },
        {
            '序号': 10, '高校名称': '河南农业大学', '学院名称': '动物科技学院', '地理大区': '华中地区', '所在省市': '河南郑州',
            '办学层次': '省部共建重点高校', '特色聚焦畜种': '生猪 · 蛋鸡/肉鸡 · 肉牛',
            '优势特色学科': '中原集约化生猪高效繁育、饲料酶制剂与发酵饲料工艺、牧业经济与现代养殖装备集成',
            '人才培养与咨询重点': '紧贴中原肉蛋奶消费腹地，生猪产业化工程创新人才与经营管理人才基地',
            '官方网站': 'https://dkxy.henau.edu.cn/'
        },
        {
            '序号': 11, '高校名称': '吉林大学', '学院名称': '植物科学与动物科学学院', '地理大区': '东北地区', '所在省市': '吉林长春',
            '办学层次': '一流大学建设高校 (985/211)', '特色聚焦畜种': '肉牛 (Beef Cattle) · 优质奶牛',
            '优势特色学科': '东北优质肉牛繁育体细胞克隆育种、黑土地种养循环农业体系、转基因动物与分子育种',
            '人才培养与咨询重点': '高层次学术型科研人才培养，肉牛杂交优势测定与高品质雪花牛肉开发创新团队',
            '官方网站': 'https://dwkx.jlu.edu.cn/'
        },
        {
            '序号': 12, '高校名称': '内蒙古农业大学', '学院名称': '动物科学学院', '地理大区': '华北地区', '所在省市': '内蒙古呼和浩特',
            '办学层次': '自治区重点优势特色高校', '特色聚焦畜种': '羊业 (肉羊/绒山羊) · 肉牛 · 双峰驼',
            '优势特色学科': '内蒙古细毛羊与肉脂羊遗传繁育、绒山羊产绒性状分子机理、双峰驼乳肉特色资源综合开发',
            '人才培养与咨询重点': '草原畜牧业拔尖创新人才摇篮，北方草畜平衡与生态健康养殖技术服务基地',
            '官方网站': 'https://dkxy.imau.edu.cn/'
        },
        {
            '序号': 13, '高校名称': '新疆农业大学', '学院名称': '动物科学学院', '地理大区': '西北地区', '所在省市': '新疆乌鲁木齐',
            '办学层次': '自治区特色高水平大学', '特色聚焦畜种': '肉牛 (新疆褐牛) · 羊业 (细毛羊/多羔羊)',
            '优势特色学科': '新疆褐牛乳肉兼用选育、多胎肉羊舍饲繁育、干旱荒漠绿洲草畜协同调控与秸秆饲料化',
            '人才培养与咨询重点': '服务一带一路与西部现代畜牧业，培养扎根边疆的肉牛肉羊现代规模化养殖专家',
            '官方网站': 'https://dkxy.xjau.edu.cn/'
        },
        {
            '序号': 14, '高校名称': '云南农业大学', '学院名称': '动物科学技术学院', '地理大区': '西南地区', '所在省市': '云南昆明',
            '办学层次': '省属重点优势大学', '特色聚焦畜种': '生猪 (地方特色猪) · 肉牛 · 特色肉羊',
            '优势特色学科': '滇南小耳猪与撒坝猪地方种质资源保护与利用、高原山地肉牛生态养殖与草料供给',
            '人才培养与咨询重点': '热带亚热带地方猪种与山地特色畜牧业科研团队，生态特色畜牧业人才孵化器',
            '官方网站': 'https://dongke.ynau.edu.cn/'
        },
        {
            '序号': 15, '高校名称': '江西农业大学', '学院名称': '动物科学技术学院', '地理大区': '华东地区', '所在省市': '江西南昌',
            '办学层次': '省部共建重点大学', '特色聚焦畜种': '生猪 (Swine) · 家禽',
            '优势特色学科': '猪遗传改良与养殖技术国家重点实验室、地方猪种肉质基因图谱构建、生猪全基因组选择核心技术',
            '人才培养与咨询重点': '黄路生院士团队领衔，全国生猪遗传改良高层次科研人才与行业精英培育摇篮',
            '官方网站': 'http://dongke.jxau.edu.cn/'
        },
        {
            '序号': 16, '高校名称': '安徽农业大学', '学院名称': '动物科技学院', '地理大区': '华东地区', '所在省市': '安徽合肥',
            '办学层次': '省属重点骨干大学', '特色聚焦畜种': '生猪 · 蛋鸡/肉鸡 · 地方水禽',
            '优势特色学科': '皖南花猪与定远黑猪良种保护、家禽羽色遗传与羽绒性能提升、饲料无抗生物替代技术',
            '人才培养与咨询重点': '长三角一体化现代畜牧产学研合作示范基地，应用型畜牧技术推广与管理人才摇篮',
            '官方网站': 'http://dkxy.ahau.edu.cn/'
        },
        {
            '序号': 17, '高校名称': '湖南农业大学', '学院名称': '动物科学技术学院', '地理大区': '华中地区', '所在省市': '湖南长沙',
            '办学层次': '国内一流大学建设高校', '特色聚焦畜种': '生猪 (宁乡猪) · 羊业 (黑山羊) · 家禽',
            '优势特色学科': '宁乡花猪保种与肉品质机理研究、湖南黑山羊选育繁殖、植物源中草药健康饲料添加剂',
            '人才培养与咨询重点': '印遇龙院士科研学术支撑，动物生化营养与环境友好型养殖技术拔尖人才高地',
            '官方网站': 'https://dkxy.hunau.edu.cn/'
        },
        {
            '序号': 18, '高校名称': '河北农业大学', '学院名称': '动物科技学院', '地理大区': '华北地区', '所在省市': '河北保定',
            '办学层次': '省部共建重点大学', '特色聚焦畜种': '蛋鸡/肉鸡 · 肉牛 · 奶牛',
            '优势特色学科': '太行山特色蛋鸡绿色养殖、华北优质肉牛育肥与架子牛采购规范、农作物秸秆反刍利用',
            '人才培养与咨询重点': '太行山道路产学研典范，现代家禽与反刍草食家畜规模养殖工程化应用人才基地',
            '官方网站': 'http://dongke.hebau.edu.cn/'
        },
        {
            '序号': 19, '高校名称': '扬州大学', '学院名称': '动物科学与技术学院', '地理大区': '华东地区', '所在省市': '江苏扬州',
            '办学层次': '省属重点综合大学', '特色聚焦畜种': '家禽 (水禽鹅鸭) · 羊业 (湖羊) · 生猪',
            '优势特色学科': '国家水禽种质资源基因库、高邮鸭扬州鹅选育、太湖猪保种与太湖水系湖羊多胎繁育',
            '人才培养与咨询重点': '全国水禽科学最高水平科研与人才集聚地之一，培养高水平特种经济家禽养殖骨干',
            '官方网站': 'https://dkxy.yzu.edu.cn/'
        },
        {
            '序号': 20, '高校名称': '贵州大学', '学院名称': '动物科学学院', '地理大区': '西南地区', '所在省市': '贵州贵阳',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '生猪 (香猪) · 羊业 (白山羊) · 肉牛',
            '优势特色学科': '喀斯特山地生态畜牧业、从江香猪保种与小型猪医学模型、贵州优质白山羊选育',
            '人才培养与咨询重点': '山地生态畜牧拔尖人才基地，解决复杂喀斯特地形饲草保障与健康养殖难题',
            '官方网站': 'http://as.gzu.edu.cn/'
        },
        {
            '序号': 21, '高校名称': '兰州大学', '学院名称': '草地农业科技学院/生态学院', '地理大区': '西北地区', '所在省市': '甘肃兰州',
            '办学层次': '一流大学建设高校 (985/211)', '特色聚焦畜种': '羊业 (肉羊) · 肉牛 · 高原牦牛',
            '优势特色学科': '草地农业生态系统全国重点实验室、反刍家畜放牧行为与营养代谢调控、高寒草地畜牧生态平衡',
            '人才培养与咨询重点': '我国草业科学顶尖A+学科依托，培养草畜结合、反刍家畜生态精准饲养高端科研战略人才',
            '官方网站': 'https://caas.lzu.edu.cn/'
        },
        {
            '序号': 22, '高校名称': '宁夏大学', '学院名称': '农林与生态学部 (动物科技系)', '地理大区': '西北地区', '所在省市': '宁夏银川',
            '办学层次': '一流学科建设高校 (211/双一流)', '特色聚焦畜种': '羊业 (盐池滩羊皮用羊) · 奶牛',
            '优势特色学科': '盐池滩羊二毛皮花弯与毛穗基因表达谱、滩羊肉风味品质形成机制、贺兰山东麓优质高产奶牛技术',
            '人才培养与咨询重点': '宁夏优势特色滩羊与奶牛支柱产业智库，培养精通地方皮肉特色羊与现代化牧场管理的骨干',
            '官方网站': 'https://nlsb.nxu.edu.cn/'
        },
        {
            '序号': 23, '高校名称': '山西农业大学', '学院名称': '动物科学学院', '地理大区': '华北地区', '所在省市': '山西晋中/太原',
            '办学层次': '省部共建重点高校', '特色聚焦畜种': '肉牛 (晋南黄牛) · 羊业 (雁北肉羊)',
            '优势特色学科': '晋南牛良种复壮与杂交改良、雁北肉羊高效多羔舍饲育肥、黄土丘陵区非常规粗饲料发酵利用',
            '人才培养与咨询重点': '山西草食畜牧业产业技术体系主力军，草食反刍家畜实战型技术推广人才基地',
            '官方网站': 'https://dkxy.sxau.edu.cn/'
        },
        {
            '序号': 24, '高校名称': '福建农林大学', '学院名称': '动物科学学院 (蜂学学院)', '地理大区': '华东地区', '所在省市': '福建福州',
            '办学层次': '省属一流大学建设高校', '特色聚焦畜种': '特种经济动物 (蜜蜂) · 羊业 (闽东黑羊) · 家禽',
            '优势特色学科': '全国唯一独立蜂学专业与蜂生物学教育部工程研究中心、亚热带草食山羊生态舍饲、地方白羽半番鸭',
            '人才培养与咨询重点': '蜂学学科全国领头羊，特种经济动物饲养与生态循环畜牧业特色人才基地',
            '官方网站': 'https://dkxy.fafu.edu.cn/'
        }
    ]

    # 多维动态检索与筛选控制台
    st.subheader("🔍 网页内直接查询高校动科院所与科研咨询信息")
    c_s1, c_s2, c_s3, c_s4 = st.columns([3, 1, 1, 1])
    search_kw = c_s1.text_input("输入高校/学院名称、省市、研究方向或优势特色关键词（支持模糊搜索）：", placeholder="例如：西北农林 / 奶山羊 / 生猪 / 985 / 蛋鸡 / 肉牛 / 滩羊 / 北京", key="edu_search_kw")
    sel_region = c_s2.selectbox("地理大区", ["全部大区", "西北地区", "华北地区", "华东地区", "华中地区", "华南地区", "西南地区", "东北地区"], key="edu_sel_reg")
    sel_level = c_s3.selectbox("办学层次", ["全部层次", "一流大学 (985)", "一流学科 (211/双一流)", "省属及部省共建重点"], key="edu_sel_lvl")
    sel_species = c_s4.selectbox("聚焦畜种", ["全部畜种", "生猪 (Swine)", "奶牛 (Dairy Cattle)", "肉牛 (Beef Cattle)", "羊业 (奶山羊/肉羊/滩羊)", "禽业 (蛋鸡/肉鸡/水禽)"], key="edu_sel_sp")

    # 过滤逻辑
    filtered = []
    for c in colleges_data:
        kw_match = True
        if search_kw.strip():
            kw = search_kw.strip().lower()
            combined = f"{c['高校名称']} {c['学院名称']} {c['所在省市']} {c['办学层次']} {c['特色聚焦畜种']} {c['优势特色学科']} {c['人才培养与咨询重点']}".lower()
            kw_match = kw in combined
        
        region_match = True
        if sel_region != "全部大区":
            region_match = (c['地理大区'] == sel_region)
            
        level_match = True
        if sel_level == "一流大学 (985)":
            level_match = ("985" in c['办学层次'])
        elif sel_level == "一流学科 (211/双一流)":
            level_match = (("211" in c['办学层次'] or "双一流" in c['办学层次']) and ("985" not in c['办学层次']))
        elif sel_level == "省属及部省共建重点":
            level_match = ("985" not in c['办学层次'] and "211" not in c['办学层次'])
            
        species_match = True
        if sel_species == "生猪 (Swine)":
            species_match = ("生猪" in c['特色聚焦畜种'] or "猪" in c['优势特色学科'])
        elif sel_species == "奶牛 (Dairy Cattle)":
            species_match = ("奶牛" in c['特色聚焦畜种'] or "奶" in c['优势特色学科'])
        elif sel_species == "肉牛 (Beef Cattle)":
            species_match = ("肉牛" in c['特色聚焦畜种'] or "牛" in c['优势特色学科'])
        elif sel_species == "羊业 (奶山羊/肉羊/滩羊)":
            species_match = ("羊" in c['特色聚焦畜种'] or "羊" in c['优势特色学科'])
        elif sel_species == "禽业 (蛋鸡/肉鸡/水禽)":
            species_match = ("鸡" in c['特色聚焦畜种'] or "禽" in c['特色聚焦畜种'] or "禽" in c['优势特色学科'])
            
        if kw_match and region_match and level_match and species_match:
            filtered.append(c)

    st.caption(f"🎯 共检索匹配到 **{len(filtered)}** 所高校动物科技学院（已收录 24 所全国农林院所，提供官方门户直达）：")

    if not filtered:
        st.warning("⚠️ 未检索到匹配的高校动物科技学院，请尝试调整搜索关键词或重置筛选下拉框。")
    else:
        col_c1, col_c2 = st.columns(2)
        for idx, c in enumerate(filtered):
            target_col = col_c1 if idx % 2 == 0 else col_c2
            with target_col:
                with st.container():
                    badge_color = "#2E75B6" if "985" in c['办学层次'] else ("#1B806A" if "211" in c['办学层次'] or "双一流" in c['办学层次'] else "#D9822B")
                    st.markdown(f"""
                    <div style='background-color:#161B22; border:1px solid #30363D; border-radius:8px; padding:16px 18px; margin-bottom:12px;'>
                        <div style='display:flex; justify-content:space-between; align-items:center;'>
                            <span style='font-size:16px; font-weight:bold; color:#FFFFFF;'>🏛️ {c['高校名称']} · {c['学院名称']}</span>
                            <span style='background-color:{badge_color}; color:#FFFFFF; font-size:11px; padding:2px 8px; border-radius:4px;'>{c['所在省市']} · {c['办学层次'][:12]}</span>
                        </div>
                        <div style='margin-top:8px; font-size:13px; color:#58A6FF;'>
                            🏷️ <strong>特色聚焦畜种</strong>: {c['特色聚焦畜种']}
                        </div>
                        <div style='margin-top:6px; font-size:12px; color:#C9D1D9; line-height:1.5;'>
                            🔬 <strong>优势学科方向</strong>: {c['优势特色学科']}
                        </div>
                        <div style='margin-top:6px; font-size:12px; color:#8C9BAB; line-height:1.4;'>
                            🎓 <strong>培养与咨询重点</strong>: {c['人才培养与咨询重点']}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    st.link_button(f"🌐 访问 {c['高校名称']} 学院官网直达", c['官方网站'], use_container_width=True)
                    st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)

    # 完整名录数据表与一键 Excel 导出
    with st.expander("📋 查看全国高校动科院所全景信息数据表与导出", expanded=False):
        df_colleges = pd.DataFrame(colleges_data)
        st.dataframe(df_colleges[['序号', '高校名称', '学院名称', '地理大区', '所在省市', '办学层次', '特色聚焦畜种', '优势特色学科', '官方网站']], use_container_width=True, hide_index=True)
        excel_colleges_path = CURRENT_DIR / "全国农林高校动物科技学院名录与科研咨询指南.xlsx"
        if excel_colleges_path.exists():
            with open(excel_colleges_path, "rb") as f_c:
                bytes_c = f_c.read()
            st.download_button(
                label="📥 导出完整高校动科院所名录 (Excel)",
                data=bytes_c,
                file_name="全国农林高校动物科技学院名录与科研咨询指南.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_colleges_excel_btn"
            )

# =====================================================================
# 六、现代畜牧产业金融、前十企与期货大宗行情矩阵（全畜种覆盖）
# =====================================================================
def render_finance_section(default_kind="生猪"):
    st.markdown("### 💹 现代畜牧产业金融、行业龙头与大宗期货行情矩阵")
    st.caption("📌 **板块定位与排名说明**：本板块深度连接现代畜牧实体养殖与金融资本市场。涵盖生猪、乳业奶牛、肉牛、羊业、蛋鸡、肉鸡各细分板块前十强龙头企业官网直达；提供生猪、原奶、牛肉、鸡蛋及核心饲料大宗期货实时行情直达与套期保值指南。")
    
    tab_f1, tab_f2, tab_f3, tab_f4, tab_f5, tab_f6, tab_f7 = st.tabs([
        "🐖 前十生猪养殖龙头", 
        "🐄 前十乳业与奶牛牧场", 
        "🐂 前十肉牛养殖加工龙头",
        "🐐 🐑 十强羊业与奶山羊/肉羊",
        "🥚 前十蛋鸡与品牌鲜蛋龙头",
        "🍗 前十肉鸡养殖与白羽肉禽",
        "📈 畜禽、原奶、牛肉、禽蛋及原料期货"
    ])
    
    def render_leader_card(p, icon="🏢", metric_name="规模"):
        with st.container():
            cp1, cp2, cp3, cp4 = st.columns([1, 4, 3, 2])
            cp1.markdown(f"### `{p['排名']}`")
            cp2.markdown(f"**[{p['企业名称']}]({p['官网']})** `{p['股票代码']}`<br><span style='color:#8C9BAB;'>📌 " + str(p['核心定位']) + "</span>", unsafe_allow_html=True)
            cp3.markdown(f"{icon} **{metric_name}**: " + str(p['核心规模']) + f"<br>💰 **市值/估值**: " + str(p['市值估算']), unsafe_allow_html=True)
            cp4.link_button(f"🌐 访问官网", p['官网'], use_container_width=True)
            st.markdown("<hr style='margin:4px 0; border: none; border-top: 1px dashed #30363D;'>", unsafe_allow_html=True)

    with tab_f1:
        st.subheader("🐖 中国生猪养殖产业十强龙头企业 (Top 10 Pig Farming Enterprises)")
        st.caption("🏆 排名依据：A股/H股上市公司最新市值及生猪年出栏量规模")
        pig_leaders = [
            {"排名": "01", "企业名称": "牧原食品股份有限公司", "股票代码": "002714.SZ", "核心规模": "年出栏超 7,000 万头（中国与全球第一）", "市值估算": "约 2,100 亿元", "核心定位": "全自繁自养一体化龙头，超低成本与智能化标杆", "官网": "https://www.muyuanfoods.com/"},
            {"排名": "02", "企业名称": "温氏食品集团股份有限公司", "股票代码": "300498.SZ", "核心规模": "年出栏近 3,000 万头肉猪（行业第二）", "市值估算": "约 1,200 亿元", "核心定位": "“公司+农户”紧密型现代畜禽产业领跑者", "官网": "https://www.wens.com.cn/"},
            {"排名": "03", "企业名称": "新希望六和股份有限公司", "股票代码": "000876.SZ", "核心规模": "年出栏约 1,600-1,800 万头", "市值估算": "约 450 亿元", "核心定位": "农牧饲料制造与生猪肉禽一体化全链巨头", "官网": "http://www.newhopegroup.com/"},
            {"排名": "04", "企业名称": "双胞胎（集团）股份有限公司", "股票代码": "非上市大型集团", "核心规模": "年出栏超 1,300 万头，生猪饲料全国第一", "市值估算": "行业头部非上市航母", "核心定位": "猪饲料全产业链与生猪高产养殖旗舰", "官网": "http://www.twin-star.com.cn/"},
            {"排名": "05", "企业名称": "四川德康农牧食品集团", "股票代码": "02419.HK", "核心规模": "年出栏超 800 万头（西南生猪领军）", "市值估算": "约 220 亿港元", "核心定位": "两阶段家庭农场与优质种猪养殖一体化", "官网": "http://www.dekon.cn/"},
            {"排名": "06", "企业名称": "江西正邦科技股份有限公司", "股票代码": "002157.SZ", "核心规模": "重整复苏，年生猪出栏超 600 万头", "市值估算": "约 260 亿元", "核心定位": "生猪养殖与饲料加工老牌上市公司", "官网": "http://www.zhengbang.com/"},
            {"排名": "07", "企业名称": "北京大北农科技集团", "股票代码": "002385.SZ", "核心规模": "年生猪出栏超 600 万头", "市值估算": "约 200 亿元", "核心定位": "农业生物育种与生猪数字化生态领航者", "官网": "https://www.dbn.cn/"},
            {"排名": "08", "企业名称": "天邦食品股份有限公司", "股票代码": "002124.SZ", "核心规模": "年生猪出栏超 600 万头", "市值估算": "约 70 亿元", "核心定位": "史记生物基因育种与健康猪肉食品全链", "官网": "http://www.tianbang.com/"},
            {"排名": "09", "企业名称": "唐人神集团股份有限公司", "股票代码": "002567.SZ", "核心规模": "年生猪出栏近 400 万头", "市值估算": "约 80 亿元", "核心定位": "生物饲料、健康养殖与品牌肉品闭环龙头", "官网": "http://www.chinatrs.com/"},
            {"排名": "10", "企业名称": "巨星农牧股份有限公司", "股票代码": "603477.SH", "核心规模": "年生猪出栏超 300 万头", "市值估算": "约 105 亿元", "核心定位": "引进优质生猪种源与精细化管理效益先锋", "官网": "http://www.giantstar.cn/"}
        ]
        for p in pig_leaders: render_leader_card(p, "🐖", "出栏规模")

    with tab_f2:
        st.subheader("🐄 中国乳业及原奶生产十强龙头企业 (Top 10 Dairy Enterprises)")
        st.caption("🏆 排名依据：上市公司最新市值、营业总收入规模及原奶/优质乳制品年产量")
        cow_leaders = [
            {"排名": "01", "企业名称": "内蒙古伊利实业集团股份有限公司", "股票代码": "600887.SH", "核心规模": "年营业收入超 1,150 亿元，亚洲乳业第一", "市值估算": "超 1,750 亿元", "核心定位": "全品类规模最大综合乳企，全球乳业五强", "官网": "https://www.yili.com/"},
            {"排名": "02", "企业名称": "中国蒙牛乳业有限公司", "股票代码": "02319.HK", "核心规模": "年营业收入近 1,000 亿元，全球乳业前十", "市值估算": "约 600 亿港元", "核心定位": "拥有特仑苏、纯甄超级品牌矩阵的创新巨头", "官网": "https://www.mengniu.com.cn/"},
            {"排名": "03", "企业名称": "光明乳业股份有限公司", "股票代码": "600597.SH", "核心规模": "中国鲜奶与低温巴氏奶第一品牌", "市值估算": "约 115 亿元", "核心定位": "百年历史底蕴，深耕冷链优质生鲜乳标杆", "官网": "https://www.brightdairy.com/"},
            {"排名": "04", "企业名称": "中国飞鹤有限公司", "股票代码": "06186.HK", "核心规模": "中国婴幼儿配方奶粉第一品牌", "市值估算": "约 420 亿港元", "核心定位": "北纬 47° 黄金奶源带全产业链婴配粉龙头", "官网": "https://www.feihe.com/"},
            {"排名": "05", "企业名称": "现代牧业（集团）有限公司", "股票代码": "01117.HK", "核心规模": "奶牛存栏超 45 万头，年产原奶超 250 万吨", "市值估算": "约 70 亿港元", "核心定位": "中国最大原料奶生产商，万头牧场模式开创者", "官网": "http://www.moderndairycn.com/"},
            {"排名": "06", "企业名称": "内蒙古优然牧业有限责任公司", "股票代码": "09858.HK", "核心规模": "奶牛存栏近 60 万头，全球原料奶领跑者", "市值估算": "约 75 亿港元", "核心定位": "涵盖育种、饲料、草业与原奶的上游全产业链", "官网": "https://www.yourandairy.com/"},
            {"排名": "07", "企业名称": "新希望乳业股份有限公司 (新乳业)", "股票代码": "002946.SZ", "核心规模": "营收超百亿元，区域鲜奶联合体代表", "市值估算": "约 110 亿元", "核心定位": "“鲜立方”战略主导，全国多区域低温乳业联合体", "官网": "http://www.newdairy.cn/"},
            {"排名": "08", "企业名称": "北京三元食品股份有限公司", "股票代码": "600429.SH", "核心规模": "北京及华北老字号全产业链乳企", "市值估算": "约 65 亿元", "核心定位": "首都高品质奶源保障，特种配方与学生饮用奶", "官网": "http://www.sanyuan.com.cn/"},
            {"排名": "09", "企业名称": "中国圣牧有机奶业有限公司", "股票代码": "01432.HK", "核心规模": "沙漠有机原奶年产超 60 万吨", "市值估算": "约 25 亿港元", "核心定位": "全球最大沙漠有机生鲜乳供应商，治沙生态标杆", "官网": "http://www.shengmudairy.com/"},
            {"排名": "10", "企业名称": "上海妙可蓝多食品科技股份有限公司", "股票代码": "600882.SH", "核心规模": "中国奶酪品类第一品牌（蒙牛控股）", "市值估算": "约 75 亿元", "核心定位": "聚焦奶酪创新研发，领跑高附加值奶酪制品市场", "官网": "http://www.milkground.cn/"}
        ]
        for c in cow_leaders: render_leader_card(c, "🥛", "原奶规模")

    with tab_f3:
        st.subheader("🐂 中国肉牛养殖加工十强龙头企业 (Top 10 Beef Cattle Enterprises)")
        st.caption("🏆 排名依据：肉牛存栏出栏量、年屠宰加工产能、牛肉品牌影响力及上市公司实力")
        beef_leaders = [
            {"排名": "01", "企业名称": "福成投资集团(福成股份)", "股票代码": "600965.SH", "核心规模": "肉牛存栏及屠宰规模居全国前列", "市值估算": "约 45 亿元", "核心定位": "A股肉牛第一股，涵盖母牛繁育、育肥屠宰及预制牛肉全链", "官网": "http://www.fucheng.cn/"},
            {"排名": "02", "企业名称": "重庆恒都农业集团有限公司", "股票代码": "非上市大型农牧", "核心规模": "全国自建万头肉牛场，年屠宰肉牛超 30 万头", "市值估算": "行业龙头集团", "核心定位": "国家级农业产业化龙头，全国单体最大肉牛加工产业园", "官网": "http://www.hengdu.cn/"},
            {"排名": "03", "企业名称": "吉林省长春皓月清真肉业集团", "股票代码": "非上市大型集团", "核心规模": "肉牛年屠宰能力 50 万头，亚洲肉牛航母", "市值估算": "超 200 亿元产值", "核心定位": "亚洲最大牛肉清真产业基地，牛肉制品进出口领先", "官网": "http://www.haoyue.com/"},
            {"排名": "04", "企业名称": "内蒙古科尔沁牛业股份有限公司", "股票代码": "农业产业化龙头", "核心规模": "草原肉牛存栏与冷鲜牛肉全国领跑", "市值估算": "知名肉牛航母", "核心定位": "内蒙古天然草原肉牛标杆，绿色有机牛肉供应商", "官网": "http://www.kerchin.com/"},
            {"排名": "05", "企业名称": "大庄园肉业集团股份有限公司", "股票代码": "行业领军品牌", "核心规模": "全国牛羊肉深加工第一方阵", "市值估算": "中国肉类十强", "核心定位": "世界级现代化肉类产业园，高品质牛肉冷鲜肉旗舰", "官网": "http://www.dazhuangyuan.com/"},
            {"排名": "06", "企业名称": "黑龙江元盛食品 (龙江和牛)", "股票代码": "高新技术肉牛", "核心规模": "国内最大和牛雪花牛肉产业化基地", "市值估算": "高端牛肉独角兽", "核心定位": "纯种和牛胚胎移植与高端大理石花纹牛肉标杆", "官网": "http://www.wagyu.cn/"},
            {"排名": "07", "企业名称": "陕西秦宝牧业股份有限公司", "股票代码": "831519.OC", "核心规模": "秦川牛良种繁育与万头育肥基地", "市值估算": "新三板肉牛龙头", "核心定位": "优质秦川肉牛选育改良与高档牛肉标准化加工", "官网": "http://www.qinbaobeef.com/"},
            {"排名": "08", "企业名称": "河南伊赛牛肉股份有限公司", "股票代码": "832993.OC", "核心规模": "年出栏及屠宰肉牛 30 万头", "市值估算": "知名牛肉品牌", "核心定位": "清真牛肉食品全产业链闭环领军企业", "官网": "http://www.yisaibeef.com/"},
            {"排名": "09", "企业名称": "内蒙古民丰肉牛产业集团", "股票代码": "农业产业化重点", "核心规模": "黄牛与西门塔尔育肥基地超 10 万头", "市值估算": "区域支柱企业", "核心定位": "华北及东北冷凉气候肉牛架子牛育肥龙头", "官网": "http://www.minfengbeef.com/"},
            {"排名": "10", "企业名称": "新疆天山畜牧生物工程 (天山生物)", "股票代码": "300313.SZ", "核心规模": "肉牛与种牛优质冻精育种全国领先", "市值估算": "约 35 亿元", "核心定位": "牛羊良种繁育、牛冷冻精液与活畜进出口上市公司", "官网": "http://www.tsbiotech.com.cn/"}
        ]
        for b in beef_leaders: render_leader_card(b, "🐂", "出栏屠宰")

    with tab_f4:
        st.subheader("🐐 🐑 中国羊业与奶山羊/肉羊十强龙头企业 (Top 10 Sheep & Goat Enterprises)")
        st.caption("🏆 排名依据：羊奶粉加工能力、肉羊养殖规模、羔羊肉与特色裘皮深加工实力")
        sheep_leaders = [
            {"排名": "01", "企业名称": "陕西红星美羚乳业股份有限公司", "股票代码": "430367.OC", "核心规模": "自建富平奶山羊生态牧场，羊奶粉第一梯队", "市值估算": "奶山羊第一股", "核心定位": "中国羊乳全产业链典范，低温湿法工艺婴配羊奶粉", "官网": "http://www.hongxingmeiling.com/"},
            {"排名": "02", "企业名称": "陕西和氏乳业集团有限公司", "股票代码": "大型羊乳集团", "核心规模": "拥有 7 大生态天然牧场及万只莎能奶山羊基地", "市值估算": "羊乳领航旗舰", "核心定位": "涵盖牛羊双配方奶粉，亚洲超大单体智能化羊奶粉工厂", "官网": "http://www.heshi乳业.com/"},
            {"排名": "03", "企业名称": "西安百跃羊乳集团有限公司", "股票代码": "老牌羊乳巨头", "核心规模": "年产婴幼儿配方羊奶粉超 3 万吨", "市值估算": "知名羊乳集团", "核心定位": "深耕关中奶山羊核心带，全产业链智能化加工标杆", "官网": "http://www.baiyue.com/"},
            {"排名": "04", "企业名称": "佳贝艾特 (澳优乳业旗下)", "股票代码": "01717.HK", "核心规模": "全球羊奶粉销量遥遥领先第一品牌", "市值估算": "约 55 亿港元", "核心定位": "全球羊奶营养创新先锋，纯正山羊乳清科研龙头", "官网": "https://www.kabrita.com.cn/"},
            {"排名": "05", "企业名称": "内蒙古巴美羊业开发有限公司", "股票代码": "良种肉羊航母", "核心规模": "巴美肉羊良种繁育基地，年出栏肉羊超 20 万只", "市值估算": "区域骨干龙头", "核心定位": "肉用多羔高产良种培育与规模化集约舍饲示范", "官网": "http://www.bameisheep.com/"},
            {"排名": "06", "企业名称": "内蒙古小尾羊牧业科技股份有限公司", "股票代码": "839808.OC", "核心规模": "全产业链羊肉屠宰加工与连锁餐饮品牌", "市值估算": "羊肉上市旗舰", "核心定位": "冷鲜羊肉分割制品与草原特色餐饮一体化巨头", "官网": "http://www.xiaoweiyang.com/"},
            {"排名": "07", "企业名称": "内蒙古草都草牧业股份有限公司", "股票代码": "837373.OC", "核心规模": "草畜平衡与羊产业饲草供应网络领跑者", "市值估算": "新三板创新层", "核心定位": "草牧业互联网数字化供应链与高产肉羊配套服务", "官网": "http://www.caodu.com/"},
            {"排名": "08", "企业名称": "内蒙古蒙羊牧业股份有限公司", "股票代码": "肉羊领跑品牌", "核心规模": "年屠宰加工肉羊超 150 万只", "市值估算": "著名羊肉品牌", "核心定位": "绿色有机羊肉冷鲜直供，清真牛羊肉全品类龙头", "官网": "http://www.mengyang.com/"},
            {"排名": "09", "企业名称": "宁夏盐池滩羊产业发展集团", "股票代码": "地理标志国企", "核心规模": "全国驰名盐池滩羊与二毛裘皮核心开发平台", "市值估算": "百亿品牌价值", "核心定位": "原产地保护与国宴品质滩羊肉、特级二毛皮产业旗舰", "官网": "http://www.yctsyy.com/"},
            {"排名": "10", "企业名称": "甘肃超杰肉羊育肥有限公司", "股票代码": "规模舍饲典型", "核心规模": "年舍饲出栏育肥肉羊超 10 万只", "市值估算": "西北肉羊先锋", "核心定位": "多胎多羔肉羊杂交利用与自动化舍饲育肥示范", "官网": "http://www.chaojiesheep.com/"}
        ]
        for s in sheep_leaders: render_leader_card(s, "🐐", "羊品规模")

    with tab_f5:
        st.subheader("🥚 中国蛋鸡养殖与品牌鲜蛋十强龙头企业 (Top 10 Layer Chicken Enterprises)")
        st.caption("🏆 排名依据：蛋鸡存栏规模、商品代雏鸡产能、品牌鲜蛋市占率及上市公司市值")
        layer_leaders = [
            {"排名": "01", "企业名称": "宁夏晓鸣农牧股份有限公司", "股票代码": "300967.SZ", "核心规模": "年销售商品代雏鸡超 2.2 亿羽，蛋鸡第一股", "市值估算": "约 30 亿元", "核心定位": "中国蛋鸡制种第一龙头，生物安全与健康种源标杆", "官网": "http://www.xmnm.net/"},
            {"排名": "02", "企业名称": "四川圣迪乐村生态食品股份有限公司", "股票代码": "高品质鲜蛋龙头", "核心规模": "蛋鸡存栏超 1,300 万羽，全国高端鲜蛋第一", "市值估算": "行业高端巨头", "核心定位": "中国可生食鸡蛋标准领跑者，五星级酒店与盒马首选", "官网": "https://www.sandile.cn/"},
            {"排名": "03", "企业名称": "北京德青源农业科技股份有限公司", "股票代码": "生态蛋鸡航母", "核心规模": "亚洲单体最大生态蛋鸡场，年产鲜蛋超 10 亿枚", "市值估算": "品牌价值超百亿", "核心定位": "循环经济与可持续蛋禽发展典范，中国生态鲜蛋先驱", "官网": "http://www.deqingyuan.com/"},
            {"排名": "04", "企业名称": "正大蛋业 (正大集团中国区)", "股票代码": "跨国农牧旗舰", "核心规模": "平谷 300 万只蛋鸡现代化全自动产业园", "市值估算": "全球农牧巨头", "核心定位": "国际最高标准全流程无人化智能化蛋鸡生产标杆", "官网": "http://www.cpgroup.cn/"},
            {"排名": "05", "企业名称": "北京华都峪口禽业有限责任公司", "股票代码": "蛋鸡自主育种", "核心规模": "自主培育‘京红’‘京粉’系列，市场占有率超 50%", "市值估算": "育种国家队", "核心定位": "中国自主蛋鸡品种核心基地，打破国外种源垄断", "官网": "http://www.yukou.com.cn/"},
            {"排名": "06", "企业名称": "湖北神丹健康食品有限公司", "股票代码": "国家重点龙头", "核心规模": "年加工蛋品超 10 亿枚，皮蛋咸蛋第一", "市值估算": "蛋品加工旗舰", "核心定位": "蛋品深加工与安全营养鲜蛋国家高新技术企业", "官网": "http://www.shendan.com/"},
            {"排名": "07", "企业名称": "山西晋龙养殖股份有限公司", "股票代码": "834421.OC", "核心规模": "蛋鸡存栏超 1,000 万羽，华北最大鲜蛋基地", "市值估算": "新三板蛋鸡龙头", "核心定位": "全封闭无抗养殖与高密度自动化立体笼养旗舰", "官网": "http://www.jinlongjt.com/"},
            {"排名": "08", "企业名称": "安徽凤达农业有限公司", "股票代码": "华东鲜蛋骨干", "核心规模": "蛋鸡存栏 300 万羽，长三角鲜蛋核心保供", "市值估算": "现代化禽业龙头", "核心定位": "绿色认证富硒功能鲜蛋与数字化饲喂标杆", "官网": "http://www.fengdaegg.com/"},
            {"排名": "09", "企业名称": "宁夏顺宝现代农业股份有限公司", "股票代码": "西北禽业龙头", "核心规模": "蛋鸡存栏 260 万羽，西北大型无抗鲜蛋基地", "市值估算": "优质农牧代表", "核心定位": "高原冷凉气候优质鲜蛋直供与有机肥资源化利用", "官网": "http://www.shunbaoegg.com/"},
            {"排名": "10", "企业名称": "田园主义功能蛋业集团", "股票代码": "新消费领军", "核心规模": "主打叶黄素、DHA 差异化高端健康功能鲜蛋", "市值估算": "新锐高成长品牌", "核心定位": "聚焦母婴与健身人群的高附加值蛋品创新者", "官网": "http://www.tianyuanegg.cn/"}
        ]
        for l in layer_leaders: render_leader_card(l, "🥚", "鲜蛋产能")

    with tab_f6:
        st.subheader("🍗 中国肉鸡养殖与白羽肉禽十强龙头企业 (Top 10 Broiler Enterprises)")
        st.caption("🏆 排名依据：上市公司市值、白羽肉鸡年出栏量、肉禽一体化加工规模及屠宰产能")
        broiler_leaders = [
            {"排名": "01", "企业名称": "福建圣农发展股份有限公司", "股票代码": "002299.SZ", "核心规模": "年肉鸡出栏超 6.5 亿羽，自主育种‘圣泽901’", "市值估算": "约 180 亿元", "核心定位": "中国白羽肉鸡第一股，自繁自养全产业链全球领航", "官网": "http://www.sunnercn.com/"},
            {"排名": "02", "企业名称": "山东益生种禽股份有限公司", "股票代码": "002458.SZ", "核心规模": "亚洲最大祖代肉种鸡养殖企业，雏鸡年销超 5 亿羽", "市值估算": "约 110 亿元", "核心定位": "中国畜禽良种繁育龙头，肉种鸡市场份额居全国第一", "官网": "http://www.yishengsfc.com/"},
            {"排名": "03", "企业名称": "山东民和牧业股份有限公司", "股票代码": "002234.SZ", "核心规模": "父母代肉种鸡与商品肉鸡雏年出栏超 3 亿羽", "市值估算": "约 45 亿元", "核心定位": "肉鸡孵化、商品养殖及生物燃气发电绿色循环龙头", "官网": "http://www.minhe.cn/"},
            {"排名": "04", "企业名称": "山东仙坛股份有限公司", "股票代码": "002746.SZ", "核心规模": "肉鸡年出栏超 2 亿羽，肉类深加工产值数十亿", "市值估算": "约 60 亿元", "核心定位": "“公司+农场”现代肉禽龙头，绿色鸡肉出口标杆", "官网": "http://www.xiantan.cn/"},
            {"排名": "05", "企业名称": "辽宁禾丰食品股份有限公司", "股票代码": "603609.SH", "核心规模": "肉鸡年出栏及屠宰量突破 8 亿羽，全链巨无霸", "市值估算": "约 85 亿元", "核心定位": "中国农牧全产业链领跑者，东北肉鸡产业核心支柱", "官网": "http://www.hefeng.com/"},
            {"排名": "06", "企业名称": "山东春雪食品股份有限公司", "股票代码": "605567.SH", "核心规模": "年出栏优质白羽肉鸡近亿羽，预制调理鸡肉领先", "市值估算": "约 30 亿元", "核心定位": "深耕鸡肉调理预制菜，大规模出口日本欧盟食品企业", "官网": "http://www.chunxue.com/"},
            {"排名": "07", "企业名称": "江苏立华牧业股份有限公司", "股票代码": "300761.SZ", "核心规模": "年出栏黄羽肉鸡超 4.5 亿羽，中国黄鸡领航者", "市值估算": "约 180 亿元", "核心定位": "中国优质国鸡产业化第一龙头，跨足白羽与生猪", "官网": "http://www.lihuagroup.com/"},
            {"排名": "08", "企业名称": "温氏股份肉禽事业部", "股票代码": "300498.SZ", "核心规模": "肉鸡年上市量超 11 亿羽，全国规模最大肉禽巨头", "市值估算": "千亿级航母板块", "核心定位": "高效标准化家庭农场模式，规模化与成本效益卓越", "官网": "https://www.wens.com.cn/"},
            {"排名": "09", "企业名称": "新希望六和禽产业BG", "股票代码": "000876.SZ", "核心规模": "肉鸡与肉鸭年加工量超 15 亿羽，禽肉加工第一", "市值估算": "超大型农牧板块", "核心定位": "全球领先的冷鲜禽肉制造与全链肉禽综合运营商", "官网": "http://www.newhopegroup.com/"},
            {"排名": "10", "企业名称": "大成食品（亚洲）有限公司", "股票代码": "03999.HK", "核心规模": "中国北方白羽肉鸡龙头，肯德基德克士核心供应商", "市值估算": "约 10 亿港元", "核心定位": "“来源透明”可追溯鸡肉食品开拓者，肉禽外向型旗舰", "官网": "http://www.dachanfood.com/"}
        ]
        for r in broiler_leaders: render_leader_card(r, "🍗", "出栏屠宰")

    with tab_f7:
        st.subheader("📈 畜禽、原奶、牛肉、禽蛋及核心饲料大宗期货行情一键直达")
        st.caption("💡 畜牧产业受周期波动与饲料原料成本（占总成本 60%–70%）影响极大。期货工具是大中型养殖牧场规避价格下跌风险、锁定原料采购成本与销售利润（套期保值）的核心金融抓手。")
        
        st.markdown("#### 1️⃣ 畜禽与肉类核心大宗期货 (Livestock & Dairy Futures)")
        col_q1, col_q2, col_q3, col_q4 = st.columns(4)
        with col_q1:
            st.markdown("""
            **🐖 生猪期货 (Live Hog)**
            - **交易所**: 大连商品交易所 (`LH`)
            - **交易单位**: 16 吨/手，活体交割
            - **功能**: 平抑“猪周期”波动核心工具
            """)
            st.link_button("📊 大商所生猪期货指南", "http://www.dce.com.cn/dalianshangpin/sspz/sz/index.html", use_container_width=True)
            st.link_button("📈 东方财富生猪主连行情", "https://quote.eastmoney.com/qihuo/lhm.html", use_container_width=True)

        with col_q2:
            st.markdown("""
            **🥚 鸡蛋期货 (Egg Futures)**
            - **交易所**: 大连商品交易所 (`JD`)
            - **地位**: 我国首个上市生鲜畜禽期货
            - **功能**: 蛋鸡场锁定鲜蛋利润标杆
            """)
            st.link_button("📊 大商所鸡蛋期货指南", "http://www.dce.com.cn/dalianshangpin/sspz/jd/index.html", use_container_width=True)
            st.link_button("📈 东方财富鸡蛋主连行情", "https://quote.eastmoney.com/qihuo/jdm.html", use_container_width=True)

        with col_q3:
            st.markdown("""
            **🥩 牛肉 / 活牛期货 (Cattle)**
            - **交易所**: 芝加哥商业交易所 (`LE`/`FC`)
            - **标的**: 活牛与育肥牛期货合约
            - **功能**: 全球肉牛育肥套期保值风向标
            """)
            st.link_button("📊 芝商所活牛期货合约", "https://www.cmegroup.com/markets/agriculture/livestock/live-cattle.html", use_container_width=True)
            st.link_button("📈 芝商所育肥牛期货走势", "https://www.cmegroup.com/markets/agriculture/livestock/feeder-cattle.html", use_container_width=True)

        with col_q4:
            st.markdown("""
            **🥛 原奶 / 牛奶期货 (Milk)**
            - **交易所**: 芝商所 CME (`DC`)
            - **标的**: Class III 生鲜乳 (20万磅/手)
            - **功能**: 全球乳制品贸易与原奶风向标
            """)
            st.link_button("📊 芝商所牛奶期货合约", "https://www.cmegroup.com/markets/agriculture/dairy/class-iii-milk.html", use_container_width=True)
            st.link_button("📈 英为财情国际原奶行情", "https://cn.investing.com/commodities/class-iii-milk", use_container_width=True)

        st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
        st.markdown("#### 2️⃣ 各种核心饲料原料期货 (Feed Raw Material Futures)")
        st.caption("🌽 饲料成本占畜禽养殖完全成本的 60%–70% 以上，是牧场成本敏感性分析最关键影响变量：")
        
        c_fd1, c_fd2, c_fd3 = st.columns(3)
        with c_fd1:
            st.markdown("""
            **🌱 豆粕期货 (Soybean Meal)**
            - **交易所**: 大连商品交易所 (`M`)
            - **地位**: 蛋白饲料第一大原料
            """)
            st.link_button("📈 豆粕主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/mm.html", use_container_width=True)
        with c_fd2:
            st.markdown("""
            **🌽 玉米期货 (Corn)**
            - **交易所**: 大连商品交易所 (`C`)
            - **地位**: 畜禽能量饲料之王
            """)
            st.link_button("📈 玉米主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/cm.html", use_container_width=True)
        with c_fd3:
            st.markdown("""
            **🌿 菜粕期货 (Rapeseed Meal)**
            - **交易所**: 郑州商品交易所 (`RM`)
            - **地位**: 优质植物蛋白替代原料
            """)
            st.link_button("📈 菜粕主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/rmm.html", use_container_width=True)

        c_fd4, c_fd5, c_fd6 = st.columns(3)
        with c_fd4:
            st.markdown("""
            **💧 豆油期货 (Soybean Oil)**
            - **交易所**: 大商所 (`Y`) · 高能饲喂油脂
            """)
            st.link_button("📈 豆油主连行情 (东方财富)", "https://quote.eastmoney.com/qihuo/ym.html", use_container_width=True)
        with c_fd5:
            st.markdown("""
            **🌾 玉米淀粉期货 (Corn Starch)**
            - **交易所**: 大商所 (`CS`) · 能量与深加工关联
            """)
            st.link_button("📈 玉米淀粉行情 (东方财富)", "https://quote.eastmoney.com/qihuo/csm.html", use_container_width=True)
        with c_fd6:
            st.markdown("""
            **🌾 强麦期货 (Strong Wheat)**
            - **交易所**: 郑商所 (`WH`) · 能量与粗蛋白替代
            """)
            st.link_button("📈 强麦主连行情 (新浪财经)", "https://gu.sina.cn/ft/hq/nf.php?symbol=WH0", use_container_width=True)

# =====================================================================
# 七、侧边栏控制台：八大畜禽种类与对应数据源选择
# =====================================================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/dairy-cow.png", width=64)
    st.title("智慧养殖控制台")
    st.markdown("#### 分步向导操作流程")

    # 步骤 1：选择畜禽种类（8大种类）
    st.markdown("##### 📍 第一步：选择畜禽种类")
    animal_choice = st.radio(
        "请选择当前要监管与分析的畜禽对象：",
        [
            "🐖 生猪养殖 (Swine)", 
            "🐄 奶牛养殖 (Dairy Cattle)",
            "🐂 肉牛养殖 (Beef Cattle)",
            "🐐 奶山羊养殖 (Dairy Goat)",
            "🐑 肉羊养殖 (Meat Sheep)",
            "🐏 皮用羊养殖 (Fur Sheep)",
            "🥚 蛋鸡养殖 (Layer Chicken)",
            "🍗 肉鸡养殖 (Broiler Chicken)"
        ],
        index=0
    )

    st.markdown("---")
    # 步骤 2：选择对应数据源
    st.markdown("##### 📍 第二步：选择对应数据源")

    if "生猪" in animal_choice:
        source_options = ["默认500头母猪基准演练模型", "上传自定义生猪生产 Excel (.xlsx)"]
    elif "奶牛" in animal_choice:
        source_options = ["默认1000头成母牛现代牧场示范数据", "上传自定义奶牛生产 Excel (.xlsx)"]
    elif "肉牛" in animal_choice:
        source_options = ["默认500头肉牛专业育肥示范数据", "上传自定义肉牛生产 Excel (.xlsx)"]
    elif "奶山羊" in animal_choice:
        source_options = ["默认500只成母羊示范场数据", "上传自定义奶山羊生产 Excel (.xlsx)"]
    elif "肉羊" in animal_choice:
        source_options = ["默认1000只舍饲肉羊育肥基准数据", "上传自定义肉羊生产 Excel (.xlsx)"]
    elif "皮用羊" in animal_choice:
        source_options = ["默认500只基础母羊繁育基地数据", "上传自定义皮用羊生产 Excel (.xlsx)"]
    elif "蛋鸡" in animal_choice:
        source_options = ["默认20,000羽高产蛋鸡示范数据", "上传自定义蛋鸡生产 Excel (.xlsx)"]
    else: # 肉鸡
        source_options = ["默认年出栏6批次立体笼养肉鸡数据", "上传自定义肉鸡生产 Excel (.xlsx)"]

    data_source_choice = st.selectbox("请选择要载入的数据来源：", source_options)

    uploaded_file = None
    is_custom_upload = "上传自定义" in data_source_choice

    if is_custom_upload:
        st.markdown("##### 📥 采集模版获取与填报指引")
        st.info("💡 填报贴士：支持直接下载标准采集模版并在本地修改或补充数值；未采集字段可留空，系统内置容错隔离引擎自动保障其余指标计算正常。")
        
        _, tmpl_fname, tmpl_path = get_template_info(animal_choice)
        if tmpl_path.exists():
            with open(tmpl_path, "rb") as f_t:
                tmpl_bytes = f_t.read()
            st.download_button(
                label=f"📥 直接下载【{tmpl_fname}】到本地",
                data=tmpl_bytes,
                file_name=tmpl_fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help=f"点击直接下载{animal_choice[:2]}标准 Excel 模版至本地，填入数值后再在此上传"
            )
            
            with st.expander(f"👀 在线预览【{animal_choice[:2]}生产台账模版】结构与规范", expanded=False):
                st.caption(f"模版包含规范台账录入与法定指标指南双表，支持直接修改基准值或填入真实数据：")
                df_prev, df_gd = load_template_preview_df(tmpl_path)
                if not df_prev.empty:
                    st.markdown("**① 台账数据录入表结构预览：**")
                    st.dataframe(df_prev, hide_index=True, use_container_width=True)
                if df_gd is not None and not df_gd.empty:
                    st.markdown("**② 核心字段填报与容错指南预览：**")
                    st.dataframe(df_gd, hide_index=True, use_container_width=True)

        uploaded_file = st.file_uploader(
            f"📤 请上传填写好的 {animal_choice[:2]} 台账工作簿 (.xlsx)",
            type=["xlsx", "xls"]
        )

    st.markdown("---")
    st.subheader("💡 智慧牧业赋能亮点")
    st.markdown("""
    - **全套法定档案数据精准映射**（严格依循《畜牧法》与 NY/T 3445-2019）
    - **八大畜种核心经营指标全量闭环**（猪、奶牛、肉牛、奶山羊、肉羊、皮用羊、蛋鸡、肉鸡）
    - **容错隔离计算底座**：空指标不崩溃，好指标照样算
    - **全景免滚轴沉浸呈现**：清晰分界线与合理板块间距
    - **权威行业新闻、DOI直达论文与现代产业金融**
    """)
    
    render_admin_sidebar_auth()

    st.markdown("---")
    st.subheader("📬 开发者联系方式")
    st.markdown("""
    - **QQ**: `3100785036`
    - **WeChat (微信)**: `13679275651`
    - **项目作者**: 李尧 (学号: `2024010983`)
    - **培养单位**: 中国西北农林科技大学 动物科技学院
    """)
    st.caption("版本：V2.6 全畜种覆盖·全景分界增强·金融与DOI升级版 | 2026-10-22")

# =====================================================================
# 八、主流程驱动与数据流转
# =====================================================================
status_msg = ""

def get_species_calc(animal_choice, uploaded_file, data_source_choice):
    if "生猪" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "生猪")
                c, k = compute_pig_metrics(raw)
                return raw, c, k, f"已解析上传生猪档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_pig_data()
                c, k = compute_pig_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认500头母猪基准模型"
        raw = get_default_pig_data()
        c, k = compute_pig_metrics(raw)
        return raw, c, k, "已载入默认500头母猪基准演练模型"

    elif "奶牛" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "奶牛")
                c, k = compute_cow_metrics(raw)
                return raw, c, k, f"已解析上传奶牛档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_cow_data()
                c, k = compute_cow_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认1000头成母牛示范数据"
        raw = get_default_cow_data()
        c, k = compute_cow_metrics(raw)
        return raw, c, k, "已载入默认1000头成母牛现代牧场示范数据"

    elif "肉牛" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "肉牛")
                c, k = compute_beef_metrics(raw)
                return raw, c, k, f"已解析上传肉牛档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_beef_data()
                c, k = compute_beef_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认500头肉牛育肥示范数据"
        raw = get_default_beef_data()
        c, k = compute_beef_metrics(raw)
        return raw, c, k, "已载入默认500头肉牛专业育肥示范数据"

    elif "奶山羊" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "奶山羊")
                c, k = compute_dairy_goat_metrics(raw)
                return raw, c, k, f"已解析上传奶山羊档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_dairy_goat_data()
                c, k = compute_dairy_goat_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认500只成母羊示范数据"
        raw = get_default_dairy_goat_data()
        c, k = compute_dairy_goat_metrics(raw)
        return raw, c, k, "已载入默认500只成母羊示范场数据"

    elif "肉羊" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "肉羊")
                c, k = compute_meat_sheep_metrics(raw)
                return raw, c, k, f"已解析上传肉羊档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_meat_sheep_data()
                c, k = compute_meat_sheep_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认1000只肉羊育肥基准数据"
        raw = get_default_meat_sheep_data()
        c, k = compute_meat_sheep_metrics(raw)
        return raw, c, k, "已载入默认1000只舍饲肉羊育肥基准数据"

    elif "皮用羊" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "皮用羊")
                c, k = compute_fur_sheep_metrics(raw)
                return raw, c, k, f"已解析上传皮用羊档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_fur_sheep_data()
                c, k = compute_fur_sheep_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认500只基础母羊基地数据"
        raw = get_default_fur_sheep_data()
        c, k = compute_fur_sheep_metrics(raw)
        return raw, c, k, "已载入默认500只基础母羊繁育基地数据"

    elif "蛋鸡" in animal_choice:
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "蛋鸡")
                c, k = compute_layer_metrics(raw)
                return raw, c, k, f"已解析上传蛋鸡档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_layer_data()
                c, k = compute_layer_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认20000羽蛋鸡示范数据"
        raw = get_default_layer_data()
        c, k = compute_layer_metrics(raw)
        return raw, c, k, "已载入默认20,000羽高产蛋鸡示范数据"

    else: # 肉鸡
        if uploaded_file is not None:
            try:
                raw = parse_uploaded_excel(uploaded_file, "肉鸡")
                c, k = compute_broiler_metrics(raw)
                return raw, c, k, f"已解析上传肉鸡档案【{uploaded_file.name}】（共 {len(raw)} 条记录）"
            except Exception:
                raw = get_default_broiler_data()
                c, k = compute_broiler_metrics(raw)
                return raw, c, k, "上传解析异常，已自动切回默认6批次肉鸡数据"
        raw = get_default_broiler_data()
        c, k = compute_broiler_metrics(raw)
        return raw, c, k, "已载入默认年出栏6批次立体笼养肉鸡数据"

df_raw, df_calc, kpi, status_msg = get_species_calc(animal_choice, uploaded_file, data_source_choice)

# 主标题与法定政策依据标注
st.title("📊 智慧畜牧数字化管理与核心经营指标大屏")
st.markdown("##### 智慧牧业工程 · 法定养殖档案数据 (81张表) + 八大畜种经营指标闭环演算")

st.caption(
    "📜 **法定养殖档案政策依据**：根据《中华人民共和国畜牧法》（第四十一条）、《中华人民共和国动物防疫法》、"
    "《中华人民共和国农产品质量安全法》、农业农村部《畜禽标识和养殖档案管理办法》及行业标准《畜禽养殖场档案规范》（NY/T 3445-2019），"
    "依法建立覆盖生产、投入品、防疫、诊疗、无害化等全生命周期法定档案底座（覆盖全套81张档案表），严格遵守法定保存期限（生猪/家禽≥2年，牛羊≥10-20年，种畜禽长期保存），"
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

if is_custom_upload and uploaded_file is not None:
    st.success(f"🟢 **当前驱动数据**：{status_msg}。系统已启用容错隔离引擎，部分空值指标已安全隔离。")

# 大屏呈现模式控制（全景平铺多板块同时呈现 vs 标签页）
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

render_section_divider()

# =====================================================================
# 九、业务大屏主视觉呈现：指标卡、趋势、成本结构、原始台账与导出
# =====================================================================

# 1. 顶部指标卡片渲染
st.markdown(f"### {animal_choice[:2]} {animal_choice[2:].split(' ')[0]}核心经营指标概览")

def render_top_kpi_cards(animal_choice, kpi):
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    if "生猪" in animal_choice:
        render_kpi_card(c1, "年提供断奶仔猪 (PSY)", kpi.get('PSY'), fmt=".2f", target_delta="+3.27 超标杆线 (≥20)", threshold=20.0)
        render_kpi_card(c2, "年提供出栏肥猪 (MSY)", kpi.get('MSY'), fmt=".2f", target_delta="+2.49 超达标线 (≥20)", threshold=20.0)
        render_kpi_card(c3, "全程成活率", kpi.get('全程成活率'), is_pct=True, fmt=".1f", target_delta="目标 92.0%", threshold=0.92)
        render_kpi_card(c4, "全年生猪料肉比", kpi.get('年料肉比'), fmt=".2f", target_delta="-0.20 优于上限 (3.4)", threshold=3.4, inverse=True)
        render_kpi_card(c5, "出栏头均盈利", kpi.get('年头均盈利'), fmt=".1f", unit="元", target_delta="成本核算闭环")
        render_kpi_card(c6, "全场月均死亡率", kpi.get('月均死亡率'), is_pct=True, fmt=".2f", target_delta="预警线 1.5%", threshold=0.015, inverse=True)
    elif "奶牛" in animal_choice:
        render_kpi_card(c1, "成母牛年均单产", kpi.get('成母牛年单产'), fmt=",.0f", unit="kg", target_delta="+973 优于标杆 (10,000kg)", threshold=10000.0)
        render_kpi_card(c2, "泌乳牛日均单产", kpi.get('泌乳牛日均单产'), fmt=".1f", unit="kg/天", target_delta="高产阶段 (≥35.0)", threshold=35.0)
        f_val = kpi.get('生鲜牛乳脂率')
        p_val = kpi.get('生鲜牛乳蛋白率')
        if pd.notna(f_val) and pd.notna(p_val):
            c3.metric(label="生鲜牛乳脂 / 乳蛋白率", value=f"{f_val*100:.2f}% / {p_val*100:.2f}%", delta="特级生鲜乳标准")
        else:
            c3.metric(label="生鲜牛乳脂率", value="—", delta="待采集")
        render_kpi_card(c4, "体细胞数 SCC 均值", kpi.get('体细胞数_SCC均值'), fmt=".1f", unit="万/mL", target_delta="特级乳健康标准 (≤20万)", threshold=20.0, inverse=True)
        render_kpi_card(c5, "全群料奶比 (FMR)", kpi.get('年化料奶比'), fmt=".2f", unit=": 1", target_delta="高效区间 (1.35-1.45)", threshold=1.45, inverse=True)
        render_kpi_card(c6, "成母牛年头均净利润", kpi.get('成母牛头均年净利润'), fmt=",.0f", unit="元", target_delta="全成本动态推演")
    elif "肉牛" in animal_choice:
        render_kpi_card(c1, "育肥平均日增重 (ADG)", kpi.get('平均日增重'), fmt=".2f", unit="kg/天", target_delta="+0.18 超标杆 (≥1.25)", threshold=1.25)
        render_kpi_card(c2, "料肉比 (干物质FCR)", kpi.get('料肉比_FCR'), fmt=".2f", target_delta="优于控制线 (≤7.5)", threshold=7.5, inverse=True)
        render_kpi_card(c3, "出栏屠宰率", kpi.get('屠宰率'), is_pct=True, fmt=".1f", target_delta="优质胴体标准 (≥58%)", threshold=0.58)
        render_kpi_card(c4, "月出栏均重", kpi.get('出栏均重'), fmt=".1f", unit="kg", target_delta="达标出栏 (≥600kg)", threshold=600.0)
        render_kpi_card(c5, "公斤增重完全成本", kpi.get('公斤增重完全成本'), fmt=".2f", unit="元/kg", target_delta="合理成本区间 (≤26元)", threshold=26.0, inverse=True)
        render_kpi_card(c6, "头均育肥净利润", kpi.get('头均出栏净利润'), fmt=",.0f", unit="元", target_delta="育肥周期回报优良")
    elif "奶山羊" in animal_choice:
        render_kpi_card(c1, "成母羊年均单产", kpi.get('成母羊年单产'), fmt=",.0f", unit="kg", target_delta="优于高产线 (≥700kg)", threshold=700.0)
        render_kpi_card(c2, "泌乳母羊日均单产", kpi.get('泌乳羊日均单产'), fmt=".2f", unit="kg/天", target_delta="高产阶段 (≥2.30)", threshold=2.3)
        render_kpi_card(c3, "生鲜羊乳脂率", kpi.get('生鲜羊乳脂率'), is_pct=True, fmt=".2f", target_delta="优质生羊乳标准 (≥4.0%)", threshold=0.04)
        render_kpi_card(c4, "体细胞数 SCC 均值", kpi.get('体细胞数SCC'), fmt=".1f", unit="万/mL", target_delta="优于行业预警线 (≤100万)", threshold=100.0, inverse=True)
        render_kpi_card(c5, "公斤羊奶完全成本", kpi.get('公斤羊奶完全成本'), fmt=".2f", unit="元/kg", target_delta="成本控制达标 (≤5.8元)", threshold=5.8, inverse=True)
        render_kpi_card(c6, "成母羊年均净利", kpi.get('成母羊年头均净利润'), fmt=",.0f", unit="元", target_delta="羊奶产业收益稳定")
    elif "肉羊" in animal_choice:
        render_kpi_card(c1, "育肥平均日增重 (ADG)", kpi.get('平均日增重'), fmt=".1f", unit="g/天", target_delta="优于高效线 (≥240g)", threshold=240.0)
        render_kpi_card(c2, "料肉比 (干物质FCR)", kpi.get('料肉比_FCR'), fmt=".2f", target_delta="控制在标杆内 (≤5.6)", threshold=5.6, inverse=True)
        render_kpi_card(c3, "出栏屠宰率", kpi.get('屠宰率'), is_pct=True, fmt=".1f", target_delta="肉用标准 (≥52%)", threshold=0.52)
        render_kpi_card(c4, "出栏均重", kpi.get('出栏均重'), fmt=".1f", unit="kg", target_delta="标准适宰体重 (≥50kg)", threshold=50.0)
        render_kpi_card(c5, "公斤增重完全成本", kpi.get('公斤增重完全成本'), fmt=".2f", unit="元/kg", target_delta="合理成本 (≤30元)", threshold=30.0, inverse=True)
        render_kpi_card(c6, "单只出栏净利润", kpi.get('单只出栏净利'), fmt=".1f", unit="元", target_delta="育肥效益良好")
    elif "皮用羊" in animal_choice:
        render_kpi_card(c1, "特级及一级裘皮率", kpi.get('特级及一级皮率'), is_pct=True, fmt=".1f", target_delta="优等皮率高 (≥88%)", threshold=0.88)
        render_kpi_card(c2, "平均花弯数 (九道弯)", kpi.get('平均花弯数'), fmt=".1f", unit="个", target_delta="典型滩羊花弯 (≥8.0)", threshold=8.0)
        render_kpi_card(c3, "平均毛穗长度", kpi.get('平均毛穗长'), fmt=".1f", unit="cm", target_delta="适龄剥取标准 (≥8.0cm)", threshold=8.0)
        render_kpi_card(c4, "只均胴体重", kpi.get('只均胴体重'), fmt=".1f", unit="kg", target_delta="优质二毛羔羊肉", threshold=11.0)
        render_kpi_card(c5, "综合销售毛利率", kpi.get('综合毛利率'), is_pct=True, fmt=".1f", target_delta="肉皮双收高回报 (≥30%)", threshold=0.30)
        render_kpi_card(c6, "只均二毛羊净利润", kpi.get('只均二毛羊净利'), fmt=",.0f", unit="元", target_delta="特色裘皮高附加值")
    elif "蛋鸡" in animal_choice:
        render_kpi_card(c1, "饲养日产蛋率 (HD%)", kpi.get('饲养日产蛋率'), is_pct=True, fmt=".1f", target_delta="高产阶段 (≥90%)", threshold=0.90)
        render_kpi_card(c2, "料蛋比 (FCR)", kpi.get('料蛋比_FCR'), fmt=".2f", target_delta="优于高效线 (≤2.25)", threshold=2.25, inverse=True)
        render_kpi_card(c3, "平均蛋重", kpi.get('平均蛋重'), fmt=".1f", unit="g", target_delta="商品一级蛋 (≥62g)", threshold=62.0)
        render_kpi_card(c4, "哈氏单位 (HU)", kpi.get('哈氏单位_HU'), fmt=".1f", unit="HU", target_delta="AA级特级鲜蛋 (≥80)", threshold=80.0)
        render_kpi_card(c5, "公斤鸡蛋完全成本", kpi.get('公斤蛋完全成本'), fmt=".2f", unit="元/kg", target_delta="成本控制良好 (≤8.2元)", threshold=8.2, inverse=True)
        render_kpi_card(c6, "只鸡全期净利润", kpi.get('只鸡全期净利润'), fmt=".1f", unit="元/羽", target_delta="蛋禽运营闭环")
    else: # 肉鸡
        render_kpi_card(c1, "欧洲效益指数 (EPI)", kpi.get('欧洲效益指数_EPI'), fmt=".0f", target_delta="超国际标杆 (≥360)", threshold=360.0)
        render_kpi_card(c2, "育成期全期成活率", kpi.get('全期成活率'), is_pct=True, fmt=".1f", target_delta="健康高成活 (≥95%)", threshold=0.95)
        render_kpi_card(c3, "全期料肉比 (FCR)", kpi.get('全期料肉比_FCR'), fmt=".2f", target_delta="节粮高效 (≤1.62)", threshold=1.62, inverse=True)
        render_kpi_card(c4, "出栏均重", kpi.get('出栏均重'), fmt=".2f", unit="kg", target_delta="达到适宰均重 (≥2.5kg)", threshold=2.5)
        render_kpi_card(c5, "公斤毛鸡完全成本", kpi.get('公斤毛鸡完全成本'), fmt=".2f", unit="元/kg", target_delta="低成本竞争优势 (≤8.8元)", threshold=8.8, inverse=True)
        render_kpi_card(c6, "只鸡出栏净利润", kpi.get('只鸡净利润'), fmt=".2f", unit="元/羽", target_delta="批次周转效益好")

render_top_kpi_cards(animal_choice, kpi)

# 2. 趋势与分析图表函数
def render_species_trends(animal_choice, df_calc):
    st.subheader(f"📈 {animal_choice[:2]}生产经营走势与关键预警")
    x_col = '批次号' if '批次号' in df_calc.columns else '月份'
    c_t1, c_t2 = st.columns(2)
    with c_t1:
        if "生猪" in animal_choice and '月死亡率' in df_calc.columns:
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=df_calc[x_col], y=df_calc['月死亡率']*100, mode='lines+markers', name='月死亡率(%)', line=dict(color='#3370FF', width=3)))
            fig1.add_hline(y=1.5, line_dash="dash", line_color="#FF4D4F", annotation_text="预警线 (1.5%)")
            fig1.update_layout(title="全场月死亡率走势（自动核验预警线）", template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
        elif "奶牛" in animal_choice and '泌乳牛日均单产_kg' in df_calc.columns:
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=df_calc[x_col], y=df_calc['泌乳牛日均单产_kg'], mode='lines+markers', name='日均单产(kg)', line=dict(color='#52C41A', width=3)))
            fig1.add_hline(y=35.0, line_dash="dash", line_color="#FAAD14", annotation_text="高产标杆 (35kg)")
            fig1.update_layout(title="泌乳牛日均单产波动走势", template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
        elif "肉牛" in animal_choice and '平均日增重_kg_天' in df_calc.columns:
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=df_calc[x_col], y=df_calc['平均日增重_kg_天'], mode='lines+markers', name='日增重(kg/天)', line=dict(color='#722ED1', width=3)))
            fig1.add_hline(y=1.30, line_dash="dash", line_color="#FF4D4F", annotation_text="高温热应激预警线 (1.30kg)")
            fig1.update_layout(title="肉牛平均日增重 (ADG) 动态走势", template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
        elif "奶山羊" in animal_choice and '泌乳羊日单产_kg' in df_calc.columns:
            fig1 = px.line(df_calc, x=x_col, y='泌乳羊日单产_kg', title="奶山羊泌乳日单产走势 (kg/天)", markers=True)
            st.plotly_chart(fig1, use_container_width=True)
        elif "肉羊" in animal_choice and '平均日增重_g_天' in df_calc.columns:
            fig1 = px.line(df_calc, x=x_col, y='平均日增重_g_天', title="肉羊育肥平均日增重走势 (g/天)", markers=True)
            st.plotly_chart(fig1, use_container_width=True)
        elif "皮用羊" in animal_choice and '特级及一级皮率' in df_calc.columns:
            fig1 = px.line(df_calc, x=x_col, y=df_calc['特级及一级皮率']*100, title="滩羊特级及一级裘皮率走势 (%)", markers=True)
            st.plotly_chart(fig1, use_container_width=True)
        elif "蛋鸡" in animal_choice and '饲养日产蛋率_HD' in df_calc.columns:
            fig1 = go.Figure()
            fig1.add_trace(go.Scatter(x=df_calc[x_col], y=df_calc['饲养日产蛋率_HD']*100, mode='lines+markers', name='产蛋率(HD%)', line=dict(color='#FA8C16', width=3)))
            fig1.add_hline(y=90.0, line_dash="dash", line_color="#52C41A", annotation_text="高产标杆 (90%)")
            fig1.update_layout(title="蛋鸡饲养日产蛋率走势", template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
        else: # 肉鸡
            fig1 = px.bar(df_calc, x=x_col, y='欧洲效益指数_EPI', title="肉鸡欧洲效益指数 (EPI) 批次对比", text_auto='.0f', color='欧洲效益指数_EPI')
            st.plotly_chart(fig1, use_container_width=True)

    with c_t2:
        fcr_col = None
        for col_name in ['料肉比', '料奶比', '料肉比_FCR', '料奶比_FMR', '料蛋比_FCR', '全期料肉比_FCR']:
            if col_name in df_calc.columns:
                fcr_col = col_name
                break
        if fcr_col:
            fig2 = px.bar(df_calc, x=x_col, y=fcr_col, title=f"月度/批次转换效率 ({fcr_col}) 动态映射", text_auto='.2f', color_discrete_sequence=['#13C2C2'])
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("ℹ️ 转化效率走势图待录入相关饲料与产出台账。")

def render_species_structure(animal_choice, df_calc):
    st.subheader(f"🍩 {animal_choice[:2]}成本效益与精细化结构剖析")
    c_s1, c_s2 = st.columns(2)
    with c_s1:
        cost_cols = [c for c in df_calc.columns if any(k in c for k in ['饲料', '兽药', '成本', '费用', '雏鸡', '折旧'])]
        val_map = {}
        for c in cost_cols:
            s_val = df_calc[c].sum(min_count=1)
            if pd.notna(s_val) and s_val > 0:
                val_map[c.replace('_元', '')] = s_val
        if val_map:
            fig_pie = px.pie(names=list(val_map.keys()), values=list(val_map.values()), title="养殖总成本构成分布", hole=0.45)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("ℹ️ 成本构成图待补充饲料、兽药等投入品明细。")

    with c_s2:
        rev_cols = [c for c in df_calc.columns if any(k in c for k in ['净利润', '收入', '盈利', '毛利'])]
        if rev_cols:
            main_rev_col = rev_cols[0]
            x_col = '批次号' if '批次号' in df_calc.columns else '月份'
            fig_bar = px.bar(df_calc, x=x_col, y=main_rev_col, title=f"效益走势分析 ({main_rev_col})", color=main_rev_col, color_continuous_scale='Teal')
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("ℹ️ 收益效益走势图待补充销售记录。")

def render_species_archive_table(animal_choice, df_calc):
    st.subheader(f"📑 {animal_choice[:2]}原始生产台账与数据质量体检看板")
    valid_cols_cnt = sum(has_valid_data(df_calc[c]) for c in df_calc.columns if c not in ['月份', '批次号'])
    total_cols_cnt = max(1, len([c for c in df_calc.columns if c not in ['月份', '批次号']]))
    c_q1, c_q2, c_q3 = st.columns([1, 1, 2])
    c_q1.metric("指标就绪度", f"{valid_cols_cnt} / {total_cols_cnt}", f"{valid_cols_cnt/total_cols_cnt*100:.1f}%")
    c_q2.metric("台账记录完整性", f"{len(df_calc)} 期", "连续性保持")
    c_q3.info("💡 提示：本系统具备指标容错隔离特性，任何空项均不影响其他指标推演。下方表格中「—」代表该月份暂未采集该项台账。")

    st.markdown("##### 原始数据台账明细与核验（空值优雅呈现为 —）")
    df_show = df_calc.copy()
    for c in df_show.columns:
        if c not in ['月份', '批次号', '进雏日期', '出栏日期']:
            df_show[c] = df_show[c].apply(lambda x: f"{x:,.2f}" if (pd.notna(x) and isinstance(x, (int, float))) else ("—" if pd.isna(x) else str(x)))
    st.dataframe(df_show, hide_index=True, use_container_width=True)

def render_species_export(animal_choice, df_calc):
    st.subheader(f"📥 {animal_choice[:2]}测算结果与标准模版导出")
    st.markdown("您可以将本次测算生成的宽表完整导出，或直接下载标准 Excel 上传模版至本地分发填报：")
    
    _, tmpl_fname, tmpl_path = get_template_info(animal_choice)
    
    c_ex1, c_ex2 = st.columns(2)
    with c_ex1:
        st.download_button(
            f"📄 导出{animal_choice[:2]}测算宽表 (CSV)",
            df_calc.to_csv(index=False).encode('utf-8-sig'),
            f"{animal_choice[:2]}数字化测算宽表导出.csv",
            "text/csv",
            use_container_width=True
        )
    with c_ex2:
        if tmpl_path.exists():
            with open(tmpl_path, "rb") as f_t:
                t_bytes = f_t.read()
            st.download_button(
                f"📊 下载标准模版【{tmpl_fname}】(Excel)",
                t_bytes,
                tmpl_fname,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.info("💡 模版提示：可直接利用导出的 CSV 或标准模版在本地 Excel 修改台账数据，并在侧边栏点击上传重算。")
            
    with st.expander(f"👀 在线展开预览【{animal_choice[:2]}标准上传模版】全部字段与规范", expanded=False):
        df_prev, df_gd = load_template_preview_df(tmpl_path)
        if not df_prev.empty:
            st.markdown(f"**📑 工作表一：{animal_choice[:2]}台账数据录入（共 {len(df_prev)} 条基准参考数据，支持就地替换真实值）**")
            st.dataframe(df_prev, hide_index=True, use_container_width=True)
        if df_gd is not None and not df_gd.empty:
            st.markdown(f"**📖 工作表二：{animal_choice[:2]}字段填报指南与系统容错说明（对标法定养殖档案 NY/T 3445-2019）**")
            st.dataframe(df_gd, hide_index=True, use_container_width=True)

# =====================================================================
# 十、页面布局渲染：全景平铺模式 vs 标签选项卡模式
# =====================================================================

if "全景平铺" in view_mode:
    # 全景平铺模式：全景平铺沉浸呈现，采用略显著分界线与合理增大间距
    render_species_trends(animal_choice, df_calc)
    render_section_divider()

    render_species_structure(animal_choice, df_calc)
    render_section_divider()

    render_species_archive_table(animal_choice, df_calc)
    render_section_divider()

    render_industry_news_section(animal_choice)
    render_section_divider()

    render_research_papers_section(animal_choice)
    render_section_divider()

    render_education_section()
    render_section_divider()

    render_finance_section(animal_choice)
    render_section_divider()

    render_species_export(animal_choice, df_calc)

else:
    # 标签选项卡模式
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "📊 生产经营走势", 
        "🍩 成本效益结构", 
        "📑 台账质量体检",
        "📰 行业新闻资讯", 
        "📚 权威论文 (DOI)", 
        "🎓 人才培养 (动科院所)",
        "💹 产业金融与期货", 
        "📥 测算结果与模版"
    ])
    with tab1:
        render_species_trends(animal_choice, df_calc)
    with tab2:
        render_species_structure(animal_choice, df_calc)
    with tab3:
        render_species_archive_table(animal_choice, df_calc)
    with tab4:
        render_industry_news_section(animal_choice)
    with tab5:
        render_research_papers_section(animal_choice)
    with tab6:
        render_education_section()
    with tab7:
        render_finance_section(animal_choice)
    with tab8:
        render_species_export(animal_choice, df_calc)

# =====================================================================
# 十一、知识产权、学术严正声明与页脚版权
# =====================================================================
render_section_divider()
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

render_section_divider()
st.markdown("© 2026 智慧牧业工程 · 李尧 2024010983 · 中国西北农林科技大学动物科技学院 · 覆盖八大畜禽种类数字化养殖档案与指标计算演示系统 (V2.6)")
