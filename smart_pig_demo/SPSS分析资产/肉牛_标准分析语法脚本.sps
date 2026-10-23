* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：肉牛 | 对应数据集：Beef_Cattle_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="肉牛_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Month A8
  InitCattle F8.0
  EndCattle F8.0
  OutCattle F8.0
  MonthGain F8.1
  ADG F8.2
  FCR F8.2
  OutWeight F8.1
  DressingPct F8.2
  MeatPct F8.2
  KgCost F8.2
  ProfitPerHead F8.2.
CACHE.
EXECUTE.
DATASET NAME Beef_Cattle_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Beef_Cattle_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Month "育肥月份 (Month)".
VARIABLE LABELS InitCattle "期初在栏肉牛 (头)".
VARIABLE LABELS EndCattle "期末在栏肉牛 (头)".
VARIABLE LABELS OutCattle "出栏肥牛头数 (头)".
VARIABLE LABELS MonthGain "全群月增重 (kg)".
VARIABLE LABELS ADG "平均日增重 (kg/天)".
VARIABLE LABELS FCR "料肉比 (Feed Conversion Ratio)".
VARIABLE LABELS OutWeight "出栏均重 (kg/头)".
VARIABLE LABELS DressingPct "抽检屠宰率 (%)".
VARIABLE LABELS MeatPct "抽检净肉率 (%)".
VARIABLE LABELS KgCost "公斤增重完全成本 (元/kg)".
VARIABLE LABELS ProfitPerHead "头均出栏净利润 (元/头)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL InitCattle EndCattle OutCattle MonthGain ADG FCR OutWeight DressingPct MeatPct KgCost ProfitPerHead (SCALE).
VARIABLE LEVEL Month (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=InitCattle EndCattle OutCattle MonthGain ADG FCR
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY InitCattle BY Month
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=InitCattle EndCattle OutCattle MonthGain ADG
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Beef_Cattle_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.