* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：奶牛 | 对应数据集：Dairy_Cattle_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="奶牛_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Month A8
  EndCow F8.0
  LactCow F8.0
  MilkYield F8.1
  DailyYield F8.2
  FatPct F8.2
  ProteinPct F8.2
  SCC F8.1
  DMI F8.2
  FMR F8.2
  KgCost F8.2
  NetProfit F8.2.
CACHE.
EXECUTE.
DATASET NAME Dairy_Cattle_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Dairy_Cattle_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Month "生产月份 (Month)".
VARIABLE LABELS EndCow "期末成母牛存栏 (头)".
VARIABLE LABELS LactCow "在挤泌乳牛头数 (头)".
VARIABLE LABELS MilkYield "生鲜牛乳产量 (kg)".
VARIABLE LABELS DailyYield "泌乳牛日均单产 (kg/头/天)".
VARIABLE LABELS FatPct "生鲜牛乳脂率 (%)".
VARIABLE LABELS ProteinPct "生鲜牛乳蛋白率 (%)".
VARIABLE LABELS SCC "体细胞数 (万/mL)".
VARIABLE LABELS DMI "干物质采食量 (kg/天)".
VARIABLE LABELS FMR "料奶比 (Feed-Milk Ratio)".
VARIABLE LABELS KgCost "公斤奶完全成本 (元/kg)".
VARIABLE LABELS NetProfit "月度总净利润 (元)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL EndCow LactCow MilkYield DailyYield FatPct ProteinPct SCC DMI FMR KgCost NetProfit (SCALE).
VARIABLE LEVEL Month (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=EndCow LactCow MilkYield DailyYield FatPct ProteinPct
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY EndCow BY Month
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=EndCow LactCow MilkYield DailyYield FatPct
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Dairy_Cattle_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.