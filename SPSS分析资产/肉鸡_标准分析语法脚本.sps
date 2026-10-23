* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：肉鸡 | 对应数据集：Broiler_Chicken_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="肉鸡_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Batch A8
  InChicks F8.0
  OutBirds F8.0
  LiveRate F8.2
  OutWeight F8.2
  FCR F8.2
  EPI F8.0
  EvisceratedPct F8.2
  BreastMeatPct F8.2
  KgCost F8.2
  ProfitPerBird F8.2.
CACHE.
EXECUTE.
DATASET NAME Broiler_Chicken_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Broiler_Chicken_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Batch "出栏批次 (Batch)".
VARIABLE LABELS InChicks "进雏羽数 (羽)".
VARIABLE LABELS OutBirds "出栏毛鸡羽数 (羽)".
VARIABLE LABELS LiveRate "全期成活率 (%)".
VARIABLE LABELS OutWeight "出栏均重 (kg/羽)".
VARIABLE LABELS FCR "全期料肉比 (FCR)".
VARIABLE LABELS EPI "欧洲效益指数 (EPI/EPEF)".
VARIABLE LABELS EvisceratedPct "抽检全净膛率 (%)".
VARIABLE LABELS BreastMeatPct "胸肌率 (%)".
VARIABLE LABELS KgCost "公斤毛鸡完全成本 (元/kg)".
VARIABLE LABELS ProfitPerBird "只鸡出栏净利 (元/羽)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL InChicks OutBirds LiveRate OutWeight FCR EPI EvisceratedPct BreastMeatPct KgCost ProfitPerBird (SCALE).
VARIABLE LEVEL Batch (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=InChicks OutBirds LiveRate OutWeight FCR EPI
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY InChicks BY Batch
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=InChicks OutBirds LiveRate OutWeight FCR
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Broiler_Chicken_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.