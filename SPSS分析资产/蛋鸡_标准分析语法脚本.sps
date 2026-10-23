* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：蛋鸡 | 对应数据集：Layer_Chicken_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="蛋鸡_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Month A8
  EndHens F8.0
  LayRateHD F8.2
  EggWeight F8.1
  FCR F8.2
  BrokenRate F8.2
  HaughUnit F8.1
  ShellStrength F8.2
  KgCost F8.2
  ProfitPerHen F8.2.
CACHE.
EXECUTE.
DATASET NAME Layer_Chicken_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Layer_Chicken_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Month "饲养月份 (Month)".
VARIABLE LABELS EndHens "在舍产蛋鸡存栏 (羽)".
VARIABLE LABELS LayRateHD "饲养日产蛋率 (HD %)".
VARIABLE LABELS EggWeight "平均蛋重 (g)".
VARIABLE LABELS FCR "料蛋比 (Feed-Egg Ratio)".
VARIABLE LABELS BrokenRate "破损蛋率 (%)".
VARIABLE LABELS HaughUnit "哈氏单位 (HU)".
VARIABLE LABELS ShellStrength "蛋壳强度 (kg/cm²)".
VARIABLE LABELS KgCost "公斤鸡蛋完全成本 (元/kg)".
VARIABLE LABELS ProfitPerHen "只鸡月净利润 (元/羽)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL EndHens LayRateHD EggWeight FCR BrokenRate HaughUnit ShellStrength KgCost ProfitPerHen (SCALE).
VARIABLE LEVEL Month (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=EndHens LayRateHD EggWeight FCR BrokenRate HaughUnit
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY EndHens BY Month
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=EndHens LayRateHD EggWeight FCR BrokenRate
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Layer_Chicken_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.