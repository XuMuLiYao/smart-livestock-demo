* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：生猪 | 对应数据集：Swine_Production_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="生猪_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Month A8
  EndSow F8.0
  PigBorn F8.0
  PigOut F8.0
  PigDead F8.0
  WeanPig F8.0
  GainWeight F8.1
  FeedWeight F8.1
  FCR F8.2
  DeathRate F8.2
  NetProfit F8.2
  ProfitPerHead F8.2.
CACHE.
EXECUTE.
DATASET NAME Swine_Production_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Swine_Production_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Month "生产月份 (Month)".
VARIABLE LABELS EndSow "期末能繁母猪存栏 (头)".
VARIABLE LABELS PigBorn "出生活仔数 (头)".
VARIABLE LABELS PigOut "出栏肉猪头数 (头)".
VARIABLE LABELS PigDead "全群死亡头数 (头)".
VARIABLE LABELS WeanPig "断奶仔猪数 (头)".
VARIABLE LABELS GainWeight "全群总增重 (kg)".
VARIABLE LABELS FeedWeight "全期耗料量 (kg)".
VARIABLE LABELS FCR "料肉比 (Feed Conversion Ratio)".
VARIABLE LABELS DeathRate "月度死亡率 (Death Rate %)".
VARIABLE LABELS NetProfit "月度净利润 (元)".
VARIABLE LABELS ProfitPerHead "头均盈利水平 (元/头)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL EndSow PigBorn PigOut PigDead WeanPig GainWeight FeedWeight FCR DeathRate NetProfit ProfitPerHead (SCALE).
VARIABLE LEVEL Month (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=EndSow PigBorn PigOut PigDead WeanPig GainWeight
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY EndSow BY Month
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=EndSow PigBorn PigOut PigDead WeanPig
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Swine_Production_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.