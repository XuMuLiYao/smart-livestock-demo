* Encoding: UTF-8.
* ====================================================================.
* IBM SPSS Statistics 语法批处理脚本.
* 项目：智慧牧业数字化工程 · 核心指标深度统计分析.
* 畜种：皮用羊 | 对应数据集：Fur_Sheep_SPSS_Dataset.
* 生成时间：2026-10-23 15:12:39.
* ====================================================================.

* 1. 自动定义与导入外部数据源.
GET DATA  /TYPE=TXT
  /FILE="皮用羊_生产台账标准数据.csv"
  /ENCODING='UTF-8'
  /DELCASE=LINE
  /DELIMITERS=","
  /QUALIFIER='"'
  /ARRANGEMENT=DELIMITED
  /FIRSTCASE=2
  /IMPORTCASE=ALL
  /VARIABLES=
  Month A8
  EweCount F8.0
  SlaughterLamb F8.0
  TopSkinRate F8.2
  AvgCurls F8.1
  HairLength F8.1
  SkinArea F8.1
  CarcassWeight F8.1
  ProfitPerSheep F8.2.
CACHE.
EXECUTE.
DATASET NAME Fur_Sheep_SPSS_Dataset WINDOW=FRONT.
DATASET ACTIVATE Fur_Sheep_SPSS_Dataset.

* 2. 赋予变量中文与学术标签定义 (VARIABLE LABELS).
VARIABLE LABELS Month "生产月份 (Month)".
VARIABLE LABELS EweCount "基础母羊存栏 (只)".
VARIABLE LABELS SlaughterLamb "适龄剥皮二毛羊 (只)".
VARIABLE LABELS TopSkinRate "特级及一级皮率 (%)".
VARIABLE LABELS AvgCurls "平均花弯数 (个/九道弯)".
VARIABLE LABELS HairLength "毛穗长度 (cm)".
VARIABLE LABELS SkinArea "抽检皮板面积 (dm²)".
VARIABLE LABELS CarcassWeight "二毛羊胴体均重 (kg)".
VARIABLE LABELS ProfitPerSheep "只均综合净利 (元/只)".

* 3. 设置变量测量尺度 (VARIABLE LEVEL: SCALE / NOMINAL / ORDINAL).
VARIABLE LEVEL EweCount SlaughterLamb TopSkinRate AvgCurls HairLength SkinArea CarcassWeight ProfitPerSheep (SCALE).
VARIABLE LEVEL Month (NOMINAL).

* 4. 自动生成基础描述性统计与正态性检验指令 (DESCRIPTIVES & EXAMINE).
DESCRIPTIVES VARIABLES=EweCount SlaughterLamb TopSkinRate AvgCurls HairLength SkinArea
  /STATISTICS=MEAN STDDEV MIN MAX KURTOSIS SKEWNESS.

* 5. 自动配置单因素方差分析与多重比较检验指令 (ONEWAY ANOVA).
ONEWAY EweCount BY Month
  /STATISTICS DESCRIPTIVES HOMOGENEITY
  /POSTHOC=TUKEY DUNCAN ALPHA(0.05).

* 6. 自动配置双变量相关性分析指令 (CORRELATIONS).
CORRELATIONS
  /VARIABLES=EweCount SlaughterLamb TopSkinRate AvgCurls HairLength
  /PRINT=TWOTAIL NOSIG
  /MISSING=PAIRWISE.

* 7. 保存为 IBM SPSS 官方原生数据集 (.sav).
SAVE OUTFILE="Fur_Sheep_SPSS_Dataset.sav"
  /COMPRESSED.
EXECUTE.
* === SPSS 语法批处理脚本生成完毕 ===.