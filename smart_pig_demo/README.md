# 智慧畜牧数字化管理与核心经营指标大屏 (Smart Livestock Demo)

本项目基于**现代畜牧标准化养殖**与数据闭环演算深度融合，构建以法定养殖档案为基础的生猪与奶牛双畜种经营决策大屏。

---

## 🌟 核心功能亮点

1. **多畜种分步导航**：
   - 步骤 1：选择畜禽种类（🐖 生猪养殖 / 🐄 奶牛养殖）
   - 步骤 2：选择数据源（默认 500 头母猪基准演练模型、默认 1000 头成母牛现代牧场示范数据、自定义 Excel 上传）
   - 步骤 3：实时动态呈现 28 项核心经营指标与交互图表
2. **生猪与奶牛核心指标全覆盖与容错隔离**：
   - 生猪：PSY、MSY、料肉比、全程成活率、头均全成本与净盈利、1.5% 月死亡率红线预警
   - 奶牛：成母牛年单产（305天）、泌乳牛日均单产、生鲜牛乳脂/乳蛋白率、体细胞数 SCC 预警（20万/mL）、全群料奶比、公斤奶完全成本
3. **💹 现代畜牧产业金融与大宗期货行情矩阵**：
   - **中国前十生猪养殖龙头**（排名依据上市公司最新市值及生猪年出栏量）：牧原、温氏、新希望、双胞胎、德康、正邦、大北农、天邦、唐人神、巨星等企业官网一键直达；
   - **中国前十乳业与奶牛牧场龙头**（排名依据上市公司最新市值、营业收入及原奶/乳制品年产量）：伊利、蒙牛、光明、飞鹤、现代牧业、优然牧业、新乳业、三元、圣牧、妙可蓝多等企业官网一键直达；
   - **畜禽、原奶、牛肉及核心饲料原料期货行情**：生猪期货（大商所 LH）、牛奶/原奶期货（CME Class III）、牛肉/活牛期货（CME Live/Feeder Cattle）、豆粕期货（大商所 M）、玉米期货（大商所 C）、菜粕期货（郑商所 RM）、豆油期货（大商所 Y）、玉米淀粉期货（大商所 CS）、强麦期货（郑商所 WH）等合约与实时行情直达。
4. **结构洞察扇形图与质量体检**：
   - 收入结构与成本利润构成环形扇形图；
   - 仔畜全生命周期存活与死淘分布扇形图；
   - 台账健康度与指标就绪度实时体检。
5. **报表与模版导出**：支持一键导出重算后的全指标宽表 CSV 与标准采集模版 XLSX。

---

## ⚖️ 项目独立进行与知识产权严正声明 (Project Declaration)

**中文声明**：  
该项目由**中国西北农林科技大学动物科技学院李尧（学号：2024010983）**于 2026 年 10 月独立进行，仅供交流学习，不得擅自进行论文发表及商用。任何未经许可的学术论文撰写发表、专利申请、商业软件封装或盈利性运营行为均属侵权。

**English Statement**:  
*This project was independently developed and conducted in October 2026 by Li Yao (Student ID: 2024010983) from the College of Animal Science and Technology, Northwest A&F University. It is strictly intended for academic communication and educational study only. Unauthorized academic publication, thesis submission, software redistribution, or commercial use is strictly prohibited.*

**📬 开发者联系方式 (Developer Contact)**:  
- **QQ**: `3100785036`  
- **WeChat (微信)**: `13679275651`  
- **机构**: 中国西北农林科技大学 动物科技学院 (College of Animal Science and Technology, Northwest A&F University)

---

## 🚀 部署至 Streamlit Community Cloud 步骤指南

1. **推送代码至 GitHub**：
   - 将本目录代码关联并推送到远程仓库：
     ```bash
     git remote add origin https://github.com/<您的GitHub用户名>/smart-livestock-demo.git
     git branch -M main
     git push -u origin main
     ```

2. **登录 Streamlit Cloud 一键部署**：
   - 访问 [Streamlit Community Cloud](https://share.streamlit.io/) 并使用 GitHub 账号登录。
   - 选择刚创建的仓库，Main file path 填写 `app.py`，点击 **「Deploy!」**。

3. **公网访问与自定义域名**：
   - 支持绑定个性化独立域名（如 `smartpastoral.com`），并通过配套的 `index.html` 实现无缝全屏嵌入。

---

© 2026 智慧牧业工程 · 李尧 2024010983 · 中国西北农林科技大学动物科技学院
