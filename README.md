# 智慧畜牧数字化管理与核心经营指标大屏 (Smart Livestock Demo)

本项目基于**图书情报学（元数据标引与本体架构）**与现代畜禽养殖规范深度融合，构建以法定养殖档案（75张表）为底座的生猪与奶牛双畜种经营决策大屏。

---

## 🌟 核心功能亮点

1. **多畜种分步导航**：
   - 步骤 1：选择畜禽种类（🐖 生猪养殖 / 🐄 奶牛养殖）
   - 步骤 2：自适应选择数据源（示范模型、本地指标库、全套档案表、随机万头规模场、自定义 Excel 上传）
   - 步骤 3：实时动态呈现 28 项核心指标与交互图表
2. **生猪与奶牛指标全覆盖**：
   - 生猪：PSY、MSY、料肉比、全程成活率、头均全成本与净盈利、1.5% 月死亡率红线预警
   - 奶牛：成母牛年单产（305天）、泌乳牛日均单产、生鲜牛乳脂/乳蛋白率、体细胞数 SCC 预警（20万/mL）、全群料奶比、公斤奶完全成本
3. **结构洞察扇形图**：
   - 收入结构与成本利润构成环形扇形图
   - 仔畜全生命周期存活与死淘分布扇形图
4. **报表导出**：支持一键导出重算后的全指标宽表 CSV。

---

## 🚀 部署至 Streamlit Community Cloud 步骤指南

1. **推送代码至 GitHub**：
   - 在 GitHub 创建一个新的代码仓库（如 `smart-livestock-demo`，设为 Public 或 Private 均可）。
   - 将本目录代码关联并推送到该远程仓库：
     ```bash
     git remote add origin https://github.com/<您的GitHub用户名>/smart-livestock-demo.git
     git branch -M main
     git push -u origin main
     ```

2. **登录 Streamlit Cloud 一键部署**：
   - 访问 [Streamlit Community Cloud](https://share.streamlit.io/) 并使用 GitHub 账号登录。
   - 点击右上角 **「New app」**（或 Deploy an app）。
   - 选择刚创建的仓库 `smart-livestock-demo`，Branch 选择 `main`，Main file path 填写 `app.py`。
   - 点击 **「Deploy!」**，等待 1-2 分钟即可完成自动化构建。

3. **公网访问与自定义域名**：
   - 构建完成后，系统会自动生成一个专属的公网免费安全网址（例如 `https://smart-livestock-demo.streamlit.app`）。
   - 如果您已在阿里云/腾讯云购买了域名，可在 Streamlit Cloud 应用后台的 **Settings ➔ Custom domain** 中直接绑定您的个性化独立域名（如 `demo.yourdomain.com`）。
