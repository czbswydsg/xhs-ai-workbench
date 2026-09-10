# AI 小红书内容增长工作台

一个面向小红书（Xiaohongshu / RED）内容运营场景的**数据采集 → 爆款分析 → AI 内容生成**一体化 Web 应用。基于 Streamlit 构建，支持多访客隔离登录、真实数据采集、Excel 报告导出，可直接部署到云端作为 SaaS 式工具使用。

> 🚀 **在线体验**：部署后填写你的 Streamlit Cloud 地址，例如 `https://xxx.streamlit.app`

## 功能一览

| 模块 | 能力 |
|---|---|
| 小红书采集 | Playwright 自动化采集竞品笔记（关键词筛选、点赞/收藏数、封面与高清原图下载），扫码登录，多访客登录态隔离，实时采集进度 |
| 帖子数据 | 当前真实帖子列表、Excel 导出、生成爆款分析表 |
| 爆款分析 | 爆款卡片 + 详情面板 + 规则提炼（`analyzer.py` 计算热度分、归纳爆款规律） |
| 内容生成 | 产品 / 用户 / 场景 → 小红书内容提示词；基于爆款批量生成笔记文案 |
| 历史记录 | 任务卡片、已生成提示词沉淀复用 |
| 设置 | Excel 数据导入、数据管理、品牌画像配置（`brand_profile.json`） |

**核心链路**：采集 → 数据 → 爆款分析 → 内容生成 → 发布提示词

## 技术栈

- **Python / Streamlit** — 全栈 Web 应用，7 页面产品化结构
- **Playwright** — 浏览器自动化采集与扫码登录（`xhs_scraper.py`）
- **pandas / openpyxl** — 数据处理与 Excel 报告生成（`build_report.py`）
- **Docker / docker-compose / nginx** — 云服务器一键部署（见 `deploy/`）

## 项目结构

```
.
├── app.py                  # Streamlit 主应用（7 个页面）
├── xhs_scraper.py          # Playwright 采集核心
├── analyzer.py             # 爆款数据分析
├── content_studio.py       # 内容创作模块
├── generate_content.py     # 基于爆款的笔记文案生成
├── generate_prompt.py      # 内容提示词构建
├── build_report.py         # Excel 报告生成
├── brand_profile.json      # 品牌画像配置
├── ai_growth_workbench.html # 单文件离线版界面
└── deploy/                 # 云服务器部署包（Docker + nginx + 裸机脚本）
```

## 本地运行

```bash
pip install -r requirements.txt
python -m playwright install chromium
streamlit run app.py   # 打开 http://localhost:8501
```

## 云端部署

### 方式一：Streamlit Community Cloud（免费，推荐演示用）

1. 把本仓库 push 到 GitHub（公开仓库）
2. 打开 <https://share.streamlit.io> → Sign in with GitHub
3. **New app** → 选择仓库 / 分支 `main` / 入口文件 `app.py` → Deploy
4. 等待数分钟后访问 `https://<仓库名>.streamlit.app`

> 应用已内置 Chromium 自动补装逻辑，云端无需手动配置浏览器。

### 方式二：云服务器 + Docker（生产级，支持域名与 HTTPS）

按 `deploy/README.md` 操作：腾讯云香港轻量服务器（2C4G，免备案）→ 防火墙放行 8501 → `docker compose up -d --build`。

## 演示提示

- **内置示例数据**：仓库根目录的 `xhs_notes.json` 为演示用构造样例（15 条，字段结构与真实采集完全一致，每条带 `"demo": true` 标记）。应用启动时会自动加载，方便面试官/访客随时完整走通「帖子数据 → 爆款分析 → 内容生成」全流程，无需先采集。
- 云端采集受小红书风控与容器内存限制可能不稳定（机房 IP / 约 1GB 内存），点击「开始采集」能采到就显示真实数据；采不到时示例数据兜底，演示不受影响。
- 本仓库不含任何登录态（`cookies.json` / `sessions/`）与真实采集数据，首次运行自动生成，已通过 `.gitignore` 排除。

## 说明

仅供个人学习与内容运营研究使用，请遵守小红书平台规则，控制采集频率。
