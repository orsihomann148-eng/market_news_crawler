# 多国家电商新闻抓取与客户材料生成工具

这是一个面向电商平台研究与客户汇报的本地 Web 工具。它可以按国家、时间范围和平台抓取新闻，使用 AI 做指标相关性筛选、去重、翻译、情绪判断和行业趋势标记，并把最终新闻用于资讯表 Excel 和客户新闻总结生成。

项目当前同时提供两套界面：

- 普通用户模式：面向非技术使用者，保留“选择范围、抓取新闻、查看结果、添加新闻、导出材料”的主流程。
- 开发者模式：面向维护者，保留 API 配置、品牌别名、补充检索关键词、AI 提示词、国家管理、来源管理和诊断信息。

## 主要功能

- 单账号登录与首次管理员密码设置。
- 多国家配置，支持意大利、法国、西班牙、德国、日本等市场。
- 按国家、时间范围、平台品牌抓取新闻。
- 支持媒体、buyer、seller 来源以及 Google/Bing News 补充检索。
- 支持平台搜索别名、补充检索关键词和 AI 筛选提示词配置。
- AI 理解筛选：根据问卷指标判断新闻是否与 NPS/消费者感知相关。
- 初步确定性去重和最终 AI 去重。
- 标题翻译、关联说明、情绪字段、行业趋势标记。
- 新闻查看：按批次、平台、关键词、星标、勾选新闻筛选。
- 手动添加新闻：可把使用者自己查到的新闻加入某次 SQLite 抓取批次。
- AI 自动填表：输入新闻 URL 后辅助生成手动新闻字段，人工确认后保存。
- 生成资讯表：基于当前筛选或勾选新闻生成 Excel。
- 生成新闻总结：基于当前筛选或勾选新闻生成客户汇报口径的 TXT 摘要。
- SQLite 本地数据库保存抓取批次、新闻结果、手动新闻和星标。

## 安装依赖

建议使用 Python 3.9+。

```powershell
cd market_news_crawler
python -m pip install -r requirements.txt
```

主要依赖包括：

- `flask`：Web 应用。
- `requests`、`beautifulsoup4`、`lxml`：网页与 RSS 抓取解析。
- `deep-translator`：翻译兜底。
- `openpyxl`：读取问卷 Excel 与生成资讯表。
- `python-dateutil`、`tzdata`：日期与时区处理。

## 启动 Web 工具

```powershell
cd market_news_crawler
python web_app.py
```

默认访问地址：

```text
http://127.0.0.1:8000
```

首次启动时，如果还没有管理员密码，页面会要求先设置密码。之后登录即可进入普通用户模式，右上角可以切换到开发者模式。

## 零成本云端部署建议

如果目标是长期零成本、数据重启后不丢失，推荐使用 **Oracle Cloud Always Free VM**，而不是 Vercel/Render 这类 Serverless 或免费 PaaS。原因是本项目目前依赖 Flask 常驻进程、SQLite、`outputs/` 文件和后台抓取线程，需要一块稳定的持久磁盘。

### 生产环境变量

云端建议至少配置：

```bash
MARKET_NEWS_DATA_DIR=/data/market_news_crawler
SECRET_KEY=一段足够长的随机字符串
PORT=8000
```

说明：

- `MARKET_NEWS_DATA_DIR`：所有运行态数据都会写入这里，包括登录配置、AI API 配置、SQLite、outputs、任务耗时记录和可编辑来源配置。
- `SECRET_KEY`：用于 Flask 会话和本地配置加密。云端建议固定设置，避免重启或换机器后无法解密已保存配置。
- `PORT`：本地开发默认 8000；PaaS 通常会自动注入。

### 生产启动命令

项目已支持 gunicorn：

```bash
cd market_news_crawler
gunicorn --workers 1 --threads 8 --timeout 600 --bind 0.0.0.0:${PORT:-8000} web_app:app
```

必须保持 `--workers 1`。当前抓取任务状态保存在 Web 进程内存里，多 worker 会导致进度轮询和后台任务状态分裂。

### Oracle Always Free VM 部署步骤

1. 创建一台 Ubuntu Always Free VM。
2. 安装 Python、Git、Nginx 和可选的 Certbot。
3. 拉取 GitHub 仓库到 `/opt/market_news_crawler/app`。
4. 创建虚拟环境并安装依赖：

   ```bash
   cd /opt/market_news_crawler/app/market_news_crawler
   python3 -m venv .venv
   . .venv/bin/activate
   pip install -r requirements.txt
   ```

5. 创建持久目录：

   ```bash
   sudo mkdir -p /data/market_news_crawler
   sudo chown -R $USER:$USER /data/market_news_crawler
   ```

6. 用 systemd 启动服务，环境变量示例：

   ```ini
   [Service]
   WorkingDirectory=/opt/market_news_crawler/app/market_news_crawler
   Environment=MARKET_NEWS_DATA_DIR=/data/market_news_crawler
   Environment=SECRET_KEY=请替换成随机长字符串
   Environment=PORT=8000
   ExecStart=/opt/market_news_crawler/app/market_news_crawler/.venv/bin/gunicorn --workers 1 --threads 8 --timeout 600 --bind 127.0.0.1:8000 web_app:app
   Restart=always
   ```

7. Nginx 反向代理到 `127.0.0.1:8000`。
8. 首次访问云端 URL，设置管理员密码，再配置 AI API。
9. 跑一次最近 7 天小样本，确认新闻查看、手动添加、资讯表和新闻总结可用。

### 云端备份

定期备份整个数据目录即可：

```bash
tar -czf market_news_backup_$(date +%Y%m%d).tar.gz /data/market_news_crawler
```

至少应包含：

- `web_app_settings.json`
- `news_crawler.db`
- `outputs/`
- `country_configs_custom.json`
- `country_data/` 下的可编辑来源配置

## 普通用户使用流程

1. 进入普通用户模式。
2. 在页面顶部配置 AI API，或由管理员提前在开发者模式配置。
3. 在“新闻抓取”中选择国家、时间范围和品牌。
4. 点击“开始抓取”，等待进度完成。
5. 在“新闻查看”中选择一次数据批次，筛选、勾选或星标新闻。
6. 如需补充人工发现的新闻，在新闻查看底部展开“添加新闻”，输入 URL 后可使用 AI 自动填写。
7. 在“导出材料”中生成资讯表 Excel 或新闻总结 TXT。

## 开发者模式

开发者模式入口为：

```text
http://127.0.0.1:8000/developer
```

开发者模式包含更多配置能力：

- AI API URL、Key、模型配置与测试。
- 国家已有品牌、品牌搜索别名、恢复上一次、恢复系统默认。
- 补充检索关键词与 AI 优化。
- AI 筛选提示词编辑、默认提示词迁移和来源状态提示。
- 国家管理和来源管理。
- 新闻查看、星标、资讯表、新闻总结、手动新闻管理。
- 抓取阶段诊断、来源产出诊断、品牌阶段数量概览。

普通用户模式和开发者模式共用同一套后端抓取逻辑。开发者模式保存的品牌别名、补充检索关键词和 AI 提示词，会影响普通用户后续抓取。

## 重要文件

```text
market_news_crawler/
  web_app.py                 Web 入口与路由
  xlsx_source_test.py        主抓取流程入口
  country_config.py          内置国家配置
  country_configs_custom.json 自定义国家配置
  survey_filter.py           指标筛选与提示词逻辑
  dedupe.py                  初步去重和最终 AI 去重候选逻辑
  article_browser.py         新闻查看、数据源和展示字段
  briefing_table.py          资讯表生成
  news_summary.py            新闻总结生成
  db_store.py                SQLite 存储
  source_manager.py          来源管理
  templates/                 Web 页面模板
  tests/                     自动化测试
  country_data/<country>/    各国家问卷和来源配置
  生成资讯表.xlsx             资讯表模板
```

## 本地数据与 GitHub 安全说明

以下文件属于本地运行数据或敏感配置，不应上传 GitHub：

- `market_news_crawler/outputs/`：抓取输出、生成资讯表、新闻总结等结果。
- `market_news_crawler/news_crawler.db`：SQLite 本地数据库。
- `market_news_crawler/web_app_settings.json`：AI API 配置、登录 hash、页面配置。
- `market_news_crawler/job_timing_history.json`：本地任务耗时记录。
- `market_news_crawler/article_star_store.json`：旧版星标备份。
- `market_news_crawler/country_data/*/site_credentials.json`：站点凭据。
- `market_news_crawler/country_data/*/source_capability_cache.json`：来源评估缓存。
- `Git_sshkey*`、`.env`、`*.key`、`*.pem`：密钥或环境配置。
- `__pycache__/`、`*.pyc`：Python 缓存。

这些内容已在 `.gitignore` 中忽略。仓库只保留源码、模板、测试、国家配置、问卷 Excel、来源适配器配置和资讯表模板。

## 测试

上传或交付前建议运行：

```powershell
python -m py_compile market_news_crawler/web_app.py market_news_crawler/xlsx_source_test.py
python -m unittest discover -s market_news_crawler/tests -v
```

如果测试涉及本地 API 配置或历史数据，缺少本地文件时应优先检查 `.gitignore` 中的运行数据是否需要重新生成。

## GitHub 协作建议

- 不把真实新闻输出、SQLite 数据库、API Key、登录状态或 SSH key 放进仓库。
- 配置逻辑、模板、测试和国家默认配置可以提交。
- 新增功能后先运行测试，再提交。
- 如果误把敏感文件提交进历史，应重建干净历史或清理 Git 历史后再推送。
