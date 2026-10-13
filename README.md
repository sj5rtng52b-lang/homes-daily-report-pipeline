# hema-homes-daily-report-pipeline

> **Hello, DWS** 竞赛参赛项目 — 基于 DWS CLI 的可复用 Skill / 工作流
> Agent 工具：千问办公（QwenWork）| DWS CLI 版本：v1.0.63

盒马门店每日营收对比表（MTD + 单日双份）自动导出全链路：从 Quark 报表数据接口取数 → openpyxl 自动填 Excel 模板 → Chrome headless 渲染长图 → 钉钉群投递，已稳定运行 30+ 天。

## 架构

```
fetch_homes.js ──▶ homes_server.py ──▶ fill_template.py ──▶ render_report_image.py
  (取数)           (接收落盘)           (填表+checks)          (渲染长图)
                                            │
                                     budget_lib.py (预算读取)
                                            │
                                     hema_login.py (登录态获取)
```

## 核心源码

| 文件 | 语言 | 功能 |
|------|------|------|
| `fetch_homes.js` / `fetch_homes_mtd.js` / `fetch_homes_1d.js` | JavaScript | 浏览器重放 Quark WidgetAction 接口取 JSON |
| `homes_server.py` | Python | 本地 HTTP 服务，接收 JS 取数结果落盘 |
| `fill_template.py` | Python | openpyxl 自动填 Excel 模板，含 6 项内置 checks |
| `render_report_image.py` | Python | HTML 模板 → Chrome headless 截图 → PIL 裁切 → PNG |
| `xlsx_to_image.py` | Python | Excel 转 PNG 辅助工具 |
| `budget_lib.py` | Python | 预算文件读取库 |
| `hema_login.py` | Python | Quark 报表登录态自动获取 |
| `RUNBOOK.md` | Markdown | 完整运维手册（9 步执行流程 + 失败处理 + 复用指南） |

## 9 步执行流程

1. 检查幂等标记（已发送则跳过）
2. 执行 fetch_homes.js 取数（reportId + QUARK_PARAMS 令牌）
3. 检查数据就绪（rows>0 且非全空数组）
4. 执行 fill_template.py 填表（MTD + 单日双份）
5. 检查 6 项内置 checks
6. 执行 render_report_image.py 渲染长图
7. 钉钉群投递：图文同框 + 文字版@负责人
8. 验证：messages-query-send-status 确认 SUCCESS
9. 写幂等标记

## 复用到其他 Quark 报表

只需改 3 处：
1. `fetch_homes*.js` 中的 `reportId` 和 `QUARK_PARAMS`
2. `fill_template.py` 中的 Excel 模板路径和列映射
3. cron payload 中的群 ID 和 @负责人列表

## DWS CLI 命令

```bash
# 群消息发送（Markdown 图文同框）
dws chat +messages-send --group <群ID> --markdown "..."

# 群消息发送（文件消息）
dws chat +messages-send --group <群ID> --file /path/to/image.png

# 消息转发（受限群接力）
dws chat +messages-forward --message-id <id> --group <目标群ID>

# 发送状态验证
dws chat +messages-query-send-status --message-id <id>

# 钉盘文件上传
dws drive +upload --file /path/to/report.png

# 获取本人 userId（自测）
dws contact +get-self
```

## 踩过的坑

| 坑 | 现象 | 解决方案 |
|----|------|----------|
| 图片+@同框渲染失败 | Markdown 中 @ 符号破坏图片渲染 | 拆为两条消息：图文同框 + 文字版@ |
| 数据未入库 | WidgetAction 返回 rows=0 但接口不报错 | 探测 rc=0 但 rows=0 时判定未就绪，间隔重试 |
| 防重复发送 | 主任务与重试任务并行险双发 | 主任务 delegate 成功后立即写标记 |
| QUARK_PARAMS 过期 | DA_TICKET 堆积致重定向循环 | 去掉旧 DA_TICKET 重导航获新令牌 |

## License

MIT
