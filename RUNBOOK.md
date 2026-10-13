# Homes 营收对比表 · 每日自动导出运行手册（RUNBOOK）

> 本手册供定时任务代理（cron agent）逐步执行。目标：每天自动从盒马 Homes「店长大盘」抓取数据，
> 生成**两份**报表——【MTD 月度累计】与【单日】，各自渲染成报表长图 PNG，经问问通道投递
> （图片为主，Excel 存档）。所有脚本位于 `D:\xiaoQ专用\Homes_MTD自动导出\`。

## 背景与原理（执行前必读）
- 「店长大盘」是 FBI/Quark 报表 reportId=312793。页面 URL 携带签名令牌 `QUARK_PARAMS`。
- 数据不靠抓 DOM，而是直接重放报表的 WidgetAction 数据接口（4 个 widget），返回纯 JSON 数组：
  - 497734 品类看板-部门（各部门 未税营收/成交/app/pos 明细）
  - 497737 商品管理-部门（损耗率/管损率）
  - 497731 门店看板-整体表现（总计行 成交/用户/客单价）
  - 497780 门店看板-分渠道成交（总计行 app/pos）
- 取数脚本 `fetch_homes.js` 在报表页上下文执行 fetch，把 4 个 widget 的 JSON POST 到本地
  `homes_server.py`（127.0.0.1），服务器落盘原始数据并调用 `fill_template.py` 生成 Excel。
- **传输层已改 XHR（2026-09-11）**：实测报表页上下文 `fetch()` 偶发全面失效（同域 GET 也报
  Failed to fetch，而 XHR 与顶层导航正常，疑似扩展/页面脚本钩住 window.fetch）。fetch_homes.js
  已整体改用 XMLHttpRequest（widget 异步、本地回环同步 POST）。若再遇 "Failed to fetch"，
  先确认登录态与页面正常后，优先怀疑 fetch 被钩，直接用脚本现状（XHR）即可，不必回退。
- **周期 period 两档**：`mtd`=月度累计；`1d`=单日。fetch_homes.js 用占位符 `__PERIOD__` 注入。
- **单日口径**（已实测验证）：单日环比=比上周同日（7 天前同星期）；单日无同比（API 返回"-"）；日均=当日值。
- **预算口径**（已对 9 月逐部门勾稽验证，差异=0.0%）：
  - 预算来自 `D:\xiaoQ专用\Homes_MTD自动导出\预算分解.xlsx`（budget_lib.py 读取）。
  - 日预算表为【含税成交】，按日参数税比折【未税】：部门当日未税预算=(线上+POS)/税比。
  - 单日预算=当日未税预算；MTD 预算=本月 1 日至统计日逐日累计未税预算。
  - 预算占比=部门预算/总预算；达成率=未税营收/预算（整体 K=E/J，POS AQ=POS未税营收/POS预算）。
  - fill_template.py 自动校验预算覆盖期（checks.budget_covered）；统计日期超出预算覆盖期时会在
    checks 里标出 budget_covered=false，需在投递文案提醒用户更新预算文件。
- 口径已用 8 月数据逐格验证（核心财务单元格 100% 精确匹配），无需再核对公式。

## 预算文件维护（用户职责）
- 预算文件 `预算分解.xlsx` 当前覆盖 **2026-09-01 ~ 2026-10-07**。
- 预算期滚动（如 10 月中旬后）时，用户需把新一期预算分解文件覆盖到该路径，
  格式与现文件一致（日预算表按 日期×部门_线上/_POS + 日参数税比）。
- 定时任务检测到统计日期超出预算覆盖期时，报表照常生成但预算列会缺失，并在投递时提醒。

## 前置条件（缺一不可）
1. 电脑开机、未休眠，Chrome 正在运行且 QwenWork 浏览器插件可用。
2. Homes/阿里 SSO 登录态有效（能打开店长大盘不跳登录页）。
若任一不满足，本次任务会失败——按「失败处理」通知用户，不要硬试。

## 执行步骤

### 第 0 步：确定统计日期
用 Bash 运行 `date` 获取今天，统计日期 = **昨天**（T-1），格式 YYYYMMDD。
例：今天 2026-09-08 → DATE=20260907。跨月第一天则取上月最后一天（即上月完整 MTD）。

### 第 1 步：打开报表页拿 QUARK_PARAMS
- `tabs_context` 查看现有标签；用 `tabs_create` 新建空标签，`navigate` 到：
  `https://portalhm.hemaos.com/rbac/quark/?reportId=312793&ts=1`
- 该地址自动重定向到 `quark.alibaba.com/dashboard/view/report.htm?...&QUARK_PARAMS=...`。
- 用 javascript_tool 执行 `document.title + ' | ' + (location.href.match(/QUARK_PARAMS=/)?'HAS_QP':'NO_QP')` 确认。
  - 标题含「登录」或返回 NO_QP → SSO 失效，转「失败处理-A」。（实测 2026-09-11：跳到登录页即真失效，
    无静默续期；老板手动登录后立即恢复。登录态有效期约 1 天，属正常现象。）
  - 记住这个 tabId（下面两个周期复用同一个标签）。

### 第 2 步：跑 MTD 周期（端口 8790）
1. Bash **run_in_background** 启动服务器：
   `"C:\Users\泇辰461427\AppData\Local\Programs\Python\Python312\python.exe" "D:\xiaoQ专用\Homes_MTD自动导出\homes_server.py" --date <DATE> --port 8790 --timeout 300`
2. 读取 `fetch_homes.js`，把 `__DATE__`→<DATE>、`__PORT__`→`8790`、`__PERIOD__`→`mtd`，
   整段作为 javascript_tool 的 text 在第 1 步 tabId 执行；期望返回 `POSTED ...B -> RECEIVED`。
3. 读后台输出（一行 JSON），校验 `status=="ok"` 且 `date_echo_ok==true`。
   产物：`out\盒马Homes自动生成MTD营收对比表_<DATE>.xlsx`。

### 第 3 步：跑单日周期（端口 8791）
与第 2 步完全相同，仅三处替换：端口 `8791`、`__PERIOD__`→`1d`。
产物：`out\盒马Homes自动生成单日营收对比表_<DATE>.xlsx`。
（单日预算=预算文件里当日未税预算；单日环比=比上周同日。）

### 第 4 步：生成两张长图（xlsx → PNG）
分别运行（把 <DATE> 换成统计日期）：
```
"...Python312\python.exe" "D:\xiaoQ专用\Homes_MTD自动导出\xlsx_to_image.py" "D:\xiaoQ专用\Homes_MTD自动导出\out\盒马Homes自动生成MTD营收对比表_<DATE>.xlsx"  "D:\xiaoQ专用\Homes_MTD自动导出\out\盒马MTD营收对比表_<DATE>.png"
"...Python312\python.exe" "D:\xiaoQ专用\Homes_MTD自动导出\xlsx_to_image.py" "D:\xiaoQ专用\Homes_MTD自动导出\out\盒马Homes自动生成单日营收对比表_<DATE>.xlsx" "D:\xiaoQ专用\Homes_MTD自动导出\out\盒马单日营收对比表_<DATE>.png"
```
每条成功输出 `IMAGE_OK ...png size=...x... bytes=...`。图片按 预算达成/整体成交/盒马APP/盒马POS/线上汇总/渠道占比 六块纵向排布，负值标红、总计加粗；单版图抬头自动标注「单日 (1d)·环比=比上周同日」。

### 第 5 步：投递（问问通道 + 存 D 盘；醒目提醒 + 两张图）
- 背景：问问聊天对 .xlsx 只能当附件、无法内嵌打开；PNG 可直接渲染。所以**投递以图片为主**。
- 用 `qwenwork_channel_list_conversations` 找到「问问」(xiaoq) 会话，再用 `qwenwork_channel_delegate_to_im`
  投递，要求问问代理**先发一条醒目提醒文字消息，再发两张图片**（MTD 长图 + 单日长图，直接发图不要只发路径）。
- 醒目提醒消息模板（尖括号换成真实数字，金额千分位；指标取自两份 fill_template 输出 JSON 的 E18/K18/AQ18/N18）：
  ```
  ━━━━━━━━━━━━━━
  📊 【每日营收报表提醒】
  📅 统计日期：<YYYY-MM-DD>
  【MTD月度累计】
  · 总未税营收 <X> 元，预算达成 <X>%
  · 总成交 <X> 元
  【单日】(环比=比上周同日)
  · 总未税营收 <X> 元，单日预算达成 <X>%
  · 总成交 <X> 元
  👇 两张报表长图见下
  ━━━━━━━━━━━━━━
  ```
- 若 checks 有异常（budget_covered=false / 达成率>200% / 营收为0 / missing_in_api 等），在提醒消息里用 ⚠️ 明确标出。
- 同时告知：两份完整 Excel 已存 `D:\xiaoQ专用\Homes_MTD自动导出\out\`。
- **发送成功后，立即用 Bash 写入已发送标记文件**（供重试任务幂等判断，防止重复发送）：
  `echo 已发送 $(date "+%Y-%m-%d %H:%M:%S") > 'D:/xiaoQ专用/Homes_MTD自动导出/out/已发送_<DATE>.txt'`
  （<DATE> 为统计日期 YYYYMMDD。只有真正发送成功才写；未发送不得写。）
- 备注：用户最终想发钉钉并要 DING 强提醒，但当前钉钉通路均被组织权限拦截（dws 报"该组织尚未开启 CLI
  数据访问权限"，DING 同样被拦；QwenWork 钉钉通道报 CORP_VALIDATION_FAILED）。现以问问通道醒目提醒消息代替；
  待组织开通权限后改为直接发钉钉 + DING 强提醒。

### 自动数据审核（每次运行内置，无需手动）
fill_template.py 生成每份报表时会自动执行一组校验，结果随 JSON 的 checks 字段输出。cron 代理必须
逐项看 checks，任何一项异常都要在处理结果里向用户说明，不得静默跳过：
- `date_echo_ok`：API 返回的统计日期与请求日期一致（防止拿到旧数据）。必须为 true。
- `budget_date_ok` / `budget_covered`：统计日期在预算文件覆盖期内。若 covered=false，报表照常出但
  预算列缺失，须提醒用户更新预算文件。
- `online_plus_pos_eq_1`：线上占比+POS占比≈1（渠道拆分自洽）。
- `diff_N18_vs_sumDept`：门店级总成交 与 部门Σ成交 的差额（正常约几千元，因门店级含未分摊部分）。
- `missing_in_api` / `new_in_api`：模板部门与 API 部门的差集（某部门当日/当月无数据时列出）。
- `K18` / `AQ18` / `E18` / `J18_budget`：总达成率、总营收、总预算回显，量级异常（如达成率>200% 或
  营收为 0）时须在文案中提示用户复核。
此外口径公式本身已用 8 月真实数据逐格验证（100% 精确匹配）、预算已对 9 月逐部门勾稽（差异 0.0%），
属于一次性验证通过的固定资产，日常运行不需重复核对公式。

### 失败处理
- A（登录态失效）：**先尝试自动登录，两级兜底**（见下节「自动登录机制」）：
  第一兜底 = MCP 坐标法（步骤 1~9）；一旦出现视口 0x0 / surface 漂移症状（见「第二兜底」小节的症状清单），
  立即切换第二兜底 = `hema_login.py` Python 原子脚本，不要用 MCP 反复重试。
  仅当两级兜底都失败（无凭据文件/凭据未填写/密码错误/出现验证码）时，才通过问问通道发消息：
  「Homes 登录态已失效且自动登录未成功，今日营收对比表未生成。请打开盒马 Homes 店长大盘完成登录后告诉我重跑。」
- B（数据未就绪 / 服务异常）：若判定为「统计日期数据尚未入库」（取数返回空、rows=0），**不要手动重跑**——
  由独立的「数据未就绪自动重试」任务每半小时自动重试，数据就绪后自动补发（见下节）。
  主任务此时只需发一条简短预告「数据暂未入库，将每半小时自动重试，好了自动发你」。
  若 MTD 成功、单日失败（或反之），先把成功的那份发出，并说明另一份失败原因。
- 任何情况下都不要伪造数据或跳过校验硬交付。

### 自动登录机制（账号+密码，2026-09-11 上线）
背景：Homes 登录态有效期约 1 天，过期后需重新登录。登录页（login.portal.hemaos.com）只需账号+密码、
无短信验证码/扫码（老板 2026-09-11 确认）。**结构要点（已实测）**：真正的登录表单位于**跨域 iframe**
（havanalogin.taobao.com/mini_login.htm）内，主页面 JS 读不到 iframe 内部 DOM，无障碍树也看不到——
必须用「坐标点击 + 剪贴板粘贴」注入凭据；登录成功后父页面可能再弹「选择商家」对话框（该对话框与
父页面同源，可正常点击）。

**凭据与纪律**
- 凭据文件：`D:\xiaoQ专用\Homes_MTD自动导出\credentials.json`，格式
  `{"username":"<账号>","password":"<密码>"}`，由用户本人填写与维护。
- 辅助脚本：`D:\xiaoQ专用\Homes_MTD自动导出\clip_cred.py`，用法
  `python clip_cred.py user|pass|clear`——把账号/密码从凭据文件直接送入 Windows 剪贴板（Win32 API，
  不回显任何内容）。
- ⚠️ 纪律：凭据内容**严禁**出现在任何聊天消息、汇报文字、日志或工具参数里。输入一律走
  「clip_cred.py 入剪贴板 → 点击输入框 → ctrl+v 粘贴」，绝不把账号密码作为文本参数传递、
  绝不 cat/echo 凭据文件。流程结束必须 `clip_cred.py clear` 清空剪贴板。
- 文件不存在或字段仍是占位符（含「请填写」）→ 跳过自动登录，直接走失败处理-A 的通知分支。

**自动登录步骤**
1. 校验凭据文件已填写（用 Python 读 JSON 判断，只输出是否就绪，不输出内容）。失败则走通知分支。
2. navigate 到
   `https://login.portal.hemaos.com/login?from=https%3A%2F%2Fportalhm.hemaos.com%2Frbac%2Fquark%2F%3FreportId%3D312793`
   等页面加载完（约 3~5 秒）。
3. 用 javascript_tool 测 iframe 位置：
   `document.querySelector('iframe[src*="mini_login"]').getBoundingClientRect()`，记下左上角 (X,Y)。
4. 计算输入框坐标（2026-09-11 实测 iframe 内布局：输入框距 iframe 左边 20px、宽 220px、高 36px；
   账号框顶边距 17px，密码框顶边距 58px）：
   - 账号框中心 = (X+130, Y+35)
   - 密码框中心 = (X+130, Y+76)
5. Bash 运行 `python clip_cred.py user` → computer left_click 账号框坐标 → computer key `ctrl+v`。
6. Bash 运行 `python clip_cred.py pass` → computer left_click 密码框坐标 → computer key `ctrl+v`。
7. computer key `Enter` 提交（勿点按钮：独立渲染时按钮位置不稳，回车提交已验证可触发 login.do）。
8. Bash 运行 `python clip_cred.py clear` 清空剪贴板。
9. 等 8 秒，按序验证：
   a. 直接 navigate 报表地址 `https://portalhm.hemaos.com/rbac/quark/?reportId=312793&ts=1`：
      拿到 QUARK_PARAMS → 登录成功，继续正常取数流程（不必打扰用户）。
   b. 若仍被弹回登录页 → 回到登录页扫描「选择商家」对话框：用 javascript_tool 遍历可见元素，
      找文本含「商家/盒马/进入」的对话框条目并点击盒马对应项（同源可点），等 5 秒后再验 a。
   c. 仍失败（停登录页/出现验证码/错误提示）→ 停止，走失败处理-A 通知分支，
      说明「自动登录未成功（可能触发风控或验证码），请手动登录一次」。

**安全阀**：同一天自动登录尝试 ≤2 次（防账号锁定/风控），超过即停改人工通知。
**维护**：用户改密码后只需更新凭据文件；删除凭据文件即关闭自动登录（回到纯人工模式）。

**实测记录（2026-09-11 凌晨）**：凭据已通过 Havana 认证（newlogin/login.do 成功、无验证码）；
坐标点击+剪贴板粘贴在真实登录页 iframe 内成功触发 account/check.do 与 login.do；
商家选择对话框环节当次未观察到弹出（疑似短时间内第 2 次登录触发风控）。首次自然过期实战时
重点观察第 9-b 步；若连续出现，考虑在两次登录间增加间隔或改回人工。

**第二兜底：Python 原子脚本 hema_login.py（2026-09-13 实战验证）**

MCP 坐标法有两种已实锤的失效模式——出现任一症状立即改用本兜底，不要用 MCP 反复重试：
- **视口 0x0**：javascript_tool 返回 innerWidth/innerHeight=0，iframe getBoundingClientRect 为负坐标，
  resize_window 无效。整页渲染到视口外，只发生在**新建的 tab** 上，已有 tab 正常。
- **surface 漂移**：存在多个 Chrome 窗口时，computer 工具的截图/点击落到 OS 前台窗口而非目标标签
  （截图显示的是别的页面；点击后 activeElement 聚焦了 iframe 但粘贴不进去；zoom 返回别的窗口画面）。

原理：绕开 MCP computer 工具，用 Win32 API 在单进程内一次完成「窗口激活+坐标点击+剪贴板粘贴+回车」。
脚本已归档为 `D:\xiaoQ专用\Homes_MTD自动导出\hema_login.py`（凭据由脚本自读 credentials.json，零回显）。

步骤：
1. 换到任一**渲染正常的已有标签**（javascript_tool 确认 innerWidth>0），navigate 到上面第 2 步的
   登录页 URL。（不要用 0x0 的标签，也不要再新建标签硬试。）
2. javascript_tool 测量：
   `JSON.stringify({f:(function(){var r=document.querySelector('iframe[src*="mini_login"]').getBoundingClientRect();return [r.x,r.y]})(),vw:window.innerWidth})`
   得 iframe 左上角 CSS 坐标 (x,y) 与 CSS 视口宽 vw。
3. Bash 执行：
   `"C:\Users\泇辰461427\AppData\Local\Programs\Python\Python312\python.exe" "D:\xiaoQ专用\Homes_MTD自动导出\hema_login.py" <x> <y> --css-vw <vw>`
   期望输出 WINDOW/VIEWPORT_PX/SCALE/POINTS 四行 + `LOGIN_SUBMITTED`。
   脚本内部：EnumWindows 找标题含 HEMAOS 的窗口 → Alt trick SetForegroundWindow 前台化 →
   EnumChildWindows 取 Chrome_RenderWidgetHostHWND（top≥100 且面积最大 = 真实 web 视口）→
   SendInput 点击（视口原点 + CSS 坐标 × scale，scale=物理视口宽/vw）→ 剪贴板粘贴账号/密码 →
   Enter 提交 → 清空剪贴板。
4. 改过脚本或排查探测问题时可加 `--dry-run`（只探测窗口/视口并输出将使用的坐标，不点击不粘贴）。
5. 提交后的验证与后续同上面第 9 步（navigate 报表地址拿 QUARK_PARAMS / 处理「选择商家」对话框 /
   仍失败则停止并走通知分支）。

**实测记录（2026-09-13 晨）**：三条流水线因登录态过期全 error；MCP 坐标法先后踩视口 0x0 与
surface 漂移全部失效；改本脚本，iframe (868,230.875)、视口原点 (0,221)、物理点击落点
(1247,553)/(1247,604)，一次成功登录（商家选择对话框当次未弹出），拿到 QUARK_PARAMS 完成当日全部交付。
**补测（2026-09-15）**：① 窗口处于后台时渲染子窗口矩形不刷新（只得 (0,0) 幻影矩形、真实视口子窗口
缺席），必须先前台化再枚举——脚本已按「先前台化后枚举+找不到真实视口延时 1 秒重试一次」实现；
② 视口原点不是固定 (0,221)，随 Chrome UI 状态变化（实测出现过 151 与 221 两种），每次必须实时读取，
严禁把原点或 POINTS 硬编码进任何脚本。
**安全阀**：与 MCP 坐标法合并计算，同一天自动登录总尝试次数 ≤2 次。

### 数据未就绪自动重试机制（独立重试任务）
除 09:00 主任务外，另有一个**重试任务**，调度为每天 `09:30~12:30 每隔半小时`（cron `0,30 9-12 * * *`，
但 09:00 那次直接跳过、交给主任务，避免同时刻抢端口）。重试任务每次运行逻辑：
1. 统计日期 = 昨天（T-1），记为 <DATE>。
2. **幂等判断**：若 `out/已发送_<DATE>.txt` 标记已存在 → 说明今天已发送成功，直接结束，不做任何事。
3. 若无标记 → 按第 1~4 步取数并校验：
   - 数据已就绪（rows>0）→ 生成 MTD+单日两份长图，按第 5 步发醒目提醒+两张图，并写已发送标记。
   - 数据仍未就绪（rows=0）→
     - 当前时间 < 12:30：**静默跳过**（不发任何消息，等下一个半小时再试）。
     - 当前时间 ≥ 12:30（当天最后一次）：发一条提醒「今日数据截至 12:30 仍未入库，已停止自动重试；
       数据就绪后对我说『重跑 Homes 营收对比表』即可补发」，然后结束。
4. 重试任务同样遵守「不伪造数据、先诊断再动作」；登录态失效按失败处理-A。

## 关键常量（勿改）
- reportId=312793；门店过滤值 19005（青岛盒马鲜生卓悦里店）；周期 mtd / 1d；端口 MTD=8790、单日=8791。
- widget 与 encryptWidgetId：
  497734/75a35406af74613c、497737/bff564290024995a（groupNo 14336961）
  497731/771f514d9159da27、497780/bd70adf8fc9f05b3（groupNo 14336959）
- 模板：`C:\Users\泇辰461427\Desktop\盒马Homes自动生成MTD营收对比表.xlsx`
  （仅提供部门行序与表格格式；预算改由 预算分解.xlsx 提供，不再读模板手工列。）
- 预算文件：`D:\xiaoQ专用\Homes_MTD自动导出\预算分解.xlsx`（当前覆盖 2026-09-01~10-07，见上）。
