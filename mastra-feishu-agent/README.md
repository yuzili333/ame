# Mastra 飞书办公智能体

独立的 TypeScript 服务。飞书官方 Node SDK 的 WebSocket 长连接处理私聊、群聊与卡片回调；Mastra AgentController 管理会话，Memory/LibSQL 保存上下文，Workspace 提供文件、命令与 Skills，Workflow 处理会议纪要，Studio 提供本机调试。

## 准备

1. Windows 11 云桌面安装 Node.js 22.13+ 和 npm，在本目录执行 `npm ci`。
2. 复制 `.env.example` 为 `.env` 并填写飞书应用 ID、密钥、私有化域名、模型代理 URL 与模型名。`FEISHU_DOMAIN` 为完整的 `https://` 内网域名；`MODEL_PROXY_BASE_URL` 填 API 基础路径（例如 `https://api.siliconflow.cn/v1`），不要附加 `/chat/completions`。Gateway 和 Studio 中注册的 Agent 共用 `MODEL_PROXY_BASE_URL`、`MODEL_PROXY_MODEL`、`MODEL_PROXY_API_KEY`。`MODEL_PROXY_MODEL` 可以写 `deepseek-ai/DeepSeek-V4.1-Flash`，也可以写 Studio 显示的 `siliconflow-cn/deepseek-ai/DeepSeek-V4.1-Flash`；调用 OpenAI 兼容接口时会自动去掉 `siliconflow-cn/` 前缀。若在 Studio 的 SiliconFlow CN 模型选择器中调试，还需将 `SILICONFLOW_CN_API_KEY` 设为同一密钥；使用 SiliconFlow CN 地址时，服务启动会校验两者一致。模型代理需要提供 OpenAI 兼容的流式 `tool_calls`，由代理完成 SSO。
3. 飞书开发者后台启用机器人、**长连接事件订阅**，订阅 `im.message.receive_v1` 和 `im.message.recalled_v1`；在卡片回调中订阅 `card.action.trigger`。撤回事件需要应用具备飞书后台要求的消息读取权限并发布新版本。把机器人加入需要服务的群，并确保应用可访问目标文档、日历与任务。
4. 在权限管理中，按照租户当前版本的接口权限清单授予：接收与发送消息、更新消息卡片，创建/读取/编辑新版文档，Wiki 节点读取，电子表格和多维表格读取/写入，读取/创建/更新日历日程，以及读取/创建/更新任务。共享资源默认使用应用身份。个人日历使用用户身份，并要求用户单独授权；私有化环境缺失相应接口时会明确报错，不会借用另一种身份代执行。具体权限以 [飞书开放平台的接口页面](https://open.feishu.cn/document/) 为准。
5. 将需要私聊机器人的用户 `open_id` 写入 `FEISHU_ALLOWED_USERS`，以逗号分隔；`FEISHU_ALLOWED_CHATS` 用于允许私聊和代发消息的目标会话。首次启动前可先留空：私聊机器人发送 `/id` 获取自己的 `open_id` 与私聊 `chat_id`。将机器人邀请进群后，群内任意成员直接 `@机器人` 即可对话，无须为新成员或新群修改 `.env`、重启网关。群聊普通消息仍须 @机器人。`/id` 只用于查询 ID，不能赋予私聊或跨会话代发权限。`open_id` 与应用绑定，须从这个机器人获取。

## 运行

分别打开两个终端：

```powershell
npm run gateway
npm run dev
```

也可用 `powershell -ExecutionPolicy Bypass -File .\scripts\start-windows.ps1 -Mode both`。Mastra Studio 与 API 只绑定 `127.0.0.1:4111`，访问 `http://127.0.0.1:4111/studio`。在 Studio 中查看 Agent、Workflow、Tools 与 Traces；飞书写工具的实地审批需要同时运行网关。生产运行前执行 `npm run build`，再使用 `npm run start` 启动已构建的本机 Studio/API。网关仍以 `npm run gateway` 单独启动。

更新网关代码后，先在原网关终端按 `Ctrl+C`，确认进程退出，再执行 `npm run gateway`。启动日志会打印新的 PID 和入口文件；`data/gateway.pid` 用于阻止第二个网关实例。如果提示已有网关运行，应先检查该 PID 对应的进程并停止旧实例，不能只启动一个新终端，否则飞书消息可能仍由旧代码处理。

在 Studio 中调试需要审批的工具时，给 Request Context 设置 `{"caller":{"userId":"<允许的 open_id>","chatId":"<允许的私聊 chat_id>","chatType":"p2p"}}`；群聊调试可使用机器人所在群的 `chat_id` 与 `"chatType":"group"`，不需要群白名单。Workflow 输入还需要相同的 `caller` 对象。Studio 使用应用凭据直接发审批卡片，网关负责接收确认回调。

## 功能与审批

- 私聊按用户独立记忆；群聊按群共享，和私聊隔离。机器人所在群的成员均可通过 @机器人使用 Agent；邀请机器人入群即意味着该群成员可访问 Agent 的对话及机器人应用有权读取的办公资源。私聊仍要求用户和私聊 ID 均获准；跨会话代发仅限 `FEISHU_ALLOWED_USERS` 内用户，目标会话仍需列入 `FEISHU_ALLOWED_CHATS`。群内写操作仍由原请求者逐项审批。消息 ID 在本地 SQLite 中去重，同一会话串行处理。
- 文档、日程与任务的读取直接用机器人应用身份执行；新建、更新、代发到其他会话的消息先发卡片，原请求者在 30 分钟内逐次确认。普通当前会话回复直接发送。文档更新针对指定文本块；新建文档可包含正文。
- 创建个人日程时，Agent 使用用户身份；若尚无用户日历授权，机器人先发文字询问选择“用户权限”或“应用权限”。选择用户权限会收到 OAuth 授权链接，授权账号必须与原请求者的 `open_id` 一致；选择应用权限时需要提供机器人可写的共享日历 ID。权限选择和授权阶段都不会发送写入审批卡片；目标身份及日历确定后才发卡片。群聊中的选择回复仍须 @机器人，且只有原请求者能继续该请求。回复“取消”会放弃。用户令牌在本地 SQLite 中加密保存，到期时使用一次性刷新令牌更新；用户身份失败不会自动回退应用身份。若旧令牌缺少日历权限，选择“用户权限”会发起新的授权链接；授权失败后 30 分钟内可直接回复“重新授权”。新申请会替代该会话的失败授权记录。
- 日程的相对日期以飞书原消息时间和北京时间为准；工具在发起审批前校正“今天／明天／后天”及完整年月日与模型时间戳不符的情况。审批卡片显示可读的北京时间，执行成功后展示实际创建的日程时间和飞书日程链接。
- 文件工具以 `AGENT_WORKSPACE_DIR` 为根目录。命令从该目录启动，但继承 Windows 进程的文件权限；应以受限服务账号运行。文件写入、编辑、删除、命令执行与定时任务使用 Mastra 工具审批。Skills 位于 `workspace/skills/`。外部 MCP 服务仅从 `MCP_SERVERS_JSON` 中连接，且只向 Agent 暴露 `allowedTools` 列表中的工具；MCP 工具调用也需审批。
- 输入 `/minutes <documentId>` 执行“读会议纪要 → 提取最多 10 个待办 → 每项任务单独提请审批 → 生成通知草稿”。草稿作为当前会话回复；若要向别处发送，需另行调用代发消息工具并审批。
- 请求“按模板生成下发材料”时，提供模板文件夹链接、材料父文件夹链接、下发编号及各字段的值。Agent 调用 `office_release_materials` 先只读扫描当前层，再逐文件发送预览，最后只发一张整批确认卡片；原请求者确认后运行 `release-materials` Workflow。也可单独使用目录查询、创建子目录及复制单个文件工具，其中写操作同样须审批。
- `fields` 中允许包含未在模板出现的字段；预览会列出这些未使用字段，它们不会作为参数传给飞书 API，也不会写入副本。若飞书返回 `99992402`，网关错误会列出请求方法、接口、查询参数名及飞书返回的字段违规信息，便于定位实际失败的接口；请在确认卡片出现前处理预检错误。
- 若下发批次失败或审批被拒绝、过期，原请求者可回复“重新发起”重新执行只读预检并获取新卡片；在群聊中必须发送“@机器人 重新发起”。机器人仅订阅群内 @ 消息时，未 @ 的群消息不会到达网关。新批次不会复用旧审批，执行结果不明的批次仍需人工核对。
- 用户撤回尚在推理或排队的消息时，网关按消息 ID 取消该轮运行、队列项和关联的待审批操作；已进入飞书写入接口的操作结果仍需人工核对。会话忙时最多暂存 20 条指令并按接收顺序处理，排队超过 30 分钟失效；网关重启后只恢复尚未开始的队列项，不重放执行中的任务。群聊指令仍需 @机器人。
- 审批记录、消息去重与 Workflow 状态位于 `data/state.db`；Mastra Memory/Controller 状态位于 `data/mastra.db`。重启后，未决的 Office 写入卡片在有效期内仍可确认；挂起的本地/MCP 工具审批与运行中的 Workflow 标记为 `interrupted`，不会重放。外部接口失败后不自动重试，以避免不确定提交造成重复写入。

### 个人日历授权配置

在飞书开发者后台为应用开通用户身份的 `calendar:calendar`、`calendar:calendar:readonly` 和 `offline_access` 权限，并在“安全设置”登记 `FEISHU_OAUTH_REDIRECT_URI`。填写 `.env` 中的 OAuth 授权页 URL（例如 `https://accounts.feishu.cn/open-apis/authen/v1/authorize`）、令牌 URL（例如 `https://open.feishu.cn/open-apis/authen/v2/oauth/token`）以及回调 URL。私有化部署应使用其实际内网域名。回调服务由网关在 `FEISHU_OAUTH_LISTEN_HOST:FEISHU_OAUTH_LISTEN_PORT` 启动；当前示例使用 `http://localhost:4000/auth/feishu/callback`，因此授权页面必须在运行网关的同一台电脑上打开。若飞书客户端运行在其他设备，应改为客户端可访问的内网 HTTPS 地址，并在开发者后台同步登记、通过反向代理转发到监听端口。Studio 仍只在本机监听。`FEISHU_TOKEN_ENCRYPTION_KEY` 应设置为长期稳定的随机密钥并妥善备份；重启后要能解密已保存的用户令牌。未配置这些变量时，应用身份的共享日历操作仍可用，选择用户权限会得到明确的配置错误。

新增日历读取权限后，已授权用户需要重新打开授权链接以取得新范围。

用户授权流程依次为：文字选择身份 → 打开授权链接 → 回调校验原请求者身份 → 查询其主日历写入权限 → 发送写入审批卡片。已有用户令牌也会在发卡前检查目标日历的写入权限；权限不足时重新追问并引导补充授权。授权链接及卡片均在 30 分钟内有效；审批依旧只接受原请求者、原会话的一次确认。授权成功本身不创建日程。联调时应额外验证回调 URL 可达、用户身份与应用身份各自的日历范围、授权拒绝、错误账号、令牌刷新和重复回调。

### 电子表格与多维表格

机器人接受 `wiki/…`、`sheets/…`、`base/…` 链接。Wiki 节点会先解析实际资源类型，电子表格使用 Sheets API，多维表格使用 Bitable API；`docx` 仍使用原文档工具。电子表格支持查看、按表头写入首个空白格、按表头追加整行、按 A1 地址或唯一匹配行修改单元格；多维表格支持查看、创建记录、按记录 ID 或唯一字段匹配修改记录。Bitable 写入目前只接受文本、数字和复选框字段；其他字段会明确拒绝。工作表、表头或记录不唯一时需要用户提供明确目标。

在开发者后台开通 Wiki 节点读取（例如 `wiki:node:read`）、Sheets 读取与写入、Bitable 读取与编辑，以及云文档“判断用户权限”接口所需权限，并将机器人应用加入目标文档或知识空间的协作者；多维表格启用高级权限时还需赋予相应数据表权限。权限名称应以私有化租户的接口文档和 `99991672` 返回信息为准。`FEISHU_DOCUMENT_HOSTS` 应包含实际表格链接的主机名。应用权限不足时，机器人会询问是否改用请求者的用户权限；只有请求者明确选择后才发送 OAuth 链接。`FEISHU_TABLE_OAUTH_SCOPES` 需与开发者后台开通的用户身份权限一致，已有仅含日历权限的令牌会要求重新授权。授权成功后才发写入审批卡片；回调或审批重复提交不会再次执行。审批时目标数据若已变化，写入会中止，远端结果不明时不自动重试。

私有化联调先用测试 Wiki 链接验证节点解析，再分别用测试电子表格和测试多维表格验证只读、审批和写入；若部署不提供某个 OpenAPI，错误会指出缺失的接口。普通电子表格首空格最多扫描表头以下 9999 行，超过时需提供明确单元格。

### 下发材料自动生成

示例请求：

> 请用模板文件夹 `https://tenant.example/drive/folder/<模板token>` 生成本次下发材料，放到 `https://tenant.example/drive/folder/<材料父目录token>`，下发编号 `20261029`，下发日期 `2026-10-29`，下发版本负责人 `张三`。

字段名要与模板中的 `{{字段名}}` 一致；“下发编号”自动取文件夹名，若显式提供则必须相同。兼容“版本负责人”和模板中常见的误写“版本负责任”，两者若同时提供必须取相同值。Docx、文件标题或多维表格中缺少需替换字段时预检停止并要求补充；电子表格中未匹配本次字段的占位符保持原值。多提供的字段会列在预览中，不会传给飞书 API，也不会写入副本。仅处理普通云空间文件夹当前层的 Docx、普通电子表格和多维表格；当前层其他类型会报错，子文件夹列入忽略清单。Docx 支持标题、正文和表格单元格中的文本块，单个占位符须位于同一个富文本元素内；电子表格和多维表格的文件标题若含占位符，会在预览中展示改名前后值，确认后仅改副本标题并回读；电子表格普通文本单元格支持整值占位符及 `L_REL_{{下发编号}}` 这样的混合文本，单元格内可替换多个已提供字段，公式不写入；多维表格仅处理普通文本字段的整值占位符。复杂格式会在审批前报错。没有匹配占位符的文件会直接复制，不调用内容写入接口。

普通电子表格允许不同工作表只包含部分字段。预检每次最多读取 1,000 个单元格；若某工作表配置的网格超过 100,000 个单元格，需明确提供该表格文件名、工作表名及 `A1:D100` 格式的扫描范围，工具会将其作为 `sheetScopes`，并在审批预览中注明未扫描的工作表。合并单元格中的占位符必须位于合并区域左上角；预览显示整个合并范围，执行前再次核对合并结构，只写入副本的左上角。读取失败会返回工作表、范围、接口和飞书的字段校验详情，便于定位私有化接口差异。

开发者后台需为**应用身份**开通并发布：目录枚举、创建与复制所需的 `drive:drive:readonly` 和 `drive:drive:write`（或租户提供的 `drive:drive`）；新版文档读取和更新所需的 `docx:document:readonly`、`docx:document`；普通电子表格读取与写入所需的 `sheets:spreadsheet:read`、`sheets:spreadsheet:write_only`（或完整读写权限）；多维表格读取与更新所需的 `bitable:app:readonly`、`bitable:app`。私有化租户的权限名称以接口页面及 `99991672` 错误提示为准。管理员须将模板文件夹与材料父目录共享给**包含机器人和请求者的协作群**，并授予机器人源目录读取/复制及目标目录写入权限；材料父目录中的当前协作群须为“可编辑”或“可管理”，仅“可查看”无法创建子文件夹。不能仅把应用直接添加为文件夹协作者。若多维表格使用高级权限，还需授予机器人相应数据表读写权。`FEISHU_DOCUMENT_HOSTS` 必须包含客户端复制链接的域名；公司内网分享域名 `cibt.co.cibfintech.com` 已加入示例与本地配置。API 域名 `FEISHU_DOMAIN` 与文档链接域名可以不同，工具从链接中提取 `folder_token` 后向配置的 API 域名发起请求；请提供完整的 `https://<文档域名>/drive/folder/<token>` 链接。请求者最终能否打开材料文件夹应由该账号在客户端验证。

预检读取完整文件清单及占位符；确认时重新扫描源文件和目标目录。同名材料文件夹会停止，源内容变化会要求重新预览。确认后只修改副本，副本尚未可读时有限等待；每次远端复制/写入最多发起一次。SQLite `release_jobs` 和 `release_files` 记录 Workflow run ID、审批、文件夹、副本 token 和逐文件状态。复制/写入超时或服务重启后，状态标记为需人工核对，保留已有材料，不自动重试、删除或回滚；卡片显示已知文件链接与未完成状态。完整的运行路径需在私有化环境中用测试目录人工确认一次，并从请求者账号验证三类副本内容和可见性。

## 调试与验收

```powershell
npm run typecheck
npm test
npm run build
```

自动化测试使用内存中的飞书传输替身，不需要应用凭据。Windows 私有化联调仍需逐项验证：长连接建立、私聊/群聊 @、卡片回调身份、文档/日程/任务读取与写入、模型代理的流式工具调用、MCP 服务以及 Studio 四个视图。若 `card.action.trigger` 或目标 OpenAPI 在私有化版本不可用，必须保持对应写操作不可用，并依据实际租户接口再适配。

参考：[Mastra Agent Harness](https://mastra.ai/templates/agent-harness)、[Mastra Harness](https://mastra.ai/docs/harness/overview)、[Mastra Workflow](https://mastra.ai/docs/workflows/overview)、[Mastra Studio](https://mastra.ai/docs/studio/overview)、[飞书 Node SDK](https://github.com/larksuite/node-sdk/blob/main/docs/channel.md)。
