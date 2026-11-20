# 存量办公操作纳入统一入口

## 检查结果

上一版仅共用底层 `OfficeApi` 请求、身份和错误处理，模型仍同时看到存量工具和通用调用工具；日历/任务写入被通用入口排除，其他部分原始写接口可绕开原业务预检。因此不能称为已完整纳入统一调用。

## 本次收敛

`office_call_api` 现在接收两种操作 ID，均从 `office_discover_apis` 查询完整参数定义：

| 业务 | 统一入口 operationId 示例 | 复用的逻辑 |
|---|---|---|
| 日历实体、日程读写 | `office_list_calendars`、`office_calendar_create/update/delete`、`office_calendar_container_*` | 原消息时间校正、日历角色检查、授权续办、审批 |
| 任务读写 | `office_list_tasks`、`office_task_create/update/delete` | 开始/截止时间、负责人、用户身份、审批 |
| 文档读写 | `office_read_document`、`office_doc_create/update` | 文档类型识别、真实 Docx ID、正文写入、审批 |
| 普通表格读写 | `office_table_inspect`、`office_table_update_cell/append_row/fill_first_empty` | 工作表/表头定位、权限预检、旧值与并发修改校验 |
| 多维表格读写 | `office_table_inspect/create_record/update_record` | 表与字段检查、唯一记录匹配、旧字段核对 |
| 云盘文件夹及文件 | `office_drive_list_folder/create_folder/copy_file/file_rename` | 当前层枚举、类型校验、复制结果检查、改名回读 |
| 指定群聊发送消息 | `office_message_send` | 请求者和目标会话白名单、操作审批 |
| 整批下发材料 | `office_release_materials` | 原批处理预检与 Workflow，不重复生成实现 |
| 其他 SDK JSON API | 如 `calendar.v4.calendar.event.attendee.create` | 目录 schema、固定域名、身份与通用审批 |

存量业务 `arguments` 直接使用原业务 schema；SDK API `arguments` 使用 `{path, params, data}`。业务注册表自动取原工具的 schema 与执行函数，并在统一入口验证参数、可信调用者、撤回状态；没有另写一套日历、表格或消息实现。

模型只暴露四个办公入口：检索、调用、近期日程、能力清单。原工具导出与 Mastra 注册作为旧客户端/工作流兼容层保留，调用的是同一实现。能力清单另外枚举全部业务操作，避免收敛工具后功能显示缺失。

`managed-operations.ts` 定义已有业务流程拥有的原始写接口。目录详情给出 `managedOperationId`，直接调用这些原始写接口会引导使用对应业务操作，避免出现两条不同的校验路径。该策略用于日历/任务、文档创建/编辑、多维表格新增/更新、云盘创建/复制、群聊消息发送，及未来目录中的旧版 Sheets 写接口。其他未覆盖的 API 不受此路由限制。

删除了工具包装层重复的日程日期校正；校正统一保留在资源操作入口，授权恢复也复用此处。保留 TableService、ReleaseMaterialsService 等必要业务服务，而非将它们误当成冗余 HTTP 代码删除。

## 验证

测试检查所有存量业务 schema 可发现、模型侧没有重复工具；覆盖统一入口下日历/任务/文档/云盘/消息审批、文档/任务/日程/目录读取、表格读取及并发修改拒绝、多维表格字段预检、目标群聊白名单、原始写 API 的受管路由。

最终结果：`npm run typecheck`、`npm test`（22 个测试文件、254 项测试）、`git diff --check` 全部通过。`npm run build` 于北京时间 2026-11-18 12:02 成功，包含 SDK 目录再生成和生成目录依赖安装。

未改动线上已创建资源，未重启网关；内网需更新并重启后验证。通用入口对存量业务沿用原身份与授权流程，不代表本次额外增加了原业务不支持的用户身份或字段。
