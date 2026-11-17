# Calendar and Task Implementation Plan

**Goal:** 完整 CRUD、授权审批和原请求续办。
**Architecture:** 新增 resource-operation 服务作为日历/任务统一入口，持久化具体意图；用现有 actions 表承载授权审批，permission_requests 存续办状态。OAuth 回调只恢复已批准且未撤回的请求。保留旧流程以兼容历史记录。
**Tech Stack:** TypeScript / Zod / Mastra / LibSQL / Feishu OpenAPI.
**Spec:** 00-requirements.md

## Tasks
- [x] 1. actions.ts / client.ts：日历实体、日程删除、用户任务 CRUD 和分页；先测试请求身份、路径和参数。
- [x] 2. resource-operation.ts / resource-permission.ts / user-auth.ts：可信调用者、授权范围、审批 state、一次性回调、读取续办/写审批、取消与恢复；先测试未授权和回调防重放。
- [x] 3. office-tools.ts / gateway.ts / cards.ts / agent.ts：工具注册与提示、授权卡片链接、回调/撤回/重启集成；测试真实网关审批路径。
- [x] 4. README / .env.example / workspace skill：配置与能力说明；全量验证、内联验收、HTML 摘要。

## Verification
`npm test -- src/tests/resource-operation.test.ts src/tests/resource-permission.test.ts`
`npm run typecheck`、`npm test`、`npm run build`

## Review Focus
跨用户确认；权限缺失与资源 ACL 区分；回调重复/过期/撤回；刷新 scope 变化；DELETE 网络超时不重放；日历只读角色和主日历保护。

## Execution
使用 writing-plans / test-driven-development / verification-before-completion；按现有授权范围直接实施。内联多角度评审，不启动子代理，不提交/发布。
