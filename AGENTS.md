# 上海松江区横山乡村振兴改造项目 — Codex 智能体配置

> Codex AI Agent 配置 — 本文件定义了 Codex 在该项目中的行为模式和能力边界

## 项目身份

- **项目名称:** 上海松江区横山乡村振兴改造项目
- **项目类型:** 顶层策划 / 区位研究 / 对标分析 / 招商与运营方案
- **工作目录:** `C:\Users\Lenovo\Documents\shanghai-hengshan-rural-revitalization`
- **默认模型:** deepseek-v4-flash

## 可用能力

### 核心能力

| 能力 | 实现方式 | 状态 |
|---|---|---|
| 文档生成 (Word) | python-docx | ✅ |
| 表格处理 (Excel) | openpyxl + pandas | ✅ |
| 演示文稿 (PPT) | pptxgenjs | ✅ |
| 网页浏览/截图 | Playwright Chromium | ✅ |
| 桌面远程控制 | Computer Use (named pipe) | ✅ |
| AI 编码智能体 | CodeWhale | ✅ |
| 网络代理 | 127.0.0.1:51008 | ✅ |

### 已安装工具索引

| 工具 | 安装位置 | 用途 |
|---|---|---|
| CodeWhale | npm global | DeepSeek 编码智能体终端 |
| Playwright | bundled runtime | 浏览器自动化、截图 |
| python-docx | bundled Python | Word 文档读写 |
| openpyxl | bundled Python | Excel 表格处理 |
| pptxgenjs | bundled Node.js | PPT 演示文稿生成 |
| codex-proxy-switcher-win | ~/.codex/skills/ | 代理启动器 (GUI) |

## 工作流规则

1. **规划先行:** 复杂任务先输出结构化计划，确认后再执行
2. **证据优先:** 工具输出优于猜测，验证是任务的一部分
3. **用户意图优先:** 当前请求优先于历史记忆和既有配置
4. **输出校验:** 生成文件后自动验证内容完整性
5. **错误处理:** DeepSeek 限流/超时时自动降级重试

## 安全边界

- 所有输出文件保存到项目 `output/` 目录
- 不修改系统级配置
- 不访问非授权目录
- 代理配置: 127.0.0.1:51008

---

*由 Codex 自动生成 — 2026-06-09*

---

Use installed `superpowers-*` skills whenever the task matches `obra/superpowers` workflows such as brainstorming, planning, execution, TDD, debugging, code review, git worktree isolation, branch finishing, verification, or skill authoring.

Codex has native skills and subagents. Prefer those primitives directly instead of translating the workflow into long ad-hoc prompts.

Default document output language is Simplified Chinese for plans, specs, reviews, summaries, postmortems, design docs, ADRs, and status updates.

When creating a new durable document-style file without an explicit user-provided path, prefer the repository's documentation directory. Use `docs/` by default, unless the repository already clearly uses another documentation directory such as `doc/`, `spec/`, or `specs/`.

When creating a new document-style file without an explicit user-provided name, prefer a concise Simplified Chinese filename that matches the document's actual purpose, such as `实施计划.md`, `代码评审.md`, `问题排查.md`, `接口设计.md`, `数据结构设计.md`, `表结构设计.md`, `Redis设计.md`, `S3设计.md`, or `字段说明.md`, unless the repository already uses an English naming convention.

If the document is specifically about Redis or S3, prefer literal names like `Redis设计.md` and `S3设计.md` over broader names such as cache design or object storage design.

Write document content in Simplified Chinese and keep it direct, concrete, and easy for Chinese-speaking teammates to read. Prefer plain language over heavy jargon. If a technical term is necessary, keep it accurate and add a brief explanation when that helps readability.

Keep source code, commands, file paths, URLs, logs, stack traces, environment variable names, schema identifiers, and existing English API terms in their original language unless the user explicitly asks for translation.

When the current workspace is already a linked worktree or detached HEAD, do not create another worktree blindly. Run setup and baseline checks first. If branch creation, push, or PR actions are blocked by the host sandbox, commit the work and hand off via the host application's native branch or local-handoff controls instead of claiming the workflow is impossible.

When the user is still exploring requirements, constraints, trade-offs, overall design, or explicitly asks to think first, prefer `superpowers-brainstorming`.

When the user already wants a concrete document deliverable such as an implementation plan, interface design, request/response contract, data structure, table structure, Redis design, S3 design, field descriptions, or an OpenAPI-style skeleton, prefer `superpowers-writing-plans` directly. Only use `superpowers-brainstorming` first if key decisions are still unresolved.

If the user provides a `TODO.md`, backlog file, worklist, or similar task list and asks for a plan, treat that file as the requirement source. Preserve the list order, start from the first actionable item, and only expand forward when the first item is too small to form a reasonable standalone work slice.

If design or planning documents are produced, explicitly capture the relevant data structures, interface contracts, field definitions, validation rules, naming conventions, retention rules, and error cases when they matter to the task.

Respect explicit user language overrides and repository-local documentation conventions.

### 视觉能力

| 能力 | 实现方式 | 状态 |
|---|---|---|
| 截屏识图 / OCR / 视频帧分析 | vision-bridge (自定义 skill) | 等待 API Key 配置 |
