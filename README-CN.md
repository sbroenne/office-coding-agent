# Office Coding Agent

[English](README.md) | 中文

一个 Office 插件，将 GitHub Copilot 直接引入 Excel、PowerPoint、Word 和 Outlook —— 完整支持 **[Copilot CLI 插件](https://docs.github.com/en/copilot/reference/cli-plugin-reference)**。安装任意插件后，其代理、技能、提示词和 MCP 服务器会立即显示在任务窗格中。无需 API 密钥，无需配置 —— 只需使用 GitHub 账号登录即可。

使用 React、Tailwind CSS 和 [GitHub Copilot SDK](https://www.npmjs.com/package/@github/copilot-sdk) 构建。架构基于 [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office)。

> **研究项目声明**
>
> 本仓库是一个独立的**研究项目**。它**不**隶属于 Microsoft 或 GitHub，也**未**获得其认可、赞助或以其他方式与 Microsoft 或 GitHub 建立官方关系。

## 工作原理

```
Office 任务窗格 (React)
      ↓ WebSocket (wss://localhost:3000/api/copilot)
Node.js 代理服务器 (src/server.mjs)
      ↓ @github/copilot-sdk (内部管理 CLI 生命周期)
GitHub Copilot API
```

代理服务器使用 `@github/copilot-sdk` 管理 Copilot CLI 生命周期，并通过 WebSocket + JSON-RPC 将其桥接到浏览器任务窗格。工具调用从服务器流回浏览器，由特定于主机的处理程序执行（例如 `Excel.run()`、`PowerPoint.run()`、`Word.run()` 或 Outlook REST API）。

## 功能特性

### 🔌 Copilot CLI 插件支持

该插件是**一流的 Copilot CLI 插件宿主**。安装任意插件后，其内容会自动显示在 UI 中 —— 无需重启，无需配置：

```bash
copilot plugin add <plugin-name>
```

- 插件中的**代理**会显示在**代理选择器**中
- 插件中的**技能**会显示在**技能选择器**中，并作为上下文注入
- 插件中的**提示词**会显示在 **`/` 斜杠命令菜单**中
- 插件中的 **MCP 服务器**会自动连接
- **插件内容自动发现** —— 已安装插件的代理、技能、提示词和 MCP 服务器会自动显示在任务窗格中

### 🤖 Office 中的 AI 聊天

- **GitHub Copilot 认证** —— 使用 GitHub 账号登录一次即可；无需 API 密钥或端点配置
- **VS Code 风格的聊天 UI** —— 与 VS Code 中的 GitHub Copilot 外观和体验完全一致（设计令牌、代码图标、闪烁的思考指示器、分阶段工作框）
- **模型选择器** —— 在支持的 Copilot 模型之间切换（Claude Sonnet、GPT-4.1、Gemini 等）
- **代理选择器** —— 在针对主机优化的代理之间切换（内置 + 来自插件）
- **技能选择器** —— 开启/关闭已安装插件中的上下文技能
- **流式响应** —— 带有 Copilot 风格进度指示器的实时令牌流式传输

### 📊 Office 宿主工具

- **10 个 Excel 工具组** —— 范围、表格、图表、工作表、工作簿、批注、条件格式、数据验证、数据透视表、范围格式 —— 涵盖约 83 个操作
- **24 个 PowerPoint 工具** —— 幻灯片、形状、文本、图片、表格、图表、备注、布局；包括使用 `get_slide_image` 区域裁剪进行溢出检测的视觉 QA
- **35 个 Word 工具** —— 文档、段落、表格、图片、页眉/页脚、样式、批注、节、域、内容控件
- **22 个 Outlook 工具** —— 邮件、日历、联系人、文件夹、附件、类别、搜索、标记、草稿
- **主机路由工具** —— 根据当前 Office 主机自动选择正确的工具集
- **网页获取工具** —— 通过本地服务器代理以避免 CORS 限制

## 前提条件

- [Node.js](https://nodejs.org/) >= 20
- Microsoft Office（Excel、PowerPoint、Word 或 Outlook —— 桌面版或 Microsoft 365 网页版）
- 有效的 **GitHub Copilot** 订阅（个人版、商业版或企业版）
- 已认证的 `@github/copilot` CLI（`gh auth login` 或等效命令）

## 快速开始

**👉 完整的设置说明请参见 [GETTING_STARTED.md](./GETTING_STARTED.md)** —— 包括认证、启动代理服务器、注册插件和旁加载到 Office。

**快速开始**（需要 [Node.js 20+](https://nodejs.org/)、[GitHub CLI](https://cli.github.com/) 和有效的 [GitHub Copilot](https://github.com/features/copilot) 订阅）：

```bash
# 1. 安装依赖
npm install

# 2. 认证 GitHub Copilot（只需一次）
gh auth login

# 3. 注册插件清单 + 信任 SSL 证书
npm run register:win    # Windows
npm run register:mac    # macOS

# 4. 终端 1 —— 启动代理服务器（保持运行）
npm run dev

# 5. 终端 2 —— 旁加载到 Office
npm run start:desktop:excel   # 或 :ppt / :word
```

代理服务器运行在 `https://localhost:3000`，同时处理 Vite 开发服务器 UI 和 Copilot WebSocket 代理。使用插件时必须保持其运行。

有关本地共享文件夹旁加载和暂存清单工作流，请参见 [docs/SIDELOADING.md](./docs/SIDELOADING.md)。

## 发布

正常的开发变更仍通过拉取请求进行，但发布由手动 GitHub Actions **Release** 工作流处理。

工作流：

- 从最新的 git 标签派生下一个版本
- 构建生产包
- 创建并推送 Git 标签
- 发布 GitHub Release 工件

从 Actions 标签页一步运行：

1. 选择版本增量（`patch`、`minor` 或 `major`）
2. 可选地提供 `custom_version`
3. 运行工作流

## 可用脚本

| 脚本                             | 说明                                                           |
| -------------------------------- | --------------------------------------------------------------------- |
| `npm run dev`                    | 启动 Copilot 代理 + Vite 开发服务器（端口 3000）                     |
| `npm run start:prod-server`      | 从 `dist/` 启动生产 HTTPS 服务器                            |
| `npm run start:tray`             | 构建 + 运行 Electron 系统托盘应用                                  |
| `npm run start:tray:desktop`     | 启动托盘应用（如需要）然后旁加载 Excel 桌面版（旧版别名） |
| `npm run start:tray:excel`       | 启动托盘应用（如需要）然后旁加载 Excel 桌面版                |
| `npm run start:tray:ppt`         | 启动托盘应用（如需要）然后旁加载 PowerPoint 桌面版           |
| `npm run start:tray:word`        | 启动托盘应用（如需要）然后旁加载 Word 桌面版                 |
| `npm run stop:tray:desktop`      | 停止桌面旁加载/调试会话和服务器端口 3000              |
| `npm run build:installer`        | 通过 electron-builder 构建桌面安装程序工件                |
| `npm run build:installer:win`    | 构建 Windows 安装程序（NSIS）                                        |
| `npm run build:installer:dir`    | 构建未打包的桌面应用目录                                  |
| `npm run build`                  | 生产构建到 `dist/`                                           |
| `npm run build:dev`              | 开发构建到 `dist/`                                           |
| `npm run start:desktop`          | 旁加载到 Excel 桌面版（旧版别名）                            |
| `npm run start:desktop:excel`    | 旁加载到 Excel 桌面版                                           |
| `npm run start:desktop:ppt`      | 旁加载到 PowerPoint 桌面版                                      |
| `npm run start:desktop:word`     | 旁加载到 Word 桌面版                                            |
| `npm run stop`                   | 停止调试 / 卸载插件                                        |
| `npm run extensions:samples`     | 生成示例 `agents` 和 `skills` ZIP 文件                       |
| `npm run sideload:share:setup`   | 在 Windows 上创建本地共享文件夹目录                         |
| `npm run sideload:share:trust`   | 将本地共享注册为受信任的 Office 目录                        |
| `npm run sideload:share:publish` | 将暂存清单复制到本地共享文件夹                              |
| `npm run sideload:share:cleanup` | 移除本地共享和受信任目录设置                              |
| `npm run register:win`           | 信任证书并为 Word/PPT/Excel 注册清单（Windows）         |
| `npm run unregister:win`         | 移除已注册的清单条目（Windows）                            |
| `npm run register:mac`           | 信任证书并为 Word/PPT/Excel 注册清单（macOS）           |
| `npm run unregister:mac`         | 从 Word/PPT/Excel WEF 文件夹中移除清单（macOS）        |
| `npm run lint`                   | 运行 ESLint                                                            |
| `npm run lint:fix`               | 自动修复 ESLint 问题                                                |
| `npm run format`                 | 使用 Prettier 格式化代码                                             |
| `npm run typecheck`              | 仅类型检查，不输出                                           |
| `npm test`                       | 运行所有 Vitest 测试套件                                                 |
| `npm run test:integration`       | 运行集成测试套件                                            |
| `npm run test:ui`                | 运行 Playwright UI 测试                                               |
| `npm run test:watch`             | 以监视模式运行测试                                          |
| `npm run test:coverage`          | 运行带覆盖率报告的测试                                      |
| `npm run test:e2e`               | 在 Excel 桌面版中运行 E2E 测试                                  |
| `npm run test:e2e:ppt`           | 在 PowerPoint 桌面版中运行 E2E 测试                             |
| `npm run test:e2e:word`          | 在 Word 桌面版中运行 E2E 测试                                   |
| `npm run test:e2e:outlook`       | 在 Outlook 桌面版中运行 E2E 测试                                |
| `npm run test:e2e:all`           | 按顺序运行所有四个 E2E 测试套件                           |
| `npm run validate`               | 验证 `manifests/manifest.dev.xml`                                 |
| `npm run validate:outlook`       | 验证 `manifests/manifest.outlook.dev.xml`                         |

## 测试

本项目使用三个活跃的测试层：

- **集成测试** (`tests/integration/`, Vitest) —— 组件连接、存储、主机/工具路由和实时 Copilot WebSocket 流
- **UI 测试** (`tests-ui/`, Playwright) —— 浏览器任务窗格行为和回归覆盖
- **E2E 测试** (`tests-e2e*`, Mocha) —— 在 Excel、PowerPoint、Word 和 Outlook 桌面版中进行真实 Office 主机验证

本仓库的新工作有意不使用单元测试。

### 运行测试

```bash
# 所有 Vitest 测试
npm test

# 监视模式
npm run test:watch

# 带覆盖率
npm run test:coverage

# 浏览器 UI 测试
npm run test:ui

# E2E 测试（需要 Office 桌面应用）
npm run test:e2e
npm run test:e2e:ppt
npm run test:e2e:word
npm run test:e2e:outlook

# 验证 Office 插件清单
npm run validate
```

集成测试作为默认 `npm test` 套件的一部分运行。

## E2E 测试

项目包含跨所有四个 Office 主机的端到端测试：约 233 个 Excel 测试、约 15 个 PowerPoint 测试、约 14 个 Word 测试和约 8 个 Outlook 测试（需要 Exchange 旁加载批准）。

### 工作原理

1. **Mocha 运行器** (`tests-e2e/runner.test.ts`) 在端口 4201 启动本地测试服务器。
2. 单独的**测试插件**在 `https://localhost:3001` 构建和提供。
3. 使用 `office-addin-debugging` 将测试插件**旁加载到 Excel 桌面版**。
4. 在 Excel 内部，`test-taskpane.ts` 运行 Excel 命令测试并**将结果发送回**测试服务器。
5. Mocha 运行器**接收结果**并进行断言。

### 工具覆盖

| 类别             | 测试数量 |
| -------------------- | ----- |
| 范围工具          | 59    |
| 表格工具          | 19    |
| 图表工具          | 19    |
| 工作表工具          | 25    |
| 工作簿工具       | 18    |
| 批注工具        | 8     |
| 条件格式   | 27    |
| 数据验证      | 21    |
| 数据透视表          | 28    |
| 设置持久化 | 4     |
| AI 往返        | 4     |

### 运行 E2E 测试

```bash
npm run test:e2e
```

此命令：

- 使用带有 `tests-e2e/tsconfig.json` 项目的 `ts-node`
- 启动测试服务器，构建测试插件，旁加载到 Excel
- 等待所有测试结果，然后拆除（关闭 Excel，停止服务器）

> **注意：** E2E 测试需要在机器上安装 Excel 桌面版。它们使用单独的清单 (`tests-e2e/test-manifest.xml`) 及其自己的 GUID，以便可以与开发插件共存。

### 架构

```
┌─────────────────────┐    结果     ┌─────────────────────┐
│  Mocha 运行器       │◄───(端口 4201)──│  测试任务窗格      │
│  (Node.js)          │                 │  (Excel 内部)     │
│                     │  旁加载       │                     │
│  - 启动服务器    │────────────────►│  - 写入范围       │
│  - 构建插件    │                 │  - 创建工作表   │
│  - 断言结果  │                 │  - 管理表格   │
│                     │                 │  - 创建图表   │
└─────────────────────┘                 └─────────────────────┘
        端口 4201                              端口 3001
     (测试服务器)                         (Vite 开发服务器)
```

### 添加新的 E2E 测试

1. 使用 `pass()`/`fail()`/`assert()` 辅助函数在 `tests-e2e/src/test-taskpane.ts` 中添加测试逻辑。
2. 在 `tests-e2e/runner.test.ts` 中添加相应的 Mocha `it()` 代码块，通过 `e2eContext.getResult('your_test_name')` 读取结果。
3. 运行 `npm run test:e2e` 进行验证。

## 聊天架构

插件通过**本地代理服务器**路由消息 —— 浏览器任务窗格由于浏览器安全限制无法直接调用 GitHub Copilot API。

```
useOfficeChat(host)
      ↓ createWebSocketClient(wss://localhost:3000/api/copilot)
BrowserCopilotSession.query({ prompt, tools })
      ↓ SessionEvent 流
assistant.message_delta / tool.* / session.idle
      ↓
ThreadMessage[] → useExternalStoreRuntime
      ↓ wss://localhost:3000/api/copilot
src/server.mjs (Express HTTPS, 端口 3000)
src/copilotProxy.mjs → @github/copilot-sdk → GitHub Copilot API
```

### 代理系统

AI 代理使用**分离式系统提示词**架构：

- **`src/services/ai/BASE_PROMPT.md`** —— 通用基础提示词（进度叙述、呈现选择）
- **`src/services/ai/prompts/*_APP_PROMPT.md`** —— 主机级应用提示词
- **`src/agents/*/AGENT.md`** —— 带有 YAML 前置事项的代理特定说明
- 说明 = `buildSystemPrompt(host) + resolvedAgent.instructions + skillContext`

`agentService` 按主机解析和过滤代理。代理通过前置事项 `hosts` 定位，并可通过 `defaultForHosts` 声明主机默认值。

### 技能和代理

插件为每个 Office 主机提供捆绑代理。额外的技能、提示词、MCP 服务器和额外代理以 **Copilot CLI 插件**形式分发。

#### Copilot CLI 插件

安装任意 Copilot CLI 插件后，其内容会自动显示在插件 UI 中：

```bash
# 安装插件
copilot plugin add <plugin-name>

# 列出已安装插件
copilot plugin list

# 更新插件
copilot plugin update <plugin-name>

# 移除插件
copilot plugin remove <plugin-name>
```

插件文件约定（Copilot CLI 规范要求）：

| 内容类型           | 文件模式               | 说明                                            |
| ---------------------- | -------------------------- | ------------------------------------------------ |
| 代理                  | `agents/<name>.agent.md`   | 带有 `hosts`、`defaultForHosts` 的 YAML 前置事项 |
| 技能                  | `skills/<name>/SKILL.md`   | 带有可选 `hosts` 的 YAML 前置事项           |
| 提示词（斜杠命令） | `prompts/<name>.prompt.md` | 显示在 `/` 斜杠菜单中                        |
| MCP 服务器配置      | `mcp.json`                 | 服务器自动发现                 |
| 插件级代理     | `agents/AGENT.md`          | 使用插件名称作为代理 ID          |

> **注意：** 扩展名错误的文件（例如 `agents/my-agent.md` 而不是 `agents/my-agent.agent.md`）会被 Copilot CLI 静默忽略。

插件代理会自动与捆绑代理一起显示在 AgentPicker 中，插件技能会显示在 SkillPicker 中。使用上面显示的 `copilot plugin` CLI 命令管理插件。

#### 捆绑代理

捆绑代理随插件一起提供，不可变（在 UI 中只读）。它们位于：

- `src/agents/*/AGENT.md` —— 带有 YAML 前置事项的代理定义

### 关键 Hooks 和组件

- **`useOfficeChat`** —— 创建 `WebSocketCopilotClient`，打开 `BrowserCopilotSession`，将 `SessionEvent` 流映射到 `useExternalStoreRuntime` 的 `ThreadMessage[]`
- **`BrowserCopilotSession.query()`** —— 异步生成器，生成 `SessionEvent` 对象（assistant.message_delta、tool.execution_start、session.idle 等）
- **`getToolsForHost(host)`** —— 返回当前 Office 主机的 `Tool[]`（Copilot SDK 格式）（Excel：约 83 个工具，PowerPoint：24 个，Word：35 个，Outlook：22 个）

状态最小化：`useSettingsStore` (Zustand) 持久化模型/代理/技能配置；聊天状态是临时的。

## UI 布局

任务窗格分为三个区域：

- **ChatHeader** —— SkillPicker、会话历史选择器、权限按钮和新建对话操作
- **ChatPanel** —— 线程/消息流、内联思考指示器、作曲器和带有 AgentPicker + ModelPicker 的输入工具栏
- **App** —— 根外壳，处理 Office 主机检测、主题同步以及连接/会话/权限横幅

## 认证

认证完全由 **GitHub Copilot CLI** (`@github/copilot` 包) 处理。运行一次 `gh auth login`，CLI 会处理 OAuth 令牌管理。无需 API 密钥或 Azure AD 配置。

## 技术栈

- **React 19** —— UI 框架
- **Radix UI + Tailwind CSS v4** —— 任务窗格 UI 组件和样式（通过 `--vscode-*` CSS 自定义属性的 VS Code 设计令牌）
- **GitHub Copilot SDK** (`@github/copilot-sdk`) —— 会话管理、流式事件、工具注册
- **WebSocket + JSON-RPC** (`vscode-jsonrpc`, `ws`) —— 浏览器到代理的传输
- **Express + HTTPS** —— 带有 Vite 开发中间件的本地代理服务器
- **Zustand 5** —— 使用 `OfficeRuntime.storage` 持久化的轻量级状态管理
- **Vite 7** —— 使用 HMR 的打包
- **TypeScript 5** —— 类型安全
- **Vitest** —— 集成测试
- **Playwright** —— 任务窗格流的浏览器 UI 测试
- **Mocha** —— Excel 桌面版内的 E2E 测试（约 233 个测试）
- **Testing Library** —— React 组件测试 (`@testing-library/react`, `user-event`)
- **ESLint + Prettier** —— 代码质量

## 项目历史

本项目经历了两个主要的架构阶段：

### 第一阶段 —— Vercel AI SDK + Azure AI Foundry（2026年2月16日）

Office Coding Agent 的初始版本基于 [Vercel AI SDK](https://ai-sdk.dev/) 构建，使用 [Azure AI Foundry](https://ai.azure.com/) 作为模型后端。它使用了 `@ai-sdk/azure` 和 `@ai-sdk/react` 以及 `@assistant-ui/react-ai-sdk` 作为聊天 UI。用户必须通过设置向导手动配置 API 端点、密钥和模型部署。

### 第二阶段 —— GitHub Copilot SDK（2026年2月20日 – 至今）

受 [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office) 启发 —— 这是由 [Patrick Nikoletich](https://github.com/patniko)、[Steve Sanderson](https://github.com/SteveSandersonMS) 和[贡献者](https://github.com/patniko/github-copilot-office/graphs/contributors)开发的项目 —— 整个 AI 后端在 [PR #25](https://github.com/sbroenne/office-coding-agent/pull/25) 中被替换为 `@github/copilot-sdk`。此次迁移：

- 用 GitHub Copilot SDK 替代了 Vercel AI SDK 和 Azure AI Foundry 后端
- 添加了 Node.js WebSocket 代理服务器（将浏览器任务窗格桥接到 Copilot CLI）
- 移除了设置向导、API 密钥配置和多提供商端点管理
- 将认证简化为通过 `gh auth login` 的单一 GitHub 账号登录

代理服务器架构 (`server.mjs` → `copilotProxy.mjs` → `@github/copilot-sdk`) 和基于 WebSocket 的浏览器传输直接采用了 [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office) 中建立的模式。

## 致谢

- **[patniko/github-copilot-office](https://github.com/patniko/github-copilot-office)** —— 本项目中使用的代理服务器架构、Copilot SDK 集成模式和 WebSocket 传输设计均来自此仓库，由 [Patrick Nikoletich](https://github.com/patniko) 和 [Steve Sanderson](https://github.com/SteveSandersonMS) 开发。他们的工作为第二阶段迁移提供了基础。
- **[@trsdn (Torsten)](https://github.com/trsdn)** 和 **[@urosstojkic](https://github.com/urosstojkic)** —— 贡献了 Word 文档编排器（规划器→工作器模式）、22 个 Outlook 工具、扩展的 PowerPoint 工具（24 个工具）、WorkIQ MCP stdio 集成、主机特定的欢迎提示词、改进的自动滚动以及新技能（Outlook 邮件/日历/起草、Word 格式化/表格/文档构建器、PowerPoint 内容/布局/动画/演示）。最初作为 [PR #33](https://github.com/sbroenne/office-coding-agent/pull/33) 提交，并在 [PR #45](https://github.com/sbroenne/office-coding-agent/pull/45) 中合并。
- **[Vercel AI SDK](https://ai-sdk.dev/)** —— 第一阶段使用的原始 AI 运行时。

## 开发

本项目使用运行在 Copilot CLI 上的 [Squad AI 团队](https://github.com/bradygaster/squad) 进行开发。Squad 通过命名代理协调协作开发，每个代理都有专门的职责。团队组成和代理配置存储在 `.squad/` 中 —— 团队成员包括：Harmony（负责人）、Ellis（PM）、Dylan（前端）、Irving（后端）、Mark（测试员）、Parker（QA）、Scribe 和 Ralph。贡献者可以查看 `.squad/team.md` 以了解当前的团队结构和职责。

## 社区与安全

- [行为准则](./CODE_OF_CONDUCT.md)
- [安全策略](./SECURITY.md)
