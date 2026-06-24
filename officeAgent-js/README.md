# Word AI 助手 — Office.js 侧边栏 (React + TypeScript)

最小 TypeScript 代码，只做必须用 Word API 的事情。所有 AI 逻辑在 Spring Boot 后端。

## 职责

```
React UI                    Office.js API               Spring Boot
───────                     ────────────                ───────────
聊天面板                    读 Word 文档                 构建 Prompt
消息列表                    搜索/定位范围                 调用 DeepSeek
设置面板                    改字体颜色/加粗               解析 JSON
确认/取消                    返回 ExecutionPlan
```

## 项目结构

```
src/
├── taskpane.html           ← HTML 壳（Office.js CDN）
├── taskpane.tsx            ← 入口（React 挂载）
├── App.tsx                 ← React UI 组件
├── agent.ts                ← 编排器（读文档 → 发后端 → 执行）
├── apiClient.ts            ← Spring Boot API 客户端
├── models.ts               ← 共享数据模型
├── core/
│   ├── docReader.ts        ← 读取 Word 文档结构
│   ├── targetResolver.ts   ← 定位 Word Range
│   └── commandExecutor.ts  ← 应用格式到 Word
```

## 快速开始

```bash
cd officeAgent-js
npm install
npm run dev    # 启动开发服务器（https://localhost:3000）
```

## 与 Spring Boot 后端协作

1. 先启动 Spring Boot：`cd ../officeAgent-spring && ./mvnw spring-boot:run`
2. 再启动侧边栏：`npm run dev`
3. 在 Word 中 sideload 插件（`manifest.xml`）
