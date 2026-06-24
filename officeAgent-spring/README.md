# Word AI 助手 — Spring Boot 后端

所有 **AI 交互逻辑**都在这里。TypeScript 侧只负责 Word 操作（读文档、改格式）。

## 职责

```
Word 侧边栏 (React/TS)          Spring Boot (Java)
─────────────────────          ──────────────────
读 Word 文档                   构建 System Prompt
改 Word 格式                   调用 DeepSeek API
展示 UI                        解析 JSON → ExecutionPlan
                               重试/错误处理
        POST /api/analyze
        ──────────────────▶
        ◀──────────────────
        ExecutionPlan (JSON)
```

## 项目结构

```
src/main/java/com/officeagent/
├── OfficeAgentApplication.java    ← Spring Boot 启动类
├── controller/
│   └── AgentController.java       ← REST API（/api/analyze）
├── service/
│   └── AgentService.java          ← 核心编排（对应 C# Agent.cs）
├── model/
│   └── Models.java                ← 数据模型（6 个类）
└── ai/
    ├── PromptBuilder.java         ← System Prompt 构建
    └── DeepSeekClient.java        ← DeepSeek HTTP 客户端
```

## 快速开始

### 1. 配置 API Key

```bash
# 方式 1：环境变量（推荐）
set DEEPSEEK_API_KEY=sk-your-key-here

# 方式 2：在 application.yml 中直接写
deepseek:
  api-key: sk-your-key-here
```

### 2. 启动

```bash
cd officeAgent-spring
./mvnw spring-boot:run
```

服务运行在 `http://localhost:8080`

### 3. 测试

```bash
curl -X POST http://localhost:8080/api/analyze \
  -H "Content-Type: application/json" \
  -d '{"userInput":"把标题改成红色","snapshot":{"paragraphs":[{"index":0,"style":"Heading 1","text":"第一章 概述"}],"selText":"","selParaIdx":-1}}'
```

## 与 C# 版的对应关系

| C# 文件 | Java 文件 | 说明 |
|---|---|---|
| `Models/Models.cs` | `model/Models.java` | 6 个数据模型 |
| `AI/PromptBuilder.cs` | `ai/PromptBuilder.java` | System Prompt 模板 |
| `AI/DeepSeekClient.cs` | `ai/DeepSeekClient.java` | HTTP + 重试 + JSON 解析 |
| `Agent.cs` | `service/AgentService.java` | 编排逻辑 |
| — | `controller/AgentController.java` | REST 入口 |
