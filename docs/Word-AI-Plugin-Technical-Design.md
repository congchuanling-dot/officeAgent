# Word AI 插件 — 技术方案设计文档

> **版本**：v1.0  
> **日期**：2026-06-21  
> **定位**：Word 原生插件（VSTO Add-in），用户用自然语言描述需求，AI Agent 自动操作 Word 文档

---

## 目录

- [1. 项目背景与目标](#1-项目背景与目标)
- [2. 方案选型](#2-方案选型)
- [3. 系统架构](#3-系统架构)
- [4. 技术栈](#4-技术栈)
- [5. 项目结构](#5-项目结构)
- [6. 核心流程设计](#6-核心流程设计)
- [7. 命令协议设计](#7-命令协议设计)
- [8. Agent 引擎设计](#8-agent-引擎设计)
- [9. 命令执行引擎设计](#9-命令执行引擎设计)
- [10. 用户界面设计](#10-用户界面设计)
- [11. LLM 集成设计](#11-llm-集成设计)
- [12. 撤销与容错设计](#12-撤销与容错设计)
- [13. 开发路线图](#13-开发路线图)
- [14. 关键风险与对策](#14-关键风险与对策)
- [15. 附录](#15-附录)

---

## 1. 项目背景与目标

### 1.1 要解决的问题

用户在日常使用 Word 时，经常需要做大量重复性、繁琐的格式调整操作（如统一标题格式、调整全文字体、插入复杂表格等）。当前只能手动逐项操作，效率低下。

### 1.2 产品目标

用户在 Word 右侧面板中输入自然语言需求（如"把所有标题加粗居中变成蓝色"），AI Agent 自动理解意图，预览将要执行的操作，用户确认后一键完成所有修改。

### 1.3 核心交互

```
用户输入需求 → AI 分析意图 → 展示执行计划 → 用户确认 → 执行修改 → 完成
```

---

## 2. 方案选型

### 2.1 Word 插件技术对比

| 维度 | **VSTO Add-in ✅ 推荐** | Office.js Web Add-in | VBA 宏 |
|------|------------------------|---------------------|--------|
| 开发语言 | C# (.NET Framework 4.8) | HTML/JS/React | VBA |
| 集成方式 | Word 进程内 DLL | 独立 WebView/浏览器 | Word 内置 |
| API 完整度 | ✅ 完整的 Word Object Model | ⚠️ 受限的子集 | ✅ 完整 |
| 性能 | ✅ 进程内调用，零延迟 | ⚠️ 跨进程通信 | ✅ 快 |
| 跨平台 | ❌ 仅 Windows（桌面版 Word） | ✅ 全平台 | ❌ 仅 Windows |
| UI 能力 | ✅ WPF/WinForms 原生控件 | ✅ Web 技术栈 | ❌ 简单弹窗 |
| 分发方式 | ClickOnce / MSI 安装包 | AppSource / 侧载 | 文件分发 |
| 开发复杂度 | 中 | 低 | 低 |
| 适合场景 | 桌面端深度集成 | 跨平台轻量需求 | 简单自动化 |

### 2.2 最终选择

> **选择 VSTO Add-in（C# WPF）**  
> 原因：API 无限制、与 Word 深度融合、原生 UI 体验、适合桌面端深度使用场景。

### 2.3 AI 引擎对比

| 维度 | **Claude API ✅ 推荐** | OpenAI API |
|------|------------------------|------------|
| 结构化输出能力 | 强（System Prompt + JSON Schema） | 强（Function Calling） |
| 中文理解 | 优秀 | 优秀 |
| 成本 | 中等 | 中等 |
| 可用性 | 全球可用 | 国内需代理 |

> **选择 Claude API**（也保留接入 OpenAI 的扩展点）

---

## 3. 系统架构

### 3.1 整体架构图

```
┌──────────────────────────────────────────────────────────────────┐
│                       WINWORD.EXE 进程                            │
│                                                                   │
│  ┌───────────────┐                    ┌────────────────────────┐  │
│  │  Word 文档区域  │                    │  WPF 自定义任务窗格      │  │
│  │               │                    │                        │  │
│  │  ┌─────────┐ │                    │  ┌──────────────────┐  │  │
│  │  │ 文档内容  │ │                    │  │ 💬 输入框        │  │  │
│  │  │         │ │                    │  │ "把标题加粗居中"  │  │  │
│  │  │ 标题一   │ │                    │  └──────────────────┘  │  │
│  │  │ 正文...  │ │                    │                        │  │
│  │  │ 标题二   │ │                    │  📋 执行计划预览       │  │
│  │  │ 正文...  │ │                    │  ├ 找到 2 个标题       │  │
│  │  └─────────┘ │                    │  ├ 设置加粗            │  │
│  │               │                    │  └ 设置居中            │  │
│  │               │                    │                        │  │
│  │               │                    │  [确认执行] [取消]     │  │
│  └───────────────┘                    └────────────────────────┘  │
│                                                                   │
│  ┌─────────────────────────────────────────────────────────────┐  │
│  │                    VSTO Add-in 内部架构                       │  │
│  │                                                              │  │
│  │  ┌─────────────┐   ┌──────────────┐   ┌──────────────────┐  │  │
│  │  │   UI 层      │   │  Agent 引擎   │   │   命令执行层      │  │  │
│  │  │             │   │              │   │                  │  │  │
│  │  │ MainVM     │──▶│ IntentParser │──▶│ CommandExecutor  │  │  │
│  │  │ PreviewVM  │   │ PromptBuilder│   │ (Word Interop)   │  │  │
│  │  │ HistoryVM  │   │ LLMService   │   │ UndoManager      │  │  │
│  │  │ SettingsVM │   │ PlanValidator│   │ SnapshotManager  │  │  │
│  │  └─────────────┘   └──────┬───────┘   └──────────────────┘  │  │
│  │                           │                                   │  │
│  └───────────────────────────┼───────────────────────────────────┘  │
│                              │                                      │
└──────────────────────────────┼──────────────────────────────────────┘
                               │ HTTPS
                    ┌──────────▼──────────┐
                    │   Claude API        │
                    │   (或自建 LLM 代理)   │
                    └─────────────────────┘
```

### 3.2 分层职责

| 层 | 职责 | 不做什么 |
|---|------|---------|
| **UI 层** (WPF) | 展示界面、接收输入、显示预览 | 不直接操作 Word |
| **Agent 引擎** | 意图理解、Prompt 构造、LLM 调用 | 不执行具体操作 |
| **命令执行层** | 翻译命令为 Word COM 调用、撤销管理 | 不理解用户意图 |
| **LLM API** | 接收 Prompt、返回结构化命令 | 不操作 Word |

---

## 4. 技术栈

| 层 | 技术 | 版本/说明 |
|---|------|----------|
| **插件框架** | VSTO (Visual Studio Tools for Office) | Visual Studio 2022 内置 |
| **目标框架** | .NET Framework 4.8 | VSTO 运行时要求 |
| **UI 框架** | WPF (Windows Presentation Foundation) | 原生桌面控件 |
| **架构模式** | MVVM | ViewModel 驱动界面 |
| **Word API** | Microsoft.Office.Interop.Word | NuGet 包，COM 互操作 |
| **MVVM 框架** | CommunityToolkit.Mvvm | 简化 MVVM 代码 |
| **JSON** | System.Text.Json | 内置，高性能 |
| **HTTP** | HttpClient | 调用 LLM API |
| **依赖注入** | Microsoft.Extensions.DependencyInjection | 模块解耦 |
| **日志** | Serilog | 调试 & 问题排查 |
| **打包** | ClickOnce 或 WiX Toolset | 安装分发 |
| **AI 服务** | Claude API (Anthropic) | 可切换 OpenAI |

### 4.1 开发环境

| 工具 | 说明 |
|------|------|
| Visual Studio 2022 | 需勾选「Office/SharePoint 开发」工作负载 |
| .NET Framework 4.8 SDK | Windows 默认自带 |
| Word 2016 / 2019 / 365 (桌面版) | 用于调试 |
| Anthropic API Key | 用于调用 Claude |

---

## 5. 项目结构

```
officeAgent/
│
├── WordAIAgent.sln                              # VS 解决方案文件
│
├── WordAIAgent/                                 # VSTO 主项目
│   ├── Properties/
│   │   ├── AssemblyInfo.cs                      # 程序集信息
│   │   └── Settings.settings                    # 用户配置（API Key 等）
│   │
│   ├── ThisAddIn.cs                             # ★ 插件入口
│   │
│   ├── TaskPane/                                # WPF 任务窗格
│   │   ├── AgentTaskPane.cs                     # 任务窗格封装
│   │   ├── AgentControl.xaml                    # WPF 用户控件（界面布局）
│   │   ├── AgentControl.xaml.cs                 # 界面代码隐藏
│   │   ├── Converters/
│   │   │   └── ValueConverters.cs               # WPF 数据转换器
│   │   └── ViewModels/
│   │       ├── MainViewModel.cs                 # ★ 主 ViewModel
│   │       └── OperationPreviewModel.cs         # 操作预览 ViewModel
│   │
│   ├── Agent/                                   # Agent 核心引擎
│   │   ├── IWordAgent.cs                        # Agent 接口定义
│   │   ├── WordAgent.cs                         # ★ Agent 核心实现
│   │   ├── IntentParser.cs                      # 意图解析器
│   │   ├── CommandPlanner.cs                    # 命令规划器
│   │   └── Prompts/
│   │       ├── SystemPrompt.txt                 # LLM 系统 Prompt 模板
│   │       └── FewShotExamples.cs               # Few-shot 学习示例
│   │
│   ├── Commands/                                # 命令定义
│   │   ├── IWordCommand.cs                      # 命令接口
│   │   ├── WordCommandBase.cs                   # 命令基类
│   │   ├── SetFontCommand.cs                    # 字体格式命令
│   │   ├── SetParagraphCommand.cs               # 段落格式命令
│   │   ├── InsertTableCommand.cs                # 插入表格命令
│   │   ├── FindReplaceCommand.cs                # 查找替换命令
│   │   ├── InsertTextCommand.cs                 # 插入文本命令
│   │   ├── SetStyleCommand.cs                   # 样式命令
│   │   ├── SetPageSetupCommand.cs               # 页面设置命令
│   │   └── DeleteContentCommand.cs              # 删除内容命令
│   │
│   ├── Execution/                               # 命令执行引擎
│   │   ├── CommandExecutor.cs                   # ★ 命令执行器
│   │   ├── CommandValidator.cs                  # 命令参数校验
│   │   ├── TargetResolver.cs                    # 目标范围解析器
│   │   ├── UndoManager.cs                       # 撤销管理器
│   │   └── DocumentSnapshot.cs                  # 文档快照
│   │
│   ├── Services/                                # 服务层
│   │   ├── ILLMService.cs                       # LLM 服务接口
│   │   ├── ClaudeService.cs                     # Claude API 实现
│   │   ├── DocumentContextReader.cs             # 文档上下文读取
│   │   ├── SettingsService.cs                   # 配置管理
│   │   └── LogService.cs                        # 日志服务
│   │
│   ├── Models/                                  # 数据模型 (DTO)
│   │   ├── WordCommand.cs                       # 命令 DTO
│   │   ├── CommandParams.cs                     # 命令参数 DTO
│   │   ├── TargetSpec.cs                        # 目标定位 DTO
│   │   ├── DocumentContext.cs                   # 文档上下文 DTO
│   │   ├── ExecutionPlan.cs                     # 执行计划 DTO
│   │   └── ChatMessage.cs                       # 对话消息 DTO
│   │
│   ├── Resources/
│   │   └── Ribbon.xml                           # Word Ribbon 按钮定义
│   │
│   └── WordAIAgent.csproj                       # 项目文件
│
├── WordAIAgent.Setup/                           # 安装项目（可选）
│   └── WordAIAgent.Setup.wixproj
│
├── docs/
│   ├── Word-AI-Plugin-Technical-Design.md       # 本文档
│   └── User-Guide.md                            # 用户使用说明
│
└── README.md
```

---

## 6. 核心流程设计

### 6.1 主流程（时序图）

```
用户                    UI (WPF)              Agent 引擎            LLM 服务           Word API
 │                         │                      │                    │                  │
 │  输入："把标题加粗居中"  │                      │                    │                  │
 │────────────────────────▶│                      │                    │                  │
 │                         │                      │                    │                  │
 │                         │  ① 发送用户输入       │                    │                  │
 │                         │─────────────────────▶│                    │                  │
 │                         │                      │                    │                  │
 │                         │                      │ ② 读取文档上下文    │                  │
 │                         │                      │───────────────────────────────────────▶│
 │                         │                      │◀────────────────────────────────────────│
 │                         │                      │  (段落数、样式列表)                    │
 │                         │                      │                    │                  │
 │                         │                      │ ③ 构造 Prompt       │                  │
 │                         │                      │  (系统提示 + 文档状  │                  │
 │                         │                      │   态 + 用户需求)     │                  │
 │                         │                      │                    │                  │
 │                         │                      │ ④ 调用 LLM          │                  │
 │                         │                      │───────────────────▶│                  │
 │                         │                      │                    │                  │
 │                         │                      │ ⑤ 返回结构化命令    │                  │
 │                         │                      │◀───────────────────│                  │
 │                         │                      │   JSON:            │                  │
 │                         │                      │   { commands: [...] }                  │
 │                         │                      │                    │                  │
 │                         │                      │ ⑥ 校验命令合法性    │                  │
 │                         │                      │                    │                  │
 │                         │ ⑦ 返回执行计划        │                    │                  │
 │                         │◀─────────────────────│                    │                  │
 │                         │                      │                    │                  │
 │ 显示预览：              │                      │                    │                  │
 │ "找到2个标题，将加粗居中"│                      │                    │                  │
 │◀────────────────────────│                      │                    │                  │
 │                         │                      │                    │                  │
 │ 点击 [确认执行]          │                      │                    │                  │
 │────────────────────────▶│                      │                    │                  │
 │                         │                      │                    │                  │
 │                         │ ⑧ 执行命令列表        │                    │                  │
 │                         │─────────────────────▶│                    │                  │
 │                         │                      │                    │                  │
 │                         │                      │ ⑨ 逐条翻译为 COM   │                  │
 │                         │                      │─────────────────────────────────────────▶│
 │                         │                      │   Range.Font.Bold = true                 │
 │                         │                      │   ParagraphFormat.Alignment = Center     │
 │                         │                      │◀────────────────────────────────────────│
 │                         │                      │                    │                  │
 │                         │ ⑩ 执行完成            │                    │                  │
 │                         │◀─────────────────────│                    │                  │
 │                         │                      │                    │                  │
 │ 显示："✅ 操作完成"      │                      │                    │                  │
 │◀────────────────────────│                      │                    │                  │
 │                         │                      │                    │                  │
```

### 6.2 状态机

```
                    ┌─────────────┐
                    │   空闲      │◀──────────────────────────────┐
                    └──────┬──────┘                               │
                           │ 用户输入                              │
                           ▼                                       │
                    ┌─────────────┐                               │
                    │   分析中...  │                               │
                    └──────┬──────┘                               │
                           │ LLM 返回结果                          │
                           ▼                                       │
                    ┌─────────────┐    用户点 [取消]               │
                    │   预览      │───────────────────────────────┘
                    └──────┬──────┘
                           │ 用户点 [确认执行]
                           ▼
                    ┌─────────────┐
                    │   执行中...  │
                    └──────┬──────┘
                           │ 执行完成
                           ▼
                    ┌─────────────┐    用户点 [撤销]
                    │   完成      │──────────────────▶ 回到「空闲」
                    └─────────────┘
```

---

## 7. 命令协议设计

### 7.1 设计原则

1. **LLM 友好**：JSON Schema 简单明确，few-shot 示例丰富
2. **可扩展**：添加新命令不影响旧命令
3. **可校验**：每个命令有严格参数校验规则
4. **可翻译**：可直接映射到 Word Object Model 调用

### 7.2 核心命令定义

```json
{
  "commands": [
    {
      "action": "setFont",
      "target": { "type": "selection" },
      "params": {
        "bold": true,
        "color": "blue",
        "fontName": "微软雅黑",
        "fontSize": 12,
        "italic": false,
        "underline": false
      }
    }
  ]
}
```

#### 支持的命令清单

| action | 说明 | 示例用户需求 |
|--------|------|-------------|
| `setFont` | 设置字体格式 | "把标题改成红色加粗" |
| `setParagraph` | 设置段落格式 | "所有段落首行缩进两字符" |
| `setStyle` | 应用 Word 样式 | "用标题1样式格式化标题" |
| `insertText` | 插入文本 | "在文档末尾加上'谢谢观看'" |
| `insertTable` | 插入表格 | "插入一个3行5列的表格" |
| `insertPageBreak` | 插入分页 | "在这里分页" |
| `findReplace` | 查找替换 | "把所有的'北京'替换成'上海'" |
| `deleteContent` | 删除内容 | "删除第二段" |
| `setPageSetup` | 页面设置 | "把页边距设为2厘米" |
| `setTableFormat` | 表格格式化 | "表格首行加灰色底色" |
| `insertImage` | 插入图片 | "在文档末尾插入这张图片" |
| `sortParagraphs` | 段落排序 | "按首字母顺序排列" |

#### 目标定位方式

| type | 说明 | 示例 |
|------|------|------|
| `selection` | 当前光标选区 | 用户在文档中选中的内容 |
| `all` | 全文 | 整个文档 |
| `paragraph` | 指定段落索引（从0开始） | 第3段 |
| `paragraphs` | 段落范围 | 第2~5段 |
| `find` | 查找匹配的文本 | 包含"标题"的段落 |
| `style` | 指定样式的段落 | 所有「标题1」样式的段落 |
| `headings` | 所有标题段落 | 一级、二级标题 |
| `last` | 文档末尾 | 最后一段之后 |

### 7.3 完整 Schema 定义

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "title": "WordOperationPlan",
  "type": "object",
  "required": ["explanation", "commands"],
  "properties": {
    "explanation": {
      "type": "string",
      "description": "用中文向用户解释你将做什么"
    },
    "commands": {
      "type": "array",
      "items": { "$ref": "#/$defs/WordCommand" }
    }
  },
  "$defs": {
    "WordCommand": {
      "type": "object",
      "required": ["action", "target"],
      "properties": {
        "action": {
          "enum": [
            "setFont", "setParagraph", "setStyle",
            "insertText", "insertTable", "insertPageBreak",
            "insertImage", "findReplace", "deleteContent",
            "setPageSetup", "setTableFormat", "sortParagraphs"
          ]
        },
        "target": { "$ref": "#/$defs/TargetSpec" },
        "params": { "$ref": "#/$defs/CommandParams" }
      }
    },
    "TargetSpec": {
      "type": "object",
      "required": ["type"],
      "properties": {
        "type": {
          "enum": ["selection", "all", "paragraph", "paragraphs", "find", "style", "headings", "last"]
        },
        "index": { "type": "integer", "description": "paragraph类型：段落索引" },
        "from": { "type": "integer", "description": "paragraphs类型：起始段落索引" },
        "to": { "type": "integer", "description": "paragraphs类型：结束段落索引" },
        "text": { "type": "string", "description": "find类型：要查找的文本" },
        "styleName": { "type": "string", "description": "style类型：样式名称" }
      }
    },
    "CommandParams": {
      "type": "object",
      "properties": {
        "fontName": { "type": "string", "description": "字体名称，如 '微软雅黑'" },
        "fontSize": { "type": "number", "description": "字号（磅）" },
        "bold": { "type": "boolean", "description": "是否加粗" },
        "italic": { "type": "boolean", "description": "是否斜体" },
        "underline": { "type": "boolean", "description": "是否下划线" },
        "color": { "type": "string", "description": "字体颜色，如 'red', 'blue', '#FF0000'" },
        "alignment": { "enum": ["Left", "Center", "Right", "Justify"] },
        "lineSpacing": { "type": "number" },
        "firstLineIndent": { "type": "number" },
        "spaceBefore": { "type": "number" },
        "spaceAfter": { "type": "number" },
        "styleName": { "type": "string" },
        "content": { "type": "string", "description": "insertText：要插入的文本" },
        "rows": { "type": "integer", "description": "insertTable：行数" },
        "cols": { "type": "integer", "description": "insertTable：列数" },
        "tableData": {
          "type": "array", "items": { "type": "array", "items": { "type": "string" } },
          "description": "insertTable：表格数据"
        },
        "findText": { "type": "string", "description": "findReplace：查找文本" },
        "replaceText": { "type": "string", "description": "findReplace：替换文本" },
        "matchCase": { "type": "boolean", "description": "findReplace：区分大小写" },
        "topMargin": { "type": "number" },
        "bottomMargin": { "type": "number" },
        "leftMargin": { "type": "number" },
        "rightMargin": { "type": "number" },
        "orientation": { "enum": ["Portrait", "Landscape"] },
        "imagePath": { "type": "string" },
        "sortOrder": { "enum": ["Ascending", "Descending"] },
        "headerRowStyle": { "type": "boolean" },
        "tableStyle": { "type": "string" }
      }
    }
  }
}
```

---

## 8. Agent 引擎设计

### 8.1 System Prompt 设计

这是最关键的环节，决定了 LLM 输出的质量。核心策略：

1. **明确角色**：你是 Word 文档操作专家
2. **明确能力边界**：列出所有可执行的操作
3. **明确输出格式**：严格的 JSON Schema
4. **Few-shot 示例**：3~5 个典型场景的输入输出

```markdown
## 系统角色

你是一个 Word 文档操作专家 Agent。用户用中文描述他想要的文档效果，
你需要生成一系列 Word 操作命令来达成目标。

## 核心原则

1. 先读取文档上下文（段落结构、样式、当前选区），再制定计划
2. 只执行用户明确要求的操作，不要添加额外步骤
3. 如果用户需求模糊，生成操作前先向用户确认
4. 操作计划应简洁高效，用最少的命令达成目标

## 可执行的操作

### setFont — 设置字体格式
参数：fontName, fontSize, bold, italic, underline, color
定位：selection | all | find(text) | paragraph(index) | style(name)

### setParagraph — 设置段落格式
参数：alignment(Left/Center/Right/Justify), lineSpacing, firstLineIndent, spaceBefore, spaceAfter
定位：同上

### setStyle — 应用样式
参数：styleName（如 "Heading 1", "Heading 2", "Normal", "Title"）
定位：同上

### insertText — 插入文本
参数：content
位置：selection | last | paragraph(index)

### insertTable — 插入表格
参数：rows, cols, tableData(可选)
位置：selection | last

### findReplace — 查找替换
参数：findText, replaceText, matchCase(可选)
范围：all（默认全文）

### deleteContent — 删除内容
定位：selection | paragraph(index) | find(text)

### setPageSetup — 页面设置
参数：topMargin, bottomMargin, leftMargin, rightMargin, orientation

## 输出格式（严格 JSON）

{
  "explanation": "用中文简述将要做什么，1-2句话",
  "commands": [
    {
      "action": "操作名",
      "target": { "type": "定位方式", ... },
      "params": { 参数键值对 }
    }
  ]
}

## 示例

### 示例1：格式修改
用户："把所有标题加粗居中变成蓝色"
{
  "explanation": "我将找到所有标题段落，将其加粗、居中、设为蓝色",
  "commands": [
    {
      "action": "setFont",
      "target": { "type": "headings" },
      "params": { "bold": true, "color": "blue" }
    },
    {
      "action": "setParagraph",
      "target": { "type": "headings" },
      "params": { "alignment": "Center" }
    }
  ]
}

### 示例2：查找替换
用户："把文档里所有的'北京'改成'上海'"
{
  "explanation": "将在全文中查找'北京'并替换为'上海'",
  "commands": [
    {
      "action": "findReplace",
      "target": { "type": "all" },
      "params": { "findText": "北京", "replaceText": "上海" }
    }
  ]
}

### 示例3：插入表格
用户："在文档末尾插入一个3行4列的表格，表头是姓名、年龄、部门、职位"
{
  "explanation": "将在文档末尾插入3行4列表格，含表头",
  "commands": [
    {
      "action": "insertTable",
      "target": { "type": "last" },
      "params": {
        "rows": 3,
        "cols": 4,
        "tableData": [
          ["姓名", "年龄", "部门", "职位"],
          ["", "", "", ""],
          ["", "", "", ""]
        ]
      }
    }
  ]
}

## 当前文档状态

（运行时动态注入）
- 总段落数：10
- 当前选区：第 3 段，"这是一段正文内容..."
- 样式列表：Normal, Heading 1, Heading 2, Title
- 各段落概要：
  [0] 样式=Title       前50字="项目报告"
  [1] 样式=Heading 1   前50字="第一章 背景介绍"
  [2] 样式=Normal      前50字="本文主要讨论..."
  [3] 样式=Normal      前50字="这是一段正文..."      ← 当前选区
  ...
```

### 8.2 Few-Shot 示例策略

在 System Prompt 中预埋 **5-7 个典型场景的示例**，覆盖：

| 场景分类 | 示例 |
|---------|------|
| 字体格式 | 加粗、改颜色、改字号、换字体 |
| 段落格式 | 居中、首行缩进、行距、段前段后 |
| 样式应用 | 应用 Heading 1/2、正文样式 |
| 文本操作 | 插入、删除、查找替换 |
| 表格操作 | 插入表格、填充数据 |
| 页面设置 | 页边距、纸张方向 |

### 8.3 上下文注入策略

发送给 LLM 的文档上下文需要精简但信息充足：

```json
{
  "totalParagraphs": 15,
  "selection": {
    "text": "当前选中的文本内容（截断100字）",
    "paragraphIndex": 3,
    "style": "Normal",
    "font": { "name": "宋体", "size": 12, "bold": false }
  },
  "styles": ["Normal", "Heading 1", "Heading 2", "Title"],
  "paragraphs": [
    { "index": 0, "style": "Title", "preview": "项目报告" },
    { "index": 1, "style": "Heading 1", "preview": "第一章 背景" },
    { "index": 2, "style": "Normal", "preview": "正文内容..." },
    ...
  ],
  "pageSetup": {
    "orientation": "Portrait",
    "topMargin": 2.54,
    "bottomMargin": 2.54
  }
}
```

---

## 9. 命令执行引擎设计

### 9.1 执行器架构

```
CommandExecutor
├── TargetResolver       ← 将 TargetSpec 解析为 Word Range 对象
├── FontFormatter        ← 字体格式操作
├── ParagraphFormatter   ← 段落格式操作  
├── StyleApplier         ← 样式操作
├── ContentInserter      ← 内容插入（文本、表格、分页）
├── ContentModifier      ← 查找替换、删除
├── PageSetupModifier    ← 页面设置
└── TableFormatter       ← 表格格式化
```

### 9.2 核心逻辑（伪代码）

```
class CommandExecutor:
    
    function execute(doc, command):
        range = resolveTarget(doc, command.target)
        
        switch command.action:
            case "setFont":
                apply range command.params:
                    if bold != null:      range.Font.Bold = bold
                    if color != null:     range.Font.Color = toWdColor(color)
                    if fontName != null:  range.Font.Name = fontName
                    if fontSize != null:  range.Font.Size = fontSize
                    if italic != null:    range.Font.Italic = italic
                    
            case "setParagraph":
                apply range.ParagraphFormat command.params:
                    if alignment != null: format.Alignment = toWdAlignment(alignment)
                    if lineSpacing != null: format.LineSpacing = lineSpacing
                    if firstLineIndent != null: format.FirstLineIndent = firstLineIndent
                    
            case "setStyle":
                range.set_Style(command.params.styleName)
                
            case "insertText":
                range.InsertAfter(command.params.content)
                
            case "insertTable":
                table = doc.Tables.Add(range, command.params.rows, command.params.cols)
                if command.params.tableData:
                    fillTableData(table, command.params.tableData)
                    
            case "findReplace":
                find = range.Find
                find.Text = command.params.findText
                find.Replacement.Text = command.params.replaceText
                find.Execute(Replace: wdReplaceAll)
                
            case "deleteContent":
                range.Delete()
                
            case "setPageSetup":
                setup = doc.PageSetup
                apply setup command.params (margins, orientation)
    
    function resolveTarget(doc, target):
        switch target.type:
            case "selection":  return doc.Application.Selection.Range
            case "all":        return doc.Content
            case "paragraph":  return doc.Paragraphs[target.index + 1].Range
            case "paragraphs": return doc.Range(doc.Paragraphs[target.from+1].Range.Start,
                                                doc.Paragraphs[target.to+1].Range.End)
            case "find":       return findRange(doc.Content, target.text)
            case "style":      return findRangesByStyle(doc, target.styleName)
            case "headings":   return findHeadingRanges(doc)
            case "last":       return doc.Content.End - 1
```

### 9.3 目标解析（难点）

这是执行引擎最复杂的部分。Word 中"所有标题段落"没有一个直接的 API，需要组合实现：

```csharp
// 伪代码：找到所有标题段落
List<Range> FindHeadingRanges(Document doc)
{
    var headings = new List<Range>();
    foreach (Paragraph p in doc.Paragraphs)
    {
        // 判断是否是标题样式
        var styleName = ((Style)p.get_Style()).NameLocal;
        if (styleName.Contains("Heading") || 
            styleName.Contains("标题") ||
            p.OutlineLevel != WdOutlineLevel.wdOutlineLevelBodyText)
        {
            headings.Add(p.Range);
        }
    }
    return headings;
}

// 伪代码：找到指定样式的所有段落
List<Range> FindRangesByStyle(Document doc, string styleName)
{
    var ranges = new List<Range>();
    foreach (Paragraph p in doc.Paragraphs)
    {
        if (((Style)p.get_Style()).NameLocal == styleName)
            ranges.Add(p.Range);
    }
    return ranges;
}
```

---

## 10. 用户界面设计

### 10.1 任务窗格布局

```
┌─────────────────────────────────┐
│  🤖 Word AI 助手               │  ← 标题栏，Office 蓝色
├─────────────────────────────────┤
│                                 │
│  ┌───────────────────────────┐  │
│  │ AI: 你好！我是你的 Word    │  │  ← 对话历史区域
│  │ 文档助手。告诉我你想要     │  │     (可滚动)
│  │ 什么效果，我来帮你实现。   │  │
│  └───────────────────────────┘  │
│                                 │
│  ┌───────────────────────────┐  │
│  │ 👤 你：把所有标题加粗居中  │  │
│  └───────────────────────────┘  │
│                                 │
│  ┌───────────────────────────┐  │
│  │ AI：我将执行以下操作：     │  │
│  │                           │  │
│  │ 📋 操作计划：             │  │
│  │ ├─ 找到 2 个标题段落      │  │
│  │ ├─ 设置加粗              │  │
│  │ └─ 设置居中对齐           │  │
│  │                           │  │
│  │ ┌────────┐ ┌──────┐      │  │
│  │ │✅ 确认 │ │❌ 取消│      │  │
│  │ └────────┘ └──────┘      │  │
│  └───────────────────────────┘  │
│                                 │
│  ┌───────────────────────────┐  │
│  │ AI: ✅ 操作完成！已修改   │  │
│  │ 2 个标题段落。            │  │
│  │ [↩ 撤销此次操作]          │  │
│  └───────────────────────────┘  │
│                                 │
├─────────────────────────────────┤
│  ┌───────────────────────────┐  │
│  │ 💬 输入你的需求...        │  │  ← 输入框
│  │                           │  │     Ctrl+Enter 发送
│  │                           │  │
│  └───────────────────────────┘  │
│                                 │
│  [📤 发送] [↩ 撤销] [⚙ 设置] │  ← 底部按钮栏
│                         就绪     │
└─────────────────────────────────┘
```

### 10.2 UI 状态说明

| UI 状态 | 界面表现 | 触发条件 |
|---------|---------|---------|
| 空闲 | 输入框可用，发绿光指示 | 初始状态 / 操作完成后 |
| 分析中 | 输入框禁用，显示 "🤔 AI 分析中..." | 用户点击发送 |
| 预览 | 显示操作计划 + 确认/取消按钮 | LLM 返回结果后 |
| 执行中 | 显示 "⏳ 执行中..." 进度条 | 用户确认后 |
| 完成 | 显示结果 + 撤销按钮 | 执行成功 |
| 出错 | 显示红色错误信息 + 重试按钮 | 执行失败 / LLM 错误 |

---

## 11. LLM 集成设计

### 11.1 API 调用设计

```
HTTP 请求: POST https://api.anthropic.com/v1/messages

请求头:
  x-api-key: {用户配置的 API Key}
  anthropic-version: 2023-06-01
  
请求体:
{
  "model": "claude-sonnet-4-6",
  "max_tokens": 4096,
  "temperature": 0.3,
  "system": "{系统 Prompt 模板 + 文档上下文}",
  "messages": [
    { "role": "user", "content": "{用户输入的自然语言}" }
  ]
}

响应:
{
  "id": "msg_xxx",
  "content": [
    {
      "type": "text",
      "text": "{\n  \"explanation\": \"...\",\n  \"commands\": [...]\n}"
    }
  ]
}

处理流程:
  1. 提取 content[0].text
  2. 从文本中提取 JSON（处理 ```json 代码块包裹）
  3. 对 JSON 做 Schema 校验
  4. 如果校验失败 → 重试（最多 3 次）
```

### 11.2 错误处理策略

| 错误类型 | 处理方式 |
|---------|---------|
| 网络超时 | 显示 "网络异常，请检查网络后重试" |
| API Key 无效 | 引导用户打开设置页面配置 Key |
| JSON 解析失败 | 自动重试一次，重写 Prompt 强调格式 |
| Schema 校验失败 | 重试，附带校验错误信息 |
| 不支持的 action | 拒绝执行，提示 "AI 返回了不支持的操作" |
| Token 超限 | 减少文档上下文内容，分批 |

### 11.3 API Key 安全存储

使用 Windows 凭据管理器（Credential Manager）存储 API Key：

```
存储位置：
  控制面板 → 凭据管理器 → Windows 凭据
  名称：WordAIAgent_ClaudeApiKey
  类型：通用凭据
  
代码读取：
  var credential = new Credential { Target = "WordAIAgent_ClaudeApiKey" };
  credential.Load();
  var apiKey = credential.Password;
```

---

## 12. 撤销与容错设计

### 12.1 两层撤销机制

```
Layer 1: Word 内置 Undo（简单可靠）
  - 每次操作后自动记录到 Word Undo 栈
  - 用户按 Ctrl+Z 即可撤销
  - 无需额外开发

Layer 2: XML 快照撤销（完整还原）
  - 批量操作前保存文档 XML 快照
  - 撤销时用 XML 完整还原
  - 支持跨多次 Word 内置 Undo 的复杂撤销

策略：默认使用 Layer 1，批量操作时使用 Layer 2
```

### 12.2 快照机制

```
执行批量操作前：
  1. snapshot = doc.Content.XML  （获取完整文档 XML）
  2. 存入 UndoStack

用户点撤销：
  1. 弹出栈顶快照
  2. doc.Content.InsertXML(snapshot)  （还原文档）
  3. 提示用户 "已撤销"

快照限制：
  - 最多保留 20 个快照
  - 每个快照最大 50MB
  - 超限时自动清理最旧的快照
```

### 12.3 部分失败处理

```
场景：执行计划包含 5 个命令，第 3 个执行失败

处理策略：
  1. 前 2 个命令已生效
  2. 第 3 个命令捕获异常
  3. 停止执行后续命令
  4. 提示用户：
     "第 3/5 步执行失败：{错误原因}
      前 2 步已执行完成。
      [还原全部] [保留已执行的] [重试失败步骤]"
```

---

## 13. 开发路线图

### Phase 1 — MVP（约 1 周）

```
目标：核心链路跑通，"一句话改格式"可用

├── ① 创建 VSTO 项目，配置 Word Interop 引用
├── ② 实现 WPF 任务窗格
│   ├── 简单输入框 + 消息列表（纯文本）
│   └── 主 ViewModel 绑定
├── ③ 实现 DocumentContextReader（读取文档状态）
├── ④ 实现 ClaudeService（LLM API 调用）
├── ⑤ 实现 System Prompt + 2 个 Few-shot 示例
├── ⑥ 实现 CommandExecutor（支持 5 个核心命令）
│   ├── setFont
│   ├── setParagraph
│   ├── setStyle
│   ├── insertText
│   └── findReplace
├── ⑦ 实现 TargetResolver（支持 5 种定位方式）
│   ├── selection
│   ├── all
│   ├── paragraph
│   ├── find
│   └── headings
├── ⑧ 实现简单预览 + 确认执行流程
└── ⑨ 联调测试
```

### Phase 2 — 增强（约 1 周）

```
目标：功能完善，体验流畅

├── ① 表格操作（insertTable, setTableFormat）
├── ② 页面设置（setPageSetup）
├── ③ 对话历史（多轮上下文）
├── ④ 撤销机制（XML 快照 + Word Undo）
├── ⑤ Ribbon 按钮集成
├── ⑥ 设置页面（API Key 配置）
├── ⑦ 错误处理 & 重试
├── ⑧ 命令校验增强
└── ⑨ Few-shot 示例扩充到 7 个
```

### Phase 3 — 打磨（约 1 周）

```
目标：产品质量化，可分发

├── ① UI 美化（Fluent 风格）
├── ② 操作进度显示
├── ③ 本地规则引擎兜底（离线时基础操作仍可用）
├── ④ 安装包制作（ClickOnce）
├── ⑤ 安装包制作（WiX MSI）
├── ⑥ 完整使用文档
├── ⑦ 异常日志收集（Serilog）
└── ⑧ 发布到 GitHub
```

---

## 14. 关键风险与对策

| 风险 | 概率 | 影响 | 对策 |
|------|------|------|------|
| **LLM 输出不稳定** | 中 | 高 | JSON Schema 校验 + 3 次重试 + Few-shot 示例校准 |
| **COM 调用慢** | 低 | 中 | 批量读取、复用 Range 对象、避免逐字符操作 |
| **Word 卡死** | 低 | 高 | 长时间操作放后台线程 + ProgressBar + 可取消 |
| **API Key 泄露** | 低 | 高 | Windows 凭据管理器加密存储 |
| **大文档性能差** | 中 | 中 | 上下文摘要（每段仅前 50 字）、分批处理 |
| **VSTO 运行时缺失** | 低 | 中 | ClickOnce 自动下载安装 VSTO Runtime |
| **Word 版本兼容** | 中 | 中 | 使用兼容 API（避免 Word 365 独占特性），覆盖 2016+ |
| **LLM 成本过高** | 低 | 中 | 缓存文档上下文、精简 Prompt、可选本地模型 |

---

## 15. 附录

### 15.1 WdColor 颜色映射表（部分）

| 颜色名 | WdColor 常量 | 颜色名 | WdColor 常量 |
|--------|-------------|--------|-------------|
| red | wdColorRed | blue | wdColorBlue |
| green | wdColorGreen | yellow | wdColorYellow |
| black | wdColorBlack | white | wdColorWhite |
| gray | wdColorGray50 | orange | wdColorOrange |
| pink | wdColorPink | brown | wdColorBrown |
| darkblue | wdColorDarkBlue | lightblue | wdColorLightBlue |
| purple | wdColorViolet | teal | wdColorTeal |

### 15.2 对齐方式映射

| LLM 输出 | WdParagraphAlignment |
|----------|---------------------|
| Left | wdAlignParagraphLeft |
| Center | wdAlignParagraphCenter |
| Right | wdAlignParagraphRight |
| Justify | wdAlignParagraphJustify |

### 15.3 关键 NuGet 包

```xml
<!-- WordAIAgent.csproj 中的包引用 -->
<PackageReference Include="Microsoft.Office.Interop.Word" Version="15.0.4797.1004" />
<PackageReference Include="CommunityToolkit.Mvvm" Version="8.2.2" />
<PackageReference Include="Microsoft.Extensions.DependencyInjection" Version="8.0.0" />
<PackageReference Include="Microsoft.Extensions.Http" Version="8.0.0" />
<PackageReference Include="Serilog" Version="3.1.1" />
<PackageReference Include="Serilog.Sinks.File" Version="5.0.0" />
<PackageReference Include="System.Text.Json" Version="8.0.4" />
```

### 15.4 参考资料

- [VSTO Add-in 官方文档](https://learn.microsoft.com/en-us/visualstudio/vsto/create-vsto-add-ins-for-office-by-using-visual-studio)
- [Word Object Model 参考](https://learn.microsoft.com/en-us/office/vba/api/overview/word/object-model)
- [Claude API 文档](https://docs.anthropic.com/en/docs)
- [WPF 开发文档](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
- [CommunityToolkit.Mvvm](https://learn.microsoft.com/en-us/dotnet/communitytoolkit/mvvm/)

---

> **文档维护者**：技术方案组  
> **最后更新**：2026-06-21  
> **下一步**：Phase 1 MVP 开发启动
