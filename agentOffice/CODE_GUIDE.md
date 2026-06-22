# 📖 agentOffice 代码阅读指南

> 总共约 650 行业务代码，按以下顺序半小时可通读一遍。

---

## 第一轮 · 数据结构（先懂"语言"）

**先看** [Models/Models.cs](Models/Models.cs) — 只有 50 行，6 个纯数据类，没有逻辑。看懂了它，后面所有文件的输入输出你就全清楚了。

| 模型 | 用途 |
|---|---|
| `DocumentSnapshot` | 文档快照：段落列表 + 选区信息 |
| `ParaInfo` | 单段信息：索引、样式名、文本(截断50字) |
| `ExecutionPlan` | AI 返回的执行计划：解释 + 命令列表 |
| `FontCommand` | 单条命令：动作 + 目标 + 参数 |
| `TargetSpec` | 目标定位：类型(find/headings/all/paragraph/selection) + 定位参数 |
| `ChatMessage` | UI 消息：角色 + 内容 + 挂载的待确认计划 |

---

## 第二轮 · 主流程（串联全局）

### [ThisAddIn.cs](ThisAddIn.cs)
插件唯一入口，只有 40 行。看懂 `Startup → TaskPaneHost.Show()` 这一行就够了。

### [Agent.cs](Agent.cs)
**核心编排器**，45 行。只有两个方法：

- `ProcessAsync` — 读文档 → 拼 Prompt → 调 AI → 返回计划
- `ExecuteAsync` — 拿着计划去改 Word

到这里你就能在脑子里跑通整个流程了。

---

## 第三轮 · 核心模块（分三叉，互不依赖）

看完 Agent 后，按任意顺序看这三个（全是 `static class`，纯函数，无状态）：

| 文件 | 行数 | 干了什么 |
|---|---|---|
| [Core/DocReader.cs](Core/DocReader.cs) | 42 | 读取 Word 文档前 20 段，记录选区位置 |
| [Core/TargetResolver.cs](Core/TargetResolver.cs) | 60 | 把 AI 说的 "find xxx" 翻译成 Word Range 对象 |
| [Core/CommandExecutor.cs](Core/CommandExecutor.cs) | 58 | 拿到 Range 后真正改字体颜色/大小/加粗 |

---

## 第四轮 · AI 层

### [AI/PromptBuilder.cs](AI/PromptBuilder.cs)
54 行，System Prompt 模板。理解 AI 被"教"了哪些指令、返回什么 JSON 格式。

### [AI/DeepSeekClient.cs](AI/DeepSeekClient.cs)
115 行，HTTP 调用 + JSON 解析。重试逻辑、Markdown 代码块剥离也在这里。

---

## 第五轮 · UI 层（最后看，因为最厚）

### [UI/TaskPaneVM.cs](UI/TaskPaneVM.cs)
113 行，ViewModel。`Send()` 和 `Confirm()` 两个方法就是用户交互的完整流程。

### [UI/TaskPaneHost.cs](UI/TaskPaneHost.cs)
46 行，负责把 WPF 控件塞进 Word 侧边栏。

### [UI/TaskPaneControl.xaml](UI/TaskPaneControl.xaml)
纯布局，看一眼截图就能对应出来。

---

## 🗺️ 总路线图

```
Models  →  ThisAddIn  →  Agent  →  DocReader / TargetResolver / CommandExecutor
                                    →  PromptBuilder  →  DeepSeekClient
                                    →  TaskPaneVM  →  TaskPaneHost  →  XAML
```

## 🔄 核心数据流

```
用户输入文字
    │
    ▼
TaskPaneVM.Send()
    │
    ▼
Agent.ProcessAsync(userInput)
    │
    ├── DocReader.Read(doc)           ← 读取文档前20段 + 当前选区
    ├── PromptBuilder.Build(snap)      ← 构建 System Prompt（含文档内容）
    └── DeepSeekClient.CallAsync()     ← HTTP POST → DeepSeek API
             │
             ▼
        ExecutionPlan                  ← AI 返回的 JSON 命令
    │
    ▼
TaskPaneVM 展示预览 → 用户点「确认」
    │
    ▼
Agent.ExecuteAsync(plan)
    │
    ├── TargetResolver.Resolve()       ← 将 target 转为 Word Range
    └── CommandExecutor.Execute()      ← 对 Range 执行字体修改
             │
             ▼
        Word 文档格式修改完成 ✅
```

## 🏗️ 技术栈

| 层 | 技术 | 用途 |
|---|---|---|
| 宿主 | VSTO (.NET Framework 4.7.2) | Word 插件框架 |
| UI | WPF (ElementHost 嵌入 WinForms) | 右侧任务面板 |
| 数据绑定 | MVVM (手写，无框架) | UI ↔ 逻辑解耦 |
| AI 通信 | HttpClient + Newtonsoft.Json | 调用 DeepSeek API |
| 文档操作 | Word Interop (COM) | 读写 Word 文档 |
