# Word AI 插件 — 开发指南

> 基于 `agentoffice` 项目（VSTO + WPF + DeepSeek v4）

## 环境

| 工具 | 说明 |
|------|------|
| VS 2022 | 需安装「Office/SharePoint 开发」工作负载 |
| .NET Framework 4.7.2 | Windows 自带 |
| Word 桌面版 | 2016 / 2019 / 365 |
| DeepSeek API Key | 设入环境变量 `DEEPSEEK_API_KEY`（已配置） |

启动方式：VS 打开 `agentoffice/agentOffice.sln` → **F5** → Word 自动启动，右侧出现助手面板。

---

## 项目结构

```
agentoffice/
├── agentOffice.sln              ← VS 解决方案（双击打开）
├── agentOffice.csproj           ← VSTO 项目文件
│
├── ThisAddIn.cs                 ← 插件入口，启动时打开 TaskPane
├── Agent.cs                     ← 核心编排（用户输入 → 读文档 → AI → 执行）
│
├── Models/
│   └── Models.cs                ← 全部 DTO：DocumentSnapshot、ExecutionPlan、FontCommand、ChatMessage
│
├── AI/
│   ├── PromptBuilder.cs         ← System Prompt 模板（告诉 DeepSeek 怎么输出）
│   └── DeepSeekClient.cs        ← HTTP 调用 DeepSeek API，解析 JSON 响应
│
├── Core/
│   ├── DocReader.cs             ← 读 Word 文档结构（段落/选区/快照位置）
│   ├── TargetResolver.cs        ← 把 target.type 翻译成 Word.Range 列表
│   └── CommandExecutor.cs       ← 执行命令（当前只有 setFont）
│
├── UI/
│   ├── TaskPaneVM.cs            ← MVVM ViewModel（绑定属性 + 命令）
│   ├── TaskPaneControl.xaml     ← WPF 布局
│   ├── TaskPaneControl.xaml.cs  ← 代码隐藏（事件 + 转换器）
│   └── TaskPaneHost.cs          ← CustomTaskPane 封装（WPF → WinForms 桥接）
│
└── Properties/                  ← VS 生成，无需修改
```

## 调用链

```
UI/TaskPaneVM.Send()
  → Agent.ProcessAsync(userInput)
    → Core/DocReader.Read(doc)            ← 读取文档快照
    → AI/PromptBuilder.Build(snapshot)    ← 构造 System Prompt
    → AI/DeepSeekClient.CallAsync()       ← 调用 DeepSeek → 返回 ExecutionPlan
  → UI 显示预览

UI/TaskPaneVM.Confirm()
  → Agent.ExecuteAsync(plan)
    → foreach cmd: Core/CommandExecutor.Execute()
      → Core/TargetResolver.Resolve()     ← 定位 Word.Range
      → Apply()                           ← 设置颜色/字体
```

## 当前能力（MVP）

支持 **1 个操作**：`setFont` — 修改字体颜色

| 参数 | 说明 |
|------|------|
| color | 颜色，英文或中文（red/红色、blue/蓝色…） |
| bold | 可选，加粗 |
| fontName | 可选，字体名 |
| fontSize | 可选，字号（磅） |

三种定位方式：`find`（搜索文字）、`headings`（所有标题）、`all`（全文）

## 如何加新操作

**示例：加一个 "setParagraph" 支持居中**

1. **`AI/PromptBuilder.cs`** — 在 Prompt 模板里加描述：
```
### setParagraph — 修改段落格式
参数: alignment(Center/Left/Right)
```

2. **`Core/CommandExecutor.cs`** — 在 `Apply()` 的 switch 里加：
```csharp
case "setParagraph":
    var al = S("alignment");
    if (!string.IsNullOrEmpty(al)) range.ParagraphFormat.Alignment = MapAlign(al);
    break;
```

3. 加 `MapAlign()` 函数。

不改 `TargetResolver`，因为定位方式已够用。不改 `Agent.cs`、`TaskPaneVM.cs`、`DeepSeekClient.cs`。

## 加新定位方式

**示例：加一个 "tables" 定位所有表格**

在 `Core/TargetResolver.cs` 的 switch 里加 case：

```csharp
case "tables":
    for (int i = 1; i <= doc.Tables.Count; i++)
        list.Add(doc.Tables[i].Range);
    break;
```

然后在 `PromptBuilder` 里说明这个 type 怎么用。

## 常见问题

- **编译失败**：关闭 Word 进程，清 `obj/` 和 `bin/`，重启 VS
- **AI 返回解析失败**：DeepSeek 偶尔输出非 JSON，`DeepSeekClient` 有 3 次重试
- **选区丢失**：点击确认按钮时选区会跳，`DocReader` 在执行前保存了快照位置

## 目录清理

| 目录 | 说明 |
|------|------|
| `agentoffice/` | 唯一活跃项目 |
| `docs/` | 本文件 |
| 其他文件（如 `WordAIAgent/`） | 已删除，历史旧版 |
