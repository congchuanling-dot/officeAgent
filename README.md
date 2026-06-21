# Word AI 助手

> Word 原生 AI 插件 — 输入自然语言，自动修改文档。基于 VSTO + WPF + DeepSeek v4。

## 快速开始

### 1. 环境

| 依赖 | 说明 |
|------|------|
| Visual Studio 2022 | 安装时勾选「Office/SharePoint 开发」 |
| Word 2016+ 桌面版 | 调试运行 |
| DeepSeek API Key | 已设入环境变量 `DEEPSEEK_API_KEY` |

### 2. 运行

```
双击打开 agentoffice/agentOffice.sln → 按 F5
```

Word 自动启动，右侧出现「🤖 Word AI 助手」面板。

### 3. 使用

在面板输入框打字，**Ctrl+Enter** 发送：

> "把 1.2.2 整体技术架构设计这段字体变成红色"
> "所有标题改成蓝色加粗"

AI 返回操作预览 → 点「✅ 确认」→ 文档自动修改。

## 项目结构

```
officeAgent/
├── README.md
├── docs/dev-guide.md            ← 开发指南（加新操作看这里）
│
└── agentoffice/                 ← 唯一项目
    ├── agentOffice.sln          ← VS 打开这个
    ├── ThisAddIn.cs             ← 插件入口
    ├── Agent.cs                 ← 核心编排
    ├── Models/Models.cs         ← 数据模型
    ├── AI/                      ← DeepSeek 调用 + Prompt
    ├── Core/                    ← Word 交互（读文档/定位/执行）
    ├── UI/                      ← WPF 界面（VM + 控件 + 宿主）
    └── Properties/              ← VS 生成
```

## 当前能力

MVP 阶段实现 **1 个命令**：

| 命令 | 作用 | 示例 |
|------|------|------|
| `setFont` | 修改字体颜色/加粗 | "把标题变成红色加粗" |

定位方式：`find`（搜索文字）、`headings`（所有标题）、`all`（全文）。

后续可自行扩展更多命令，详见 [开发指南](docs/dev-guide.md)。

## API Key

Key 通过 Windows 用户环境变量 `DEEPSEEK_API_KEY` 读取，无需写在配置文件里，不会泄漏到 git。

去 [platform.deepseek.com](https://platform.deepseek.com/) 获取 Key。
