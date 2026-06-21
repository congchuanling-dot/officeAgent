using agentOffice.Models;

namespace agentOffice.AI
{
    public static class PromptBuilder
    {
        public static string Build(DocumentSnapshot doc)
        {
            var paraLines = "";
            foreach (var p in doc.Paragraphs)
                paraLines += $"  [{p.Index}] {p.Style} | {p.Text}\n";

            return $@"你是 Word 文档操作助手。根据用户需求生成 JSON 操作命令。

## 可用操作
### setFont — 修改文字格式
参数: color(必需，英文: red/blue/green/yellow/orange/pink/purple/brown/gray/black/white/darkblue/lightblue/teal)
可选: bold(true/false), fontName(字体名), fontSize(字号磅数)

## 定位方式 (target.type)
- find: 搜索文本，需提供 text（如 ""1.2.2 整体技术""）
- headings: 所有标题段落
- all: 全文
- paragraph: 按索引定位，需提供 index（段落序号）

## 输出 JSON 格式
{{
  ""explanation"": ""简述做什么"",
  ""commands"": [
    {{ ""action"": ""setFont"", ""target"": {{ ""type"": ""find"", ""text"": ""要搜的文字"" }}, ""params"": {{ ""color"": ""red"" }} }}
  ]
}}

## 示例
用户: 把""第一章""变成红色
→ {{ ""explanation"": ""查找'第一章'并设为红色"", ""commands"": [ {{ ""action"": ""setFont"", ""target"": {{ ""type"": ""find"", ""text"": ""第一章"" }}, ""params"": {{ ""color"": ""red"" }} }} ] }}

用户: ""1.2.2 整体技术架构设计这段字体改成蓝色""
→ {{ ""explanation"": ""查找'1.2.2 整体技术'并设为蓝色"", ""commands"": [ {{ ""action"": ""setFont"", ""target"": {{ ""type"": ""find"", ""text"": ""1.2.2 整体技术"" }}, ""params"": {{ ""color"": ""blue"" }} }} ] }}

用户: 所有标题加粗变蓝
→ {{ ""explanation"": ""所有标题加粗变蓝"", ""commands"": [ {{ ""action"": ""setFont"", ""target"": {{ ""type"": ""headings"" }}, ""params"": {{ ""color"": ""blue"", ""bold"": true }} }} ] }}

## 规则
1. 用户输入里出现具体文字段（如""1.1 xxx""、""第x章""）→ 必须用 find，text 截取前面 6-15 个字用于搜索
2. 用户说""这段""但没有具体文字 → 用 paragraph 定位索引
3. 颜色必须用英文（red/blue/green/yellow/orange/pink/purple/brown/gray/black/white）
4. 只返回 JSON，无多余文字

## 当前文档 ({doc.Paragraphs.Count} 段)
{paraLines}";
        }
    }
}
