using agentOffice.Models;

namespace agentOffice.AI
{
    // ============================================================
    // 📝 Prompt 构建器 — 将文档上下文注入 System Prompt
    //
    // 职责：把 DocumentSnapshot 转换为 AI 能理解的"说明书"。
    //
    // 这是决定 AI 输出质量的关键：
    //   - 告诉 AI 它能做什么操作（当前仅 setFont）
    //   - 告诉 AI 如何定位目标（5 种 type + 示例）
    //   - 告诉 AI 返回什么 JSON 格式
    //   - 注入当前文档的结构（段落列表），让 AI "看到"文档
    //
    // 设计要点：
    //   - 用 few-shot 示例教 AI 正确行为（中文指令→JSON）
    //   - 明确规则防止 AI 瞎编（颜色必须用英文、文字段必须用 find）
    //   - 文档内容注入让 AI 能精确匹配用户提到的文字
    // ============================================================

    public static class PromptBuilder
    {
        /// <summary>
        /// 根据文档快照构建 System Prompt。
        ///
        /// 每次调用都会重新构建（因为文档内容可能变了），
        /// 所以不需要缓存。
        /// </summary>
        /// <param name="doc">DocReader 读取的文档快照</param>
        /// <returns>完整的 System Prompt 字符串</returns>
        public static string Build(DocumentSnapshot doc)
        {
            // ── 拼接段落列表（JSON 友好格式）──
            // 格式：  [0] Normal | 这是一段文字...
            // AI 通过这个了解文档结构和内容，判断用户想改哪一段
            var paraLines = "";
            foreach (var p in doc.Paragraphs)
                paraLines += $"  [{p.Index}] {p.Style} | {p.Text}\n";

            // ── System Prompt 模板 ──
            // 使用 verbatim 字符串 ($@"...") 保持可读性
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
{{{{
  ""explanation"": ""简述做什么"",
  ""commands"": [
    {{{{ ""action"": ""setFont"", ""target"": {{{{ ""type"": ""find"", ""text"": ""要搜的文字"" }}}}, ""params"": {{{{ ""color"": ""red"" }}}} }}}}
  ]
}}}}

## 示例
用户: 把""第一章""变成红色
→ {{{{ ""explanation"": ""查找'第一章'并设为红色"", ""commands"": [ {{{{ ""action"": ""setFont"", ""target"": {{{{ ""type"": ""find"", ""text"": ""第一章"" }}}}, ""params"": {{{{ ""color"": ""red"" }}}} }}}}] }}}}

用户: ""1.2.2 整体技术架构设计这段字体改成蓝色""
→ {{{{ ""explanation"": ""查找'1.2.2 整体技术'并设为蓝色"", ""commands"": [ {{{{ ""action"": ""setFont"", ""target"": {{{{ ""type"": ""find"", ""text"": ""1.2.2 整体技术"" }}}}, ""params"": {{{{ ""color"": ""blue"" }}}} }}}}] }}}}

用户: 所有标题加粗变蓝
→ {{{{ ""explanation"": ""所有标题加粗变蓝"", ""commands"": [ {{{{ ""action"": ""setFont"", ""target"": {{{{ ""type"": ""headings"" }}}}, ""params"": {{{{ ""color"": ""blue"", ""bold"": true }}}} }}}}] }}}}

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
