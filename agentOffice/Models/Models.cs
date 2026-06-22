using System.Collections.Generic;

namespace agentOffice.Models
{
    // ============================================================
    // 📦 数据模型层 — 整个项目的"共同语言"
    //
    // 所有类都是纯数据容器（POCO），不含任何逻辑。
    // 理解这 6 个类后，后续所有文件的输入输出就一目了然。
    // ============================================================

    /// <summary>
    /// 📄 文档快照 — 从 Word 提取的结构化数据，供 AI 理解当前文档状态。
    /// 由 DocReader.Read() 生成，传给 PromptBuilder 注入 System Prompt。
    /// </summary>
    public class DocumentSnapshot
    {
        /// <summary>文档前 20 段的摘要（索引 + 样式 + 截断文本）</summary>
        public List<ParaInfo> Paragraphs { get; set; } = new List<ParaInfo>();

        /// <summary>用户当前选中的文字内容</summary>
        public string SelText { get; set; } = "";

        /// <summary>选区所在的段落索引（从 0 开始）</summary>
        public int SelParaIdx { get; set; }

        /// <summary>选区在文档中的起始字符位置（Word Range.Start）</summary>
        public int SelStart { get; set; }

        /// <summary>选区在文档中的结束字符位置（Word Range.End）</summary>
        public int SelEnd { get; set; }
    }

    /// <summary>
    /// 📝 段落摘要 — 不传全文，只传结构信息给 AI：
    ///   - 索引：让 AI 能按段落号定位（对应 paragraph 类型的 target）
    ///   - 样式：判断是正文/标题/列表（标题样式命中 headings 定位）
    ///   - 文本：截断到 50 字，够 AI 做语义匹配即可（find 定位的依据）
    /// </summary>
    public class ParaInfo
    {
        /// <summary>段落索引，从 0 开始，对应 Word.Paragraphs[i+1]</summary>
        public int Index { get; set; }

        /// <summary>段落样式名称（如 "Normal"、"Heading 1"、"标题 1"）</summary>
        public string Style { get; set; } = "";

        /// <summary>段落文本，换行符已清理，超过 50 字截断</summary>
        public string Text { get; set; } = "";
    }

    /// <summary>
    /// 🎯 执行计划 — DeepSeek API 返回的 JSON 反序列化结果。
    /// 包含 AI 的解释说明和一串可执行命令。
    ///
    /// 流程中的位置：
    ///   DeepSeekClient.CallAsync() 返回 → TaskPaneVM 展示预览
    ///   → 用户确认 → Agent.ExecuteAsync() 执行
    /// </summary>
    public class ExecutionPlan
    {
        /// <summary>AI 对人类友好的操作描述，展示在预览消息中</summary>
        public string Explanation { get; set; } = "";

        /// <summary>JSON 解析是否成功（false 时看 Error 字段）</summary>
        public bool IsValid { get; set; }

        /// <summary>解析失败或 API 错误时的错误信息</summary>
        public string Error { get; set; } = "";

        /// <summary>AI 生成的具体操作命令列表，按顺序执行</summary>
        public List<FontCommand> Commands { get; set; } = new List<FontCommand>();
    }

    /// <summary>
    /// ✏️ 格式命令 — AI 指定 "对什么东西、做什么操作、用什么参数"。
    ///
    /// 示例解构：
    ///   用户说："把标题改成红色加粗"
    ///   → Action = "setFont"
    ///   → Target = { Type="headings" }
    ///   → Params = { {"color":"red"}, {"bold":true} }
    ///
    /// 当前只支持 Action="setFont"，后续可扩展更多 action 类型。
    /// </summary>
    public class FontCommand
    {
        /// <summary>动作类型，当前仅 "setFont"，未来可扩展 "insertText"、"setParagraph" 等</summary>
        public string Action { get; set; } = "";

        /// <summary>操作目标：对 Word 文档的哪个范围执行动作</summary>
        public TargetSpec Target { get; set; } = new TargetSpec();

        /// <summary>操作参数，键值对形式。setFont 支持：color, bold, italic, fontName, fontSize</summary>
        public Dictionary<string, object?> Params { get; set; } = new Dictionary<string, object?>();
    }

    /// <summary>
    /// 🎯 目标定位 — 描述 "要修改文档的哪一部分"。
    /// 由 TargetResolver.Resolve() 翻译为具体的 Word.Range 列表。
    ///
    /// Type 的 5 种取值（对应 TargetResolver 中的 switch 分支）：
    ///   "all"       — 全文范围
    ///   "headings"  — 所有标题段落（样式含 "Heading" 或 "标题"）
    ///   "find"      — 在全文搜索 text 字段指定的字符串
    ///   "paragraph" — 按 Index 字段定位特定段落
    ///   "selection" — 用户在 Word 中的当前选区
    /// </summary>
    public class TargetSpec
    {
        /// <summary>
        /// 定位方式：all | headings | find | paragraph | selection
        /// PromptBuilder 的 System Prompt 中明确列出了这 5 种类型的用法
        /// </summary>
        public string Type { get; set; } = "";

        /// <summary>find 类型时：要搜索的文本片段（6-15 字，模糊匹配）</summary>
        public string? Text { get; set; }

        /// <summary>paragraph 类型时：段落索引（从 0 开始）</summary>
        public int? Index { get; set; }
    }

    /// <summary>
    /// 💬 聊天消息 — UI 面板中一条气泡的数据。
    /// 角色区分 + 内容 + 可能挂载一个待确认的执行计划。
    ///
    /// PendingPlan 的设计意图：
    ///   AI 返回命令后不立即执行，而是挂在消息上等用户点「确认」。
    ///   点确认 → TaskPaneVM.Confirm() 取出 PendingPlan 执行。
    ///   点取消 → 清空 PendingPlan，消息后标注（已取消）。
    /// </summary>
    public class ChatMessage
    {
        /// <summary>消息角色标识：👤用户 | 🤖AI预览 | ✅成功 | ❌错误</summary>
        public string Role { get; set; } = "";

        /// <summary>消息正文，显示在气泡中</summary>
        public string Content { get; set; } = "";

        /// <summary>待确认的执行计划（仅 AI 回复消息有值），确认后置 null</summary>
        public ExecutionPlan? PendingPlan { get; set; }
    }
}
