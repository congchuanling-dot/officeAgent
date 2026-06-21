using System.Collections.Generic;

namespace agentOffice.Models
{
    public class DocumentSnapshot
    {
        public List<ParaInfo> Paragraphs { get; set; } = new List<ParaInfo>();
        public string SelText { get; set; } = "";
        public int SelParaIdx { get; set; }
        public int SelStart { get; set; }
        public int SelEnd { get; set; }
    }

    public class ParaInfo
    {
        public int Index { get; set; }
        public string Style { get; set; } = "";
        public string Text { get; set; } = "";
    }

    public class ExecutionPlan
    {
        public string Explanation { get; set; } = "";
        public bool IsValid { get; set; }
        public string Error { get; set; } = "";
        public List<FontCommand> Commands { get; set; } = new List<FontCommand>();
    }

    public class FontCommand
    {
        public string Action { get; set; } = "";
        public TargetSpec Target { get; set; } = new TargetSpec();
        public Dictionary<string, object?> Params { get; set; } = new Dictionary<string, object?>();
    }

    public class TargetSpec
    {
        public string Type { get; set; } = "";
        public string? Text { get; set; }
        public int? Index { get; set; }
    }

    public class ChatMessage
    {
        public string Role { get; set; } = "";
        public string Content { get; set; } = "";
        public ExecutionPlan? PendingPlan { get; set; }
    }
}
