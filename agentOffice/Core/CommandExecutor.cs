using System;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    // ============================================================
    // ✏️ 命令执行器 — 把 FontCommand 真正写到 Word 文档
    //
    // 职责：数据流的最后一站，将 AI 的结构化命令转译为 Word COM 调用。
    //
    // 执行路径：
    //   FontCommand → TargetResolver.Resolve() → List&lt;Range&gt;
    //               → 对每个 Range 逐条应用参数
    //
    // 当前唯一支持的动作：setFont（修改字体格式）
    // 扩展方式：在 switch(cmd.Action) 中添加新 case 即可。
    // ============================================================

    public static class CommandExecutor
    {
        /// <summary>
        /// 执行单条命令：先解析目标 → 再对每个目标应用格式。
        /// </summary>
        /// <param name="doc">Word 文档对象</param>
        /// <param name="cmd">AI 生成的格式命令</param>
        public static void Execute(Word.Document doc, FontCommand cmd)
        {
            var ranges = TargetResolver.Resolve(doc, cmd.Target);
            foreach (var range in ranges)
            {
                Apply(range, cmd);
            }
        }

        /// <summary>
        /// 对单个 Range 应用字体格式修改。
        ///
        /// 参数提取辅助方法：
        ///   S(k) — 提取字符串参数（如 color, fontName）
        ///   F(k) — 提取浮点数参数（如 fontSize，单位磅）
        ///   B(k) — 提取布尔参数（如 bold, italic）
        /// </summary>
        static void Apply(Word.Range range, FontCommand cmd)
        {
            // ── 参数提取辅助函数（局部函数，简洁高效）──
            // 从 Params 字典取值并转为对应类型，失败返回 null
            string? S(string k) =>
                cmd.Params.TryGetValue(k, out var v) ? v?.ToString() : null;

            float? F(string k)
            {
                if (cmd.Params.TryGetValue(k, out var v) && v != null)
                    try { return Convert.ToSingle(v); } catch { }
                return null;
            }

            bool? B(string k)
            {
                if (cmd.Params.TryGetValue(k, out var v) && v != null)
                    try { return Convert.ToBoolean(v); } catch { }
                return null;
            }

            // ── 按 Action 分发 ──
            switch (cmd.Action)
            {
                // ================================================
                // setFont — 修改字体格式
                //
                // 支持参数：
                //   color    — 颜色（英文名或 #RRGGBB，必填）
                //   bold     — 加粗（true/false）
                //   italic   — 斜体（true/false）
                //   fontName — 字体名称（如 "宋体", "Arial"）
                //   fontSize — 字号磅数（如 12, 14）
                //
                // Word Interop 注意：
                //   Font.Bold/Italic 是 int 类型（-1=true, 0=false），不是 bool
                // ================================================
                case "setFont":
                    var c = S("color");
                    if (!string.IsNullOrEmpty(c))
                        range.Font.Color = ParseColor(c);

                    var bold = B("bold");
                    if (bold.HasValue)
                        range.Font.Bold = bold.Value ? -1 : 0;

                    var it = B("italic");
                    if (it.HasValue)
                        range.Font.Italic = it.Value ? -1 : 0;

                    var fn = S("fontName");
                    if (!string.IsNullOrEmpty(fn))
                        range.Font.Name = fn;

                    var fs = F("fontSize");
                    if (fs.HasValue)
                        range.Font.Size = fs.Value;
                    break;

                // 未来扩展入口：
                // case "insertText": ...
                // case "setParagraphSpacing": ...
                // case "setTableStyle": ...
            }
        }

        // ----------------------------------------------------------
        // 🎨 颜色解析 — 中英文颜色名 → Word WdColor 枚举
        // ----------------------------------------------------------

        /// <summary>
        /// 将用户/AI 指定的颜色字符串转为 Word 颜色常量。
        ///
        /// 支持三种格式：
        ///   1. 英文名：red, blue, darkblue, lightblue...
        ///   2. 中文名：红色, 蓝色, 深蓝...
        ///   3. 十六进制：#RRGGBB（BGR 顺序，因 Word WdColor 是 BGR 编码）
        ///
        /// 匹配不到的默认返回 wdColorAutomatic（自动颜色）。
        /// </summary>
        static Word.WdColor ParseColor(string c)
        {
            switch ((c ?? "").Trim().ToLower())
            {
                // ── 基础色 ──
                case "red":    case "红色":     return Word.WdColor.wdColorRed;
                case "blue":   case "蓝色":     return Word.WdColor.wdColorBlue;
                case "green":  case "绿色":     return Word.WdColor.wdColorGreen;
                case "yellow": case "黄色":     return Word.WdColor.wdColorYellow;
                case "orange": case "橙色": case "橘色": return Word.WdColor.wdColorOrange;
                case "pink":   case "粉色": case "粉红": case "粉红色": return Word.WdColor.wdColorPink;
                case "purple": case "紫色":     return Word.WdColor.wdColorViolet;
                case "brown":  case "棕色": case "褐色": return Word.WdColor.wdColorBrown;
                case "gray": case "grey": case "灰色":   return Word.WdColor.wdColorGray50;
                case "white":  case "白色":     return Word.WdColor.wdColorWhite;
                case "black":  case "黑色":     return Word.WdColor.wdColorBlack;

                // ── 衍生色 ──
                case "darkblue":  case "深蓝": case "深蓝色": return Word.WdColor.wdColorDarkBlue;
                case "lightblue": case "浅蓝": case "浅蓝色": return Word.WdColor.wdColorLightBlue;
                case "teal":      case "青色":     return Word.WdColor.wdColorTeal;

                // ── 十六进制 #RRGGBB ──
                // Word WdColor 编码是 0xBBGGRR（BGR 顺序），所以字节要翻转
                default:
                    if (c.StartsWith("#") && c.Length == 7)
                    {
                        try
                        {
                            return (Word.WdColor)(
                                Convert.ToInt32(c.Substring(5, 2), 16) << 16  // R → 高字节
                                | Convert.ToInt32(c.Substring(3, 2), 16) << 8 // G → 中字节
                                | Convert.ToInt32(c.Substring(1, 2), 16)       // B → 低字节
                            );
                        }
                        catch { /* 解析失败回退到自动色 */ }
                    }
                    return Word.WdColor.wdColorAutomatic;
            }
        }
    }
}
