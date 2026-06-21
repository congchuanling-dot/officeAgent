using System;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    public static class CommandExecutor
    {
        public static void Execute(Word.Document doc, FontCommand cmd)
        {
            var ranges = TargetResolver.Resolve(doc, cmd.Target);
            foreach (var range in ranges) Apply(range, cmd);
        }

        static void Apply(Word.Range range, FontCommand cmd)
        {
            string? S(string k) => cmd.Params.TryGetValue(k, out var v) ? v?.ToString() : null;
            float? F(string k) { if (cmd.Params.TryGetValue(k, out var v) && v != null) try { return Convert.ToSingle(v); } catch { } return null; }
            bool? B(string k) { if (cmd.Params.TryGetValue(k, out var v) && v != null) try { return Convert.ToBoolean(v); } catch { } return null; }

            switch (cmd.Action)
            {
                case "setFont":
                    var c = S("color");   if (!string.IsNullOrEmpty(c)) range.Font.Color = ParseColor(c);
                    var bold = B("bold");  if (bold.HasValue) range.Font.Bold = bold.Value ? -1 : 0;
                    var it = B("italic");  if (it.HasValue) range.Font.Italic = it.Value ? -1 : 0;
                    var fn = S("fontName"); if (!string.IsNullOrEmpty(fn)) range.Font.Name = fn;
                    var fs = F("fontSize"); if (fs.HasValue) range.Font.Size = fs.Value;
                    break;
            }
        }

        static Word.WdColor ParseColor(string c)
        {
            switch ((c ?? "").Trim().ToLower())
            {
                case "red": case "红色": return Word.WdColor.wdColorRed;
                case "blue": case "蓝色": return Word.WdColor.wdColorBlue;
                case "green": case "绿色": return Word.WdColor.wdColorGreen;
                case "yellow": case "黄色": return Word.WdColor.wdColorYellow;
                case "orange": case "橙色": case "橘色": return Word.WdColor.wdColorOrange;
                case "pink": case "粉色": case "粉红": case "粉红色": return Word.WdColor.wdColorPink;
                case "purple": case "紫色": return Word.WdColor.wdColorViolet;
                case "brown": case "棕色": case "褐色": return Word.WdColor.wdColorBrown;
                case "gray": case "grey": case "灰色": return Word.WdColor.wdColorGray50;
                case "white": case "白色": return Word.WdColor.wdColorWhite;
                case "black": case "黑色": return Word.WdColor.wdColorBlack;
                case "darkblue": case "深蓝": case "深蓝色": return Word.WdColor.wdColorDarkBlue;
                case "lightblue": case "浅蓝": case "浅蓝色": return Word.WdColor.wdColorLightBlue;
                case "teal": case "青色": return Word.WdColor.wdColorTeal;
                default:
                    if (c.StartsWith("#") && c.Length == 7)
                        try { return (Word.WdColor)(Convert.ToInt32(c.Substring(5, 2), 16) << 16 | Convert.ToInt32(c.Substring(3, 2), 16) << 8 | Convert.ToInt32(c.Substring(1, 2), 16)); }
                        catch { }
                    return Word.WdColor.wdColorAutomatic;
            }
        }
    }
}
