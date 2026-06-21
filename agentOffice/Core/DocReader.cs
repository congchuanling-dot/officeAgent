using System;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    public static class DocReader
    {
        public static int SavedSelStart, SavedSelEnd;

        public static DocumentSnapshot Read(Word.Document doc)
        {
            var snap = new DocumentSnapshot();
            try
            {
                int n = Math.Min(doc.Paragraphs.Count, 20);
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        var p = doc.Paragraphs[i];
                        var txt = (p.Range.Text ?? "").Replace("\r", "").Replace("\n", " ").Trim();
                        var style = (p.get_Style() as Word.Style)?.NameLocal ?? "";
                        snap.Paragraphs.Add(new ParaInfo { Index = i - 1, Style = style, Text = txt.Length > 50 ? txt.Substring(0, 50) : txt });
                    }
                    catch { }
                }

                var sel = doc.Application.Selection;
                snap.SelText = (sel.Text ?? "").Trim();
                try { snap.SelStart = sel.Range.Start; snap.SelEnd = sel.Range.End; } catch { }
                SavedSelStart = snap.SelStart; SavedSelEnd = snap.SelEnd;

                int selIdx = 0;
                try { for (int i = 1; i <= doc.Paragraphs.Count; i++) { var r = doc.Paragraphs[i].Range; if (snap.SelStart >= r.Start && snap.SelStart <= r.End) { selIdx = i - 1; break; } } } catch { }
                snap.SelParaIdx = selIdx;
            }
            catch { }
            return snap;
        }
    }
}
