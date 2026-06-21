using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    public static class TargetResolver
    {
        public static List<Word.Range> Resolve(Word.Document doc, TargetSpec t)
        {
            var list = new List<Word.Range>();
            switch (t.Type)
            {
                case "all":
                    list.Add(doc.Content);
                    break;
                case "paragraph":
                    if (t.Index != null && t.Index >= 0 && t.Index < doc.Paragraphs.Count)
                        list.Add(doc.Paragraphs[t.Index.Value + 1].Range);
                    break;
                case "headings":
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        try
                        {
                            var p = doc.Paragraphs[i];
                            var s = (p.get_Style() as Word.Style)?.NameLocal ?? "";
                            if (s.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                s.IndexOf("标题", StringComparison.OrdinalIgnoreCase) >= 0)
                                list.Add(p.Range);
                        }
                        catch { }
                    }
                    break;
                case "find":
                    if (!string.IsNullOrEmpty(t.Text))
                    {
                        var rng = doc.Content;
                        rng.Find.Text = t.Text;
                        rng.Find.Forward = true;
                        rng.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        while (rng.Find.Execute())
                        {
                            var dup = doc.Application.Selection.Range.Duplicate as Word.Range;
                            if (dup != null) list.Add(dup);
                        }
                    }
                    break;
                case "selection":
                    if (DocReader.SavedSelStart > 0 && DocReader.SavedSelEnd > DocReader.SavedSelStart)
                        list.Add(doc.Range(DocReader.SavedSelStart, DocReader.SavedSelEnd));
                    else
                        list.Add(doc.Application.Selection.Range);
                    break;
            }
            return list;
        }
    }
}
