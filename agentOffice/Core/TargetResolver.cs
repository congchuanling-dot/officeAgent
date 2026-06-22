using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    // ============================================================
    // 🎯 目标解析器 — 把 AI 的 TargetSpec 翻译成 Word.Range
    //
    // 职责：充当 AI 自然语言和 Word COM 对象之间的"翻译官"。
    //
    // AI 返回的 target 是语义描述（"find/标题/全文/段落3/选区"），
    // 这个类负责把它们转换为 Word 可以操作的 Range 对象列表。
    //
    // 每种 Type 对应一个定位策略：
    //   all        → doc.Content（整个文档）
    //   paragraph  → 按段落索引定位
    //   headings   → 遍历所有样式含 "Heading"/"标题" 的段落
    //   find       → 利用 Word 内置 Find 功能搜索文本
    //   selection  → 使用保存的选区位置或实时选区
    // ============================================================

    public static class TargetResolver
    {
        /// <summary>
        /// 根据 TargetSpec 解析出一个或多个 Word.Range。
        ///
        /// 为什么返回 List？
        ///   - headings 类型可能匹配多个段落
        ///   - find 类型可能找到多处匹配
        ///   - CommandExecutor 会对每个 Range 分别执行命令
        /// </summary>
        /// <param name="doc">Word 文档对象</param>
        /// <param name="t">AI 指定的目标描述</param>
        /// <returns>匹配的 Word Range 列表（可能为空）</returns>
        public static List<Word.Range> Resolve(Word.Document doc, TargetSpec t)
        {
            var list = new List<Word.Range>();

            switch (t.Type)
            {
                // ── all：全文 ──
                // doc.Content 返回整个文档正文的 Range
                case "all":
                    list.Add(doc.Content);
                    break;

                // ── paragraph：按索引定位 ──
                // Index 是 0-based（与 DocReader 中 ParaInfo.Index 对齐）
                // Word.Paragraphs 是 1-based，所以 +1
                case "paragraph":
                    if (t.Index != null
                        && t.Index >= 0
                        && t.Index < doc.Paragraphs.Count)
                    {
                        list.Add(doc.Paragraphs[t.Index.Value + 1].Range);
                    }
                    break;

                // ── headings：所有标题段落 ──
                // 匹配条件：样式名包含 "Heading"（英文 Word）或 "标题"（中文 Word）
                // 中英文双匹配确保兼容不同语言环境的 Word
                case "headings":
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        try
                        {
                            var p = doc.Paragraphs[i];
                            var s = (p.get_Style() as Word.Style)?.NameLocal ?? "";
                            if (s.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0
                                || s.IndexOf("标题", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                list.Add(p.Range);
                            }
                        }
                        catch { /* 某段样式读取失败则跳过 */ }
                    }
                    break;

                // ── find：全文搜索文本 ──
                // 利用 Word 内置 Find 功能，支持循环查找所有匹配项
                // wdFindStop：找到文档末尾就停止（不从头循环）
                case "find":
                    if (!string.IsNullOrEmpty(t.Text))
                    {
                        var rng = doc.Content;
                        rng.Find.Text = t.Text;
                        rng.Find.Forward = true;           // 向前搜索
                        rng.Find.Wrap = Word.WdFindWrap.wdFindStop;  // 到尾停止

                        // Word Find 循环：每次 Execute() 找到下一处
                        // 注意：Find 会改变 Selection，这里取 Duplicate
                        // 防止多次查找后 Range 引用失效
                        while (rng.Find.Execute())
                        {
                            var dup = doc.Application.Selection.Range.Duplicate as Word.Range;
                            if (dup != null) list.Add(dup);
                        }
                    }
                    break;

                // ── selection：用户当前选区 ──
                // 优先使用保存的选区（DocReader 读取时的快照），
                // 因为在 DocReader 调用后到 TargetResolver 执行期间，
                // 用户的选区可能已经变了。
                case "selection":
                    if (DocReader.SavedSelStart > 0
                        && DocReader.SavedSelEnd > DocReader.SavedSelStart)
                    {
                        // 用保存的位置还原选区 Range
                        list.Add(doc.Range(DocReader.SavedSelStart, DocReader.SavedSelEnd));
                    }
                    else
                    {
                        // 后备方案：取实时选区
                        list.Add(doc.Application.Selection.Range);
                    }
                    break;
            }
            return list;
        }
    }
}
