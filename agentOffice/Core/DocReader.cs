using System;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.Models;

namespace agentOffice.Core
{
    // ============================================================
    // 📖 文档读取器 — 从 Word 提取结构化的文档快照
    //
    // 职责：将 Word COM 对象转为 AI 可以理解的纯数据。
    // 为什么只读前 20 段？
    //   - 控制 Prompt 长度（每段截断 50 字，20 段约 1KB）
    //   - AI 主要需要上下文而非全文，20 段足够定位用户意图
    //   - 后续可扩展为滑动窗口或按需加载
    //
    // SavedSelStart/End 是静态字段，用于在 AI 异步返回后
    // 仍能还原选区位置（用户可能在等待期间改变了选区）。
    // ============================================================

    public static class DocReader
    {
        /// <summary>保存选区起始位置（Word 内部字符偏移），供 TargetResolver 定位 "selection" 类型使用</summary>
        public static int SavedSelStart;

        /// <summary>保存选区结束位置</summary>
        public static int SavedSelEnd;

        /// <summary>
        /// 读取当前活动文档的结构化快照。
        ///
        /// 提取内容：
        ///   - 前 20 段：每段的索引、样式名、截断文本（≤50 字）
        ///   - 当前选区：文本内容 + 起始/结束位置 + 所在段落索引
        ///
        /// 容错设计：每个段落独立 try-catch，某段出错不影响其他段。
        /// </summary>
        /// <param name="doc">Word 活动文档对象</param>
        /// <returns>文档快照</returns>
        public static DocumentSnapshot Read(Word.Document doc)
        {
            var snap = new DocumentSnapshot();
            try
            {
                // ── 遍历前 20 段（或文档总段数，取小值）──
                int n = Math.Min(doc.Paragraphs.Count, 20);
                for (int i = 1; i <= n; i++)   // Word.Paragraphs 索引从 1 开始
                {
                    try
                    {
                        var p = doc.Paragraphs[i];

                        // 清理换行符 → 替换为空格，避免 JSON 解析问题
                        var txt = (p.Range.Text ?? "")
                            .Replace("\r", "").Replace("\n", " ").Trim();

                        // 获取样式名（如 "Normal", "Heading 1", "标题 1"）
                        var style = (p.get_Style() as Word.Style)?.NameLocal ?? "";

                        // 截断到 50 字，足够 AI 做语义匹配
                        snap.Paragraphs.Add(new ParaInfo
                        {
                            Index = i - 1,    // 转成 0-based，与外界的 paragraph target.index 对齐
                            Style = style,
                            Text = txt.Length > 50 ? txt.Substring(0, 50) : txt
                        });
                    }
                    catch { /* 单段失败不中断整体 */ }
                }

                // ── 读取选区信息 ──
                var sel = doc.Application.Selection;
                snap.SelText = (sel.Text ?? "").Trim();

                // Range.Start/End 是 Word 内部的绝对字符偏移
                try { snap.SelStart = sel.Range.Start; snap.SelEnd = sel.Range.End; }
                catch { }

                // 保存静态副本（稍后 AI 返回时选区可能已变化）
                SavedSelStart = snap.SelStart;
                SavedSelEnd = snap.SelEnd;

                // ── 定位选区所在段落索引 ──
                int selIdx = 0;
                try
                {
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        var r = doc.Paragraphs[i].Range;
                        // 判断选区起始位置是否落在该段范围内
                        if (snap.SelStart >= r.Start && snap.SelStart <= r.End)
                        {
                            selIdx = i - 1;
                            break;
                        }
                    }
                }
                catch { }
                snap.SelParaIdx = selIdx;
            }
            catch { /* 整体异常时返回空快照 */ }
            return snap;
        }
    }
}
