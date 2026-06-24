// ============================================================
// 📖 文档读取器 — 从 Word 提取结构化的文档快照
//
// 职责：将 Word 文档转译为 AI 可以理解的纯数据。
// 为什么只读前 20 段？
//   - 控制 Prompt 长度（每段截断 50 字，20 段约 1KB）
//   - AI 主要需要上下文而非全文，20 段足够定位用户意图
//
// Office.js 与 VSTO 的差异：
//   - 不需要保存选区位置（selection 在执行时实时获取）
//   - 所有操作必须在 Word.run() 内执行
// ============================================================

import { DocumentSnapshot, ParaInfo } from '../models';

/**
 * 读取当前活动文档的结构化快照。
 *
 * 提取内容：
 *   - 前 20 段：每段的索引、样式名、截断文本（≤50 字）
 *   - 当前选区：文本内容 + 所在段落索引
 *
 * 容错设计：整体 try-catch，出错返回空快照。
 */
export async function readDocument(): Promise<DocumentSnapshot> {
  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      context.load(paragraphs, 'items');

      const selection = context.document.getSelection();
      context.load(selection, ['text']);

      await context.sync();

      const totalCount = paragraphs.items.length;
      const count = Math.min(totalCount, 20);

      // 批量加载前 20 段的 text 和 style 属性
      const firstParagraphs = paragraphs.items.slice(0, count);
      for (const p of firstParagraphs) {
        context.load(p, ['text', 'style']);
      }
      await context.sync();

      // ── 构建段落摘要列表 ──
      const paraInfos: ParaInfo[] = [];
      for (const p of firstParagraphs) {
        try {
          const txt = (p.text || '')
            .replace(/\r/g, '')
            .replace(/\n/g, ' ')
            .trim();
          paraInfos.push({
            index: paraInfos.length, // 0-based index
            style: p.style || '',
            text: txt.length > 50 ? txt.substring(0, 50) : txt,
          });
        } catch {
          // 单段失败不中断整体
        }
      }

      // ── 选区信息 ──
      const selText = (selection.text || '').trim();

      // ── 选区所在段落索引 ──
      // Office.js 中判断选区所在段落比较复杂，置为 -1 让 AI 用 selection 类型处理
      const selParaIdx = -1;

      return {
        paragraphs: paraInfos,
        selText,
        selParaIdx,
      };
    });
  } catch {
    // 整体异常时返回空快照
    return { paragraphs: [], selText: '', selParaIdx: -1 };
  }
}
