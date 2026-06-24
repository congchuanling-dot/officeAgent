// ============================================================
// 🎯 目标解析器 — 把 AI 的 TargetSpec 翻译成 Word.Range[]
//
// 职责：充当 AI 自然语言和 Word 对象之间的"翻译官"。
//
// AI 返回的 target 是语义描述（"find/标题/全文/段落3/选区"），
// 这个模块把它们转换为 Word 可以操作的 Range 对象。
//
// 每种 type 对应一个定位策略：
//   all        → body.getRange()（整个文档）
//   paragraph  → 按段落索引定位
//   headings   → 遍历所有样式含 "Heading"/"标题" 的段落
//   find       → 利用 Word 内置 search 功能搜索文本
//   selection  → 实时获取用户当前选区
//
// Office.js 注意：
//   - 所有操作必须在 Word.run() 上下文内
//   - search() 返回的 RangeCollection 遍历后 Range 依然可用
// ============================================================

import { TargetSpec } from '../models';

/**
 * 根据 TargetSpec 解析出一个或多个 Word.Range。
 * 必须在 Word.run() 回调内调用。
 *
 * @returns 匹配的 Word Range 数组（可能为空）
 */
export async function resolve(
  context: Word.RequestContext,
  t: TargetSpec
): Promise<Word.Range[]> {
  const list: Word.Range[] = [];

  switch (t.type) {
    // ── all：全文 ──
    case 'all':
      list.push(context.document.body.getRange());
      break;

    // ── paragraph：按索引定位 ──
    // index 是 0-based（与 DocReader 中 ParaInfo.index 对齐）
    case 'paragraph':
      if (
        t.index != null &&
        t.index >= 0
      ) {
        const paragraphs = context.document.body.paragraphs;
        context.load(paragraphs, 'items');
        await context.sync();

        if (t.index < paragraphs.items.length) {
          list.push(paragraphs.items[t.index].getRange());
        }
      }
      break;

    // ── headings：所有标题段落 ──
    // 匹配条件：样式名包含 "Heading"（英文 Word）或 "标题"（中文 Word）
    case 'headings': {
      const paragraphs = context.document.body.paragraphs;
      context.load(paragraphs, 'items');
      await context.sync();

      // 批量加载所有段落的 style 属性
      for (const p of paragraphs.items) {
        context.load(p, ['style']);
      }
      await context.sync();

      for (const p of paragraphs.items) {
        const s = p.style || '';
        if (
          s.toLowerCase().includes('heading') ||
          s.includes('标题')
        ) {
          list.push(p.getRange());
        }
      }
      break;
    }

    // ── find：全文搜索文本 ──
    case 'find':
      if (t.text && t.text.trim()) {
        try {
          const searchResults = context.document.body.search(t.text.trim(), {
            matchCase: false,
            matchWholeWord: false,
          });
          context.load(searchResults, 'items');
          await context.sync();

          for (const range of searchResults.items) {
            list.push(range);
          }
        } catch {
          // 搜索失败时返回空列表
        }
      }
      break;

    // ── selection：用户当前选区 ──
    // Office.js 中直接实时获取，不需要像 C# 那样保存选区快照
    case 'selection':
      list.push(context.document.getSelection());
      break;
  }

  return list;
}
