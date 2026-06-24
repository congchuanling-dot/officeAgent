// ============================================================
// ✏️ 命令执行器 — 把 FontCommand 真正写到 Word 文档
//
// 职责：数据流的最后一站，将 AI 的结构化命令转译为 Word API 调用。
//
// 执行路径：
//   FontCommand → TargetResolver.resolve() → Range[]
//               → 对每个 Range 逐条应用参数
//
// 当前唯一支持的动作：setFont（修改字体格式）
// 扩展方式：在 applyFont() 中添加新 case 即可。
// ============================================================

import { FontCommand } from '../models';
import { resolve } from './targetResolver';

/**
 * 执行单条命令：先解析目标 → 再对每个目标应用格式。
 * 必须在 Word.run() 回调内调用。
 */
export async function execute(
  context: Word.RequestContext,
  cmd: FontCommand
): Promise<void> {
  const ranges = await resolve(context, cmd.target);
  for (const range of ranges) {
    applyFont(range, cmd);
  }
}

/**
 * 对单个 Range 应用字体格式修改。
 * 不调用 context.sync() — 由外层统一 sync，减少网络往返。
 */
function applyFont(range: Word.Range, cmd: FontCommand): void {
  if (cmd.action !== 'setFont') return;

  // ── 辅助：从 params 提取值 ──
  const params = cmd.params;
  const color = stringParam(params, 'color');
  const fontName = stringParam(params, 'fontName');
  const bold = boolParam(params, 'bold');
  const italic = boolParam(params, 'italic');
  const fontSize = floatParam(params, 'fontSize');

  // ── 逐项应用 ──
  if (color) {
    range.font.color = parseColorToJs(color);
  }
  if (bold != null) {
    range.font.bold = bold;
  }
  if (italic != null) {
    range.font.italic = italic;
  }
  if (fontName) {
    range.font.name = fontName;
  }
  if (fontSize != null) {
    range.font.size = fontSize;
  }
}

// ----------------------------------------------------------
// 参数提取辅助函数
// ----------------------------------------------------------

function stringParam(params: Record<string, unknown>, key: string): string | null {
  const v = params[key];
  if (v == null) return null;
  return String(v);
}

function boolParam(params: Record<string, unknown>, key: string): boolean | null {
  const v = params[key];
  if (v == null) return null;
  if (typeof v === 'boolean') return v;
  const s = String(v).toLowerCase();
  if (s === 'true' || s === '1') return true;
  if (s === 'false' || s === '0') return false;
  return null;
}

function floatParam(params: Record<string, unknown>, key: string): number | null {
  const v = params[key];
  if (v == null) return null;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v));
  return isNaN(n) ? null : n;
}

// ----------------------------------------------------------
// 🎨 颜色解析 — 中英文颜色名 → Office.js 颜色值
// ----------------------------------------------------------

const COLOR_MAP: Record<string, string> = {
  // ── 基础色 ──
  'red': '#FF0000',    '红色': '#FF0000',
  'blue': '#0000FF',   '蓝色': '#0000FF',
  'green': '#00FF00',  '绿色': '#00FF00',
  'yellow': '#FFFF00', '黄色': '#FFFF00',
  'orange': '#FFA500', '橙色': '#FFA500', '橘色': '#FFA500',
  'pink': '#FFC0CB',   '粉色': '#FFC0CB', '粉红': '#FFC0CB', '粉红色': '#FFC0CB',
  'purple': '#800080', '紫色': '#800080',
  'brown': '#A52A2A',  '棕色': '#A52A2A', '褐色': '#A52A2A',
  'gray': '#808080',   'grey': '#808080', '灰色': '#808080',
  'white': '#FFFFFF',  '白色': '#FFFFFF',
  'black': '#000000',  '黑色': '#000000',

  // ── 衍生色 ──
  'darkblue': '#00008B',  '深蓝': '#00008B', '深蓝色': '#00008B',
  'lightblue': '#ADD8E6', '浅蓝': '#ADD8E6', '浅蓝色': '#ADD8E6',
  'teal': '#008080',      '青色': '#008080',
};

/**
 * 将用户/AI 指定的颜色字符串转为 Office.js 可接受的颜色值。
 *
 * 支持三种格式：
 *   1. 英文名：red, blue, darkblue, lightblue...
 *   2. 中文名：红色, 蓝色, 深蓝...
 *   3. 十六进制：#RRGGBB（直接透传）
 *
 * 匹配不到的默认返回 'auto'（自动颜色）。
 */
function parseColorToJs(c: string): string {
  const key = (c || '').trim().toLowerCase();

  // 查到映射表
  if (COLOR_MAP[key]) {
    return COLOR_MAP[key];
  }

  // 十六进制直接透传（Office.js 支持 #RRGGBB 格式）
  if (key.startsWith('#') && key.length === 7) {
    return c.trim();
  }

  // 兜底：自动颜色
  return 'auto';
}
