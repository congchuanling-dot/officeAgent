package com.officeagent.ai;

import com.officeagent.model.DocumentSnapshot;
import com.officeagent.model.ParaInfo;

// ============================================================
// 📝 Prompt 构建器 — 将文档上下文注入 System Prompt
//
// 与 C# 版 PromptBuilder.cs 逻辑完全一致。
// 决定 AI 输出质量的关键：告诉 AI 能做什么、怎么定位、返回什么。
// ============================================================

public class PromptBuilder {

    /**
     * 根据文档快照构建 System Prompt。
     * 每次调用重新构建（文档内容可能变了）。
     */
    public static String build(DocumentSnapshot doc) {
        // ── 拼接段落列表 ──
        // 格式：  [0] Normal | 这是一段文字...
        StringBuilder paraLines = new StringBuilder();
        for (ParaInfo p : doc.getParagraphs()) {
            paraLines.append(String.format("  [%d] %s | %s\n",
                p.getIndex(), p.getStyle(), p.getText()));
        }

        // ── System Prompt 模板 ──
        return String.format("""
            你是 Word 文档操作助手。根据用户需求生成 JSON 操作命令。

            ## 可用操作
            ### setFont — 修改文字格式
            参数: color(必需，英文: red/blue/green/yellow/orange/pink/purple/brown/gray/black/white/darkblue/lightblue/teal)
            可选: bold(true/false), fontName(字体名), fontSize(字号磅数)

            ## 定位方式 (target.type)
            - find: 搜索文本，需提供 text（如 "1.2.2 整体技术"）
            - headings: 所有标题段落
            - all: 全文
            - paragraph: 按索引定位，需提供 index（段落序号）
            - selection: 用户当前在 Word 中的选区

            ## 输出 JSON 格式
            {
              "explanation": "简述做什么",
              "commands": [
                { "action": "setFont", "target": { "type": "find", "text": "要搜的文字" }, "params": { "color": "red" } }
              ]
            }

            ## 示例
            用户: 把"第一章"变成红色
            → { "explanation": "查找'第一章'并设为红色", "commands": [ { "action": "setFont", "target": { "type": "find", "text": "第一章" }, "params": { "color": "red" } } ] }

            用户: "1.2.2 整体技术架构设计这段字体改成蓝色"
            → { "explanation": "查找'1.2.2 整体技术'并设为蓝色", "commands": [ { "action": "setFont", "target": { "type": "find", "text": "1.2.2 整体技术" }, "params": { "color": "blue" } } ] }

            用户: 所有标题加粗变蓝
            → { "explanation": "所有标题加粗变蓝", "commands": [ { "action": "setFont", "target": { "type": "headings" }, "params": { "color": "blue", "bold": true } } ] }

            ## 规则
            1. 用户输入里出现具体文字段（如"1.1 xxx"、"第x章"）→ 必须用 find，text 截取前面 6-15 个字用于搜索
            2. 用户说"这段"但没有具体文字 → 用 selection 定位
            3. 颜色必须用英文（red/blue/green/yellow/orange/pink/purple/brown/gray/black/white）
            4. 只返回 JSON，无多余文字

            ## 当前文档 (%d 段)
            %s""", doc.getParagraphs().size(), paraLines.toString());
    }
}
