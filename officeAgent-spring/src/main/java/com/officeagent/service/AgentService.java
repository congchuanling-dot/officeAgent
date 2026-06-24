package com.officeagent.service;

import com.officeagent.ai.DeepSeekClient;
import com.officeagent.ai.PromptBuilder;
import com.officeagent.model.AnalyzeRequest;
import com.officeagent.model.ExecutionPlan;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

// ============================================================
// 🎯 核心编排服务 — 与 C# 版 Agent.cs 逻辑完全一致
//
// 职责：
//   Phase 1: analyze()
//     文档快照 → 构建 Prompt → 调 DeepSeek → 返回 ExecutionPlan
//
// 这是你以后扩展功能的主要入口：
//   - 换模型、加 Prompt 规则 → 改 PromptBuilder
//   - 换 API 提供商 → 改 DeepSeekClient
//   - 加新操作类型 → 改 PromptBuilder + 这里
// ============================================================

@Service
public class AgentService {

    private final DeepSeekClient ai;

    public AgentService(
        @Value("${deepseek.api-key}") String apiKey,
        @Value("${deepseek.api-url}") String apiUrl,
        @Value("${deepseek.model}") String model
    ) {
        this.ai = new DeepSeekClient(apiKey, apiUrl, model);
    }

    /**
     * Phase 1: 分析用户输入，返回执行计划（不修改文档）。
     *
     * 内部流程：
     *   1. PromptBuilder.build() → 将文档上下文注入 System Prompt
     *   2. DeepSeekClient.call()  → 发给 AI，拿回 JSON 命令
     */
    public ExecutionPlan analyze(AnalyzeRequest req) {
        // 守卫：必须有文档内容
        if (req.getSnapshot() == null
            || req.getSnapshot().getParagraphs() == null
            || req.getSnapshot().getParagraphs().isEmpty()) {
            return ExecutionPlan.error("文档为空，请先在 Word 中输入内容");
        }

        // 步骤 1：构建 System Prompt（含文档上下文）
        String prompt = PromptBuilder.build(req.getSnapshot());

        // 步骤 2：调用 AI
        return ai.call(prompt, req.getUserInput());
    }
}
