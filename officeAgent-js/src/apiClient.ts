// ============================================================
// 🌐 Spring Boot API 客户端
//
// 唯一职责：把请求发给 Spring Boot，拿回 ExecutionPlan。
// 所有 AI 逻辑（Prompt 构建、DeepSeek 调用、JSON 解析）都在 Java 侧。
// ============================================================

import { AnalyzeRequest, ExecutionPlan } from './models';

/** Spring Boot 后端地址（开发时指向 localhost:8080） */
const API_BASE = 'http://localhost:8080/api';

/**
 * 发送分析请求到 Spring Boot。
 * Spring Boot 负责：构建 Prompt → 调 DeepSeek → 返回 ExecutionPlan
 */
export async function analyze(req: AnalyzeRequest): Promise<ExecutionPlan> {
  const resp = await fetch(`${API_BASE}/analyze`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req),
  });

  if (!resp.ok) {
    const text = await resp.text();
    return {
      explanation: '',
      isValid: false,
      error: `后端错误 (${resp.status}): ${text}`,
      commands: [],
    };
  }

  return resp.json();
}
