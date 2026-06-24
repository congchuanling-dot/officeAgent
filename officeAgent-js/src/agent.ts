// ============================================================
// 🎯 核心编排器（TS 薄层）
//
// TS 侧只做三件事：
//   1. 读 Word 文档 → DocumentSnapshot
//   2. 发给 Spring Boot → ExecutionPlan
//   3. 确认后 → 写到 Word 文档
//
// 跟 C# 版 Agent.cs 一样的两个阶段，但 AI 逻辑都委托给 Spring Boot
// ============================================================

import { ExecutionPlan } from './models';
import { analyze } from './apiClient';
import { readDocument } from './core/docReader';
import { execute } from './core/commandExecutor';

export class Agent {
  // ----------------------------------------------------------
  // Phase 1: 分析 — 读文档 → 发 Spring Boot → 拿计划
  // ----------------------------------------------------------
  async processAsync(userInput: string): Promise<ExecutionPlan> {
    // 步骤 1：读 Word 文档快照
    let snapshot;
    try {
      snapshot = await readDocument();
    } catch {
      return { explanation: '', isValid: false, error: '无法读取 Word 文档', commands: [] };
    }

    if (snapshot.paragraphs.length === 0) {
      return { explanation: '', isValid: false, error: '文档为空，请先输入内容', commands: [] };
    }

    // 步骤 2：发给 Spring Boot（它负责 Prompt + DeepSeek）
    return analyze({ userInput, snapshot });
  }

  // ----------------------------------------------------------
  // Phase 2: 执行 — 把命令写到 Word 文档
  // ----------------------------------------------------------
  async executeAsync(plan: ExecutionPlan): Promise<string> {
    try {
      return await Word.run(async (context) => {

        let ok = 0;
        for (const cmd of plan.commands) {
          try {
            await execute(context, cmd);
            ok++;
          } catch (ex: any) {
            console.warn(`命令执行失败: ${ex.message}`, cmd);
          }
        }

        await context.sync();
        return `✅ 完成 ${ok} 个操作`;
      });
    } catch (ex: any) {
      return `❌ 执行失败: ${ex.message || ex}`;
    }
  }
}
