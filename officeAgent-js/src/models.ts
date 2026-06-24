// ============================================================
// 📦 共享数据模型 — TypeScript 侧
//
// 这些 interface 和 Spring Boot 端的 model 类一一对应。
// TS 侧只负责：读 Word → 发给 Spring Boot → 执行返回的命令
// ============================================================

/** 📝 段落摘要 */
export interface ParaInfo {
  index: number;
  style: string;
  text: string;
}

/** 📄 文档快照 */
export interface DocumentSnapshot {
  paragraphs: ParaInfo[];
  selText: string;
  selParaIdx: number;
}

/** 🎯 执行计划 */
export interface ExecutionPlan {
  explanation: string;
  isValid: boolean;
  error: string;
  commands: FontCommand[];
}

/** ✏️ 格式命令 */
export interface FontCommand {
  action: string;
  target: TargetSpec;
  params: Record<string, unknown>;
}

/** 🎯 目标定位 */
export interface TargetSpec {
  type: string;
  text?: string;
  index?: number;
}

/** 💬 聊天消息 */
export interface ChatMessage {
  role: string;
  content: string;
  pendingPlan?: ExecutionPlan;
}

/** Spring Boot 请求体 */
export interface AnalyzeRequest {
  userInput: string;
  snapshot: DocumentSnapshot;
}
