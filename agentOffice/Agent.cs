using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.AI;
using agentOffice.Core;
using agentOffice.Models;

namespace agentOffice
{
    // ============================================================
    // 🎯 核心编排器 — 项目的"大脑"
    //
    // 这是整个数据流的中枢，连接 AI 层和 Core 执行层。
    // 职责只有两个方法，对应一个完整操作的两个阶段：
    //
    //   Phase 1: ProcessAsync()
    //     用户输入自然语言 → 读取文档快照 → 构建 System Prompt
    //     → 调用 DeepSeek API → 返回结构化命令（预览）
    //
    //   Phase 2: ExecuteAsync()
    //     用户确认预览 → 遍历命令 → 逐条执行到 Word 文档
    //
    // 设计原则：预览-确认两阶段，避免 AI 直接误改文档。
    // ============================================================

    public class Agent
    {
        readonly Word.Application _app;   // Word 应用程序对象，操作文档的入口
        readonly DeepSeekClient _ai;      // DeepSeek API 客户端

        public Agent(Word.Application app, DeepSeekClient ai)
        {
            _app = app;
            _ai = ai;
        }

        /// <summary>设置 API Key，由 UI 设置面板传入</summary>
        public void SetApiKey(string k) => _ai.SetApiKey(k);

        // ----------------------------------------------------------
        // Phase 1: 分析 — 自然语言 → ExecutionPlan
        // ----------------------------------------------------------

        /// <summary>
        /// 分析用户输入，返回带命令的执行计划（不修改文档）。
        ///
        /// 内部流程：
        ///   1. 获取当前活动文档
        ///   2. DocReader.Read()       → 提取文档前 20 段 + 选区信息
        ///   3. PromptBuilder.Build()  → 将文档上下文注入 System Prompt
        ///   4. DeepSeekClient.CallAsync() → 发给 AI，拿回 JSON 命令
        ///
        /// 返回的 ExecutionPlan 由 UI 展示为预览消息，等待用户确认。
        /// </summary>
        /// <param name="userInput">用户在聊天框输入的自然语言指令</param>
        /// <returns>AI 解析后的执行计划（IsValid=false 表示失败）</returns>
        public async Task<ExecutionPlan> ProcessAsync(string userInput)
        {
            // 守卫：必须有活动文档才能读写
            var doc = _app.ActiveDocument;
            if (doc == null) return PlanErr("请先打开一个文档");

            // 步骤 1+2：读取文档快照 → 构建 System Prompt（含文档上下文）
            var snap = DocReader.Read(doc);
            var prompt = PromptBuilder.Build(snap);

            // 步骤 3：调用 AI，返回结构化 ExecutionPlan
            return await _ai.CallAsync(prompt, userInput);
        }

        // ----------------------------------------------------------
        // Phase 2: 执行 — ExecutionPlan → Word 文档修改
        // ----------------------------------------------------------

        /// <summary>
        /// 执行已确认的计划，将命令逐条应用到 Word 文档。
        ///
        /// 每条命令的执行路径：
        ///   FontCommand → TargetResolver.Resolve() → List&lt;Word.Range&gt;
        ///               → CommandExecutor.Execute() → 修改格式
        ///
        /// Word COM 操作必须在 STA 线程，所以用 Task.Run 包裹。
        /// </summary>
        /// <param name="plan">用户已确认的执行计划</param>
        /// <returns>执行结果描述（成功数 / 错误信息）</returns>
        public async Task<string> ExecuteAsync(ExecutionPlan plan)
        {
            var doc = _app.ActiveDocument;
            if (doc == null) return "没有打开的文档";

            int ok = 0;
            foreach (var cmd in plan.Commands)
            {
                // Task.Run 确保 COM 调用在正确的线程上下文执行
                await Task.Run(() => CommandExecutor.Execute(doc, cmd));
                ok++;
            }
            return $"✅ 完成 {ok} 个操作";
        }

        /// <summary>快速构建一个错误状态的 ExecutionPlan</summary>
        static ExecutionPlan PlanErr(string msg) =>
            new ExecutionPlan { IsValid = false, Error = msg };
    }
}
