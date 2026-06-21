using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.AI;
using agentOffice.Core;
using agentOffice.Models;

namespace agentOffice
{
    /// <summary>核心编排: 用户输入 → AI分析 → 预览 → 执行</summary>
    public class Agent
    {
        readonly Word.Application _app;
        readonly DeepSeekClient _ai;

        public Agent(Word.Application app, DeepSeekClient ai) { _app = app; _ai = ai; }

        public void SetApiKey(string k) => _ai.SetApiKey(k);

        public async Task<ExecutionPlan> ProcessAsync(string userInput)
        {
            var doc = _app.ActiveDocument;
            if (doc == null) return PlanErr("请先打开一个文档");

            var snap = DocReader.Read(doc);
            var prompt = PromptBuilder.Build(snap);
            return await _ai.CallAsync(prompt, userInput);
        }

        public async Task<string> ExecuteAsync(ExecutionPlan plan)
        {
            var doc = _app.ActiveDocument;
            if (doc == null) return "没有打开的文档";

            int ok = 0;
            foreach (var cmd in plan.Commands)
            {
                await Task.Run(() => CommandExecutor.Execute(doc, cmd));
                ok++;
            }
            return $"✅ 完成 {ok} 个操作";
        }

        static ExecutionPlan PlanErr(string msg) => new ExecutionPlan { IsValid = false, Error = msg };
    }
}
