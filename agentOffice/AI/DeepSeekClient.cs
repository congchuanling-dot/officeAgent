using System;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using agentOffice.Models;

namespace agentOffice.AI
{
    // ============================================================
    // 🤖 DeepSeek API 客户端 — 与 AI 的通信管道
    //
    // 职责：
    //   1. 封装 HTTP 请求（认证、超时、重试）
    //   2. 发送 System Prompt + 用户输入
    //   3. 解析 AI 返回的 JSON → ExecutionPlan
    //
    // 配置来源优先级：环境变量 → 代码默认值 → UI 设置
    // 环境变量（用户级）：
    //   DEEPSEEK_API_KEY — API 密钥
    //   DEEPSEEK_API_URL — 自定义 API 地址
    //   DEEPSEEK_MODEL    — 模型名称
    // ============================================================

    public class DeepSeekClient
    {
        readonly HttpClient _http;

        // ── 配置项（可由环境变量覆盖）──
        string _apiKey = "";                                          // API 密钥
        string _apiUrl = "https://api.deepseek.com/chat/completions"; // API 端点
        string _model = "deepseek-v4-pro";                            // 模型名

        public DeepSeekClient()
        {
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };

            // 从用户环境变量加载配置（优先级高于硬编码默认值）
            var key = Environment.GetEnvironmentVariable(
                "DEEPSEEK_API_KEY", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(key)) _apiKey = key;

            var url = Environment.GetEnvironmentVariable(
                "DEEPSEEK_API_URL", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(url)) _apiUrl = url;

            var mdl = Environment.GetEnvironmentVariable(
                "DEEPSEEK_MODEL", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(mdl)) _model = mdl;
        }

        /// <summary>运行时设置 API Key（UI 设置面板调用）</summary>
        public void SetApiKey(string k) => _apiKey = k;

        /// <summary>
        /// 核心方法：发送消息到 DeepSeek，返回结构化执行计划。
        ///
        /// 调用方：Agent.ProcessAsync()
        ///
        /// 特性：
        ///   - 最多重试 3 次（网络瞬断 / 临时服务器错误）
        ///   - 401 不重试（Key 无效，重试无意义）
        ///   - 自动剥离 Markdown 代码块（AI 有时会用 ```json 包裹）
        ///   - temperature=0.1：低随机性，保证输出格式稳定
        /// </summary>
        /// <param name="systemPrompt">PromptBuilder.Build() 构建的 System Prompt</param>
        /// <param name="userInput">用户在聊天框输入的原始文字</param>
        /// <returns>ExecutionPlan，失败时 IsValid=false 且 Error 有描述</returns>
        public async Task<ExecutionPlan> CallAsync(string systemPrompt, string userInput)
        {
            // 前置检查：没配置 Key 则直接失败
            if (string.IsNullOrWhiteSpace(_apiKey))
                return PlanErr("请先配置 API Key");

            // 构建请求体（Chat Completions 标准格式）
            var body = new
            {
                model = _model,
                max_tokens = 2048,     // 命令 JSON 不会很长，2048 足够
                temperature = 0.1,     // 低温度 = 稳定格式输出
                messages = new[] {
                    new { role = "system", content = systemPrompt },
                    new { role = "user",   content = userInput }
                }
            };

            // ── 重试循环（最多 3 次）──
            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    var req = new HttpRequestMessage(HttpMethod.Post, _apiUrl)
                    {
                        Content = new StringContent(
                            JsonConvert.SerializeObject(body),
                            Encoding.UTF8,
                            "application/json")
                    };
                    req.Headers.Add("Authorization", "Bearer " + _apiKey);

                    var resp = await _http.SendAsync(req);
                    var respText = await resp.Content.ReadAsStringAsync();

                    if (!resp.IsSuccessStatusCode)
                    {
                        // 401：密钥无效，不重试
                        if ((int)resp.StatusCode == 401)
                            return PlanErr("API Key 无效");

                        // 其他错误：等 1 秒重试
                        if (retry < 2) { await Task.Delay(1000); continue; }
                        return PlanErr("API 错误: " + respText);
                    }

                    // 成功：解析 JSON 为 ExecutionPlan
                    return ParseJson(respText);
                }
                catch (TaskCanceledException)
                {
                    // 超时：最后机会不重试
                    if (retry == 2) return PlanErr("请求超时");
                }
                catch (Exception ex)
                {
                    // 其他异常：网络/序列化等
                    if (retry == 2) return PlanErr(ex.Message);
                }
                await Task.Delay(1000);   // 重试前等待
            }
            return PlanErr("未知错误");
        }

        /// <summary>
        /// 解析 AI 返回的原始 JSON → ExecutionPlan。
        ///
        /// 处理流程：
        ///   1. 提取 choices[0].message.content
        ///   2. 剥离可能的 ```json ... ``` Markdown 包裹
        ///   3. 反序列化为 ExecutionPlan（JObject 手动解析，容错性好）
        /// </summary>
        ExecutionPlan ParseJson(string respText)
        {
            try
            {
                // 步骤 1：从 Chat Completions 响应中取 content
                var root = JObject.Parse(respText);
                var text = root["choices"]?[0]?["message"]?["content"]?.ToString();
                if (string.IsNullOrWhiteSpace(text))
                    return PlanErr("AI 返回空内容");

                // 步骤 2：剥离 Markdown 代码块
                // AI 有时会在 JSON 外加 ```json 和 ```，需要去掉
                var t = text.Trim();
                if (t.StartsWith("```"))
                {
                    // 跳过第一行（```json 或 ```）
                    var i = t.IndexOf('\n');
                    t = t.Substring(i < 0 ? 3 : i + 1);

                    // 去掉结尾的 ```
                    i = t.LastIndexOf("```");
                    if (i > 0) t = t.Substring(0, i).Trim();
                }

                // 步骤 3：手动解析 JSON（比直接反序列化更容错）
                var json = JObject.Parse(t);
                var plan = new ExecutionPlan
                {
                    Explanation = json["explanation"]?.ToString() ?? "",
                    IsValid = true
                };

                // 遍历 commands 数组，逐条构建 FontCommand
                foreach (var cmd in json["commands"] ?? new JArray())
                {
                    var fc = new FontCommand
                    {
                        Action = cmd["action"]?.ToString() ?? "",
                        Target = new TargetSpec
                        {
                            Type  = cmd["target"]?["type"]?.ToString() ?? "",
                            Text  = cmd["target"]?["text"]?.ToString(),
                            Index = cmd["target"]?["index"]?.Value<int?>()
                        }
                    };

                    // params 是动态键值对，遍历填充
                    if (cmd["params"] is JObject po)
                    {
                        foreach (var kv in po)
                        {
                            fc.Params[kv.Key] = ((JToken)kv.Value).ToObject<object?>();
                        }
                    }

                    plan.Commands.Add(fc);
                }

                return plan;
            }
            catch (Exception ex)
            {
                return PlanErr("解析失败: " + ex.Message);
            }
        }

        /// <summary>快速构建错误状态的 ExecutionPlan</summary>
        static ExecutionPlan PlanErr(string msg) =>
            new ExecutionPlan { IsValid = false, Error = msg };
    }
}
