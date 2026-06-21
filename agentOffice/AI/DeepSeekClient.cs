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
    public class DeepSeekClient
    {
        readonly HttpClient _http;
        string _apiKey = "";
        string _apiUrl = "https://api.deepseek.com/chat/completions";
        string _model = "deepseek-v4-pro";

        public DeepSeekClient()
        {
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            var key = Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(key)) _apiKey = key;
            var url = Environment.GetEnvironmentVariable("DEEPSEEK_API_URL", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(url)) _apiUrl = url;
            var mdl = Environment.GetEnvironmentVariable("DEEPSEEK_MODEL", EnvironmentVariableTarget.User);
            if (!string.IsNullOrEmpty(mdl)) _model = mdl;
        }

        public void SetApiKey(string k) => _apiKey = k;

        public async Task<ExecutionPlan> CallAsync(string systemPrompt, string userInput)
        {
            if (string.IsNullOrWhiteSpace(_apiKey))
                return PlanErr("请先配置 API Key");

            var body = new
            {
                model = _model,
                max_tokens = 2048,
                temperature = 0.1,
                messages = new[] {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = userInput }
                }
            };

            for (int retry = 0; retry < 3; retry++)
            {
                try
                {
                    var req = new HttpRequestMessage(HttpMethod.Post, _apiUrl)
                    {
                        Content = new StringContent(JsonConvert.SerializeObject(body), Encoding.UTF8, "application/json")
                    };
                    req.Headers.Add("Authorization", "Bearer " + _apiKey);

                    var resp = await _http.SendAsync(req);
                    var respText = await resp.Content.ReadAsStringAsync();

                    if (!resp.IsSuccessStatusCode)
                    {
                        if ((int)resp.StatusCode == 401) return PlanErr("API Key 无效");
                        if (retry < 2) { await Task.Delay(1000); continue; }
                        return PlanErr("API 错误: " + respText);
                    }

                    return ParseJson(respText);
                }
                catch (TaskCanceledException) { if (retry == 2) return PlanErr("请求超时"); }
                catch (Exception ex) { if (retry == 2) return PlanErr(ex.Message); }
                await Task.Delay(1000);
            }
            return PlanErr("未知错误");
        }

        ExecutionPlan ParseJson(string respText)
        {
            try
            {
                var root = JObject.Parse(respText);
                var text = root["choices"]?[0]?["message"]?["content"]?.ToString();
                if (string.IsNullOrWhiteSpace(text)) return PlanErr("AI 返回空内容");

                var t = text.Trim();
                if (t.StartsWith("```")) { var i = t.IndexOf('\n'); t = t.Substring(i < 0 ? 3 : i + 1); i = t.LastIndexOf("```"); if (i > 0) t = t.Substring(0, i).Trim(); }

                var json = JObject.Parse(t);
                var plan = new ExecutionPlan { Explanation = json["explanation"]?.ToString() ?? "", IsValid = true };

                foreach (var cmd in json["commands"] ?? new JArray())
                {
                    var fc = new FontCommand
                    {
                        Action = cmd["action"]?.ToString() ?? "",
                        Target = new TargetSpec
                        {
                            Type = cmd["target"]?["type"]?.ToString() ?? "",
                            Text = cmd["target"]?["text"]?.ToString(),
                            Index = cmd["target"]?["index"]?.Value<int?>()
                        }
                    };
                    if (cmd["params"] is JObject po)
                        foreach (var kv in po) fc.Params[kv.Key] = ((JToken)kv.Value).ToObject<object?>();
                    plan.Commands.Add(fc);
                }
                return plan;
            }
            catch (Exception ex) { return PlanErr("解析失败: " + ex.Message); }
        }

        static ExecutionPlan PlanErr(string msg) => new ExecutionPlan { IsValid = false, Error = msg };
    }
}
