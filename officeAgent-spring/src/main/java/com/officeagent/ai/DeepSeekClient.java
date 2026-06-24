package com.officeagent.ai;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.officeagent.model.ExecutionPlan;
import com.officeagent.model.FontCommand;
import com.officeagent.model.TargetSpec;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.time.Duration;
import java.util.Map;

// ============================================================
// 🤖 DeepSeek API 客户端 — 与 AI 的通信管道
//
// 与 C# 版 DeepSeekClient.cs 逻辑完全一致：
//   - HTTP 请求 + Bearer 认证
//   - 3 次重试，401 不重试
//   - 剥离 Markdown 代码块
//   - 手动解析 JSON → ExecutionPlan
// ============================================================

public class DeepSeekClient {

    private final String apiKey;
    private final String apiUrl;
    private final String model;
    private final HttpClient http;
    private final ObjectMapper mapper;

    public DeepSeekClient(String apiKey, String apiUrl, String model) {
        this.apiKey = apiKey;
        this.apiUrl = apiUrl;
        this.model = model;
        this.http = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(10))
            .build();
        this.mapper = new ObjectMapper();
    }

    /**
     * 核心方法：发送消息到 DeepSeek，返回结构化执行计划。
     *
     * @param systemPrompt PromptBuilder 构建的 System Prompt
     * @param userInput    用户在聊天框输入的原始文字
     * @return ExecutionPlan，失败时 isValid=false 且 error 有描述
     */
    public ExecutionPlan call(String systemPrompt, String userInput) {
        // 前置检查：没配置 Key 则直接失败
        if (apiKey == null || apiKey.isBlank()) {
            return ExecutionPlan.error("请先配置 API Key（环境变量 DEEPSEEK_API_KEY）");
        }

        // 构建请求体
        String body;
        try {
            body = buildRequestBody(systemPrompt, userInput);
        } catch (Exception e) {
            return ExecutionPlan.error("构建请求失败: " + e.getMessage());
        }

        // ── 重试循环（最多 3 次）──
        for (int retry = 0; retry < 3; retry++) {
            try {
                HttpRequest req = HttpRequest.newBuilder()
                    .uri(URI.create(apiUrl))
                    .header("Content-Type", "application/json")
                    .header("Authorization", "Bearer " + apiKey)
                    .timeout(Duration.ofSeconds(60))
                    .POST(HttpRequest.BodyPublishers.ofString(body))
                    .build();

                HttpResponse<String> resp = http.send(req, HttpResponse.BodyHandlers.ofString());

                if (resp.statusCode() != 200) {
                    // 401：密钥无效，不重试
                    if (resp.statusCode() == 401) {
                        return ExecutionPlan.error("API Key 无效");
                    }
                    if (retry < 2) { sleep(1000); continue; }
                    return ExecutionPlan.error("API 错误: " + resp.body());
                }

                // 成功：解析 JSON
                return parseJson(resp.body());

            } catch (java.net.http.HttpTimeoutException e) {
                if (retry == 2) return ExecutionPlan.error("请求超时");
            } catch (Exception e) {
                if (retry == 2) return ExecutionPlan.error(e.getMessage());
            }

            if (retry < 2) sleep(1000);
        }
        return ExecutionPlan.error("未知错误");
    }

    // ----------------------------------------------------------
    // 构建请求体 JSON
    // ----------------------------------------------------------

    private String buildRequestBody(String systemPrompt, String userInput) throws Exception {
        var root = mapper.createObjectNode();
        root.put("model", model);
        root.put("max_tokens", 2048);
        root.put("temperature", 0.1);

        var messages = mapper.createArrayNode();
        var sysMsg = mapper.createObjectNode();
        sysMsg.put("role", "system");
        sysMsg.put("content", systemPrompt);
        messages.add(sysMsg);

        var userMsg = mapper.createObjectNode();
        userMsg.put("role", "user");
        userMsg.put("content", userInput);
        messages.add(userMsg);

        root.set("messages", messages);
        return mapper.writeValueAsString(root);
    }

    // ----------------------------------------------------------
    // 解析 AI 返回的 JSON → ExecutionPlan
    // ----------------------------------------------------------

    private ExecutionPlan parseJson(String respText) {
        try {
            JsonNode root = mapper.readTree(respText);

            // 步骤 1：取 choices[0].message.content
            JsonNode contentNode = root.path("choices").path(0).path("message").path("content");
            if (contentNode.isMissingNode() || contentNode.asText().isBlank()) {
                return ExecutionPlan.error("AI 返回空内容");
            }

            // 步骤 2：剥离 Markdown 代码块
            String t = contentNode.asText().trim();
            if (t.startsWith("```")) {
                int i = t.indexOf('\n');
                t = i < 0 ? t.substring(3) : t.substring(i + 1);
                int j = t.lastIndexOf("```");
                if (j > 0) t = t.substring(0, j).trim();
            }

            // 步骤 3：手动解析 JSON
            JsonNode json = mapper.readTree(t);
            ExecutionPlan plan = new ExecutionPlan();
            plan.setExplanation(json.path("explanation").asText(""));
            plan.setValid(true);

            // 遍历 commands 数组
            JsonNode commandsNode = json.path("commands");
            if (commandsNode.isArray()) {
                for (JsonNode cmd : commandsNode) {
                    FontCommand fc = new FontCommand();
                    fc.setAction(cmd.path("action").asText(""));

                    // target
                    TargetSpec target = new TargetSpec();
                    JsonNode targetNode = cmd.path("target");
                    target.setType(targetNode.path("type").asText(""));
                    if (targetNode.has("text") && !targetNode.get("text").isNull()) {
                        target.setText(targetNode.get("text").asText());
                    }
                    if (targetNode.has("index") && !targetNode.get("index").isNull()) {
                        target.setIndex(targetNode.get("index").asInt());
                    }
                    fc.setTarget(target);

                    // params（动态键值对）
                    JsonNode paramsNode = cmd.path("params");
                    if (paramsNode.isObject()) {
                        var iter = paramsNode.fields();
                        while (iter.hasNext()) {
                            Map.Entry<String, JsonNode> kv = iter.next();
                            JsonNode v = kv.getValue();
                            if (v.isTextual()) {
                                fc.getParams().put(kv.getKey(), v.asText());
                            } else if (v.isBoolean()) {
                                fc.getParams().put(kv.getKey(), v.asBoolean());
                            } else if (v.isNumber()) {
                                fc.getParams().put(kv.getKey(), v.numberValue());
                            }
                        }
                    }

                    plan.getCommands().add(fc);
                }
            }

            return plan;
        } catch (Exception e) {
            return ExecutionPlan.error("解析失败: " + e.getMessage());
        }
    }

    private static void sleep(long ms) {
        try { Thread.sleep(ms); } catch (InterruptedException ignored) { }
    }
}
