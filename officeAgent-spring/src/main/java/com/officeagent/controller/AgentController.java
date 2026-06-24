package com.officeagent.controller;

import com.officeagent.model.AnalyzeRequest;
import com.officeagent.model.ExecutionPlan;
import com.officeagent.service.AgentService;

import org.springframework.web.bind.annotation.*;

// ============================================================
// 🌐 REST Controller — 给 Word 侧边栏调用的 API
//
// 当前只有一个接口，后续可扩展：
//   - GET  /api/health    健康检查
//   - POST /api/analyze   分析用户输入
// ============================================================

@RestController
@RequestMapping("/api")
@CrossOrigin(origins = "*") // Office 侧边栏跨域（生产环境应限制具体域名）
public class AgentController {

    private final AgentService service;

    public AgentController(AgentService service) {
        this.service = service;
    }

    /**
     * POST /api/analyze
     *
     * 请求体（JSON）：
     *   {
     *     "userInput": "把标题改成红色加粗",
     *     "snapshot": { "paragraphs": [...], "selText": "...", "selParaIdx": 0 }
     *   }
     *
     * 返回（JSON）：ExecutionPlan，与 C#/TS 版结构完全一致
     */
    @PostMapping("/analyze")
    public ExecutionPlan analyze(@RequestBody AnalyzeRequest req) {
        return service.analyze(req);
    }

    /** 健康检查 */
    @GetMapping("/health")
    public String health() {
        return "OK";
    }
}
