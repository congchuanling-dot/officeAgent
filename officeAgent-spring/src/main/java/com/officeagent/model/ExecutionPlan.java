package com.officeagent.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 🎯 执行计划 — DeepSeek API 返回的 JSON 反序列化结果
 */
public class ExecutionPlan {
    private String explanation = "";
    private boolean isValid;
    private String error = "";
    private List<FontCommand> commands = new ArrayList<>();

    public String getExplanation() { return explanation; }
    public void setExplanation(String explanation) { this.explanation = explanation; }

    public boolean isValid() { return isValid; }
    public void setValid(boolean isValid) { this.isValid = isValid; }

    public String getError() { return error; }
    public void setError(String error) { this.error = error; }

    public List<FontCommand> getCommands() { return commands; }
    public void setCommands(List<FontCommand> commands) { this.commands = commands; }

    /** 工厂：构建错误计划 */
    public static ExecutionPlan error(String msg) {
        ExecutionPlan p = new ExecutionPlan();
        p.isValid = false;
        p.error = msg;
        return p;
    }
}
