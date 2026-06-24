package com.officeagent.model;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * ✏️ 格式命令 — AI 指定 "对什么东西、做什么操作、用什么参数"
 */
public class FontCommand {
    private String action = "";
    private TargetSpec target = new TargetSpec();
    private Map<String, Object> params = new LinkedHashMap<>();

    public String getAction() { return action; }
    public void setAction(String action) { this.action = action; }

    public TargetSpec getTarget() { return target; }
    public void setTarget(TargetSpec target) { this.target = target; }

    public Map<String, Object> getParams() { return params; }
    public void setParams(Map<String, Object> params) { this.params = params; }
}
