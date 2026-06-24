package com.officeagent.model;

/**
 * 🎯 目标定位 — 描述 "要修改文档的哪一部分"
 */
public class TargetSpec {
    private String type = "";
    private String text;
    private Integer index;

    public String getType() { return type; }
    public void setType(String type) { this.type = type; }

    public String getText() { return text; }
    public void setText(String text) { this.text = text; }

    public Integer getIndex() { return index; }
    public void setIndex(Integer index) { this.index = index; }
}
