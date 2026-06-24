package com.officeagent.model;

/**
 * 📝 段落摘要 — 不传全文，只传结构信息给 AI
 */
public class ParaInfo {
    private int index;
    private String style = "";
    private String text = "";

    public int getIndex() { return index; }
    public void setIndex(int index) { this.index = index; }

    public String getStyle() { return style; }
    public void setStyle(String style) { this.style = style; }

    public String getText() { return text; }
    public void setText(String text) { this.text = text; }
}
