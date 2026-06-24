package com.officeagent.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 📄 文档快照 — 从 Word 提取的结构化数据
 */
public class DocumentSnapshot {
    private List<ParaInfo> paragraphs = new ArrayList<>();
    private String selText = "";
    private int selParaIdx = -1;

    public List<ParaInfo> getParagraphs() { return paragraphs; }
    public void setParagraphs(List<ParaInfo> paragraphs) { this.paragraphs = paragraphs; }

    public String getSelText() { return selText; }
    public void setSelText(String selText) { this.selText = selText; }

    public int getSelParaIdx() { return selParaIdx; }
    public void setSelParaIdx(int selParaIdx) { this.selParaIdx = selParaIdx; }
}
