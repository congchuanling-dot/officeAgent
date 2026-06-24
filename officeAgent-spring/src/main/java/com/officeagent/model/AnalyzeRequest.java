package com.officeagent.model;

/**
 * 📥 请求体 — TS 侧发过来的 JSON
 */
public class AnalyzeRequest {
    private String userInput;
    private DocumentSnapshot snapshot;

    public String getUserInput() { return userInput; }
    public void setUserInput(String userInput) { this.userInput = userInput; }

    public DocumentSnapshot getSnapshot() { return snapshot; }
    public void setSnapshot(DocumentSnapshot snapshot) { this.snapshot = snapshot; }
}
