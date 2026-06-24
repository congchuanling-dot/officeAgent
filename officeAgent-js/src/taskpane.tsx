// ============================================================
// 🚀 入口文件 — Office Add-in 启动
//
// 流程：
//   1. Office.onReady() 等待 Office.js 初始化
//   2. 渲染 React App 到页面
// ============================================================

import React from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const root = createRoot(document.getElementById('root')!);
    root.render(<App />);
    console.log('✅ Word AI 助手已启动');
  }
});
