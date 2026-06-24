// ============================================================
// 🎨 React 主组件 — Word AI 助手侧边栏
//
// 这个文件就是你以后要改 UI 时主要看的地方。
// 所有关键逻辑委托给 Agent（TS 侧）→ Spring Boot（Java 侧）
// ============================================================

import React, { useState, useRef, useEffect, useCallback } from 'react';
import { Agent } from './agent';
import { ChatMessage } from './models';

const agent = new Agent();

export const App: React.FC = () => {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState('');
  const [status, setStatus] = useState('就绪');
  const [busy, setBusy] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [apiKey, setApiKey] = useState('');
  const scrollRef = useRef<HTMLDivElement>(null);

  // 滚动到底部
  useEffect(() => {
    scrollRef.current?.scrollTo(0, scrollRef.current.scrollHeight);
  }, [messages]);

  // ============================================================
  // 📤 Send — 分析阶段
  // ============================================================
  const send = useCallback(async () => {
    const txt = input.trim();
    if (!txt || busy) return;

    setInput('');
    setMessages((prev) => [...prev, { role: '👤', content: txt }]);
    setBusy(true);
    setStatus('分析中...');

    try {
      const plan = await agent.processAsync(txt);

      if (plan.isValid) {
        let preview = '📋 ' + plan.explanation + '\n\n操作:';
        let n = 1;
        for (const cmd of plan.commands) {
          const targetDesc =
            cmd.target.type === 'find' ? '「' + cmd.target.text + '」' : cmd.target.type;
          preview += `\n${n++}. 修改 ${targetDesc} 的颜色`;
        }

        setMessages((prev) => [
          ...prev,
          { role: '🤖', content: preview, pendingPlan: plan },
        ]);
        setStatus('请确认');
      } else {
        setMessages((prev) => [...prev, { role: '❌', content: plan.error || '分析失败' }]);
        setStatus('就绪');
      }
    } catch (ex: any) {
      setMessages((prev) => [...prev, { role: '❌', content: ex.message || '未知错误' }]);
      setStatus('就绪');
    } finally {
      setBusy(false);
    }
  }, [input, busy]);

  // ============================================================
  // ✅ Confirm — 执行阶段
  // ============================================================
  const confirm = useCallback(async () => {
    const msg = messages.find((m) => m.pendingPlan);
    if (!msg?.pendingPlan) return;

    const plan = msg.pendingPlan;
    msg.pendingPlan = undefined;

    setBusy(true);
    setStatus('执行中...');
    setMessages([...messages]);

    try {
      const result = await agent.executeAsync(plan);
      setMessages((prev) => [...prev, { role: '✅', content: result }]);
      setStatus('就绪');
    } catch (ex: any) {
      setMessages((prev) => [...prev, { role: '❌', content: ex.message || '执行失败' }]);
      setStatus('就绪');
    } finally {
      setBusy(false);
    }
  }, [messages, busy]);

  // ============================================================
  // ❌ Cancel
  // ============================================================
  const cancel = useCallback(() => {
    setMessages((prev) =>
      prev.map((m) => {
        if (m.pendingPlan) {
          return { ...m, pendingPlan: undefined, content: m.content + '（已取消）' };
        }
        return m;
      })
    );
    setStatus('就绪');
  }, []);

  // ============================================================
  // ⌨️ 快捷键
  // ============================================================
  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.ctrlKey && e.key === 'Enter') {
      e.preventDefault();
      send();
    }
  };

  // ============================================================
  // 🎨 渲染
  // ============================================================
  const hasPending = messages.some((m) => m.pendingPlan);
  const canSend = !busy && input.trim().length > 0;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%' }}>
      {/* ===== Header ===== */}
      <div style={styles.header}>
        <span style={{ fontSize: 15, fontWeight: 'bold' }}>🤖 Word AI 助手</span>
        <button style={styles.gearBtn} onClick={() => setShowSettings(!showSettings)} title="设置">
          ⚙
        </button>
      </div>

      {/* ===== Main ===== */}
      <div style={{ flex: 1, overflow: 'hidden', position: 'relative' }}>
        {/* 消息区 */}
        <div ref={scrollRef} style={{ height: '100%', overflowY: 'auto', padding: 8 }}>
          {messages.map((m, i) => (
            <div key={i} style={styles.msgCard}>
              <div style={{ fontWeight: 'bold', fontSize: 12, marginBottom: 3 }}>{m.role}</div>
              <div style={{ fontSize: 13, lineHeight: 1.5, whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>
                {m.content.split('\n').map((line, j) => (
                  <React.Fragment key={j}>
                    {line}
                    {j < m.content.split('\n').length - 1 && <br />}
                  </React.Fragment>
                ))}
              </div>
            </div>
          ))}
        </div>

        {/* 设置面板 */}
        {showSettings && (
          <div style={{ position: 'absolute', inset: 0, background: '#fff', padding: 16, overflowY: 'auto' }}>
            <h3 style={{ fontSize: 15, marginBottom: 14 }}>⚙ 设置</h3>
            <label style={{ fontSize: 12, color: '#666' }}>Spring Boot 后端 Key（传给后端处理）</label>
            <input
              type="password"
              placeholder="sk-..."
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              style={styles.input}
            />
            <button
              style={{ ...styles.btnPrimary }}
              onClick={() => {
                setMessages((prev) => [...prev, { role: '🤖', content: '✅ Key 已保存' }]);
                setShowSettings(false);
              }}
            >
              💾 保存
            </button>
            <div style={{ marginTop: 10 }}>
              <a
                href="https://platform.deepseek.com/api_keys"
                target="_blank"
                rel="noopener"
                style={{ color: '#2B579A', fontSize: 12 }}
              >
                如何获取 Key？
              </a>
            </div>
          </div>
        )}
      </div>

      {/* ===== Input ===== */}
      <div style={{ borderTop: '1px solid #ddd', padding: '6px 8px' }}>
        <textarea
          rows={2}
          placeholder="输入指令，例如：把标题改成红色加粗..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={busy}
          style={styles.textarea}
        />
      </div>

      {/* ===== Footer ===== */}
      <div style={{ display: 'flex', alignItems: 'center', padding: '4px 8px 8px', gap: 10 }}>
        <button disabled={!canSend} onClick={send} style={{ ...styles.btn, ...styles.btnSend }}>
          📤 发送
        </button>
        <span style={{ flex: 1, fontSize: 11, color: '#888' }}>{status}</span>
        <button
          disabled={busy || !hasPending}
          onClick={cancel}
          style={{ ...styles.btn, background: '#ddd', color: '#333' }}
        >
          取消
        </button>
        <button
          disabled={busy || !hasPending}
          onClick={confirm}
          style={{ ...styles.btn, ...styles.btnConfirm }}
        >
          ✅ 确认
        </button>
      </div>
    </div>
  );
};

// ============================================================
// 🎨 内联样式（可提取到 .css 文件，这里保持简单）
// ============================================================

const styles: Record<string, React.CSSProperties> = {
  header: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    background: '#2B579A',
    color: '#fff',
    padding: '10px 12px',
    flexShrink: 0,
  },
  gearBtn: {
    background: 'transparent',
    color: '#fff',
    border: 'none',
    fontSize: 16,
    cursor: 'pointer',
    width: 28,
    height: 28,
    borderRadius: 4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  msgCard: {
    background: '#F8F8F8',
    borderRadius: 6,
    padding: '8px 10px',
    marginBottom: 4,
  },
  input: {
    width: '100%',
    padding: '8px',
    fontSize: 13,
    border: '1px solid #ccc',
    borderRadius: 4,
    marginBottom: 10,
    outline: 'none',
  },
  btnPrimary: {
    background: '#2B579A',
    color: '#fff',
    border: 'none',
    padding: '7px 20px',
    fontSize: 13,
    borderRadius: 4,
    cursor: 'pointer',
  },
  textarea: {
    width: '100%',
    minHeight: 36,
    maxHeight: 64,
    fontSize: 13,
    fontFamily: 'inherit',
    border: 'none',
    outline: 'none',
    resize: 'none',
    background: 'transparent',
    lineHeight: 1.4,
  },
  btn: {
    height: 32,
    padding: '0 16px',
    fontSize: 13,
    border: 'none',
    borderRadius: 4,
    cursor: 'pointer',
    whiteSpace: 'nowrap',
  },
  btnSend: { background: '#2B579A', color: '#fff', minWidth: 72 },
  btnConfirm: { background: '#388E3C', color: '#fff', minWidth: 72 },
};
