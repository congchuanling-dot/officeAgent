using System;
using System.Windows.Forms.Integration;
using Microsoft.Office.Tools;
using agentOffice.AI;

namespace agentOffice.UI
{
    // ============================================================
    // 🪟 任务面板宿主 — WPF ↔ Word CustomTaskPane 的桥梁
    //
    // 职责：把 WPF 控件（TaskPaneControl）嵌入 Word 侧边栏。
    //
    // 技术难点：VSTO 的 CustomTaskPane 原生只接受 WinForms 控件，
    // 但我们的 UI 是 WPF（现代、数据绑定、样式灵活），
    // 所以需要 ElementHost（WinForms → WPF 桥接器）做中转。
    //
    // 层级嵌套：
    //   CustomTaskPane → WinForms UserControl → ElementHost → WPF UserControl
    //
    // 生命周期：
    //   Show()  → 创建全栈对象，挂到 Word 侧边栏
    //   Hide()  → 仅设为不可见（保留状态）
    //   Toggle()→ 显隐切换
    //   Close() → 销毁全部对象（插件卸载时调用）
    // ============================================================

    public class TaskPaneHost
    {
        CustomTaskPane? _pane;     // VSTO 自定义任务面板
        TaskPaneControl? _ctrl;    // WPF 用户控件（视图）
        TaskPaneVM? _vm;           // ViewModel（逻辑）
        Agent? _agent;             // 核心编排器

        /// <summary>
        /// 显示任务面板。首次调用创建全栈，后续调用仅设为可见。
        ///
        /// 创建顺序（依赖注入，从内到外）：
        ///   DeepSeekClient → Agent → TaskPaneVM → TaskPaneControl
        ///   → ElementHost → WinForms UserControl → CustomTaskPane
        /// </summary>
        /// <param name="panes">VSTO 提供的面板集合（来自 ThisAddIn.CustomTaskPanes）</param>
        /// <param name="app">Word Application 对象</param>
        public void Show(
            CustomTaskPaneCollection panes,
            Microsoft.Office.Interop.Word.Application app)
        {
            // 已创建则只恢复显示，不重建
            if (_pane != null) { _pane.Visible = true; return; }

            // ── 1. 创建 AI 客户端和核心 Agent ──
            var ai = new DeepSeekClient();
            _agent = new Agent(app, ai);

            // ── 2. 创建 ViewModel ──
            _vm = new TaskPaneVM(_agent);

            // ── 3. 创建 WPF 控件并绑定 DataContext ──
            _ctrl = new TaskPaneControl { DataContext = _vm };

            // ── 4. 用 ElementHost 桥接 WPF → WinForms ──
            var host = new ElementHost
            {
                Child = _ctrl,
                Dock = System.Windows.Forms.DockStyle.Fill   // 填满整个面板
            };

            // ── 5. 包一层 WinForms UserControl ──
            var container = new System.Windows.Forms.UserControl();
            container.Controls.Add(host);

            // ── 6. 挂到 Word 侧边栏 ──
            _pane = panes.Add(container, "Word AI 助手");
            _pane.Visible = true;
            _pane.Width = 360;
        }

        /// <summary>隐藏面板（不销毁，下次 Show() 恢复）</summary>
        public void Hide()
        {
            if (_pane != null) _pane.Visible = false;
        }

        /// <summary>切换面板显隐状态</summary>
        public void Toggle(
            CustomTaskPaneCollection panes,
            Microsoft.Office.Interop.Word.Application app)
        {
            if (_pane == null)
                Show(panes, app);     // 首次创建
            else
                _pane.Visible = !_pane.Visible;   // 切换
        }

        /// <summary>
        /// 销毁面板 — 插件卸载时调用。
        /// CustomTaskPane.Dispose() 释放 COM 资源。
        /// </summary>
        public void Close()
        {
            if (_pane != null)
            {
                try { _pane.Dispose(); }
                catch { /* COM 释放可能抛异常，忽略 */ }
                _pane = null;
            }
            _ctrl = null;
            _vm = null;
            _agent = null;
        }
    }
}
