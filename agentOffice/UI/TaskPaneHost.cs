using System;
using System.Windows.Forms.Integration;
using Microsoft.Office.Tools;
using agentOffice.AI;

namespace agentOffice.UI
{
    public class TaskPaneHost
    {
        CustomTaskPane? _pane;
        TaskPaneControl? _ctrl;
        TaskPaneVM? _vm;
        Agent? _agent;

        public void Show(CustomTaskPaneCollection panes, Microsoft.Office.Interop.Word.Application app)
        {
            if (_pane != null) { _pane.Visible = true; return; }

            var ai = new DeepSeekClient();
            _agent = new Agent(app, ai);
            _vm = new TaskPaneVM(_agent);
            _ctrl = new TaskPaneControl { DataContext = _vm };

            var host = new ElementHost { Child = _ctrl, Dock = System.Windows.Forms.DockStyle.Fill };
            var container = new System.Windows.Forms.UserControl();
            container.Controls.Add(host);

            _pane = panes.Add(container, "Word AI 助手");
            _pane.Visible = true;
            _pane.Width = 360;
        }

        public void Hide() { if (_pane != null) _pane.Visible = false; }

        public void Toggle(CustomTaskPaneCollection panes, Microsoft.Office.Interop.Word.Application app)
        {
            if (_pane == null) Show(panes, app);
            else _pane.Visible = !_pane.Visible;
        }

        public void Close()
        {
            if (_pane != null) { try { _pane.Dispose(); } catch { } _pane = null; }
            _ctrl = null; _vm = null; _agent = null;
        }
    }
}
