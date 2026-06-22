using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.UI;

namespace agentOffice
{
    // ============================================================
    // 🔌 插件入口 — VSTO 加载项的生命周期管理
    //
    // 这是 Word 加载 agentOffice 时最先执行的代码。
    // 职责极简：
    //   1. 启动时创建右侧任务面板（TaskPaneHost）
    //   2. 监听文档打开事件，新文档也自动显示面板
    //   3. 关闭时清理资源
    //
    // VSTO 框架：ThisAddIn.Designer.cs 中 InternalStartup()
    // 将 Startup/Shutdown 事件绑定到以下两个方法。
    // ============================================================

    public partial class ThisAddIn
    {
        /// <summary>任务面板宿主，管理 CustomTaskPane 的创建/显示/隐藏/销毁</summary>
        TaskPaneHost? _host;

        /// <summary>
        /// 插件启动入口 — Word 加载此插件时调用。
        /// 流程：创建 TaskPaneHost → 初始化 AI 和 UI → 挂载到 Word 侧边栏。
        /// </summary>
        void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _host = new TaskPaneHost();
                // TaskPaneHost.Show() 内部创建 DeepSeekClient、Agent、TaskPaneVM、WPF 控件
                _host.Show(CustomTaskPanes, Application);

                // 监听文档打开事件：打开新文档时自动显示面板
                Application.DocumentOpen += OnDocOpen;
            }
            catch (Exception ex) { MessageBox.Show("启动失败: " + ex.Message); }
        }

        /// <summary>
        /// 插件卸载入口 — Word 关闭或禁用插件时调用。
        /// 解绑事件 + 释放 COM 资源（CustomTaskPane）。
        /// </summary>
        void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.DocumentOpen -= OnDocOpen;
            _host?.Close();   // 释放 CustomTaskPane + ViewModel + Agent
        }

        /// <summary>
        /// 文档打开回调 — 用户打开已有文档或新建文档时触发。
        /// 确保任何 Word 窗口都能看到 AI 面板。
        /// </summary>
        void OnDocOpen(Word.Document doc) => _host?.Show(CustomTaskPanes, Application);

        /// <summary>
        /// 手动切换面板显隐（预留，目前未绑定 UI 按钮）。
        /// </summary>
        public void TogglePane() => _host?.Toggle(CustomTaskPanes, Application);

        // ============================================================
        // 🔧 VSTO 自动生成区域
        // InternalStartup() 在 ThisAddIn.Designer.cs 中自动生成，
        // 将 VSTO 生命周期事件连接到上面的处理方法。
        // ============================================================
        #region VSTO generated
        void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
