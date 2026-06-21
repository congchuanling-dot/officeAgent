using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using agentOffice.UI;

namespace agentOffice
{
    public partial class ThisAddIn
    {
        TaskPaneHost? _host;

        void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                _host = new TaskPaneHost();
                _host.Show(CustomTaskPanes, Application);
                Application.DocumentOpen += OnDocOpen;
            }
            catch (Exception ex) { MessageBox.Show("启动失败: " + ex.Message); }
        }

        void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Application.DocumentOpen -= OnDocOpen;
            _host?.Close();
        }

        void OnDocOpen(Word.Document doc) => _host?.Show(CustomTaskPanes, Application);

        public void TogglePane() => _host?.Toggle(CustomTaskPanes, Application);

        #region VSTO generated
        void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion
    }
}
