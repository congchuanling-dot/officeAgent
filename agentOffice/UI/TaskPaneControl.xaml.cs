using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace agentOffice.UI
{
    public partial class TaskPaneControl : UserControl
    {
        TaskPaneVM? VM => DataContext as TaskPaneVM;

        public static readonly IValueConverter InvBool = new InvertBoolConv();
        public static readonly IValueConverter BoolToVis = new BoolToVisConv();

        public TaskPaneControl() { InitializeComponent(); }

        void Settings_Click(object sender, RoutedEventArgs e) { if (VM != null) VM.ShowSettings = !VM.ShowSettings; }
        void SaveSettings_Click(object sender, RoutedEventArgs e) { if (VM != null) { VM.ShowSettings = false; MessageBox.Show("已保存"); } }
        void GetKey_Click(object sender, RoutedEventArgs e) { Process.Start(new ProcessStartInfo { FileName = "https://platform.deepseek.com/", UseShellExecute = true }); }

        class InvertBoolConv : IValueConverter
        {
            public object Convert(object v, Type t, object p, CultureInfo c) => v is bool b && !b;
            public object ConvertBack(object v, Type t, object p, CultureInfo c) => v is bool b && !b;
        }
        class BoolToVisConv : IValueConverter
        {
            public object Convert(object v, Type t, object p, CultureInfo c) => v is bool b && b ? Visibility.Visible : Visibility.Collapsed;
            public object ConvertBack(object v, Type t, object p, CultureInfo c) => v is Visibility vis && vis == Visibility.Visible;
        }
    }
}
