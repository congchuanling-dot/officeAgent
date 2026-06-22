using System;
using System.Diagnostics;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace agentOffice.UI
{
    // ============================================================
    // 🎨 视图代码后置 — XAML 的事件处理和值转换器
    //
    // 职责极简：
    //   - 按钮点击事件 → 转发给 ViewModel
    //   - 值转换器（bool ↔ Visibility / Invert）→ XAML 绑定用
    //
    // 设计原则：逻辑全在 ViewModel，这里只做"接线"。
    // ============================================================

    public partial class TaskPaneControl : UserControl
    {
        /// <summary>便捷访问 ViewModel</summary>
        TaskPaneVM? VM => DataContext as TaskPaneVM;

        // ── 供 XAML 静态引用转换器 ──
        /// <summary>布尔取反转换器（IsBusy → IsEnabled 反向绑定）</summary>
        public static readonly IValueConverter InvBool = new InvertBoolConv();

        /// <summary>布尔 → 可见性转换器（ShowSettings → 设置面板显隐）</summary>
        public static readonly IValueConverter BoolToVis = new BoolToVisConv();

        public TaskPaneControl() { InitializeComponent(); }

        /// <summary>⚙ 按钮 — 切换设置面板显隐</summary>
        void Settings_Click(object sender, RoutedEventArgs e)
        {
            if (VM != null) VM.ShowSettings = !VM.ShowSettings;
        }

        /// <summary>💾 保存 — 隐藏设置面板并弹出提示</summary>
        void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            if (VM != null)
            {
                VM.ShowSettings = false;
                // ApiKey 已通过双向绑定自动写入 VM.ApiKey，
                // Agent.SetApiKey() 在下次 Send() 时自动生效
                MessageBox.Show("已保存");
            }
        }

        /// <summary>「如何获取 Key？」链接 — 打开 DeepSeek 官网</summary>
        void GetKey_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://platform.deepseek.com/",
                UseShellExecute = true
            });
        }

        // ============================================================
        // 🔄 值转换器 — 供 XAML 绑定的 {Binding ..., Converter=...}
        // ============================================================

        /// <summary>
        /// 布尔取反：true → false, false → true。
        /// 用于 IsEnabled="{Binding IsBusy, Converter={x:Static local:TaskPaneControl.InvBool}}"
        /// 效果：IsBusy=true 时输入框禁用
        /// </summary>
        class InvertBoolConv : IValueConverter
        {
            public object Convert(object v, Type t, object p, CultureInfo c) =>
                v is bool b && !b;

            public object ConvertBack(object v, Type t, object p, CultureInfo c) =>
                v is bool b && !b;
        }

        /// <summary>
        /// 布尔 → 可见性：true → Visible, false → Collapsed。
        /// 用于 Visibility="{Binding ShowSettings, Converter={x:Static local:TaskPaneControl.BoolToVis}}"
        /// 效果：ShowSettings=true 时设置面板可见
        /// </summary>
        class BoolToVisConv : IValueConverter
        {
            public object Convert(object v, Type t, object p, CultureInfo c) =>
                v is bool b && b ? Visibility.Visible : Visibility.Collapsed;

            public object ConvertBack(object v, Type t, object p, CultureInfo c) =>
                v is Visibility vis && vis == Visibility.Visible;
        }
    }
}
