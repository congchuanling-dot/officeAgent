using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;
using agentOffice.Models;

namespace agentOffice.UI
{
    // ============================================================
    // 🧠 ViewModel — 任务面板的"大脑"
    //
    // 职责：连接 UI（XAML 绑定）和业务逻辑（Agent）。
    //
    // 核心交互流程（两个关键方法）：
    //   Send()    — 用户输入文字 → Agent.ProcessAsync() → 展示预览
    //   Confirm() — 用户点确认 → Agent.ExecuteAsync()  → 执行修改
    //
    // 设计要点：
    //   - MVVM 无框架手写：INotifyPropertyChanged + RelayCmd
    //   - 预览-确认两阶段：AI 返回的命令先挂 PendingPlan，用户确认才执行
    //   - IsBusy 互斥锁：发送和确认期间禁用输入，防止并发操作
    // ============================================================

    public class TaskPaneVM : INotifyPropertyChanged
    {
        readonly Agent _agent;

        public TaskPaneVM(Agent agent)
        {
            _agent = agent;

            // Ctrl+Enter 发送
            SendCmd = new RelayCmd(
                async _ => await Send(),
                _ => !IsBusy && !string.IsNullOrWhiteSpace(Input));

            // 确认按钮：取出 latest 消息中挂载的 PendingPlan 执行
            ConfirmCmd = new RelayCmd(async _ => await Confirm());

            // 取消按钮：清除 PendingPlan，消息标注（已取消）
            CancelCmd = new RelayCmd(_ => Cancel());
        }

        // ────────── 绑定属性 ──────────

        /// <summary>用户输入框内容，双向绑定</summary>
        string _input = "";
        public string Input
        {
            get => _input;
            set { _input = value; OnPropChanged(); OnPropChanged(nameof(CanSend)); }
        }

        /// <summary>是否正在处理中（发送/执行），用于禁用输入框和按钮</summary>
        bool _busy;
        public bool IsBusy
        {
            get => _busy;
            set { _busy = value; OnPropChanged(); OnPropChanged(nameof(CanSend)); }
        }

        /// <summary>状态栏文字（就绪 / 分析中 / 执行中 / 请确认）</summary>
        string _status = "就绪";
        public string Status
        {
            get => _status;
            set { _status = value; OnPropChanged(); }
        }

        /// <summary>API Key，设置面板输入，保存时传给 Agent</summary>
        string _apiKey = "";
        public string ApiKey
        {
            get => _apiKey;
            set { _apiKey = value; OnPropChanged(); }
        }

        /// <summary>设置面板显隐</summary>
        bool _showSettings;
        public bool ShowSettings
        {
            get => _showSettings;
            set { _showSettings = value; OnPropChanged(); }
        }

        // ────────── 集合属性 ──────────

        /// <summary>聊天消息列表，DataTemplate 绑定到 ItemsControl</summary>
        public ObservableCollection<ChatMessage> Messages { get; }
            = new ObservableCollection<ChatMessage>();

        /// <summary>是否可以发送（输入非空 + 非忙碌状态）</summary>
        public bool CanSend => !IsBusy && !string.IsNullOrWhiteSpace(Input);

        // ────────── 命令 ──────────

        public ICommand SendCmd    { get; }   // 发送 / Ctrl+Enter
        public ICommand ConfirmCmd { get; }   // 确认执行
        public ICommand CancelCmd  { get; }   // 取消挂起的计划

        // ============================================================
        // 📤 Send — 分析阶段
        // ============================================================

        /// <summary>
        /// 发送用户输入给 AI 分析。
        ///
        /// 流程：
        ///   1. 取输入文本并清空输入框
        ///   2. 添加用户消息到列表
        ///   3. 如果设置了 API Key，传入 Agent
        ///   4. 调用 Agent.ProcessAsync() → 获取 ExecutionPlan
        ///   5. 将 AI 解释和操作列表作为预览消息展示
        ///   6. 把 ExecutionPlan 挂到消息的 PendingPlan 上等待确认
        ///
        /// 错误处理：API 失败/解析失败 → 添加 ❌ 消息
        /// </summary>
        async Task Send()
        {
            var txt = Input.Trim();
            Input = "";     // 清空输入框

            // 添加用户消息气泡
            Messages.Add(new ChatMessage { Role = "👤", Content = txt });

            // 如果用户在设置面板填了 Key，实时生效
            if (!string.IsNullOrWhiteSpace(ApiKey))
                _agent.SetApiKey(ApiKey);

            IsBusy = true;
            Status = "分析中...";

            try
            {
                // ── 核心调用：自然语言 → AI → ExecutionPlan ──
                var plan = await _agent.ProcessAsync(txt);

                if (plan.IsValid)
                {
                    // 构建人类可读的预览消息
                    var preview = "📋 " + plan.Explanation + "\n\n操作:";
                    int n = 1;
                    foreach (var c in plan.Commands)
                    {
                        // 根据 target 类型生成友好的预览描述
                        var targetDesc = c.Target.Type == "find"
                            ? "「" + c.Target.Text + "」"
                            : c.Target.Type;
                        preview += $"\n{n++}. 修改 {targetDesc} 的颜色";
                    }

                    // 添加预览消息，并挂载 PendingPlan 等待确认
                    Messages.Add(new ChatMessage
                    {
                        Role = "🤖",
                        Content = preview,
                        PendingPlan = plan
                    });
                    Status = "请确认";
                }
                else
                {
                    // AI 返回错误（如 Key 无效、解析失败）
                    Messages.Add(new ChatMessage
                    {
                        Role = "❌",
                        Content = plan.Error ?? "分析失败"
                    });
                    Status = "就绪";
                }
            }
            catch (Exception ex)
            {
                Messages.Add(new ChatMessage { Role = "❌", Content = ex.Message });
                Status = "就绪";
            }
            finally { IsBusy = false; }
        }

        // ============================================================
        // ✅ Confirm — 执行阶段
        // ============================================================

        /// <summary>
        /// 用户点击「确认」— 执行最后一个待确认的计划。
        ///
        /// 流程：
        ///   1. 遍历消息列表，找到第一个有 PendingPlan 的消息
        ///   2. 取出 PendingPlan 并清空（防止重复执行）
        ///   3. 调用 Agent.ExecuteAsync() 执行命令
        ///   4. 添加 ✅ 成功消息
        ///
        /// 设计选择：确认最新一条计划而非全部，
        /// 因为正常流程中同时只有一个待确认计划。
        /// </summary>
        async Task Confirm()
        {
            // 找到第一个挂载了 PendingPlan 的消息
            ChatMessage? msg = null;
            foreach (var m in Messages)
                if (m.PendingPlan != null) { msg = m; break; }

            if (msg?.PendingPlan == null) return;   // 没有待确认的计划

            var plan = msg.PendingPlan;
            msg.PendingPlan = null;   // 取出后立即清空，防重复执行

            IsBusy = true;
            Status = "执行中...";
            try
            {
                // ── 核心调用：ExecutionPlan → Word 文档修改 ──
                var result = await _agent.ExecuteAsync(plan);
                Messages.Add(new ChatMessage { Role = "✅", Content = result });
                Status = "就绪";
            }
            catch (Exception ex)
            {
                Messages.Add(new ChatMessage { Role = "❌", Content = ex.Message });
                Status = "就绪";
            }
            finally { IsBusy = false; }
        }

        // ============================================================
        // ❌ Cancel — 取消等待中的计划
        // ============================================================

        /// <summary>
        /// 取消所有待确认的计划，消息后标注（已取消）。
        /// </summary>
        void Cancel()
        {
            foreach (var m in Messages)
            {
                if (m.PendingPlan != null)
                {
                    m.PendingPlan = null;
                    m.Content += "（已取消）";
                }
            }
            Status = "就绪";
        }

        // ============================================================
        // INotifyPropertyChanged 实现
        // ============================================================

        public event PropertyChangedEventHandler? PropertyChanged;
        void OnPropChanged([CallerMemberName] string? n = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    // ============================================================
    // 🔘 RelayCmd — 轻量级 ICommand 实现
    //
    // 无需第三方 MVVM 框架，20 行搞定的委托命令：
    //   - exec:  点击按钮时的执行逻辑
    //   - can:   控制按钮是否可点击（IsEnabled 绑定）
    //   - CommandManager.RequerySuggested：WPF 自动检查按钮状态
    // ============================================================

    public class RelayCmd : ICommand
    {
        readonly Action<object?> _exec;
        readonly Func<object?, bool>? _can;

        public RelayCmd(Action<object?> exec, Func<object?, bool>? can = null)
        {
            _exec = exec;
            _can = can;
        }

        public bool CanExecute(object? p) => _can?.Invoke(p) ?? true;

        public void Execute(object? p) => _exec(p);

        /// <summary>
        /// 订阅 CommandManager.RequerySuggested 让 WPF 自动刷新按钮状态。
        /// 控制焦点变化等时机触发 CanExecute 重新评估。
        /// </summary>
        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }
}
