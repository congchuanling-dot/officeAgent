using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows.Input;
using agentOffice.Models;

namespace agentOffice.UI
{
    public class TaskPaneVM : INotifyPropertyChanged
    {
        readonly Agent _agent;

        public TaskPaneVM(Agent agent)
        {
            _agent = agent;
            SendCmd = new RelayCmd(async _ => await Send(), _ => !IsBusy && !string.IsNullOrWhiteSpace(Input));
            ConfirmCmd = new RelayCmd(async _ => await Confirm());
            CancelCmd = new RelayCmd(_ => Cancel());
        }

        string _input = "";
        public string Input { get => _input; set { _input = value; OnPropChanged(); OnPropChanged(nameof(CanSend)); } }

        bool _busy;
        public bool IsBusy { get => _busy; set { _busy = value; OnPropChanged(); OnPropChanged(nameof(CanSend)); } }

        string _status = "就绪";
        public string Status { get => _status; set { _status = value; OnPropChanged(); } }

        string _apiKey = "";
        public string ApiKey { get => _apiKey; set { _apiKey = value; OnPropChanged(); } }

        bool _showSettings;
        public bool ShowSettings { get => _showSettings; set { _showSettings = value; OnPropChanged(); } }

        public ObservableCollection<ChatMessage> Messages { get; } = new ObservableCollection<ChatMessage>();
        public bool CanSend => !IsBusy && !string.IsNullOrWhiteSpace(Input);

        public ICommand SendCmd { get; }
        public ICommand ConfirmCmd { get; }
        public ICommand CancelCmd { get; }

        async Task Send()
        {
            var txt = Input.Trim(); Input = "";
            Messages.Add(new ChatMessage { Role = "👤", Content = txt });

            if (!string.IsNullOrWhiteSpace(ApiKey)) _agent.SetApiKey(ApiKey);
            IsBusy = true; Status = "分析中...";

            try
            {
                var plan = await _agent.ProcessAsync(txt);
                if (plan.IsValid)
                {
                    var preview = "📋 " + plan.Explanation + "\n\n操作:";
                    int n = 1;
                    foreach (var c in plan.Commands)
                        preview += $"\n{n++}. 修改 {(c.Target.Type == "find" ? "「" + c.Target.Text + "」" : c.Target.Type)} 的颜色";
                    Messages.Add(new ChatMessage { Role = "🤖", Content = preview, PendingPlan = plan });
                    Status = "请确认";
                }
                else
                {
                    Messages.Add(new ChatMessage { Role = "❌", Content = plan.Error ?? "分析失败" });
                    Status = "就绪";
                }
            }
            catch (Exception ex) { Messages.Add(new ChatMessage { Role = "❌", Content = ex.Message }); Status = "就绪"; }
            finally { IsBusy = false; }
        }

        async Task Confirm()
        {
            ChatMessage? msg = null;
            foreach (var m in Messages) if (m.PendingPlan != null) { msg = m; break; }
            if (msg?.PendingPlan == null) return;
            var plan = msg.PendingPlan;
            msg.PendingPlan = null;

            IsBusy = true; Status = "执行中...";
            try
            {
                var result = await _agent.ExecuteAsync(plan);
                Messages.Add(new ChatMessage { Role = "✅", Content = result });
                Status = "就绪";
            }
            catch (Exception ex) { Messages.Add(new ChatMessage { Role = "❌", Content = ex.Message }); Status = "就绪"; }
            finally { IsBusy = false; }
        }

        void Cancel()
        {
            foreach (var m in Messages) if (m.PendingPlan != null) { m.PendingPlan = null; m.Content += "（已取消）"; }
            Status = "就绪";
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        void OnPropChanged([CallerMemberName] string? n = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(n));
    }

    public class RelayCmd : ICommand
    {
        readonly Action<object?> _exec;
        readonly Func<object?, bool>? _can;
        public RelayCmd(Action<object?> exec, Func<object?, bool>? can = null) { _exec = exec; _can = can; }
        public bool CanExecute(object? p) => _can?.Invoke(p) ?? true;
        public void Execute(object? p) => _exec(p);
        public event EventHandler? CanExecuteChanged { add => CommandManager.RequerySuggested += value; remove => CommandManager.RequerySuggested -= value; }
    }
}
