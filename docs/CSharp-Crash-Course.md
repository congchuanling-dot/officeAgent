# C# 速成指南 — 面向 Word AI 插件项目

> **目标**：看懂技术方案中的 C# 代码，能开始写 VSTO 插件  
> **策略**：只学项目里用到的，不学不用的  
> **前提**：你懂任意一门编程语言（如 JS/Python/Java）

---

## 目录

- [1. 为什么你 2 小时就能上手](#1-为什么你-2-小时就能上手)
- [2. 基础语法（10 分钟）](#2-基础语法10-分钟)
- [3. 面向对象（15 分钟）](#3-面向对象15-分钟)
- [4. 集合与 LINQ（10 分钟）](#4-集合与-linq10-分钟)
- [5. async/await（10 分钟）](#5-异步编程15-分钟)
- [6. WPF 界面（15 分钟）](#6-wpf-界面15-分钟)
- [7. Word Interop（15 分钟）](#7-word-互操作15-分钟)
- [8. 项目实战：对着技术方案写代码](#8-对着技术方案写代码)
- [A. 附录：术语速查表](#a-附录术语速查表)

---

## 1. 为什么你 2 小时就能上手

如果你会 JavaScript：

| JavaScript | C# | 区别 |
|-----------|-----|------|
| `let x = 1` | `int x = 1` | C# 需要声明类型 |
| `const obj = {}` | `var obj = new {}` | C# 是强类型 |
| `arr.map(x => x*2)` | `arr.Select(x => x*2)` | 方法名不同，思路一样 |
| `await fetch(url)` | `await httpClient.GetAsync(url)` | 异步写法几乎一样 |
| `class Dog extends Animal` | `class Dog : Animal` | 继承用 `:` 不用 `extends` |
| `interface` | `interface` | 概念一模一样 |

**核心差异只有一个**：C# 是静态类型语言，变量要声明类型。习惯了就好。

---

## 2. 基础语法（10 分钟）

### 2.1 变量声明（必须声明类型）

```csharp
// 基本类型
int count = 10;                    // 整数
double price = 9.99;               // 小数
string name = "Word文档";           // 字符串
bool isBold = true;                // 布尔

// var 自动推断类型（偷懒写法，实际类型不变）
var count = 10;     // 编译器推断为 int
var name = "Hello"; // 编译器推断为 string

// ? 表示可以为 null
int? maybeNumber = null;           // 可空整数
string? maybeNull = null;          // 可空字符串（C# 8.0+）
```

### 2.2 字符串

```csharp
// 双引号
string s1 = "你好";

// $ 字符串插值（等于 JS 的模板字符串 `${}`）
string s2 = $"今天日期是 {DateTime.Now}";

// @ 原样字符串（不转义，等于 Python 的 r""）
string path = @"C:\Users\文档\";

// """ 多行字符串（C# 11+，等于 Python 的 """ """）
string multi = """
    第一行
    第二行
    """;
```

### 2.3 条件与循环

```csharp
// if — 和所有语言一样
if (fontSize > 12)
{
    // 做点什么
}
else if (fontSize == 12)
{
    // ...
}
else
{
    // ...
}

// switch — 和所有语言一样
switch (action)
{
    case "setFont":
        // 设置字体
        break;
    case "setParagraph":
        // 设置段落
        break;
    default:
        // 未知操作
        break;
}

// for 循环
for (int i = 0; i < 10; i++)
{
    Console.WriteLine(i);
}

// foreach 循环（等于 Python 的 for x in list）
foreach (var paragraph in paragraphs)
{
    Console.WriteLine(paragraph.Text);
}

// while
while (condition)
{
    // ...
}
```

### 2.4 空值处理

```csharp
// ?. 安全导航（等于 JS 的 ?.）
string? text = paragraph?.Range?.Text;

// ?? 空值合并（等于 JS 的 ??）
string result = text ?? "默认值";   // 如果 text 是 null，用默认值

// ??= 赋默认值
text ??= "默认值";                  // 如果 text 是 null，赋默认值

// ! 断言非空（你确定不是 null 时用）
string notNull = text!;
```

### 2.5 表达式体方法（语法糖）

```csharp
// 完整写法
public int Add(int a, int b)
{
    return a + b;
}

// 简化写法（=> 糖）
public int Add(int a, int b) => a + b;

// 只读属性也用 =>
public string FullName => $"{FirstName} {LastName}";
```

---

## 3. 面向对象（15 分钟）

### 3.1 类（Class）

```csharp
// 定义一个类
public class WordCommand
{
    // 自动属性（等于 JS 的 this.action = action）
    public string Action { get; set; }          // 可读可写
    public string Target { get; init; }         // 初始化后只读
    public CommandParams Params { get; set; }
    
    // 构造函数（等于 JS 的 constructor）
    public WordCommand(string action)
    {
        Action = action;
    }
    
    // 方法
    public void Execute()
    {
        Console.WriteLine($"执行 {Action}");
    }
    
    // 只读计算属性
    public bool IsValid => !string.IsNullOrEmpty(Action);
}

// 使用类
var cmd = new WordCommand("setFont");
cmd.Execute();
```

### 3.2 接口（Interface）

接口 = 定义"必须有什么方法"，不写实现。等于其他语言中的 interface / protocol。

```csharp
// 定义接口
public interface IWordAgent
{
    // 接口中只声明方法签名，不写实现
    Task<ExecutionPlan> ProcessAsync(string userInput);
    Task ExecuteAsync(ExecutionPlan plan);
    void Undo();
}

// 实现接口（用 : 不用 implements）
public class WordAgent : IWordAgent
{
    // 必须实现接口中声明的所有方法
    public async Task<ExecutionPlan> ProcessAsync(string userInput)
    {
        // 具体实现...
    }
    
    public async Task ExecuteAsync(ExecutionPlan plan)
    {
        // 具体实现...
    }
    
    public void Undo()
    {
        // 具体实现...
    }
}
```

**为什么用接口？** 比如 `ILLMService` 接口可以有两个实现：
- `ClaudeService` — 调用 Claude API
- `OpenAIService` — 调用 OpenAI API

切换时只改一行，不用改其他代码。

### 3.3 继承（Inheritance）

```csharp
// 基类
public class WordCommandBase
{
    public string Action { get; set; }
    
    // virtual = 子类可以重写
    public virtual void Execute()
    {
        Console.WriteLine("执行命令");
    }
}

// 派生类（用 : 表示继承）
public class SetFontCommand : WordCommandBase
{
    public bool Bold { get; set; }
    public string Color { get; set; }
    
    // override = 重写基类方法
    public override void Execute()
    {
        // 先调用基类方法
        base.Execute();
        // 再做自己的事
        Console.WriteLine($"设置字体：加粗={Bold}，颜色={Color}");
    }
}
```

### 3.4 泛型（Generic）

等于 Java 的泛型、TypeScript 的泛型，就是"类型的参数"。

```csharp
// List<T> — 最常用的泛型
List<string> names = new List<string> { "标题", "正文" };
names.Add("备注");
string first = names[0];  // 直接取，不用 names.First()

// Dictionary<TKey, TValue> — 等于 JS 的 Map
Dictionary<string, int> scores = new Dictionary<string, int>
{
    { "张三", 90 },
    { "李四", 85 }
};
int zhangScore = scores["张三"];  // 90
```

### 3.5 静态成员

```csharp
public class Settings
{
    // static = 属于类本身，不属于实例
    public static string ApiKey { get; set; }
    public static int MaxRetries => 3;
    
    public static void Load()
    {
        // 不用 new Settings() 就能调用
    }
}

// 使用
Settings.ApiKey = "sk-xxx";
Settings.Load();           // 直接类名调用
int retries = Settings.MaxRetries;
```

### 3.6 记录类型（Record）

C# 9.0+ 新增，用于定义不可变的数据模型（DTO）。

```csharp
// 传统写法（类）
public class DocumentContext
{
    public int TotalParagraphs { get; set; }
    public string SelectionText { get; set; }
}

// 简写（record）— 自动生成比较、ToString 等
public record DocumentContext(
    int TotalParagraphs,
    string SelectionText,
    List<string> StyleNames
);

// 使用
var ctx = new DocumentContext(10, "这是正文", new List<string>());
Console.WriteLine($"{ctx.TotalParagraphs} 段");  // "10 段"
```

---

## 4. 集合与 LINQ（10 分钟）

### 4.1 常用集合

```csharp
// List<T> — 动态数组（最常用，等于 JS 的 Array）
var list = new List<string> { "a", "b", "c" };
list.Add("d");
list.Remove("a");

// Dictionary<TKey,TValue> — 哈希表
var dict = new Dictionary<string, int>();
dict["key1"] = 100;

// HashSet<T> — 去重集合
var set = new HashSet<string> { "a", "b", "a" };  // 只剩 "a", "b"

// IEnumerable<T> — 所有集合的父接口（惰性求值）
IEnumerable<string> filtered = list.Where(x => x.Length > 1);
```

### 4.2 LINQ — 链式数据处理

LINQ 是 C# 的数据查询语言，等于 JS 中 `map/filter/reduce` 那一套。

```csharp
var paragraphs = new List<ParagraphInfo>
{
    new("标题", "Heading 1"),
    new("正文", "Normal"),
    new("标题二", "Heading 1"),
    new("内容", "Normal"),
};

// ── 等于 JavaScript ──
// 对照表：
// JS:  arr.filter(x => ...)    → C#: list.Where(x => ...)
// JS:  arr.map(x => ...)       → C#: list.Select(x => ...)
// JS:  arr.find(x => ...)      → C#: list.FirstOrDefault(x => ...)
// JS:  arr.some(x => ...)      → C#: list.Any(x => ...)
// JS:  arr.every(x => ...)     → C#: list.All(x => ...)
// JS:  arr.sort((a,b) => ...)  → C#: list.OrderBy(x => ...)
// JS:  arr.slice(0,3)          → C#: list.Take(3)
// JS:  arr.length              → C#: list.Count
// JS:  [...new Set(arr)]       → C#: list.Distinct()

// 举例：找到所有标题段落，提取前50字
var headings = paragraphs
    .Where(p => p.Style == "Heading 1")   // 过滤
    .Select(p => p.Text[..Math.Min(50, p.Text.Length)])  // 映射（截50字）
    .ToList();                             // 转为 List

// 等价 JavaScript：
// const headings = paragraphs
//     .filter(p => p.Style === "Heading 1")
//     .map(p => p.Text.slice(0, 50))

// 更多 LINQ 方法
var firstHeading = paragraphs.FirstOrDefault(p => p.Style == "Heading 1");
var hasNormal = paragraphs.Any(p => p.Style == "Normal");
var count = paragraphs.Count(p => p.Style == "Heading 1");
var sorted = paragraphs.OrderBy(p => p.Text);
```

**记忆口诀**：LINQ 就是 C# 版的 `map/filter/reduce`，名字不一样，逻辑完全一样。

---

## 5. 异步编程（15 分钟）

### 5.1 async/await — 和 JS 几乎一样

```csharp
// 异步方法：返回 Task<T>（等于 JS 的 Promise<T>）
public async Task<string> FetchDataAsync()
{
    // await = 等待异步操作完成（等于 JS 的 await）
    var response = await httpClient.GetAsync("https://api.example.com");
    var content = await response.Content.ReadAsStringAsync();
    return content;
}

// 等价 JavaScript：
// async function fetchData() {
//     const response = await fetch("https://api.example.com");
//     const content = await response.text();
//     return content;
// }

// 无返回值的异步方法：返回 Task（等于 JS 的 Promise<void>）
public async Task ProcessAsync()
{
    await Task.Delay(1000);  // 等 1 秒（等于 JS 的 setTimeout Promise 版）
    Console.WriteLine("完成");
}
```

### 5.2 命名约定

- 异步方法名以 `Async` 结尾：`GetAsync()`, `ProcessAsync()`
- 这不是语法要求，是 C# 社区约定

### 5.3 常见异步操作

```csharp
// HTTP 请求
var response = await httpClient.PostAsJsonAsync(url, data);

// 文件读写
var text = await File.ReadAllTextAsync("path/to/file.txt");

// 延迟
await Task.Delay(1000);  // 毫秒

// 并行等待多个任务
var task1 = FetchFromApi1Async();
var task2 = FetchFromApi2Async();
await Task.WhenAll(task1, task2);  // 等于 JS 的 Promise.all
```

### 5.4 Task 不等于 void

```csharp
// ❌ 异步方法不能返回 void（除了事件处理）
// public async void BadMethod() { }

// ✅ 有返回值 → Task<T>
public async Task<int> GetCountAsync() => 42;

// ✅ 无返回值 → Task
public async Task DoSomethingAsync() { await Task.Delay(100); }
```

---

## 6. WPF 界面（15 分钟）

### 6.1 WPF 是什么

WPF 是 Windows 原生桌面 UI 框架。分两部分：
- **XAML**（读作 zammel）：写界面布局的 XML 文件
- **C# 代码隐藏**：写交互逻辑

等于：XAML 是 HTML，C# 代码隐藏是 JS。

### 6.2 XAML 基础

```xml
<!-- XAML = 用 XML 写界面布局 -->
<UserControl x:Class="WordAIAgent.TaskPane.AgentControl"
             Width="350" Height="500">
    
    <!-- Grid = 网格布局（等于 CSS Grid） -->
    <Grid>
        <!-- 定义 4 行 -->
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>    <!-- 自适应内容 -->
            <RowDefinition Height="*"/>       <!-- 填充剩余空间 -->
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Row 0: 标题栏 -->
        <Border Grid.Row="0" Background="#2B579A" Padding="12">
            <TextBlock Text="AI 助手" Foreground="White" FontSize="16"/>
        </Border>
        
        <!-- Row 1: 内容区域（自动滚动） -->
        <ScrollViewer Grid.Row="1">
            <StackPanel>
                <TextBlock Text="这里是消息列表" TextWrapping="Wrap"/>
            </StackPanel>
        </ScrollViewer>
        
        <!-- Row 2: 输入框 -->
        <TextBox Grid.Row="2" 
                 Text="{Binding UserInput}"    <!-- 双向绑定到 ViewModel -->
                 AcceptsReturn="True"/>
        
        <!-- Row 3: 按钮 -->
        <StackPanel Grid.Row="3" Orientation="Horizontal">
            <Button Content="发送" Command="{Binding SendCommand}"/>
            <Button Content="撤销" Command="{Binding UndoCommand}"/>
        </StackPanel>
    </Grid>
</UserControl>
```

### 6.3 常用布局控件

```xml
<!-- StackPanel — 垂直或水平堆叠（等于 Flexbox 的一维方向） -->
<StackPanel Orientation="Vertical">  <!-- 或 Horizontal -->
    <TextBlock Text="第一行"/>
    <TextBlock Text="第二行"/>
</StackPanel>

<!-- Grid — 网格（等于 CSS Grid） -->
<Grid>
    <Grid.ColumnDefinitions>
        <ColumnDefinition Width="100"/>   <!-- 固定宽度 -->
        <ColumnDefinition Width="*"/>     <!-- 弹性宽度 -->
    </Grid.ColumnDefinitions>
    <!-- 第 0 列 -->
    <TextBlock Grid.Column="0" Text="标签"/>
    <!-- 第 1 列 -->
    <TextBox Grid.Column="1" Text="值"/>
</Grid>

<!-- DockPanel — 停靠布局 -->
<DockPanel>
    <Button DockPanel.Dock="Top" Content="顶部"/>
    <Button DockPanel.Dock="Bottom" Content="底部"/>
    <TextBlock Text="填充"/>
</DockPanel>

<!-- Border — 带边框/背景的容器 -->
<Border BorderBrush="Gray" BorderThickness="1" CornerRadius="8" 
        Background="White" Padding="12">
    <TextBlock Text="内容"/>
</Border>
```

### 6.4 常用控件速查

```xml
<!-- 文本显示 -->
<TextBlock Text="静态文本"/>
<TextBlock Text="{Binding PropertyName}"/>              <!-- 绑定到 ViewModel -->

<!-- 文本输入 -->
<TextBox Text="{Binding UserInput, UpdateSourceTrigger=PropertyChanged}"/>

<!-- 按钮 -->
<Button Content="确认" Command="{Binding ConfirmCommand}"/>
<Button Content="取消" Click="CancelButton_Click"/>     <!-- 事件处理方式 -->

<!-- 复选框 -->
<CheckBox Content="区分大小写" IsChecked="{Binding MatchCase}"/>

<!-- 下拉框 -->
<ComboBox ItemsSource="{Binding Options}" SelectedItem="{Binding Selected}"/>

<!-- 列表 -->
<ListBox ItemsSource="{Binding Items}" SelectedItem="{Binding SelectedItem}"/>

<!-- 滚动区域 -->
<ScrollViewer VerticalScrollBarVisibility="Auto">
    <!-- 内容 -->
</ScrollViewer>
```

### 6.5 MVVM 模式（重点）

MVVM = Model-View-ViewModel，是 WPF 的标准架构模式。

```
View (XAML)  ←──绑定──→  ViewModel (C#)  ────→  Model (数据)
     界面                   逻辑+状态                数据模型
```

**核心：数据绑定**。View 不直接操作控件，而是绑定到 ViewModel 的属性。

```csharp
// MainViewModel.cs — ViewModel 示例
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

public partial class MainViewModel : ObservableObject
{
    // 可绑定属性：值变化时自动通知 UI 更新
    [ObservableProperty]
    private string _userInput = "";
    
    [ObservableProperty]
    private string _status = "就绪";
    
    [ObservableProperty]
    private bool _isProcessing = false;
    
    // 命令：绑定到按钮的 Command
    [RelayCommand]
    private async Task SendAsync()
    {
        IsProcessing = true;
        Status = "AI 分析中...";
        
        try
        {
            // 调用 Agent
            var plan = await _agent.ProcessAsync(UserInput);
            // 更新 UI...
            Status = "就绪";
        }
        finally
        {
            IsProcessing = false;
        }
    }
    
    [RelayCommand]
    private void Undo()
    {
        _agent.Undo();
    }
}
```

**CommunityToolkit.Mvvm** 帮你自动生成样板代码：
- `[ObservableProperty]` → 自动生成通知 UI 的属性
- `[RelayCommand]` → 自动生成 `SendCommand` 用于 XAML 绑定
- 不需要手写 INotifyPropertyChanged

```xml
<!-- View (XAML) 中绑定 -->
<TextBox Text="{Binding UserInput}"/>
<Button Command="{Binding SendCommand}" Content="发送"/>
<TextBlock Text="{Binding Status}"/>
```

---

## 7. Word 互操作（15 分钟）

### 7.1 Word 对象模型（核心概念）

Word 通过 COM 对象模型暴露所有功能，操作 Word 就是操作这些对象。

```
Application                  ← Word 应用程序
  └── Documents              ← 所有打开的文档
       └── Document          ← 一个文档
            ├── Paragraphs   ← 所有段落
            │    └── Paragraph → 一个段落
            │         └── Range      ← 段落范围
            │              ├── Font       (字体)
            │              ├── ParagraphFormat (段落格式)
            │              └── Style      (样式)
            ├── Tables       ← 所有表格
            │    └── Table → 一个表格
            ├── Sections     ← 所有节
            ├── Content      ← 全文 Range
            ├── PageSetup    ← 页面设置
            └── Styles       ← 所有样式
```

### 7.2 获取对象

```csharp
// 在 VSTO 插件中，通过 ThisAddIn 获取 Word 实例
// ThisAddIn.cs
public partial class ThisAddIn
{
    // Application 就是 Word 应用程序对象
    public Microsoft.Office.Interop.Word.Application WordApp => this.Application;
    
    // 当前活动文档
    public Document ActiveDoc => this.Application.ActiveDocument;
}
```

### 7.3 读取文档内容

```csharp
// 读取文档信息
Document doc = Globals.ThisAddIn.Application.ActiveDocument;

// 总段落数
int count = doc.Paragraphs.Count;

// 遍历所有段落
foreach (Paragraph paragraph in doc.Paragraphs)
{
    // 段落文本
    string text = paragraph.Range.Text;
    
    // 段落样式名称
    string styleName = ((Style)paragraph.get_Style()).NameLocal;
    
    // 段落对齐
    WdParagraphAlignment alignment = paragraph.ParagraphFormat.Alignment;
    
    // 字体信息
    Font font = paragraph.Range.Font;
    string fontName = font.Name;
    float fontSize = font.Size;
    int isBold = font.Bold;  // -1=true, 0=false (COM 返回值)
}

// 获取当前选区
Selection selection = doc.Application.Selection;
string selectedText = selection.Text;

// 获取全文
Range wholeDoc = doc.Content;
```

### 7.4 修改文档内容

```csharp
// ── 修改字体 ──
Range range = doc.Paragraphs[1].Range;  // 第1段
range.Font.Bold = -1;                    // -1 表示 true (COM 约定)
range.Font.Name = "微软雅黑";
range.Font.Size = 14;
range.Font.Color = WdColor.wdColorBlue;

// ── 修改段落格式 ──
range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
range.ParagraphFormat.FirstLineIndent = 28;  // 约两个中文字符

// ── 应用样式 ──
range.set_Style("Heading 1");

// ── 插入文本 ──
Selection sel = doc.Application.Selection;
sel.Range.InsertAfter("插入的内容");

// 在文档末尾插入
Range endRange = doc.Content;
endRange.Collapse(WdCollapseDirection.wdCollapseEnd);
endRange.InsertAfter("\n新内容");

// ── 查找替换 ──
Find find = doc.Content.Find;
find.Text = "北京";
find.Replacement.Text = "上海";
find.Execute(Replace: WdReplace.wdReplaceAll);

// ── 插入表格 ──
Range location = doc.Application.Selection.Range;
Table table = doc.Tables.Add(location, 3, 4);  // 3行4列
table.Cell(1, 1).Range.Text = "表头1";
table.Cell(1, 2).Range.Text = "表头2";

// ── 页面设置 ──
doc.PageSetup.TopMargin = 56.7f;     // 2cm (单位：磅)
doc.PageSetup.LeftMargin = 56.7f;
doc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
```

### 7.5 COM 互操作的坑

```csharp
// 坑1：布尔值用 int 返回（-1=true, 0=false）
int bold = range.Font.Bold;
bool isBold = bold == -1;  // 需要手动转换

// 坑2：索引从 1 开始（不是 0！）
Paragraph first = doc.Paragraphs[1];  // 不是 [0]！

// 坑3：遍历时不要用 foreach + Count，用 for 反向遍历
for (int i = doc.Paragraphs.Count; i >= 1; i--)
{
    Paragraph p = doc.Paragraphs[i];
    // 处理...
}

// 坑4：用完要释放 COM 对象（避免内存泄漏）
// 简单规则：操作完大文档后调用
Marshal.ReleaseComObject(comObject);

// 坑5：cm 转 磅（Word 内单位是磅）
float CmToPoints(float cm) => cm * 28.35f;
float PointsToCm(float points) => points / 28.35f;
```

---

## 8. 对着技术方案写代码

现在打开 [Word-AI-Plugin-Technical-Design.md](Word-AI-Plugin-Technical-Design.md)，你已经能看懂所有关键代码片段了：

### 你能看懂的

| 技术方案中的代码 | 对应本指南章节 |
|----------------|--------------|
| `public class WordAgent` | §3.1 类 |
| `public interface IWordAgent` | §3.2 接口 |
| `public class SetFontCommand : WordCommandBase` | §3.3 继承 |
| `public async Task<ExecutionPlan> ProcessAsync()` | §5 异步 |
| `List<string>` / `Dictionary<string, int>` | §3.4 泛型 |
| `paragraphs.Where(p => ...)` | §4.2 LINQ |
| `<Grid><Button Content="发送"/></Grid>` | §6.2 XAML |
| `{Binding UserInput}` | §6.5 MVVM |
| `range.Font.Bold = -1` | §7.4 Word Interop |
| `doc.Paragraphs[i].Range.Text` | §7.3 Word 读取 |

### 从哪开始写

```
Step 1: 在 VS 2022 中新建 VSTO Word Add-in 项目
Step 2: 打开 ThisAddIn.cs，看 Startup 方法
Step 3: 写一个最简单的测试：
         this.Application.ActiveDocument.Content.Text = "Hello!";
         运行 → Word 自动打开 → 文档内容被替换
Step 4: 开始按技术方案搭建结构
```

---

## A. 附录：术语速查表

| 术语 | 解释 | 等于 JS/TS 里的 |
|------|------|----------------|
| `class` | 类 | `class` |
| `interface` | 接口 | `interface` |
| `struct` | 值类型结构体 | 无（C# 特有） |
| `record` | 不可变数据类 | 类似 `Object.freeze` |
| `namespace` | 命名空间 | 类似 ES Module |
| `using` | 引入命名空间 | `import` |
| `new` | 创建实例 | `new` |
| `var` | 自动推断类型 | `let`/`const` |
| `async/await` | 异步 | 完全一样 |
| `Task<T>` | 异步任务 | `Promise<T>` |
| `List<T>` | 动态数组 | `Array<T>` |
| `Dictionary<K,V>` | 字典/哈希表 | `Map<K,V>` |
| `IEnumerable<T>` | 可迭代集合 | `Iterable<T>` |
| `LINQ` | 数据查询 | `map/filter/reduce` |
| `=>` | Lambda 表达式 | `=>` 箭头函数 |
| `XAML` | 界面标记语言 | 等于 HTML |
| `Binding` | 数据绑定 | 等于 React state |
| `ViewModel` | 视图模型 | 等于 React 组件的 state+handler |
| `Command` | WPF 命令 | 等于 onClick handler |
| `COM` | 组件对象模型 | 等于 Windows 版 RPC |
| `Interop` | 互操作 | 调用外部 DLL |
| `this` | 当前实例 | `this` |
| `base` | 基类 | `super` |
| `null` | 空值 | `null`/`undefined` |
| `?.` | 安全导航 | `?.` |
| `??` | 空值合并 | `??` |
| `!` (后缀) | 断言非空 | `!` (TS) |
| `partial class` | 分部类（可分文件定义） | 无直接等价 |

---

> **下一步**：打开 [Word-AI-Plugin-Technical-Design.md](Word-AI-Plugin-Technical-Design.md)，对照本指南阅读。  
> 两篇文档配合，技术方案里的每一行代码你都能看懂了。
