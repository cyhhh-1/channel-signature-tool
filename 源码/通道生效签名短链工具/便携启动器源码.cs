using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

/// <summary>
/// 版本号由 生成安装包.ps1 在构建时写入程序集元数据，运行时只读。
/// 源码中不再保存版本常量，避免 C#、NSIS 和打包脚本三处手工同步。
/// </summary>
internal static class AppInfo
{
    private static string ReadMetadata(string key)
    {
        object[] attributes = typeof(AppInfo).Assembly.GetCustomAttributes(
            typeof(AssemblyMetadataAttribute), false);
        foreach (AssemblyMetadataAttribute attribute in attributes)
        {
            if (string.Equals(attribute.Key, key, StringComparison.Ordinal) &&
                !string.IsNullOrEmpty(attribute.Value))
                return attribute.Value;
        }
        return string.Empty;
    }

    internal static string ProductVersion
    {
        get
        {
            string value = ReadMetadata("ProductVersion");
            return string.IsNullOrEmpty(value) ? "0.0.0" : value;
        }
    }

    internal static string RuntimeVersion
    {
        get
        {
            string value = ReadMetadata("RuntimeVersion");
            return string.IsNullOrEmpty(value) ? ProductVersion : value;
        }
    }

    internal static string MappingSeedVersion
    {
        get
        {
            string value = ReadMetadata("MappingSeedVersion");
            return string.IsNullOrEmpty(value) ? "mapping-seed" : value;
        }
    }
}

internal static class Program
{
    private static string RuntimeVersion
    {
        get { return AppInfo.RuntimeVersion; }
    }

    private const string RuntimeResource = "runtime_payload.zip";

    private static string MappingSeedVersion
    {
        get { return AppInfo.MappingSeedVersion; }
    }

    internal static string ApplicationDirectory
    {
        get { return AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar); }
    }

    internal static string ResourcesDirectory
    {
        get { return Path.Combine(EnsureRuntime(), "resources"); }
    }

    internal static string MappingFile
    {
        get
        {
            string overrideHome = Environment.GetEnvironmentVariable("CHANNEL_TOOL_DATA_HOME");
            string home = string.IsNullOrWhiteSpace(overrideHome)
                ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "通道生效签名短链程序")
                : overrideHome;
            Directory.CreateDirectory(home);
            string mapping = Path.Combine(home, "通道编号及对应主码号.xlsx");
            string seed = Path.Combine(ResourcesDirectory, "通道编号及对应主码号.xlsx");
            string seedMarker = Path.Combine(home, ".mapping-seed-" + MappingSeedVersion);
            if (!File.Exists(seedMarker))
            {
                if (File.Exists(mapping))
                {
                    string backup = Path.Combine(
                        home, "通道编号及对应主码号_更新前备份_" + MappingSeedVersion + ".xlsx");
                    File.Copy(mapping, backup, true);
                }
                File.Copy(seed, mapping, true);
                File.WriteAllText(seedMarker, MappingSeedVersion, Encoding.UTF8);
            }
            return mapping;
        }
    }

    [STAThread]
    private static int Main(string[] args)
    {
        EnableHighDpiRendering();
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        int mappingSnapshotIndex = Array.IndexOf(args, "--mapping-snapshot");
        if (mappingSnapshotIndex >= 0 && mappingSnapshotIndex + 1 < args.Length)
        {
            using (MappingForm form = new MappingForm(false))
            using (Bitmap bitmap = new Bitmap(form.ClientSize.Width, form.ClientSize.Height))
            {
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(-32000, -32000);
                form.Show();
                Application.DoEvents();
                form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.ClientSize));
                bitmap.Save(args[mappingSnapshotIndex + 1], System.Drawing.Imaging.ImageFormat.Png);
                form.Hide();
                return 0;
            }
        }

        int snapshotIndex = Array.IndexOf(args, "--ui-snapshot");
        if (snapshotIndex >= 0 && snapshotIndex + 1 < args.Length)
        {
            using (MainForm form = new MainForm())
            using (Bitmap bitmap = new Bitmap(form.ClientSize.Width, form.ClientSize.Height))
            {
                if (snapshotIndex + 2 < args.Length) form.SetSnapshotMode(args[snapshotIndex + 2]);
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(-32000, -32000);
                form.Show();
                Application.DoEvents();
                form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.ClientSize));
                bitmap.Save(args[snapshotIndex + 1], System.Drawing.Imaging.ImageFormat.Png);
                form.Hide();
                return 0;
            }
        }

        // 隐藏测试入口：只创建窗口，不弹出或进入消息循环。
        if (args.Contains("--ui-smoke-test"))
        {
            using (MainForm form = new MainForm())
            {
                return form.Text == "通道生效签名短链程序" ? 0 : 1;
            }
        }

        // 保留命令行入口，便于发布前自动验证完整的便携运行环境。
        if (args.Contains("--cli"))
        {
            try
            {
                PythonResult result = ExecutePython(args);
                if (!string.IsNullOrWhiteSpace(result.Output)) Console.Out.WriteLine(result.Output);
                if (!string.IsNullOrWhiteSpace(result.Error)) Console.Error.WriteLine(result.Error);
                return result.ExitCode;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.GetType().FullName);
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                return 2;
            }
        }

        Application.Run(new MainForm());
        return 0;
    }

    [DllImport("user32.dll", EntryPoint = "SetProcessDpiAwarenessContext")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);

    internal static void EnableHighDpiRendering()
    {
        try
        {
            // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
            if (SetProcessDpiAwarenessContext(new IntPtr(-4))) return;
        }
        catch (EntryPointNotFoundException) { }
        catch (DllNotFoundException) { }

        try { SetProcessDPIAware(); }
        catch (EntryPointNotFoundException) { }
        catch (DllNotFoundException) { }
    }

    internal static void ApplyWindowIdentity(Form form)
    {
        form.AutoScaleDimensions = new SizeF(96F, 96F);
        form.AutoScaleMode = AutoScaleMode.Dpi;
        try
        {
            using (Icon executableIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath))
            {
                if (executableIcon != null) form.Icon = (Icon)executableIcon.Clone();
            }
        }
        catch { }
        form.HandleCreated += delegate
        {
            try
            {
                const int captionColorAttribute = 35;
                const int textColorAttribute = 36;
                int white = 0x00FFFFFF;
                int black = 0x00000000;
                DwmSetWindowAttribute(form.Handle, captionColorAttribute, ref white, sizeof(int));
                DwmSetWindowAttribute(form.Handle, textColorAttribute, ref black, sizeof(int));
            }
            catch (DllNotFoundException) { }
            catch (EntryPointNotFoundException) { }
        };
    }

    internal static PythonResult ExecutePython(IEnumerable<string> arguments)
    {
        return ExecutePython(arguments, null);
    }

    internal static PythonResult ExecutePython(IEnumerable<string> arguments, Action<string> progressLineCallback)
    {
        string runtime = EnsureRuntime();
        string python = Path.Combine(runtime, "python.exe");
        string app = Path.Combine(runtime, "app.py");

        List<string> completeArguments = new List<string>();
        completeArguments.Add(Quote(app));
        completeArguments.AddRange(arguments.Select(Quote));

        // 直接设置当前进程环境，子进程会继承；避免某些 Windows 环境同时存在 Path/PATH 时冲突。
        Environment.SetEnvironmentVariable("CHANNEL_TOOL_HOME", ResourcesDirectory);
        Environment.SetEnvironmentVariable("CHANNEL_TOOL_MAPPING_FILE", MappingFile);
        Environment.SetEnvironmentVariable("PYTHONIOENCODING", "utf-8");
        Environment.SetEnvironmentVariable("PYTHONUTF8", "1");

        ProcessStartInfo info = new ProcessStartInfo();
        info.FileName = python;
        info.Arguments = string.Join(" ", completeArguments.ToArray());
        info.WorkingDirectory = ApplicationDirectory;
        info.UseShellExecute = false;
        info.CreateNoWindow = true;
        info.RedirectStandardOutput = true;
        info.RedirectStandardError = true;
        info.StandardOutputEncoding = Encoding.UTF8;
        info.StandardErrorEncoding = Encoding.UTF8;

        using (Process process = Process.Start(info))
        {
            StringBuilder output = new StringBuilder();
            StringBuilder error = new StringBuilder();
            object syncRoot = new object();
            process.OutputDataReceived += delegate(object sender, DataReceivedEventArgs eventArgs)
            {
                if (eventArgs.Data == null) return;
                lock (syncRoot) output.AppendLine(eventArgs.Data);
            };
            process.ErrorDataReceived += delegate(object sender, DataReceivedEventArgs eventArgs)
            {
                if (eventArgs.Data == null) return;
                if (progressLineCallback != null && eventArgs.Data.StartsWith("@@TABLE_PROGRESS\t", StringComparison.Ordinal))
                {
                    progressLineCallback(eventArgs.Data);
                    return;
                }
                lock (syncRoot) error.AppendLine(eventArgs.Data);
            };
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            process.WaitForExit();
            // .NET Framework 的异步输出事件可能略晚于进程退出，再等待一次确保全部行已接收。
            process.WaitForExit();
            lock (syncRoot)
                return new PythonResult(process.ExitCode, output.ToString(), error.ToString());
        }
    }

    private static string EnsureRuntime()
    {
        string runtime = Path.Combine(Path.GetTempPath(), "通道生效签名短链程序运行环境", RuntimeVersion);
        string marker = Path.Combine(runtime, ".ready");
        string mutexName = "Local\\ChannelImportTool-" + RuntimeVersion;

        using (Mutex mutex = new Mutex(false, mutexName))
        {
            if (!mutex.WaitOne(TimeSpan.FromMinutes(2)))
                throw new InvalidOperationException("等待程序运行环境准备超时，请重新运行程序。");

            try
            {
                if (File.Exists(marker) && File.Exists(Path.Combine(runtime, "python.exe")) &&
                    File.Exists(Path.Combine(runtime, "app.py")))
                    return runtime;

                Directory.CreateDirectory(runtime);
                using (Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(RuntimeResource))
                {
                    if (resource == null)
                        throw new InvalidOperationException("程序内置运行环境缺失，请重新下载完整工具包。");

                    using (ZipArchive archive = new ZipArchive(resource, ZipArchiveMode.Read))
                    {
                        string safeRoot = Path.GetFullPath(runtime) + Path.DirectorySeparatorChar;
                        foreach (ZipArchiveEntry entry in archive.Entries)
                        {
                            string destination = Path.GetFullPath(Path.Combine(runtime, entry.FullName));
                            if (!destination.StartsWith(safeRoot, StringComparison.OrdinalIgnoreCase))
                                throw new InvalidDataException("内置运行环境包含无效路径。");

                            if (string.IsNullOrEmpty(entry.Name))
                            {
                                Directory.CreateDirectory(destination);
                                continue;
                            }

                            string parent = Path.GetDirectoryName(destination);
                            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
                            using (Stream input = entry.Open())
                            using (FileStream output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None))
                                input.CopyTo(output);
                        }
                    }
                }

                File.WriteAllText(marker, RuntimeVersion, Encoding.UTF8);
                return runtime;
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }
    }

    private static string Quote(string value)
    {
        if (string.IsNullOrEmpty(value)) return "\"\"";
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}

internal sealed class PythonResult
{
    internal PythonResult(int exitCode, string output, string error)
    {
        ExitCode = exitCode;
        Output = output ?? string.Empty;
        Error = error ?? string.Empty;
    }

    internal int ExitCode { get; private set; }
    internal string Output { get; private set; }
    internal string Error { get; private set; }
}

internal static class UiFonts
{
    internal static Font Chinese(float points)
    {
        return Chinese(points, FontStyle.Regular);
    }

    internal static Font Chinese(float points, FontStyle style)
    {
        return new Font("Microsoft YaHei UI", points, style, GraphicsUnit.Point);
    }

    internal static Font Latin(float points, FontStyle style)
    {
        return new Font("Segoe UI", points, style, GraphicsUnit.Point);
    }
}

internal static class UiTheme
{
    internal static readonly Color Canvas = Color.White;
    internal static readonly Color Surface = Color.White;
    internal static readonly Color SurfaceMuted = Color.FromArgb(250, 250, 250);
    internal static readonly Color SurfaceBlue = Color.FromArgb(245, 245, 245);
    internal static readonly Color Border = Color.FromArgb(222, 222, 222);
    internal static readonly Color BorderStrong = Color.FromArgb(198, 198, 198);
    internal static readonly Color Text = Color.FromArgb(23, 23, 23);
    internal static readonly Color TextSecondary = Color.FromArgb(64, 64, 64);
    internal static readonly Color TextMuted = Color.FromArgb(115, 115, 115);
    internal static readonly Color Primary = Color.FromArgb(23, 23, 23);
    internal static readonly Color PrimaryHover = Color.FromArgb(38, 38, 38);
    internal static readonly Color PrimarySoft = Color.FromArgb(245, 245, 245);
    internal static readonly Color Success = Color.FromArgb(22, 130, 82);
    internal static readonly Color SuccessSoft = Color.FromArgb(240, 253, 244);
    internal static readonly Color Warning = Color.FromArgb(161, 98, 7);
    internal static readonly Color Danger = Color.FromArgb(185, 28, 28);
    internal static readonly Color DangerSoft = Color.FromArgb(254, 242, 242);
    internal const int SidebarWidth = 250;
    internal const int NavigationGapWidth = 12;
    internal const int Radius = 8;

    internal static void StyleTextBox(TextBox box)
    {
        box.BorderStyle = BorderStyle.FixedSingle;
        box.BackColor = Surface;
        box.ForeColor = Text;
        box.Font = UiFonts.Chinese(10F, FontStyle.Bold);
    }

    internal static void StyleCombo(ComboBox combo)
    {
        combo.FlatStyle = FlatStyle.Standard;
        combo.BackColor = Surface;
        combo.ForeColor = Text;
        combo.Font = UiFonts.Chinese(10F, FontStyle.Bold);
    }

    internal static void StyleInput(AntdUI.Input input)
    {
        input.BackColor = Surface;
        input.ForeColor = Text;
        input.BorderColor = BorderStrong;
        input.BorderHover = Color.FromArgb(163, 163, 163);
        input.BorderActive = Primary;
        input.PlaceholderColor = TextMuted;
        input.Radius = 7;
        input.BorderWidth = 1F;
        input.Font = UiFonts.Chinese(10F, FontStyle.Bold);
        input.Padding = new Padding(10, 0, 10, 0);
    }

    internal static void StyleSelect(AntdUI.Select select)
    {
        StyleInput(select);
        select.DropDownRadius = 8;
        select.DropDownPadding = new Size(6, 6);
        select.ListAutoWidth = false;
    }

    internal static void StyleSecondaryButton(AntdUI.Button button)
    {
        button.Type = AntdUI.TTypeMini.Default;
        button.BackColor = Surface;
        button.BackHover = SurfaceBlue;
        button.ForeColor = Text;
        button.ForeHover = Text;
        button.DefaultBorderColor = BorderStrong;
        button.Radius = 7;
        button.WaveSize = 0;
        button.Font = UiFonts.Chinese(9.5F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
    }
}

internal sealed class MainForm : Form
{
    private readonly RadioButton shortlinkMode = new RadioButton();
    private readonly RadioButton signatureMode = new RadioButton();
    private readonly RadioButton mixedMode = new RadioButton();
    private readonly RadioButton selfMobileMode = new RadioButton();
    private readonly RadioButton selfTelecomMode = new RadioButton();
    private readonly RadioButton tableToolsMode = new RadioButton();
    private readonly RadioButton signatureConsistencyMode = new RadioButton();
    private readonly RadioButton portAllocationMode = new RadioButton();
    private readonly AntdUI.Input sourceBox = new AntdUI.Input();
    private readonly AntdUI.Input masterBox = new AntdUI.Input();
    private readonly AntdUI.Input channelBox = new AntdUI.Input();
    private readonly AntdUI.Input startPortBox = new AntdUI.Input();
    private readonly AntdUI.Select operatorBox = new AntdUI.Select();
    private readonly AntdUI.Select businessBox = new AntdUI.Select();
    private readonly AntdUI.Select importantBox = new AntdUI.Select();
    private readonly AntdUI.Select tableActionBox = new AntdUI.Select();
    private readonly AntdUI.Select signatureMatchModeBox = new AntdUI.Select();
    private readonly AntdUI.Input maxRowsBox = new AntdUI.Input();
    private readonly Label modeHint = new Label();
    private readonly Label templateStatus = new Label();
    private readonly Label progressStatus = new Label();
    private readonly ProgressBar tableProgress = new ProgressBar();
    private readonly AntdUI.Button browseButton = new AntdUI.Button();
    private readonly AntdUI.Button browseMasterButton = new AntdUI.Button();
    private readonly AntdUI.Input outputDirectoryBox = new AntdUI.Input();
    private readonly AntdUI.Button browseOutputButton = new AntdUI.Button();
    private readonly AntdUI.Button processButton = new AntdUI.Button();
    private readonly AntdUI.Button downloadPortTemplateButton = new AntdUI.Button();
    private readonly AntdUI.Button cancelButton = new AntdUI.Button();
    private readonly BackgroundWorker worker = new BackgroundWorker();
    private readonly System.Windows.Forms.Timer progressTimer = new System.Windows.Forms.Timer();
    private readonly Stopwatch operationTimer = new Stopwatch();
    private readonly Stopwatch rateTimer = new Stopwatch();
    private string currentCancelFile = string.Empty;
    private string lastProgressPhase = string.Empty;
    private string lastProgressMessage = string.Empty;
    private long lastProgressCompleted;
    private long lastProgressTotal;
    private bool cancelRequested;
    private Label sourceLabel;
    private Label masterLabel;
    private Label channelLabel;
    private Label startPortLabel;
    private Label operatorLabel;
    private Label businessLabel;
    private Label importantLabel;
    private Label tableActionLabel;
    private Label signatureMatchModeLabel;
    private Label maxRowsLabel;
    private Label formatNote;
    private readonly Dictionary<string, NavigationButton> modernNavButtons = new Dictionary<string, NavigationButton>();
    private Label modernPageTitle;
    private Label modernBreadcrumb;
    private Label modernEyebrow;
    private Control modernLocalBadge;
    private Label modernSourceDropTitle;
    private Label modernMasterDropTitle;
    private Panel modernMainArea;
    private Panel modernTopbar;
    private Panel modernContent;
    private Panel modernInputCard;
    private Panel modernStatusCard;
    private Panel modernFooterCard;
    private Panel modernSourceDropZone;
    private Panel modernMasterDropZone;
    private Panel modernSidebar;
    private Panel modernNavigationGap;
    private NavigationButton modernGuideButton;
    private Label modernReadyStatus;
    private Label modernMappingStatus;
    private Panel modernParameterCard;
    private Panel modernResourceCard;
    private Panel modernTaskCard;
    private Panel modernSecurityCard;
    private FlowLayoutPanel modernResourcesList;
    private Label modernActionHint;
    private Label modernSourceDropMeta;
    private Label modernMasterDropMeta;

    internal MainForm()
    {
        Text = "通道生效签名短链程序";
        Program.ApplyWindowIdentity(this);
        ClientSize = new Size(860, 720);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = true;
        Font = UiFonts.Chinese(10F);
        BackColor = Color.FromArgb(244, 247, 251);

        Panel header = new Panel();
        header.SetBounds(0, 0, 860, 96);
        header.BackColor = Color.FromArgb(19, 43, 70);
        Controls.Add(header);
        Label title = MakeLabel("通道生效签名短链程序", 27, 28, 570, 34);
        title.Font = new Font("Microsoft YaHei UI", 18F, FontStyle.Bold);
        title.ForeColor = Color.White;
        header.Controls.Add(title);

        Label subtitle = MakeLabel("通道处理 · 表格整理 · 子端口分配", 29, 64, 720, 23);
        subtitle.ForeColor = Color.FromArgb(180, 202, 224);
        header.Controls.Add(subtitle);
        AntdUI.Button manageButton = new AntdUI.Button();
        manageButton.Text = "主码号管理";
        manageButton.SetBounds(690, 35, 140, 42);
        manageButton.Click += delegate
        {
            if (worker.IsBusy) return;
            using (MappingForm form = new MappingForm()) form.ShowDialog(this);
            UpdateModeDisplay();
        };
        header.Controls.Add(manageButton);
        Panel accent = new Panel();
        accent.SetBounds(0, 92, 860, 4);
        accent.BackColor = Color.FromArgb(27, 156, 178);
        header.Controls.Add(accent);

        GroupBox modeGroup = new GroupBox();
        modeGroup.Text = "处理方式";
        modeGroup.SetBounds(25, 106, 810, 130);
        modeGroup.BackColor = Color.White;
        modeGroup.ForeColor = Color.FromArgb(32, 57, 83);
        modeGroup.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
        Controls.Add(modeGroup);

        shortlinkMode.Text = "生效短链";
        shortlinkMode.SetBounds(24, 28, 145, 28);
        shortlinkMode.Checked = true;
        signatureMode.Text = "生效签名";
        signatureMode.SetBounds(190, 28, 145, 28);
        mixedMode.Text = "混合表自动分流";
        mixedMode.SetBounds(355, 28, 190, 28);
        selfMobileMode.Text = "自家移动";
        selfMobileMode.SetBounds(24, 60, 145, 28);
        selfTelecomMode.Text = "自家电信";
        selfTelecomMode.SetBounds(190, 60, 145, 28);
        tableToolsMode.Text = "拆分/合并表";
        tableToolsMode.SetBounds(355, 60, 190, 28);
        signatureConsistencyMode.Text = "签名一致性判定";
        signatureConsistencyMode.SetBounds(560, 60, 220, 28);
        portAllocationMode.Text = "子端口分配";
        portAllocationMode.SetBounds(560, 28, 220, 28);
        modeGroup.Controls.Add(shortlinkMode);
        modeGroup.Controls.Add(signatureMode);
        modeGroup.Controls.Add(mixedMode);
        modeGroup.Controls.Add(selfMobileMode);
        modeGroup.Controls.Add(selfTelecomMode);
        modeGroup.Controls.Add(tableToolsMode);
        modeGroup.Controls.Add(signatureConsistencyMode);
        modeGroup.Controls.Add(portAllocationMode);
        foreach (RadioButton modeChoice in new RadioButton[] {
            shortlinkMode, signatureMode, mixedMode, selfMobileMode,
            selfTelecomMode, tableToolsMode, signatureConsistencyMode, portAllocationMode })
        {
            modeChoice.BackColor = Color.White;
            modeChoice.ForeColor = Color.FromArgb(38, 64, 91);
            modeChoice.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold);
            modeChoice.Cursor = Cursors.Hand;
        }

        modeHint.SetBounds(24, 94, 760, 25);
        modeHint.ForeColor = Color.FromArgb(55, 90, 145);
        modeGroup.Controls.Add(modeHint);

        sourceLabel = MakeLabel("客户文件：", 28, 246, 105, 28);
        Controls.Add(sourceLabel);
        // 左侧标签的可见区域会随 Windows DPI 缩放而扩大；预留间隔，
        // 防止在 125% 等缩放比例下遮住输入框的左边框。
        sourceBox.SetBounds(140, 241, 575, 31);
        UiTheme.StyleInput(sourceBox);
        sourceBox.AllowDrop = true;
        sourceBox.DragEnter += InputDragEnter;
        sourceBox.DragOver += InputDragOver;
        sourceBox.DragLeave += InputDragLeave;
        sourceBox.DragDrop += SourceDragDrop;
        Controls.Add(sourceBox);
        browseButton.Text = "选择文件";
        browseButton.Radius = 8;
        browseButton.SetBounds(725, 240, 105, 33);
        browseButton.Click += BrowseFile;
        Controls.Add(browseButton);

        masterLabel = MakeLabel("历史总表：", 28, 296, 105, 28);
        Controls.Add(masterLabel);
        masterBox.SetBounds(140, 291, 575, 31);
        UiTheme.StyleInput(masterBox);
        masterBox.ReadOnly = true;
        masterBox.AllowDrop = true;
        masterBox.DragEnter += InputDragEnter;
        masterBox.DragOver += InputDragOver;
        masterBox.DragLeave += InputDragLeave;
        masterBox.DragDrop += MasterDragDrop;
        Controls.Add(masterBox);
        browseMasterButton.Text = "选择总表";
        browseMasterButton.Radius = 8;
        browseMasterButton.SetBounds(725, 290, 105, 33);
        browseMasterButton.Click += BrowseMasterFile;
        Controls.Add(browseMasterButton);

        channelLabel = MakeLabel("通道编号：", 28, 296, 105, 28);
        Controls.Add(channelLabel);
        channelBox.SetBounds(140, 291, 260, 31);
        UiTheme.StyleInput(channelBox);
        Controls.Add(channelBox);

        startPortLabel = MakeLabel("起始子端口：", 440, 296, 115, 28);
        Controls.Add(startPortLabel);
        startPortBox.SetBounds(565, 291, 265, 31);
        UiTheme.StyleInput(startPortBox);
        startPortBox.Text = "1000";
        Controls.Add(startPortBox);

        operatorLabel = MakeLabel("运营商：", 28, 346, 105, 28);
        Controls.Add(operatorLabel);
        ConfigureCombo(operatorBox, new string[] { "移动", "联通", "电信" }, 140, 341);
        Controls.Add(operatorBox);

        businessLabel = MakeLabel("业务类型：", 440, 346, 105, 28);
        Controls.Add(businessLabel);
        ConfigureCombo(businessBox, new string[] { "行业", "营销" }, 565, 341);
        Controls.Add(businessBox);

        importantLabel = MakeLabel("是否重保：", 28, 396, 105, 28);
        Controls.Add(importantLabel);
        ConfigureCombo(importantBox, new string[] { "否", "是" }, 140, 391);
        Controls.Add(importantBox);

        tableActionLabel = MakeLabel("操作方式：", 28, 296, 105, 28);
        Controls.Add(tableActionLabel);
        ConfigureCombo(tableActionBox, new string[] { "拆分总表", "合并分表" }, 140, 291);
        tableActionBox.SelectedIndexChanged += TableActionChanged;
        Controls.Add(tableActionBox);

        signatureMatchModeLabel = MakeLabel("判定规则：", 28, 296, 105, 28);
        Controls.Add(signatureMatchModeLabel);
        ConfigureCombo(signatureMatchModeBox, new string[] { "完整规则", "只匹配地区" }, 140, 291);
        signatureMatchModeBox.SelectedIndexChanged += SignatureMatchModeChanged;
        Controls.Add(signatureMatchModeBox);

        maxRowsLabel = MakeLabel("分表最大行数：", 440, 296, 125, 28);
        Controls.Add(maxRowsLabel);
        maxRowsBox.SetBounds(565, 291, 265, 31);
        UiTheme.StyleInput(maxRowsBox);
        maxRowsBox.Text = "1000";
        Controls.Add(maxRowsBox);

        formatNote = MakeLabel("支持 .xlsx、.csv；生效签名模式另外支持 .txt（不读取旧版 .xls）。", 28, 443, 800, 27);
        formatNote.ForeColor = Color.DimGray;
        formatNote.Font = UiFonts.Chinese(10F);
        Controls.Add(formatNote);

        templateStatus.SetBounds(28, 480, 802, 32);
        templateStatus.BorderStyle = BorderStyle.FixedSingle;
        templateStatus.TextAlign = ContentAlignment.MiddleLeft;
        templateStatus.BackColor = Color.White;
        templateStatus.Font = UiFonts.Chinese(10F);
        Controls.Add(templateStatus);

        downloadPortTemplateButton.Text = "下载分配模板";
        downloadPortTemplateButton.SetBounds(165, 540, 125, 46);
        downloadPortTemplateButton.Radius = 10;
        downloadPortTemplateButton.Cursor = Cursors.Hand;
        downloadPortTemplateButton.Visible = false;
        downloadPortTemplateButton.Click += DownloadPortTemplate;
        Controls.Add(downloadPortTemplateButton);

        processButton.Text = "开始处理";
        processButton.SetBounds(300, 540, 260, 46);
        processButton.Type = AntdUI.TTypeMini.Primary;
        processButton.Radius = 10;
        processButton.Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Bold);
        processButton.Cursor = Cursors.Hand;
        processButton.Click += StartProcessing;
        Controls.Add(processButton);

        cancelButton.Text = "取消处理";
        cancelButton.SetBounds(500, 540, 140, 46);
        cancelButton.Type = AntdUI.TTypeMini.Error;
        cancelButton.Radius = 10;
        cancelButton.Cursor = Cursors.Hand;
        cancelButton.Visible = false;
        cancelButton.Click += CancelProcessing;
        Controls.Add(cancelButton);

        tableProgress.SetBounds(140, 600, 580, 22);
        tableProgress.Minimum = 0;
        tableProgress.Maximum = 100;
        tableProgress.Visible = false;
        Controls.Add(tableProgress);

        progressStatus.SetBounds(28, 628, 802, 58);
        progressStatus.TextAlign = ContentAlignment.MiddleCenter;
        progressStatus.ForeColor = Color.DimGray;
        progressStatus.Font = UiFonts.Chinese(10F);
        Controls.Add(progressStatus);

        shortlinkMode.CheckedChanged += ModeChanged;
        signatureMode.CheckedChanged += ModeChanged;
        mixedMode.CheckedChanged += ModeChanged;
        selfMobileMode.CheckedChanged += ModeChanged;
        selfTelecomMode.CheckedChanged += ModeChanged;
        tableToolsMode.CheckedChanged += ModeChanged;
        signatureConsistencyMode.CheckedChanged += ModeChanged;
        portAllocationMode.CheckedChanged += ModeChanged;
        worker.WorkerReportsProgress = true;
        worker.DoWork += DoProcessing;
        worker.ProgressChanged += ProcessingProgressChanged;
        worker.RunWorkerCompleted += ProcessingCompleted;
        progressTimer.Interval = 1000;
        progressTimer.Tick += ProgressTimerTick;
        ApplyModernLayout(header, modeGroup);
        UpdateModeDisplay();
    }

    private void ApplyModernLayout(Panel legacyHeader, GroupBox legacyModes)
    {
        SuspendLayout();
        Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
        int minimumWidth = Math.Min(1120, Math.Max(960, workingArea.Width - 40));
        int minimumHeight = Math.Min(700, Math.Max(640, workingArea.Height - 40));
        int defaultWidth = Math.Max(minimumWidth, Math.Min(1680, (int)Math.Round(workingArea.Width * 0.86)));
        int defaultHeight = Math.Max(minimumHeight, Math.Min(960, (int)Math.Round(workingArea.Height * 0.84)));
        ClientSize = new Size(defaultWidth, defaultHeight);
        MinimumSize = new Size(minimumWidth, minimumHeight);
        MaximumSize = Size.Empty;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = true;
        BackColor = UiTheme.Canvas;
        legacyHeader.Visible = false;
        legacyModes.Visible = false;

        modernSidebar = new RoundedSidebarPanel { BackColor = UiTheme.Surface, Dock = DockStyle.Left, Width = UiTheme.SidebarWidth };
        Controls.Add(modernSidebar);
        modernSidebar.BringToFront();

        PictureBox brandIcon = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom };
        brandIcon.SetBounds(18, 18, 36, 36);
        try
        {
            using (Image image = Image.FromFile(Path.Combine(Program.ResourcesDirectory, "应用图标.png")))
                brandIcon.Image = new Bitmap(image);
        }
        catch { }
        modernSidebar.Controls.Add(brandIcon);
        Label brand = MakeLabel("工作台", 66, 11, 118, 26);
        brand.Font = UiFonts.Chinese(11F, FontStyle.Bold);
        brand.AutoSize = true;
        brand.ForeColor = UiTheme.Text;
        modernSidebar.Controls.Add(brand);
        Label brandEnglish = MakeLabel("Dashboard", 68, 37, 112, 18);
        brandEnglish.Font = UiFonts.Chinese(7.5F);
        brandEnglish.AutoSize = true;
        brandEnglish.ForeColor = UiTheme.TextMuted;
        modernSidebar.Controls.Add(brandEnglish);
        VersionBadge versionBadge = new VersionBadge { Text = "V" + AppInfo.ProductVersion };
        int versionBadgeWidth = TextRenderer.MeasureText(versionBadge.Text, versionBadge.Font).Width + 16;
        versionBadge.SetBounds(238 - versionBadgeWidth, 24, versionBadgeWidth, 24);
        modernSidebar.Controls.Add(versionBadge);

        AddSidebarSection(modernSidebar, "通道业务", 96);
        AddNavigation(modernSidebar, "self_mobile", "自家移动", "mobile", 134, selfMobileMode);
        AddNavigation(modernSidebar, "self_telecom", "自家电信", "telecom", 182, selfTelecomMode);
        AddNavigation(modernSidebar, "port_allocation", "子端口分配", "port", 230, portAllocationMode);
        AddNavigation(modernSidebar, "shortlink", "生效短链", "link", 278, shortlinkMode);
        AddNavigation(modernSidebar, "signature", "生效签名", "signature", 326, signatureMode);
        AddNavigation(modernSidebar, "mixed", "混合表自动分流", "route", 374, mixedMode);
        AddSidebarSection(modernSidebar, "数据工具", 418);
        AddNavigation(modernSidebar, "table_tools", "拆分 / 合并表", "table", 444, tableToolsMode);
        AddNavigation(modernSidebar, "signature_consistency", "签名一致性", "check", 492, signatureConsistencyMode);

        NavigationButton manage = CreateNavigationButton("主码号管理", 540, "database");
        manage.Click += delegate
        {
            if (worker.IsBusy) return;
            using (MappingForm form = new MappingForm()) form.ShowDialog(this);
            UpdateModeDisplay();
        };
        manage.Tag = "mapping";
        modernSidebar.Controls.Add(manage);

        modernGuideButton = CreateNavigationButton("使用说明", 580, "guide");
        modernGuideButton.Click += delegate
        {
            if (worker.IsBusy) return;
            using (UsageGuideForm form = new UsageGuideForm()) form.ShowDialog(this);
        };
        modernGuideButton.Tag = "guide";
        modernSidebar.Controls.Add(modernGuideButton);

        modernNavigationGap = new Panel { BackColor = UiTheme.Surface };
        modernNavigationGap.SetBounds(UiTheme.SidebarWidth, 0, UiTheme.NavigationGapWidth, defaultHeight);
        Controls.Add(modernNavigationGap);
        modernNavigationGap.BringToFront();

        modernMainArea = new Panel { BackColor = UiTheme.Canvas };
        modernMainArea.SetBounds(UiTheme.SidebarWidth + UiTheme.NavigationGapWidth, 0,
            defaultWidth - UiTheme.SidebarWidth - UiTheme.NavigationGapWidth, defaultHeight);
        Controls.Add(modernMainArea);
        modernMainArea.BringToFront();
        modernTopbar = new RoundedMainHeaderPanel { BackColor = UiTheme.Surface };
        modernTopbar.SetBounds(0, 0, modernMainArea.Width, 72);
        modernMainArea.Controls.Add(modernTopbar);
        modernBreadcrumb = MakeLabel("通道业务  /  生效短链", 24, 22, 620, 28);
        modernBreadcrumb.ForeColor = UiTheme.TextSecondary;
        modernBreadcrumb.Font = UiFonts.Chinese(9.5F);
        modernTopbar.Controls.Add(modernBreadcrumb);
        modernLocalBadge = new SecurityBadge();
        modernLocalBadge.SetBounds(18, defaultHeight - 68, 214, 48);
        modernSidebar.Controls.Add(modernLocalBadge);

        modernContent = new RoundedMainContentPanel { BackColor = UiTheme.Canvas };
        modernContent.AutoScroll = true;
        modernContent.SetBounds(0, 72, modernMainArea.Width, defaultHeight - 72);
        modernMainArea.Controls.Add(modernContent);
        modernContent.BringToFront();
        modernEyebrow = MakeLabel("短链导入任务", 24, 22, 500, 20);
        modernEyebrow.Font = UiFonts.Chinese(8F, FontStyle.Bold);
        modernEyebrow.ForeColor = UiTheme.TextMuted;
        modernEyebrow.Visible = false;
        modernContent.Controls.Add(modernEyebrow);
        modernPageTitle = MakeLabel("生效短链", 24, 22, 600, 40);
        modernPageTitle.Font = UiFonts.Chinese(18F, FontStyle.Bold);
        modernPageTitle.ForeColor = UiTheme.Text;
        modernContent.Controls.Add(modernPageTitle);
        modeHint.Parent = modernContent;
        modeHint.SetBounds(24, 86, 960, 26);
        modeHint.ForeColor = UiTheme.TextSecondary;
        modeHint.Font = UiFonts.Chinese(9.5F);
        modeHint.Visible = false;

        modernInputCard = CreateCard(new Rectangle(30, 132, 760, 206));
        modernContent.Controls.Add(modernInputCard);
        Label inputTitle = MakeLabel("待处理文件", 20, 16, 260, 26);
        inputTitle.Font = UiFonts.Chinese(10.5F, FontStyle.Bold);
        inputTitle.ForeColor = UiTheme.Text;
        modernInputCard.Controls.Add(inputTitle);
        modernSourceDropZone = new DashedPanel();
        modernSourceDropZone.SetBounds(20, 62, 712, 102);
        modernSourceDropZone.AllowDrop = true;
        modernSourceDropZone.DragEnter += InputDragEnter;
        modernSourceDropZone.DragOver += InputDragOver;
        modernSourceDropZone.DragLeave += InputDragLeave;
        modernSourceDropZone.DragDrop += SourceDragDrop;
        modernInputCard.Controls.Add(modernSourceDropZone);
        modernSourceDropTitle = MakeLabel("拖入待处理文件", 56, 12, 330, 22);
        modernSourceDropTitle.Font = UiFonts.Chinese(10F, FontStyle.Bold);
        modernSourceDropTitle.ForeColor = UiTheme.Text;
        modernSourceDropZone.Controls.Add(modernSourceDropTitle);
        modernSourceDropZone.Controls.Add(new VectorIcon { IconKey = "upload", ForeColor = UiTheme.Text, Bounds = new Rectangle(18, 15, 26, 26) });
        modernSourceDropMeta = MakeLabel("支持 .xlsx、.csv", 56, 33, 430, 20);
        modernSourceDropMeta.Font = UiFonts.Chinese(8.5F);
        modernSourceDropMeta.ForeColor = UiTheme.TextMuted;
        modernSourceDropZone.Controls.Add(modernSourceDropMeta);
        modernMasterDropZone = new DashedPanel();
        modernMasterDropZone.SetBounds(20, 62, 344, 102);
        modernMasterDropZone.AllowDrop = true;
        modernMasterDropZone.DragEnter += InputDragEnter;
        modernMasterDropZone.DragOver += InputDragOver;
        modernMasterDropZone.DragLeave += InputDragLeave;
        modernMasterDropZone.DragDrop += MasterDragDrop;
        modernInputCard.Controls.Add(modernMasterDropZone);
        modernMasterDropTitle = MakeLabel("拖入总表文件夹", 56, 12, 250, 22);
        modernMasterDropTitle.Font = UiFonts.Chinese(10F, FontStyle.Bold);
        modernMasterDropTitle.ForeColor = UiTheme.Text;
        modernMasterDropZone.Controls.Add(modernMasterDropTitle);
        modernMasterDropZone.Controls.Add(new VectorIcon { IconKey = "folder", ForeColor = UiTheme.Text, Bounds = new Rectangle(18, 15, 26, 26) });
        modernMasterDropMeta = MakeLabel("选择最新总表资料", 56, 33, 250, 20);
        modernMasterDropMeta.Font = UiFonts.Chinese(8.5F);
        modernMasterDropMeta.ForeColor = UiTheme.TextMuted;
        modernMasterDropZone.Controls.Add(modernMasterDropMeta);

        modernParameterCard = CreateCard(new Rectangle(30, 356, 760, 250));
        modernContent.Controls.Add(modernParameterCard);
        Label parameterTitle = MakeLabel("参数配置", 20, 16, 240, 26);
        parameterTitle.Font = UiFonts.Chinese(10.5F, FontStyle.Bold);
        parameterTitle.ForeColor = UiTheme.Text;
        modernParameterCard.Controls.Add(parameterTitle);

        modernFooterCard = CreateCard(new Rectangle(30, 624, 760, 88));
        modernContent.Controls.Add(modernFooterCard);
        Label actionTitle = MakeLabel("另存位置", 20, 22, 112, 24);
        actionTitle.Font = UiFonts.Chinese(10F, FontStyle.Bold);
        actionTitle.ForeColor = UiTheme.Text;
        modernFooterCard.Controls.Add(actionTitle);
        outputDirectoryBox.SetBounds(132, 15, 270, 40);
        UiTheme.StyleInput(outputDirectoryBox);
        outputDirectoryBox.PlaceholderText = "默认：原文件所在目录";
        modernFooterCard.Controls.Add(outputDirectoryBox);
        browseOutputButton.Text = "选择位置";
        browseOutputButton.SetBounds(412, 15, 104, 40);
        UiTheme.StyleSecondaryButton(browseOutputButton);
        browseOutputButton.Click += BrowseOutputDirectory;
        modernFooterCard.Controls.Add(browseOutputButton);
        modernActionHint = MakeLabel("确认文件与参数后执行，原文件不会被覆盖。", 20, 34, 420, 20);
        modernActionHint.Font = UiFonts.Chinese(8.5F);
        modernActionHint.ForeColor = UiTheme.TextMuted;
        modernActionHint.Visible = false;
        modernFooterCard.Controls.Add(modernActionHint);

        modernStatusCard = CreateCard(new Rectangle(810, 132, 320, 158));
        modernContent.Controls.Add(modernStatusCard);
        Label statusTitle = MakeLabel("系统状态", 18, 13, 220, 24);
        statusTitle.Font = UiFonts.Chinese(10.5F, FontStyle.Bold);
        statusTitle.ForeColor = UiTheme.Text;
        modernStatusCard.Controls.Add(statusTitle);
        modernStatusCard.Controls.Add(new StatusDot { Bounds = new Rectangle(22, 58, 14, 14), StateColor = UiTheme.Success });
        modernStatusCard.Controls.Add(MakeStatusLabel("系统已就绪", 44, 51));
        modernReadyStatus = MakeStatusLabel("业务模板检查中", 44, 82);
        modernStatusCard.Controls.Add(new StatusDot { Bounds = new Rectangle(22, 89, 14, 14), StateColor = UiTheme.Success, Tag = "template-dot" });
        modernStatusCard.Controls.Add(modernReadyStatus);
        modernMappingStatus = MakeStatusLabel("主码号资料检查中", 44, 113);
        modernStatusCard.Controls.Add(new StatusDot { Bounds = new Rectangle(22, 120, 14, 14), StateColor = UiTheme.Success, Tag = "mapping-dot" });
        modernStatusCard.Controls.Add(modernMappingStatus);

        modernResourceCard = CreateCard(new Rectangle(810, 310, 320, 240));
        modernContent.Controls.Add(modernResourceCard);
        Label resourceTitle = MakeLabel("所需资料", 18, 13, 220, 24);
        resourceTitle.Font = UiFonts.Chinese(10.5F, FontStyle.Bold);
        resourceTitle.ForeColor = UiTheme.Text;
        modernResourceCard.Controls.Add(resourceTitle);
        modernResourcesList = new FlowLayoutPanel();
        modernResourcesList.FlowDirection = FlowDirection.TopDown;
        modernResourcesList.WrapContents = false;
        modernResourcesList.AutoScroll = false;
        modernResourcesList.BackColor = UiTheme.Surface;
        modernResourcesList.SetBounds(18, 54, 284, 170);
        modernResourceCard.Controls.Add(modernResourcesList);
        templateStatus.Parent = modernResourceCard;
        templateStatus.Visible = false;

        modernTaskCard = CreateCard(new Rectangle(810, 570, 320, 132));
        modernContent.Controls.Add(modernTaskCard);
        Label taskTitle = MakeLabel("当前任务", 18, 13, 220, 24);
        taskTitle.Font = UiFonts.Chinese(10.5F, FontStyle.Bold);
        taskTitle.ForeColor = UiTheme.Text;
        modernTaskCard.Controls.Add(taskTitle);

        modernSecurityCard = new SecurityNotice();
        modernSecurityCard.SetBounds(810, 722, 320, 80);
        modernContent.Controls.Add(modernSecurityCard);

        ReparentInputControls();

        downloadPortTemplateButton.Parent = modernFooterCard;
        downloadPortTemplateButton.SetBounds(390, 16, 126, 38);
        UiTheme.StyleSecondaryButton(downloadPortTemplateButton);
        processButton.Parent = modernFooterCard;
        processButton.SetBounds(536, 15, 164, 40);
        processButton.Radius = 7;
        processButton.BackColor = UiTheme.Primary;
        processButton.BackHover = UiTheme.PrimaryHover;
        processButton.ForeColor = Color.White;
        processButton.ForeHover = Color.White;
        processButton.Font = UiFonts.Chinese(9.5F, FontStyle.Bold);
        processButton.WaveSize = 0;
        cancelButton.Parent = modernFooterCard;
        cancelButton.SetBounds(396, 16, 116, 38);
        cancelButton.Radius = 7;
        tableProgress.Parent = modernTaskCard;
        tableProgress.SetBounds(22, 96, 276, 8);
        progressStatus.Parent = modernTaskCard;
        progressStatus.SetBounds(22, 52, 276, 42);
        progressStatus.ForeColor = UiTheme.TextSecondary;
        progressStatus.TextAlign = ContentAlignment.MiddleLeft;
        progressStatus.Font = UiFonts.Chinese(9F);
        progressStatus.Text = "尚未开始，选择文件并确认参数后即可处理。";
        Resize += delegate { LayoutModernContent(); };
        Shown += delegate { BeginInvoke((MethodInvoker)LayoutModernContent); };
        LayoutModernContent();
        ResumeLayout(false);
    }

    private static Panel CreateCard(Rectangle bounds)
    {
        Panel card = new SurfacePanel { BackColor = UiTheme.Surface, BorderStyle = BorderStyle.None };
        card.SetBounds(bounds.X, bounds.Y, bounds.Width, bounds.Height);
        return card;
    }

    private static void AddSidebarSection(Panel sidebar, string text, int y)
    {
        Label label = MakeLabel(text, 18, y, 210, 28);
        label.Font = UiFonts.Chinese(12F, FontStyle.Bold);
        label.ForeColor = UiTheme.TextMuted;
        sidebar.Controls.Add(label);
    }

    private NavigationButton CreateNavigationButton(string text, int y, string iconKey)
    {
        NavigationButton button = new NavigationButton();
        button.Text = text;
        button.IconKey = iconKey;
        button.TextAlign = ContentAlignment.MiddleLeft;
        button.BackColor = UiTheme.Surface;
        button.ForeColor = UiTheme.TextSecondary;
        button.Font = UiFonts.Chinese(12F, FontStyle.Bold);
        button.Cursor = Cursors.Hand;
        button.TabStop = false;
        button.SetBounds(12, y, 226, 40);
        return button;
    }

    private void AddNavigation(Panel sidebar, string mode, string text, string iconKey, int y, RadioButton target)
    {
        NavigationButton button = CreateNavigationButton(text, y, iconKey);
        button.Click += delegate { if (!worker.IsBusy) target.Checked = true; };
        modernNavButtons[mode] = button;
        sidebar.Controls.Add(button);
    }

    private void ReparentInputControls()
    {
        sourceLabel.Parent = modernInputCard;
        sourceLabel.Visible = false;
        masterLabel.Parent = modernInputCard;
        masterLabel.Visible = false;
        sourceBox.Parent = modernSourceDropZone;
        browseButton.Parent = modernSourceDropZone;
        masterBox.Parent = modernMasterDropZone;
        browseMasterButton.Parent = modernMasterDropZone;
        Control[] settings = new Control[] { channelLabel, channelBox, startPortLabel, startPortBox,
            operatorLabel, operatorBox, businessLabel, businessBox, importantLabel, importantBox,
            tableActionLabel, tableActionBox, signatureMatchModeLabel, signatureMatchModeBox,
            maxRowsLabel, maxRowsBox, formatNote };
        foreach (Control control in settings) control.Parent = modernParameterCard;
        sourceLabel.Font = UiFonts.Chinese(11F, FontStyle.Bold);
        masterLabel.Font = UiFonts.Chinese(11F, FontStyle.Bold);
        UiTheme.StyleInput(sourceBox);
        sourceBox.ReadOnly = true;
        sourceBox.BackColor = UiTheme.Surface;
        sourceBox.ForeColor = UiTheme.TextSecondary;
        sourceBox.PlaceholderText = "尚未选择文件";
        sourceBox.Font = UiFonts.Chinese(10F);
        UiTheme.StyleInput(masterBox);
        masterBox.ReadOnly = true;
        masterBox.BackColor = UiTheme.Surface;
        masterBox.ForeColor = UiTheme.TextSecondary;
        masterBox.PlaceholderText = "尚未选择资料";
        masterBox.Font = UiFonts.Chinese(10F);
        foreach (AntdUI.Input box in new AntdUI.Input[] { channelBox, startPortBox, maxRowsBox }) UiTheme.StyleInput(box);
        foreach (AntdUI.Select combo in new AntdUI.Select[] { operatorBox, businessBox, importantBox, tableActionBox, signatureMatchModeBox }) UiTheme.StyleSelect(combo);
        foreach (Label label in new Label[] { sourceLabel, masterLabel, channelLabel, startPortLabel, operatorLabel,
            businessLabel, importantLabel, tableActionLabel, signatureMatchModeLabel, maxRowsLabel })
            label.ForeColor = UiTheme.TextSecondary;
        UiTheme.StyleSecondaryButton(browseButton);
        UiTheme.StyleSecondaryButton(browseMasterButton);
        sourceBox.SetBounds(56, 66, 470, 42);
        browseButton.SetBounds(550, 65, 104, 42);
        masterBox.SetBounds(56, 66, 200, 42);
        browseMasterButton.SetBounds(278, 65, 104, 42);
        formatNote.ForeColor = UiTheme.TextMuted;
        formatNote.Font = UiFonts.Chinese(8.5F);
        foreach (Label label in new Label[] { channelLabel, startPortLabel, operatorLabel, businessLabel,
            importantLabel, tableActionLabel, signatureMatchModeLabel, maxRowsLabel })
            label.Text = label.Text.TrimEnd('：');
    }

    private void LayoutModernContent()
    {
        if (modernMainArea == null || modernContent == null) return;
        int mainWidth = Math.Max(700, ClientSize.Width - UiTheme.SidebarWidth - UiTheme.NavigationGapWidth);
        int contentHeight = Math.Max(560, ClientSize.Height - 72);
        modernNavigationGap.SetBounds(UiTheme.SidebarWidth, 0, UiTheme.NavigationGapWidth, ClientSize.Height);
        modernMainArea.SetBounds(UiTheme.SidebarWidth + UiTheme.NavigationGapWidth, 0, mainWidth, ClientSize.Height);
        modernTopbar.SetBounds(0, 0, mainWidth, 72);
        modernContent.SetBounds(0, 72, mainWidth, contentHeight);
        modernLocalBadge.Top = Math.Max(630, ClientSize.Height - modernLocalBadge.Height - 20);
        modernGuideButton.Top = modernLocalBadge.Top - modernGuideButton.Height - 12;

        int workspaceWidth = Math.Min(1240, Math.Max(680, mainWidth - 40));
        int originX = Math.Max(18, (mainWidth - workspaceWidth) / 2);
        modernBreadcrumb.Left = originX;
        modernEyebrow.Left = originX;
        modernPageTitle.Left = originX;
        modeHint.Left = originX;
        modeHint.Width = workspaceWidth;

        int gap = 12;
        int rightWidth = Math.Max(270, Math.Min(316, (int)Math.Round(workspaceWidth * 0.27)));
        int leftWidth = workspaceWidth - rightWidth - gap;
        int leftX = originX;
        int rightX = leftX + leftWidth + gap;
        int top = 84;

        bool twoDropZones = modernMasterDropZone.Visible;
        int uploadHeight = 190;
        modernInputCard.SetBounds(leftX, top, leftWidth, uploadHeight);
        LayoutDropZones(leftWidth, twoDropZones);

        int parameterTop = top + uploadHeight + gap;
        int parameterHeight = LayoutParameterCard(leftX, parameterTop, leftWidth);
        int actionTop = parameterTop + parameterHeight + gap;
        modernFooterCard.SetBounds(leftX, actionTop, leftWidth, 70);
        processButton.Left = leftWidth - processButton.Width - 18;
        cancelButton.Left = processButton.Left - cancelButton.Width - 12;
        downloadPortTemplateButton.Left = processButton.Left - downloadPortTemplateButton.Width - 10;
        int outputControlsRight = processButton.Left;
        if (downloadPortTemplateButton.Visible) outputControlsRight = downloadPortTemplateButton.Left;
        else if (cancelButton.Visible) outputControlsRight = cancelButton.Left;
        browseOutputButton.Left = outputControlsRight - browseOutputButton.Width - 10;
        outputDirectoryBox.Width = Math.Max(120, browseOutputButton.Left - outputDirectoryBox.Left - 10);

        modernStatusCard.SetBounds(rightX, top, rightWidth, 142);
        int resourceHeight = Math.Max(126, Math.Min(226, 54 + (modernResourcesList.Controls.Count * 40)));
        modernResourceCard.SetBounds(rightX, top + 154, rightWidth, resourceHeight);
        modernResourcesList.SetBounds(14, 44, rightWidth - 28, resourceHeight - 54);
        foreach (Control row in modernResourcesList.Controls) row.Width = modernResourcesList.Width - 2;
        int taskTop = modernResourceCard.Bottom + gap;
        modernTaskCard.SetBounds(rightX, taskTop, rightWidth, 116);
        progressStatus.Width = rightWidth - 36;
        tableProgress.Width = rightWidth - 36;
        modernSecurityCard.SetBounds(rightX, modernTaskCard.Bottom + gap, rightWidth, 68);

        int bottom = Math.Max(modernFooterCard.Bottom, modernSecurityCard.Bottom) + 28;
        modernContent.AutoScrollMinSize = new Size(0, bottom);
    }

    private void LayoutDropZones(int cardWidth, bool twoDropZones)
    {
        int available = cardWidth - 40;
        int zoneWidth = twoDropZones ? (available - 12) / 2 : available;
        modernSourceDropZone.SetBounds(20, 52, zoneWidth, 120);
        modernMasterDropZone.SetBounds(20 + zoneWidth + 12, 52, zoneWidth, 120);
        LayoutDropZone(modernSourceDropZone, modernSourceDropTitle, modernSourceDropMeta, sourceBox, browseButton);
        LayoutDropZone(modernMasterDropZone, modernMasterDropTitle, modernMasterDropMeta, masterBox, browseMasterButton);
    }

    private static void LayoutDropZone(Panel zone, Label title, Label meta, AntdUI.Input input, AntdUI.Button button)
    {
        title.Width = Math.Max(120, zone.Width - 92);
        meta.Width = Math.Max(120, zone.Width - 92);
        button.Left = zone.Width - button.Width - 12;
        input.Width = Math.Max(96, button.Left - input.Left - 10);
    }

    private int LayoutParameterCard(int x, int y, int width)
    {
        List<KeyValuePair<Label, Control>> fields = new List<KeyValuePair<Label, Control>>();
        AddVisibleField(fields, channelLabel, channelBox);
        AddVisibleField(fields, startPortLabel, startPortBox);
        AddVisibleField(fields, operatorLabel, operatorBox);
        AddVisibleField(fields, businessLabel, businessBox);
        AddVisibleField(fields, importantLabel, importantBox);
        AddVisibleField(fields, tableActionLabel, tableActionBox);
        AddVisibleField(fields, signatureMatchModeLabel, signatureMatchModeBox);
        AddVisibleField(fields, maxRowsLabel, maxRowsBox);

        int rows = Math.Max(1, (fields.Count + 1) / 2);
        int height = 52 + (rows * 70) + 38;
        modernParameterCard.SetBounds(x, y, width, height);
        int columnGap = 16;
        int columnWidth = (width - 40 - columnGap) / 2;
        for (int index = 0; index < fields.Count; index++)
        {
            int column = index % 2;
            int row = index / 2;
            int fieldX = 20 + (column * (columnWidth + columnGap));
            int fieldY = 52 + (row * 70);
            Label label = fields[index].Key;
            Control control = fields[index].Value;
            label.SetBounds(fieldX, fieldY, columnWidth, 18);
            label.Font = UiFonts.Chinese(8.5F, FontStyle.Bold);
            label.ForeColor = UiTheme.TextSecondary;
            control.SetBounds(fieldX, fieldY + 21, columnWidth, 38);
        }
        formatNote.SetBounds(20, height - 32, width - 40, 20);
        return height;
    }

    private static void AddVisibleField(List<KeyValuePair<Label, Control>> fields, Label label, Control control)
    {
        if (label.Visible && control.Visible) fields.Add(new KeyValuePair<Label, Control>(label, control));
    }

    private static Label MakeStatusLabel(string text, int x, int y)
    {
        Label label = MakeLabel(text, x, y, 240, 28);
        label.Font = UiFonts.Chinese(9.5F);
        label.ForeColor = UiTheme.TextSecondary;
        return label;
    }

    private static Label MakeLabel(string text, int x, int y, int width, int height)
    {
        Label label = new Label();
        label.Text = text;
        label.SetBounds(x, y, width, height);
        label.TextAlign = ContentAlignment.MiddleLeft;
        return label;
    }

    private static void ConfigureCombo(AntdUI.Select combo, string[] values, int x, int y)
    {
        combo.SetBounds(x, y, 240, 32);
        combo.Items.AddRange(values);
        combo.SelectedIndex = 0;
        UiTheme.StyleSelect(combo);
    }

    private string SelectedMode
    {
        get
        {
            if (tableToolsMode.Checked) return "table_tools";
            if (signatureConsistencyMode.Checked) return "signature_consistency";
            if (portAllocationMode.Checked) return "port_allocation";
            if (selfMobileMode.Checked) return "self_mobile";
            if (selfTelecomMode.Checked) return "self_telecom";
            if (mixedMode.Checked) return "mixed";
            return signatureMode.Checked ? "signature" : "shortlink";
        }
    }

    internal void SetSnapshotMode(string mode)
    {
        if (mode == "self_mobile") selfMobileMode.Checked = true;
        else if (mode == "self_telecom") selfTelecomMode.Checked = true;
        else if (mode == "port_allocation") portAllocationMode.Checked = true;
        else if (mode == "signature") signatureMode.Checked = true;
        else if (mode == "mixed") mixedMode.Checked = true;
        else if (mode == "table_tools") tableToolsMode.Checked = true;
        else if (mode == "signature_consistency") signatureConsistencyMode.Checked = true;
        else shortlinkMode.Checked = true;
    }

    private string SelectedTableAction
    {
        get { return tableActionBox.SelectedIndex == 1 ? "merge" : "split"; }
    }

    private string SelectedSignatureMatchMode
    {
        get { return signatureMatchModeBox.SelectedIndex == 1 ? "region_only" : "full_rule"; }
    }

    private void ModeChanged(object sender, EventArgs e)
    {
        if (((RadioButton)sender).Checked) UpdateModeDisplay();
    }

    private void TableActionChanged(object sender, EventArgs e)
    {
        if (SelectedMode == "table_tools")
        {
            sourceBox.Clear();
            UpdateModeDisplay();
        }
    }

    private void SignatureMatchModeChanged(object sender, EventArgs e)
    {
        if (SelectedMode == "signature_consistency") UpdateModeDisplay();
    }

    private void UpdateModeDisplay()
    {
        bool selfMobile = SelectedMode == "self_mobile";
        bool selfTelecom = SelectedMode == "self_telecom";
        bool selfOwned = selfMobile || selfTelecom;
        bool tableTools = SelectedMode == "table_tools";
        bool consistency = SelectedMode == "signature_consistency";
        bool portAllocation = SelectedMode == "port_allocation";
        bool merging = tableTools && SelectedTableAction == "merge";
        UpdateModernChrome();
        bool showBusinessFields = !selfOwned && !tableTools && !consistency && !portAllocation;
        channelBox.Enabled = showBusinessFields;
        startPortBox.Enabled = showBusinessFields;
        businessBox.Enabled = showBusinessFields;
        channelBox.Visible = showBusinessFields;
        channelLabel.Visible = showBusinessFields;
        startPortBox.Visible = showBusinessFields;
        startPortLabel.Visible = showBusinessFields;
        operatorBox.Visible = showBusinessFields;
        operatorLabel.Visible = showBusinessFields;
        businessBox.Visible = showBusinessFields;
        businessLabel.Visible = showBusinessFields;
        importantBox.Visible = !tableTools && !consistency && !portAllocation;
        importantLabel.Visible = !tableTools && !consistency && !portAllocation;
        masterBox.Visible = portAllocation || selfMobile;
        masterLabel.Visible = false;
        sourceLabel.Visible = false;
        browseMasterButton.Visible = portAllocation || selfMobile;
        if (modernMasterDropZone != null) modernMasterDropZone.Visible = portAllocation || selfMobile;
        masterLabel.Text = selfMobile ? "总表文件夹：" : "历史总表：";
        browseMasterButton.Text = selfMobile ? "选择文件夹" : "选择总表";
        tableActionBox.Visible = tableTools;
        tableActionLabel.Visible = tableTools;
        signatureMatchModeBox.Visible = consistency;
        signatureMatchModeLabel.Visible = consistency;
        maxRowsBox.Visible = tableTools && !merging;
        maxRowsLabel.Visible = tableTools && !merging;
        sourceLabel.Text = tableTools
            ? (merging ? "分表文件夹：" : "总表文件：")
            : portAllocation ? "待分配表："
            : consistency ? "待判定文件："
            : selfTelecom ? "批次文件夹：" : selfMobile ? "关键信息表：" : "客户文件：";
        browseButton.Text = selfTelecom || merging ? "选择文件夹" : "选择文件";
        if (modernSourceDropTitle != null)
        {
            bool sharedImport = SelectedMode == "shortlink" || SelectedMode == "signature" || SelectedMode == "mixed";
            modernSourceDropTitle.Text = sharedImport ? "拖入待处理文件" : "拖入" + sourceLabel.Text.TrimEnd('：');
        }
        if (modernMasterDropTitle != null)
            modernMasterDropTitle.Text = selfMobile ? "拖入最新移动签名总表文件夹" : "拖入最新历史总表";
        modernSourceDropMeta.Text = tableTools
            ? (merging ? "选择包含全部分表的文件夹" : "支持 .xlsx、.csv")
            : consistency ? "支持 .xlsx、.xlsm"
            : selfTelecom ? "选择完整电信批次文件夹"
            : selfMobile ? "支持 .xlsx 关键信息表"
            : SelectedMode == "signature" ? "支持 .xlsx、.csv、.txt"
            : "支持 .xlsx、.csv";
        modernMasterDropMeta.Text = selfMobile ? "选择最新移动签名总表文件夹" : "选择最新历史总表文件";
        processButton.Text = portAllocation ? "开始分配"
            : tableTools ? (merging ? "开始合并" : "开始拆分")
            : consistency ? "开始判定" : "开始处理";
        downloadPortTemplateButton.Visible = portAllocation;
        formatNote.Text = tableTools
            ? (merging
                ? "快速流式合并；保留数据和首个分表表头格式，XLSX 与 CSV 不能混合。"
                : "支持 .xlsx、.csv；设置的最大行数包含每个分表重复保留的表头。")
            : portAllocation
            ? "分配类型留空时同时分配；填“仅行业”或“仅营销”时另一列留空。"
            : consistency
            ? "支持 .xlsx、.xlsm；源文件不会被修改，结果保存在原文件夹。"
            : selfMobile
            ? "关键信息表仅支持 .xlsx；可拖入表格和总表文件夹，处理时直接读取最新文件。"
            : "支持 .xlsx、.csv；生效签名模式另外支持 .txt（不读取旧版 .xls）。";
        bool mappingExists = File.Exists(Program.MappingFile);
        if (modernMappingStatus != null)
        {
            modernMappingStatus.Text = mappingExists ? "主码号资料可用" : "主码号资料缺失";
            modernMappingStatus.ForeColor = mappingExists ? UiTheme.Success : UiTheme.Danger;
        }
        UpdateStatusDot("mapping-dot", mappingExists ? UiTheme.Success : UiTheme.Danger);

        if (portAllocation)
            modeHint.Text = "按主码号与签名编号核对最新总表，分别生成已分配表和更新总表。";
        else if (tableTools)
            modeHint.Text = merging
                ? "严格校验表头并显示进度、用时和剩余时间；可安全取消，不保留半成品。"
                : "按指定最大行数拆分总表，表头和单元格格式保持一致。";
        else if (selfMobile)
            modeHint.Text = "选择关键信息表和最新总表文件夹；按主码号及消息类型分流普通移动/5G。";
        else if (selfTelecom)
            modeHint.Text = "汇总提交表和CSV/XLSX失败结果；错误原因表包含全部子端口、签名、上报状态和错误原因。";
        else if (SelectedMode == "mixed")
            modeHint.Text = "所有签名去重后进入签名模板；有引流/链接/号码的记录另外进入短链模板，并连续分配子端口。";
        else if (consistency)
            modeHint.Text = SelectedSignatureMatchMode == "region_only"
                ? "只校验签名与企业名称是否连续包含同一地区；不检查签名其余文字。"
                : "签名可按原顺序跳字；地区须在签名和企业名称中连续出现，省市可用简称。";
        else if (SelectedMode == "signature")
            modeHint.Text = "将签名写入生效签名模板；没有签名表头时默认读取第一列。";
        else
            modeHint.Text = "将签名、链接和号码写入生效短链模板，等级固定为 2。";

        string[] required;
        if (portAllocation)
            required = new string[] { "通道编号及对应主码号.xlsx", "子端口分配.xlsx" };
        else if (tableTools)
            required = new string[] { "无需外置模板" };
        else if (consistency)
            required = new string[] { "地区库.xlsx" };
        else if (selfMobile)
            required = new string[] { "通道生效签名短链导入模板.xlsx", "通道编号及对应主码号.xlsx" };
        else if (selfTelecom)
            required = new string[] { "通道生效签名短链导入模板.xlsx", "通道生效签名导入模板.xlsx", "最终报备表模板.xlsx", "通道编号及对应主码号.xlsx" };
        else if (SelectedMode == "mixed")
            required = new string[] { "通道生效签名短链导入模板.xlsx", "通道生效签名导入模板.xlsx" };
        else
            required = new string[] { SelectedMode == "signature" ? "通道生效签名导入模板.xlsx" : "通道生效签名短链导入模板.xlsx" };

        List<string> missing = required.Where(name => name != "无需外置模板" &&
            !File.Exists(name == "通道编号及对应主码号.xlsx"
                ? Program.MappingFile : Path.Combine(Program.ResourcesDirectory, name))).ToList();
        bool ready = missing.Count == 0;
        templateStatus.Text = ready ? "ready" : "missing";
        templateStatus.ForeColor = ready ? UiTheme.TextSecondary : UiTheme.Danger;
        modernReadyStatus.Text = tableTools ? "当前工具无需模板" : ready ? "业务资料已就绪" : "业务资料不完整";
        modernReadyStatus.ForeColor = ready ? UiTheme.Success : UiTheme.Danger;
        UpdateStatusDot("template-dot", ready ? UiTheme.Success : UiTheme.Danger);
        SetResourceRows(required, missing);
        modernActionHint.Text = portAllocation ? "确认两份表格后开始分配，原文件不会被覆盖。"
            : tableTools ? "处理过程可取消，不会保留未完成的结果文件。"
            : "确认文件与参数后执行，原文件不会被覆盖。";
        if (!worker.IsBusy)
        {
            progressStatus.Text = string.IsNullOrWhiteSpace(sourceBox.Text)
                ? "尚未开始，选择文件并确认参数后即可处理。"
                : "文件已选择，等待开始处理。";
            progressStatus.ForeColor = UiTheme.TextSecondary;
        }
        LayoutModernContent();
    }

    private void UpdateStatusDot(string tag, Color color)
    {
        if (modernStatusCard == null) return;
        foreach (Control control in modernStatusCard.Controls)
        {
            StatusDot dot = control as StatusDot;
            if (dot != null && Convert.ToString(dot.Tag) == tag)
            {
                dot.StateColor = color;
                dot.Invalidate();
            }
        }
    }

    private void SetResourceRows(IEnumerable<string> names, IEnumerable<string> missingNames)
    {
        if (modernResourcesList == null) return;
        HashSet<string> missing = new HashSet<string>(missingNames);
        modernResourcesList.SuspendLayout();
        modernResourcesList.Controls.Clear();
        foreach (string name in names)
        {
            ResourceRow row = new ResourceRow();
            row.ResourceName = name;
            row.Available = !missing.Contains(name);
            row.Margin = new Padding(0, 0, 0, 4);
            row.Size = new Size(Math.Max(220, modernResourcesList.Width - 2), 40);
            modernResourcesList.Controls.Add(row);
        }
        modernResourcesList.ResumeLayout();
    }

    private void UpdateModernChrome()
    {
        if (modernPageTitle == null || modernBreadcrumb == null) return;
        Dictionary<string, string> titles = new Dictionary<string, string> {
            { "self_mobile", "自家移动" }, { "self_telecom", "自家电信" },
            { "port_allocation", "子端口分配" }, { "shortlink", "生效短链" },
            { "signature", "生效签名" }, { "mixed", "混合表自动分流" },
            { "table_tools", "拆分 / 合并表" }, { "signature_consistency", "签名一致性" }
        };
        string title;
        if (!titles.TryGetValue(SelectedMode, out title)) title = "业务处理";
        modernPageTitle.Text = title;
        bool dataTool = SelectedMode == "table_tools" || SelectedMode == "signature_consistency";
        modernBreadcrumb.Text = (dataTool ? "数据工具" : "通道业务") + "   /   " + title;
        Dictionary<string, string> eyebrows = new Dictionary<string, string> {
            { "self_mobile", "移动通道任务" }, { "self_telecom", "电信通道任务" },
            { "port_allocation", "端口配置任务" }, { "shortlink", "短链导入任务" },
            { "signature", "签名导入任务" }, { "mixed", "自动分流任务" },
            { "table_tools", "表格整理工具" }, { "signature_consistency", "数据校验工具" }
        };
        string eyebrow;
        modernEyebrow.Text = eyebrows.TryGetValue(SelectedMode, out eyebrow) ? eyebrow : "任务工作区";
        foreach (KeyValuePair<string, NavigationButton> item in modernNavButtons)
        {
            bool selected = item.Key == SelectedMode;
            item.Value.Selected = selected;
            item.Value.BackColor = selected ? UiTheme.Primary : UiTheme.Surface;
            item.Value.ForeColor = selected ? Color.White : UiTheme.TextSecondary;
            item.Value.Font = UiFonts.Chinese(12F, FontStyle.Bold);
            item.Value.Invalidate();
        }
    }

    private void DownloadPortTemplate(object sender, EventArgs e)
    {
        string template = Path.Combine(Program.ResourcesDirectory, "子端口分配.xlsx");
        if (!File.Exists(template))
        {
            MessageBox.Show(this, "内置子端口分配模板缺失，请重新安装程序。", "模板缺失",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        using (SaveFileDialog dialog = new SaveFileDialog())
        {
            dialog.Title = "保存子端口分配模板";
            dialog.Filter = "Excel 文件|*.xlsx";
            dialog.FileName = "子端口分配.xlsx";
            dialog.OverwritePrompt = true;
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            try
            {
                File.Copy(template, dialog.FileName, true);
                MessageBox.Show(this, "模板已保存到：\r\n" + dialog.FileName, "下载完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "模板保存失败：\r\n" + ex.Message, "保存失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void BrowseFile(object sender, EventArgs e)
    {
        if (SelectedMode == "port_allocation")
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择待分配表";
                dialog.Filter = "Excel 文件|*.xlsx";
                if (dialog.ShowDialog(this) == DialogResult.OK) sourceBox.Text = dialog.FileName;
            }
            return;
        }
        if (SelectedMode == "table_tools")
        {
            if (SelectedTableAction == "merge")
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "选择包含全部待合并分表的文件夹";
                    if (dialog.ShowDialog(this) == DialogResult.OK) sourceBox.Text = dialog.SelectedPath;
                }
                return;
            }
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择需要拆分的总表";
                dialog.Filter = "支持的文件|*.xlsx;*.csv|Excel 文件|*.xlsx|CSV 文件|*.csv";
                if (dialog.ShowDialog(this) == DialogResult.OK) sourceBox.Text = dialog.FileName;
            }
            return;
        }
        if (SelectedMode == "self_telecom")
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择自家电信批次文件夹";
                if (dialog.ShowDialog(this) == DialogResult.OK) sourceBox.Text = dialog.SelectedPath;
            }
            return;
        }
        using (OpenFileDialog dialog = new OpenFileDialog())
        {
            dialog.Title = "选择客户文件";
            dialog.Filter = SelectedMode == "self_mobile"
                ? "运营商报备表|*.xlsx|Excel 文件|*.xlsx"
                : SelectedMode == "signature_consistency"
                ? "Excel 文件|*.xlsx;*.xlsm|XLSX 文件|*.xlsx"
                : SelectedMode == "signature"
                ? "支持的文件|*.xlsx;*.csv;*.txt|Excel 文件|*.xlsx|CSV 文件|*.csv|文本文件|*.txt"
                : "支持的文件|*.xlsx;*.csv|Excel 文件|*.xlsx|CSV 文件|*.csv";
            if (dialog.ShowDialog(this) == DialogResult.OK) sourceBox.Text = dialog.FileName;
        }
    }

    private void BrowseMasterFile(object sender, EventArgs e)
    {
        if (SelectedMode == "self_mobile")
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择包含移动签名总表的文件夹";
                if (dialog.ShowDialog(this) == DialogResult.OK) masterBox.Text = dialog.SelectedPath;
            }
            return;
        }
        using (OpenFileDialog dialog = new OpenFileDialog())
        {
            dialog.Title = "选择最新历史总表";
            dialog.Filter = "Excel 文件|*.xlsx";
            if (dialog.ShowDialog(this) == DialogResult.OK) masterBox.Text = dialog.FileName;
        }
    }

    private void BrowseOutputDirectory(object sender, EventArgs e)
    {
        using (FolderBrowserDialog dialog = new FolderBrowserDialog())
        {
            dialog.Description = "选择本次处理结果的另存文件夹；留空则使用默认位置";
            string current = outputDirectoryBox.Text.Trim();
            if (Directory.Exists(current)) dialog.SelectedPath = current;
            else
            {
                string source = sourceBox.Text.Trim();
                string suggested = File.Exists(source) ? Path.GetDirectoryName(source) : source;
                if (Directory.Exists(suggested)) dialog.SelectedPath = suggested;
            }
            if (dialog.ShowDialog(this) == DialogResult.OK)
                outputDirectoryBox.Text = dialog.SelectedPath;
        }
    }

    private void InputDragEnter(object sender, DragEventArgs e)
    {
        e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        SetDropZoneDragState(sender, e.Effect == DragDropEffects.Copy);
    }

    private void InputDragOver(object sender, DragEventArgs e)
    {
        e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
        SetDropZoneDragState(sender, e.Effect == DragDropEffects.Copy);
    }

    private void InputDragLeave(object sender, EventArgs e)
    {
        SetDropZoneDragState(sender, false);
    }

    private void SetDropZoneDragState(object sender, bool active)
    {
        Control control = sender as Control;
        DashedPanel panel = control as DashedPanel;
        if (panel == null && control != null) panel = control.Parent as DashedPanel;
        if (panel == null) return;
        panel.DragActive = active;

        Label title = panel == modernSourceDropZone ? modernSourceDropTitle : modernMasterDropTitle;
        if (title == null) return;
        if (active)
        {
            if (title.Tag == null) title.Tag = title.Text;
            bool expectsFolder = panel == modernMasterDropZone || SelectedMode == "self_telecom" ||
                (SelectedMode == "table_tools" && SelectedTableAction == "merge");
            title.Text = expectsFolder ? "松开以选择此文件夹" : "松开以选择此文件";
        }
        else if (title.Tag is string)
        {
            title.Text = (string)title.Tag;
            title.Tag = null;
        }
    }

    private void SourceDragDrop(object sender, DragEventArgs e)
    {
        SetDropZoneDragState(sender, false);
        string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
        if (paths != null && paths.Length == 1 && (File.Exists(paths[0]) || Directory.Exists(paths[0])))
            sourceBox.Text = paths[0];
    }

    private void MasterDragDrop(object sender, DragEventArgs e)
    {
        SetDropZoneDragState(sender, false);
        string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
        if (paths == null || paths.Length != 1) return;
        if (SelectedMode == "self_mobile" && Directory.Exists(paths[0])) masterBox.Text = paths[0];
        else if (SelectedMode == "port_allocation" && File.Exists(paths[0])) masterBox.Text = paths[0];
    }

    private void CancelProcessing(object sender, EventArgs e)
    {
        if (!worker.IsBusy || string.IsNullOrWhiteSpace(currentCancelFile)) return;
        DialogResult answer = MessageBox.Show(
            this,
            "确定取消当前拆分/合并处理吗？\r\n\r\n源文件不会被修改，未完成的临时结果会被删除。",
            "确认取消",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;
        try
        {
            File.WriteAllText(currentCancelFile, "cancel", Encoding.UTF8);
            cancelRequested = true;
            cancelButton.Enabled = false;
            cancelButton.Text = "正在取消…";
            lastProgressMessage = "正在安全取消，请稍候……";
            UpdateProgressText();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "无法发送取消请求：\r\n" + ex.Message, "取消失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void ProcessingProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        string line = e.UserState as string;
        if (string.IsNullOrWhiteSpace(line)) return;
        string[] parts = line.Split(new char[] { '\t' }, 5);
        if (parts.Length != 5 || parts[0] != "@@TABLE_PROGRESS") return;
        long completed;
        long total;
        if (!long.TryParse(parts[2], out completed)) completed = 0;
        if (!long.TryParse(parts[3], out total)) total = 0;

        string phase = parts[1];
        if ((phase == "merge" || phase == "split") && lastProgressPhase != phase)
            rateTimer.Restart();
        lastProgressPhase = phase;
        lastProgressCompleted = completed;
        lastProgressTotal = total;
        lastProgressMessage = parts[4];

        if ((phase == "merge" || phase == "split" || phase == "done") && total > 0)
        {
            tableProgress.Style = ProgressBarStyle.Continuous;
            int percent = (int)Math.Max(0, Math.Min(100, completed * 100L / total));
            tableProgress.Value = phase == "done" ? 100 : percent;
        }
        else
        {
            tableProgress.Style = ProgressBarStyle.Marquee;
            tableProgress.MarqueeAnimationSpeed = 25;
        }
        UpdateProgressText();
    }

    private void ProgressTimerTick(object sender, EventArgs e)
    {
        UpdateProgressText();
    }

    private void UpdateProgressText()
    {
        if (!operationTimer.IsRunning) return;
        string elapsed = FormatDuration(operationTimer.Elapsed);
        string remaining = "正在计算……";
        if (lastProgressPhase == "saving")
            remaining = "正在完成保存";
        else if ((lastProgressPhase == "merge" || lastProgressPhase == "split") &&
                 lastProgressTotal > 0 && lastProgressCompleted >= lastProgressTotal)
            remaining = "不足1秒";
        else if ((lastProgressPhase == "merge" || lastProgressPhase == "split") &&
                 lastProgressCompleted > 0 && lastProgressTotal > lastProgressCompleted &&
                 rateTimer.IsRunning && rateTimer.Elapsed.TotalSeconds >= 3)
        {
            double rowsPerSecond = lastProgressCompleted / rateTimer.Elapsed.TotalSeconds;
            if (rowsPerSecond > 0)
            {
                double seconds = (lastProgressTotal - lastProgressCompleted) / rowsPerSecond;
                remaining = FormatDuration(TimeSpan.FromSeconds(seconds));
            }
        }
        progressStatus.Text = lastProgressMessage + "\r\n已用时：" + elapsed + "；预计剩余：" + remaining;
        progressStatus.ForeColor = cancelRequested ? Color.DarkOrange : Color.FromArgb(29, 112, 184);
    }

    private static string FormatDuration(TimeSpan duration)
    {
        int totalSeconds = Math.Max(0, (int)Math.Round(duration.TotalSeconds));
        if (totalSeconds < 60) return totalSeconds + "秒";
        return (totalSeconds / 60) + "分" + (totalSeconds % 60) + "秒";
    }

    private void FinishTableProgressUi()
    {
        progressTimer.Stop();
        operationTimer.Stop();
        rateTimer.Stop();
        tableProgress.Visible = false;
        cancelButton.Visible = false;
        cancelButton.Enabled = true;
        processButton.Loading = false;
        processButton.Size = new Size(196, 48);
        LayoutModernContent();
        if (!string.IsNullOrWhiteSpace(currentCancelFile))
        {
            try
            {
                if (File.Exists(currentCancelFile)) File.Delete(currentCancelFile);
            }
            catch { }
        }
        currentCancelFile = string.Empty;
    }

    private void StartProcessing(object sender, EventArgs e)
    {
        if (worker.IsBusy) return;
        string source = sourceBox.Text.Trim();
        int startPort;
        int maxRows = 1000;
        bool selfTelecom = SelectedMode == "self_telecom";
        bool tableTools = SelectedMode == "table_tools";
        bool consistency = SelectedMode == "signature_consistency";
        bool portAllocation = SelectedMode == "port_allocation";
        bool tableMerge = tableTools && SelectedTableAction == "merge";
        bool requiresDirectory = selfTelecom || tableMerge;
        if ((requiresDirectory && !Directory.Exists(source)) || (!requiresDirectory && !File.Exists(source)))
        {
            string prompt = selfTelecom ? "请选择有效的批次文件夹。"
                : tableMerge ? "请选择有效的分表文件夹。"
                : tableTools ? "请选择有效的总表文件。"
                : "请选择有效的客户文件。";
            MessageBox.Show(this, prompt, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (portAllocation && !File.Exists(masterBox.Text.Trim()))
        {
            MessageBox.Show(this, "请选择最新的历史总表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (SelectedMode == "self_mobile" && !Directory.Exists(masterBox.Text.Trim()))
        {
            MessageBox.Show(this, "请选择或拖入有效的移动签名总表文件夹。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        bool selfOwned = SelectedMode == "self_mobile" || selfTelecom;
        if (!selfOwned && !tableTools && !consistency && !portAllocation && string.IsNullOrWhiteSpace(channelBox.Text))
        {
            MessageBox.Show(this, "请输入通道编号。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (selfOwned || tableTools || consistency || portAllocation)
            startPort = 0;
        else if (!int.TryParse(startPortBox.Text.Trim(), out startPort) || startPort < 0)
        {
            MessageBox.Show(this, "起始子端口必须是大于或等于 0 的整数。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (tableTools && !tableMerge &&
            (!int.TryParse(maxRowsBox.Text.Trim(), out maxRows) || maxRows < 2))
        {
            MessageBox.Show(this, "每个分表最大行数必须是大于或等于 2 的整数（包含表头）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string extension = requiresDirectory ? "" : Path.GetExtension(source).ToLowerInvariant();
        bool supported = requiresDirectory || (SelectedMode == "self_mobile"
            ? extension == ".xlsx"
            : consistency
            ? extension == ".xlsx" || extension == ".xlsm"
            : extension == ".xlsx" || extension == ".csv" ||
                         (SelectedMode == "signature" && extension == ".txt"));
        if (!supported)
        {
            MessageBox.Show(this, "当前处理方式不支持该文件格式。请先把旧版 .xls 另存为 .xlsx。", "格式不支持", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        UpdateModeDisplay();
        if (templateStatus.ForeColor == UiTheme.Danger)
        {
            MessageBox.Show(this, "内置资料缺失，请重新安装程序。", "缺少资料", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ProcessingRequest request = new ProcessingRequest();
        request.Mode = SelectedMode;
        request.Source = source;
        request.Master = masterBox.Text.Trim();
        request.Channel = channelBox.Text.Trim();
        request.StartPort = startPort;
        request.Operator = SelectedMode == "self_mobile" ? "移动" : selfTelecom ? "电信" : operatorBox.Text;
        request.Business = businessBox.Text;
        request.Important = importantBox.Text;
        request.TableAction = tableTools ? SelectedTableAction : "";
        request.SignatureMatchMode = consistency ? SelectedSignatureMatchMode : "full_rule";
        request.MaxRows = maxRows;
        request.OutputDirectory = outputDirectoryBox.Text.Trim();
        if (!string.IsNullOrWhiteSpace(request.OutputDirectory))
        {
            try { request.OutputDirectory = Path.GetFullPath(request.OutputDirectory); }
            catch (Exception)
            {
                MessageBox.Show(this, "另存位置不是有效的文件夹路径。", "路径无效",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        request.CancelFile = tableTools
            ? Path.Combine(Path.GetTempPath(), "channel-table-cancel-" + Guid.NewGuid().ToString("N") + ".flag")
            : string.Empty;

        currentCancelFile = request.CancelFile;
        cancelRequested = false;
        lastProgressPhase = string.Empty;
        lastProgressMessage = tableMerge ? "正在准备合并分表……" : "正在准备拆分总表……";
        lastProgressCompleted = 0;
        lastProgressTotal = 0;
        operationTimer.Restart();
        rateTimer.Reset();
        if (tableTools)
        {
            cancelButton.Text = "取消处理";
            cancelButton.Enabled = true;
            cancelButton.Visible = true;
            tableProgress.Style = ProgressBarStyle.Marquee;
            tableProgress.MarqueeAnimationSpeed = 25;
            tableProgress.Visible = true;
            progressTimer.Start();
        }

        processButton.Enabled = false;
        processButton.Loading = true;
        browseButton.Enabled = false;
        browseMasterButton.Enabled = false;
        outputDirectoryBox.Enabled = false;
        browseOutputButton.Enabled = false;
        LayoutModernContent();
        progressStatus.Text = tableTools
            ? (tableMerge ? "正在读取并合并分表，请稍候……" : "正在拆分总表，请稍候……")
            : "正在处理，请稍候……首次使用时会先准备运行环境。";
        progressStatus.ForeColor = Color.FromArgb(29, 112, 184);
        worker.RunWorkerAsync(request);
    }

    private void DoProcessing(object sender, DoWorkEventArgs e)
    {
        ProcessingRequest request = (ProcessingRequest)e.Argument;
        List<string> args = new List<string>
        {
            "--cli", "--mode", request.Mode, "--source", request.Source,
            "--channel", request.Channel, "--start-sub-port", request.StartPort.ToString(),
            "--operator", request.Operator, "--business-type", request.Business,
            "--important-support", request.Important, "--template-dir", Program.ResourcesDirectory
        };
        // 拆分/合并参数只属于新增模式。原五种模式不传空值，避免命令行
        // 参数选择校验在业务处理开始前将它们全部拦截。
        if (request.Mode == "table_tools")
        {
            args.Add("--table-action");
            args.Add(request.TableAction);
            args.Add("--max-rows");
            args.Add(request.MaxRows.ToString());
            args.Add("--cancel-file");
            args.Add(request.CancelFile);
        }
        if (request.Mode == "signature_consistency")
        {
            args.Add("--signature-match-mode");
            args.Add(request.SignatureMatchMode);
        }
        if (request.Mode == "port_allocation")
        {
            args.Add("--master");
            args.Add(request.Master);
        }
        if (request.Mode == "self_mobile")
        {
            args.Add("--signature-total-dir");
            args.Add(request.Master);
        }
        if (!string.IsNullOrWhiteSpace(request.OutputDirectory))
        {
            args.Add("--output-dir");
            args.Add(request.OutputDirectory);
        }
        BackgroundWorker currentWorker = (BackgroundWorker)sender;
        PythonResult result = Program.ExecutePython(
            args,
            delegate(string line)
            {
                if (request.Mode == "table_tools") currentWorker.ReportProgress(0, line);
            });
        e.Result = new ProcessingResponse(request.Mode, request.OutputDirectory, result);
    }

    private void ProcessingCompleted(object sender, RunWorkerCompletedEventArgs e)
    {
        processButton.Enabled = true;
        processButton.Loading = false;
        browseButton.Enabled = true;
        browseMasterButton.Enabled = true;
        outputDirectoryBox.Enabled = true;
        browseOutputButton.Enabled = true;
        bool tableOperation = tableProgress.Visible || !string.IsNullOrWhiteSpace(currentCancelFile);
        if (tableOperation) FinishTableProgressUi();

        if (e.Error != null)
        {
            if (cancelRequested)
            {
                progressStatus.Text = "处理已取消，未生成不完整结果。";
                progressStatus.ForeColor = Color.DarkOrange;
                cancelRequested = false;
                return;
            }
            progressStatus.Text = "处理失败";
            progressStatus.ForeColor = Color.Firebrick;
            MessageBox.Show(this, e.Error.Message, "处理失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        ProcessingResponse response = (ProcessingResponse)e.Result;
        if (response.Result.ExitCode != 0)
        {
            if (cancelRequested || response.Result.Error.Contains("处理已取消"))
            {
                progressStatus.Text = "处理已取消，源文件未修改，未完成结果已删除。";
                progressStatus.ForeColor = Color.DarkOrange;
                cancelRequested = false;
                return;
            }
            progressStatus.Text = "处理失败";
            progressStatus.ForeColor = Color.Firebrick;
            string message = string.IsNullOrWhiteSpace(response.Result.Error) ? "程序未能完成处理。" : response.Result.Error.Trim();
            MessageBox.Show(this, message, "处理失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            Dictionary<string, object> data = serializer.Deserialize<Dictionary<string, object>>(response.Result.Output.Trim());
            long generated = ToLong(data, "generated_rows");
            long duplicates = ToLong(data, "duplicate_rows");
            long errors = ToLong(data, "error_rows");
            StringBuilder message = new StringBuilder();
            message.AppendLine("处理完成！");
            bool tableResult = response.Mode == "table_tools";
            if (response.Mode == "port_allocation")
            {
                message.AppendLine("已分配：" + generated + " 行；每行各有行业、营销一个号码");
                AppendPath(message, data, "output_path", "分配结果");
                AppendPath(message, data, "secondary_output_path", "更新总表");
                message.AppendLine("下次分配请使用这份更新总表。");
            }
            else if (tableResult)
            {
                object actionValue;
                string action = data.TryGetValue("table_action", out actionValue) && actionValue != null
                    ? actionValue.ToString() : "";
                if (action == "split")
                {
                    message.AppendLine("总表拆分完成");
                    message.AppendLine("数据行数：" + generated + " 条");
                    message.AppendLine("生成分表：" + ToLong(data, "output_file_count") + " 个");
                    AppendPath(message, data, "output_dir", "结果文件夹");
                }
                else
                {
                    message.AppendLine("分表合并完成");
                    message.AppendLine("读取分表：" + ToLong(data, "source_file_count") + " 个");
                    message.AppendLine("合并数据：" + generated + " 条");
                    AppendPath(message, data, "output_path", "合并结果");
                }
            }
            else if (response.Mode == "self_mobile")
            {
                message.AppendLine("生效引流：" + ToLong(data, "shortlink_rows") + " 条");
                message.AppendLine("非已生效并跳过：" + ToLong(data, "skipped_rows") + " 条");
                message.AppendLine("非本公司主码号并忽略：" + ToLong(data, "dirty_rows") + " 条");
            }
            else if (response.Mode == "self_telecom")
            {
                message.AppendLine("已识别提交表：" + ToLong(data, "batch_count") + " 份");
                message.AppendLine("失败回执剔除：" + ToLong(data, "failed_rows") + " 条");
                message.AppendLine("生效短链：" + ToLong(data, "shortlink_rows") + " 条");
                message.AppendLine("生效签名：" + ToLong(data, "signature_rows") + " 条");
                message.AppendLine("最终引流报备：" + ToLong(data, "final_diversion_rows") + " 条");
                long unrelated = ToLong(data, "dirty_rows");
                if (unrelated > 0) message.AppendLine("回执中非本批次子端口并忽略：" + unrelated + " 条");
            }
            else if (response.Mode == "mixed")
            {
                message.AppendLine("生效短链：" + ToLong(data, "shortlink_rows") + " 条");
                message.AppendLine("生效签名：" + ToLong(data, "signature_rows") + " 条");
            }
            else if (response.Mode == "signature_consistency")
            {
                message.AppendLine("读取数据：" + ToLong(data, "source_rows") + " 条");
                message.AppendLine("第一层通过：" + ToLong(data, "first_layer_rows") + " 条");
                message.AppendLine("第一层过滤：" + ToLong(data, "first_layer_filtered_rows") + " 条");
                message.AppendLine("第二层通过：" + ToLong(data, "second_layer_rows") + " 条");
                message.AppendLine("第二层过滤：" + ToLong(data, "second_layer_filtered_rows") + " 条");
                message.AppendLine("企业类型判定失败：" + ToLong(data, "enterprise_type_filtered_rows") + " 条");
                message.AppendLine("处理工作表：" + ToLong(data, "processed_sheets") + " 个");
            }
            else
                message.AppendLine("生成数据：" + generated + " 条");
            if (!tableResult && response.Mode != "signature_consistency" && response.Mode != "port_allocation")
            {
                message.AppendLine("去除完全重复：" + duplicates + " 条");
                message.AppendLine("异常数据：" + errors + " 条");
            }

            if (!tableResult && response.Mode != "port_allocation")
            {
                AppendPath(message, data, "shortlink_output_path", "短链结果");
                if (response.Mode != "self_mobile") AppendPath(message, data, "signature_output_path", "签名结果");
                if (response.Mode == "self_telecom")
                {
                    AppendPath(message, data, "final_diversion_output_path", "最终引流报备表");
                    AppendPath(message, data, "failure_reason_output_path", "签名错误原因");
                }
                if (response.Mode != "mixed" && response.Mode != "self_mobile" && response.Mode != "self_telecom") AppendPath(message, data, "output_path", "结果文件");
                AppendPath(message, data, "error_path", "错误明细");
            }

            progressStatus.Text = !string.IsNullOrWhiteSpace(response.OutputDirectory)
                ? "处理完成，结果已保存到指定文件夹。"
                : tableResult
                ? "处理完成，结果已保存到“生成结果”文件夹。"
                : "处理完成，结果已保存到原文件所在文件夹。";
            progressStatus.ForeColor = Color.DarkGreen;
            cancelRequested = false;
            MessageBox.Show(this, message.ToString().Trim(), "处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            progressStatus.Text = "处理已完成，但结果提示解析失败。";
            MessageBox.Show(this, "文件已生成，但无法显示完整结果信息：\r\n" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static long ToLong(Dictionary<string, object> data, string key)
    {
        object value;
        if (!data.TryGetValue(key, out value) || value == null) return 0;
        return Convert.ToInt64(value);
    }

    private static void AppendPath(StringBuilder message, Dictionary<string, object> data, string key, string label)
    {
        object value;
        if (data.TryGetValue(key, out value) && value != null && !string.IsNullOrWhiteSpace(value.ToString()))
            message.AppendLine(label + "：" + value);
    }
}

internal sealed class UsageGuideForm : Form
{
    private const string GuideText =
        "一、基本操作\r\n" +
        "1. 从左侧菜单选择需要的功能。\r\n" +
        "2. 点击选择按钮，或把一个文件 / 文件夹拖入待处理区域。\r\n" +
        "3. 按页面要求填写参数；需要指定输出目录时，点击底部“选择位置”。\r\n" +
        "4. 点击开始处理。完成弹窗会显示全部结果路径。\r\n\r\n" +
        "二、通道业务\r\n" +
        "自家移动：选择 .xlsx 关键信息表和最新移动签名总表文件夹，再选择是否重保。\r\n\r\n" +
        "自家电信：选择包含提交表和失败回执的完整电信批次文件夹，再选择是否重保。\r\n\r\n" +
        "子端口分配：选择待分配表和最新历史总表。下一次分配必须使用本次生成的更新总表。\r\n\r\n" +
        "生效短链：支持 .xlsx、.csv；填写通道编号、起始子端口、运营商、业务类型和是否重保。\r\n\r\n" +
        "生效签名：支持 .xlsx、.csv、.txt；没有签名表头时默认读取第一列。\r\n\r\n" +
        "混合表自动分流：支持 .xlsx、.csv；签名与短链数据会自动分流并连续分配子端口。\r\n\r\n" +
        "三、数据工具\r\n" +
        "拆分总表：选择 .xlsx 或 .csv，并设置包含表头的最大行数。\r\n\r\n" +
        "合并分表：选择包含全部分表的文件夹；文件格式和表头必须一致。处理中可以安全取消。\r\n\r\n" +
        "签名一致性：支持 .xlsx、.xlsm，可选择完整规则或只匹配地区。\r\n\r\n" +
        "四、结果保存\r\n" +
        "另存位置留空时使用默认目录。自家移动、自家电信、子端口分配、混合表自动分流选择自定义目录后，会建立“生成结果_功能名_mmdd-HH时mm分”文件夹。\r\n" +
        "程序不会覆盖原始文件；同名结果会自动增加 (1)、(2)。\r\n\r\n" +
        "五、主码号管理\r\n" +
        "用于新增、修改、删除和导出通道编号及对应主码号资料。建议修改前先导出表格备份。\r\n\r\n" +
        "六、注意事项\r\n" +
        "· 一次只拖入一个文件或文件夹。\r\n" +
        "· 保存失败时，请先关闭 Excel 中正在打开的同名文件。\r\n" +
        "· 所有数据仅在本机处理，不经过网络传输。";

    internal UsageGuideForm()
    {
        Text = "使用说明";
        Program.ApplyWindowIdentity(this);
        ClientSize = new Size(880, 720);
        MinimumSize = new Size(720, 560);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = UiTheme.Surface;
        Font = UiFonts.Chinese(10F);

        Panel header = new Panel { Dock = DockStyle.Top, Height = 88, BackColor = UiTheme.Surface };
        header.Paint += delegate(object sender, PaintEventArgs e)
        {
            using (Pen pen = new Pen(UiTheme.Border, 1F))
                e.Graphics.DrawLine(pen, 0, header.Height - 1, header.Width, header.Height - 1);
        };
        Controls.Add(header);

        Label title = new Label {
            Text = "使用说明", AutoSize = true, Left = 28, Top = 18,
            Font = UiFonts.Chinese(17F, FontStyle.Bold), ForeColor = UiTheme.Text
        };
        header.Controls.Add(title);
        Label subtitle = new Label {
            Text = "通道业务、数据工具与结果保存操作指南", AutoSize = true, Left = 30, Top = 53,
            Font = UiFonts.Chinese(9F), ForeColor = UiTheme.TextMuted
        };
        header.Controls.Add(subtitle);

        Panel body = new Panel { Dock = DockStyle.Fill, BackColor = UiTheme.Surface, Padding = new Padding(28, 22, 28, 26) };
        Controls.Add(body);
        header.BringToFront();

        RichTextBox guide = new RichTextBox {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = UiTheme.Surface,
            ForeColor = UiTheme.TextSecondary,
            Font = UiFonts.Chinese(10F),
            ReadOnly = true,
            DetectUrls = false,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            Text = GuideText
        };
        body.Controls.Add(guide);
        FormatHeadings(guide);
        guide.SelectionStart = 0;
        guide.SelectionLength = 0;
    }

    private static void FormatHeadings(RichTextBox guide)
    {
        string[] headings = new string[] {
            "一、基本操作", "二、通道业务", "三、数据工具",
            "四、结果保存", "五、主码号管理", "六、注意事项"
        };
        foreach (string heading in headings)
        {
            int index = guide.Text.IndexOf(heading, StringComparison.Ordinal);
            if (index < 0) continue;
            guide.Select(index, heading.Length);
            guide.SelectionFont = UiFonts.Chinese(12F, FontStyle.Bold);
            guide.SelectionColor = UiTheme.Text;
        }
    }
}

internal sealed class SurfacePanel : Panel
{
    internal SurfacePanel()
    {
        BackColor = UiTheme.Surface;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    private static System.Drawing.Drawing2D.GraphicsPath Rounded(Rectangle rectangle, int radius)
    {
        int diameter = radius * 2;
        System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }

    protected override void OnResize(EventArgs eventArgs)
    {
        base.OnResize(eventArgs);
        if (Width <= 0 || Height <= 0) return;
        using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(new Rectangle(0, 0, Width, Height), UiTheme.Radius))
            Region = new Region(path);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        Rectangle border = new Rectangle(0, 0, Width - 1, Height - 1);
        using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(border, UiTheme.Radius))
        using (Pen pen = new Pen(UiTheme.Border, 1F))
            e.Graphics.DrawPath(pen, path);
        base.OnPaint(e);
    }
}

internal sealed class RoundedSidebarPanel : Panel
{
    internal RoundedSidebarPanel()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        const int headerY = 72;
        const int radius = 6;
        int right = Width - 1;
        int bottom = Height - 1;
        using (Pen pen = new Pen(UiTheme.BorderStrong, 1F))
        using (System.Drawing.Drawing2D.GraphicsPath outline = new System.Drawing.Drawing2D.GraphicsPath())
        {
            outline.AddLine(0, 0, right - radius, 0);
            outline.AddArc(right - (radius * 2), 0, radius * 2, radius * 2, 270, 90);
            outline.AddLine(right, radius, right, bottom);
            outline.AddLine(right, bottom, 0, bottom);
            outline.AddLine(0, bottom, 0, 0);
            e.Graphics.DrawPath(pen, outline);
            e.Graphics.DrawLine(pen, 0, headerY, right, headerY);
        }
    }
}

internal sealed class RoundedMainHeaderPanel : Panel
{
    internal RoundedMainHeaderPanel()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        const int radius = 6;
        int bottom = Height - 1;
        using (Pen pen = new Pen(UiTheme.BorderStrong, 1F))
        using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
        {
            path.AddLine(Width, 0, radius, 0);
            path.AddArc(0, 0, radius * 2, radius * 2, 270, -90);
            path.AddLine(0, radius, 0, bottom);
            path.AddLine(0, bottom, Width, bottom);
            e.Graphics.DrawPath(pen, path);
        }
    }
}

internal sealed class RoundedMainContentPanel : Panel
{
    internal RoundedMainContentPanel()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (Pen pen = new Pen(UiTheme.BorderStrong, 1F))
        {
            e.Graphics.DrawLine(pen, 0, 0, Width, 0);
            e.Graphics.DrawLine(pen, 0, 0, 0, Height);
        }
    }
}

internal sealed class VersionBadge : Control
{
    internal VersionBadge()
    {
        Font = UiFonts.Chinese(8.5F, FontStyle.Bold);
        ForeColor = UiTheme.Text;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        Rectangle bounds = new Rectangle(0, 0, Width - 1, Height - 1);
        int radius = Math.Min(10, Height / 2);
        int diameter = radius * 2;
        using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
        using (SolidBrush fill = new SolidBrush(UiTheme.SurfaceMuted))
        using (Pen border = new Pen(UiTheme.Border, 1F))
        {
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            e.Graphics.FillPath(fill, path);
            e.Graphics.DrawPath(border, path);
        }
        TextRenderer.DrawText(e.Graphics, Text, Font, bounds, ForeColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
    }
}

internal sealed class NavigationButton : System.Windows.Forms.Button
{
    private bool hovered;
    private bool pressed;
    internal bool Selected { get; set; }
    internal string IconKey { get; set; }

    internal NavigationButton()
    {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        FlatAppearance.MouseDownBackColor = UiTheme.Border;
        FlatAppearance.MouseOverBackColor = UiTheme.SurfaceBlue;
        UseVisualStyleBackColor = false;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer, true);
    }

    protected override void OnMouseEnter(EventArgs eventArgs)
    {
        hovered = true;
        Invalidate();
        base.OnMouseEnter(eventArgs);
    }

    protected override void OnMouseLeave(EventArgs eventArgs)
    {
        hovered = false;
        Invalidate();
        base.OnMouseLeave(eventArgs);
    }

    protected override void OnMouseDown(MouseEventArgs eventArgs)
    {
        pressed = true;
        Invalidate();
        base.OnMouseDown(eventArgs);
    }

    protected override void OnMouseUp(MouseEventArgs eventArgs)
    {
        pressed = false;
        Invalidate();
        base.OnMouseUp(eventArgs);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        e.Graphics.Clear(UiTheme.Surface);
        Rectangle fill = new Rectangle(0, 0, Width - 1, Height - 1);
        Color background = Selected ? UiTheme.Primary : pressed ? UiTheme.Border
            : hovered ? UiTheme.SurfaceBlue : UiTheme.Surface;
        if (Selected || pressed || hovered)
        {
            using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(fill, 6))
            using (SolidBrush brush = new SolidBrush(background))
                e.Graphics.FillPath(brush, path);
        }
        Color foreground = Selected ? Color.White : UiTheme.TextSecondary;
        TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(18, pressed ? 1 : 0, Width - 30, Height),
            foreground,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private static System.Drawing.Drawing2D.GraphicsPath Rounded(Rectangle rectangle, int radius)
    {
        int diameter = radius * 2;
        System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}

internal static class UiIconPainter
{
    internal static void Draw(Graphics graphics, string key, Rectangle bounds, Color color, float width)
    {
        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (Pen pen = new Pen(color, width))
        {
            pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
            pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
            int left = bounds.Left + 2;
            int top = bounds.Top + 2;
            int right = bounds.Right - 3;
            int bottom = bounds.Bottom - 3;
            int midX = bounds.Left + bounds.Width / 2;
            int midY = bounds.Top + bounds.Height / 2;
            if (key == "upload")
            {
                graphics.DrawLine(pen, midX, top, midX, bottom - 4);
                graphics.DrawLine(pen, midX, top, midX - 5, top + 5);
                graphics.DrawLine(pen, midX, top, midX + 5, top + 5);
                graphics.DrawLine(pen, left, bottom, right, bottom);
            }
            else if (key == "folder" || key == "table")
            {
                graphics.DrawRectangle(pen, left, top + 4, right - left, bottom - top - 4);
                graphics.DrawLine(pen, left + 2, top + 4, left + 7, top);
                graphics.DrawLine(pen, left + 7, top, midX + 2, top);
            }
            else if (key == "mobile")
            {
                graphics.DrawRectangle(pen, left + 4, top, right - left - 8, bottom - top);
                graphics.DrawLine(pen, midX - 2, bottom - 3, midX + 2, bottom - 3);
            }
            else if (key == "telecom")
            {
                graphics.DrawEllipse(pen, midX - 3, top, 6, 6);
                graphics.DrawArc(pen, midX - 7, top + 4, 14, 12, 210, 120);
                graphics.DrawArc(pen, midX - 10, top + 2, 20, 18, 210, 120);
            }
            else if (key == "port")
            {
                graphics.DrawLine(pen, left, midY, right, midY);
                graphics.DrawEllipse(pen, left, midY - 3, 6, 6);
                graphics.DrawEllipse(pen, right - 6, midY - 3, 6, 6);
                graphics.DrawLine(pen, midX, midY, midX, bottom);
            }
            else if (key == "link")
            {
                graphics.DrawArc(pen, left, midY - 5, 11, 10, 45, 270);
                graphics.DrawArc(pen, right - 11, midY - 5, 11, 10, 225, 270);
                graphics.DrawLine(pen, midX - 4, midY, midX + 4, midY);
            }
            else if (key == "signature")
            {
                graphics.DrawLine(pen, left, bottom, right, bottom);
                graphics.DrawLine(pen, left + 3, bottom - 4, right - 2, top + 1);
                graphics.DrawLine(pen, right - 2, top + 1, right, top + 5);
            }
            else if (key == "route")
            {
                graphics.DrawLine(pen, left, top + 3, midX, top + 3);
                graphics.DrawLine(pen, midX, top + 3, midX, bottom - 3);
                graphics.DrawLine(pen, midX, midY, right, midY);
                graphics.DrawLine(pen, midX, bottom - 3, right, bottom - 3);
            }
            else if (key == "check")
            {
                graphics.DrawEllipse(pen, left, top, right - left, bottom - top);
                graphics.DrawLine(pen, left + 4, midY, midX - 1, bottom - 4);
                graphics.DrawLine(pen, midX - 1, bottom - 4, right - 3, top + 4);
            }
            else if (key == "database")
            {
                graphics.DrawEllipse(pen, left, top, right - left, 6);
                graphics.DrawLine(pen, left, top + 3, left, bottom - 3);
                graphics.DrawLine(pen, right, top + 3, right, bottom - 3);
                graphics.DrawArc(pen, left, bottom - 7, right - left, 7, 0, 180);
            }
            else
            {
                graphics.DrawRectangle(pen, left, top, right - left, bottom - top);
            }
        }
    }
}

internal sealed class VectorIcon : Control
{
    internal string IconKey { get; set; }

    internal VectorIcon()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        UiIconPainter.Draw(e.Graphics, IconKey, ClientRectangle, ForeColor, 2F);
    }
}

internal sealed class StatusDot : Control
{
    internal Color StateColor { get; set; }

    internal StatusDot()
    {
        StateColor = UiTheme.Success;
        SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (SolidBrush halo = new SolidBrush(Color.FromArgb(35, StateColor)))
            e.Graphics.FillEllipse(halo, 0, 0, Width - 1, Height - 1);
        using (SolidBrush core = new SolidBrush(StateColor))
            e.Graphics.FillEllipse(core, 4, 4, Math.Max(2, Width - 8), Math.Max(2, Height - 8));
    }
}

internal sealed class SecurityBadge : Control
{
    internal SecurityBadge()
    {
        Font = UiFonts.Chinese(11F, FontStyle.Bold);
        SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(new Rectangle(0, 0, Width - 1, Height - 1), 7))
        using (SolidBrush brush = new SolidBrush(UiTheme.Surface))
        using (Pen border = new Pen(UiTheme.Border, 1F))
        {
            e.Graphics.FillPath(brush, path);
            e.Graphics.DrawPath(border, path);
        }
        DrawShield(e.Graphics, new Rectangle(16, 13, 22, 22), UiTheme.TextSecondary);
        TextRenderer.DrawText(e.Graphics, "本地安全处理", Font, new Rectangle(48, 0, Width - 58, Height),
            UiTheme.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
    }

    internal static void DrawShield(Graphics graphics, Rectangle bounds, Color color)
    {
        Point[] points = new Point[] {
            new Point(bounds.Left + bounds.Width / 2, bounds.Top),
            new Point(bounds.Right - 2, bounds.Top + 4),
            new Point(bounds.Right - 4, bounds.Bottom - 5),
            new Point(bounds.Left + bounds.Width / 2, bounds.Bottom - 1),
            new Point(bounds.Left + 4, bounds.Bottom - 5),
            new Point(bounds.Left + 2, bounds.Top + 4)
        };
        using (Pen pen = new Pen(color, 1.6F)) graphics.DrawPolygon(pen, points);
        using (Pen pen = new Pen(color, 1.6F))
        {
            graphics.DrawLine(pen, bounds.Left + 6, bounds.Top + 9, bounds.Left + 9, bounds.Top + 12);
            graphics.DrawLine(pen, bounds.Left + 9, bounds.Top + 12, bounds.Right - 5, bounds.Top + 6);
        }
    }

    private static System.Drawing.Drawing2D.GraphicsPath Rounded(Rectangle rectangle, int radius)
    {
        int diameter = radius * 2;
        System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}

internal sealed class SecurityNotice : Panel
{
    internal SecurityNotice()
    {
        BackColor = UiTheme.SurfaceMuted;
        SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer |
            ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (Pen border = new Pen(UiTheme.Border, 1F)) e.Graphics.DrawLine(border, 0, 0, Width, 0);
        SecurityBadge.DrawShield(e.Graphics, new Rectangle(16, 19, 20, 22), UiTheme.TextSecondary);
        TextRenderer.DrawText(e.Graphics, "数据仅在本机处理", UiFonts.Chinese(9.5F, FontStyle.Bold),
            new Rectangle(46, 10, Width - 58, 24), UiTheme.Text,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        TextRenderer.DrawText(e.Graphics, "不经过网络传输", UiFonts.Chinese(8.5F),
            new Rectangle(46, 32, Width - 58, 20), UiTheme.TextSecondary,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
    }
}

internal sealed class ResourceRow : Control
{
    internal string ResourceName { get; set; }
    internal bool Available { get; set; }

    internal ResourceRow()
    {
        ResourceName = string.Empty;
        Available = true;
        Font = UiFonts.Chinese(8.5F);
        SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        Color state = Available ? UiTheme.Success : UiTheme.Danger;
        using (Pen pen = new Pen(UiTheme.BorderStrong, 1.4F))
        {
            e.Graphics.DrawRectangle(pen, 3, 7, 18, 22);
            e.Graphics.DrawLine(pen, 14, 7, 21, 14);
        }
        using (SolidBrush brush = new SolidBrush(state)) e.Graphics.FillEllipse(brush, 15, 22, 11, 11);
        using (Pen pen = new Pen(Color.White, 1.4F))
        {
            if (Available)
            {
                e.Graphics.DrawLine(pen, 18, 27, 20, 29);
                e.Graphics.DrawLine(pen, 20, 29, 24, 24);
            }
            else
            {
                e.Graphics.DrawLine(pen, 18, 25, 23, 30);
                e.Graphics.DrawLine(pen, 23, 25, 18, 30);
            }
        }
        TextRenderer.DrawText(e.Graphics, ResourceName, Font, new Rectangle(34, 3, Width - 88, Height - 6),
            UiTheme.TextSecondary, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        TextRenderer.DrawText(e.Graphics, Available ? "就绪" : "缺失", UiFonts.Chinese(8F, FontStyle.Bold),
            new Rectangle(Width - 50, 3, 46, Height - 6), state,
            TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
    }
}

internal sealed class DashedPanel : Panel
{
    private bool dragActive;

    internal bool DragActive
    {
        get { return dragActive; }
        set
        {
            if (dragActive == value) return;
            dragActive = value;
            BackColor = value ? Color.FromArgb(242, 242, 242) : UiTheme.SurfaceMuted;
            Invalidate();
        }
    }

    internal DashedPanel()
    {
        BackColor = UiTheme.SurfaceMuted;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (Pen pen = new Pen(DragActive ? UiTheme.Text : UiTheme.BorderStrong, DragActive ? 1.8F : 1F))
        {
            pen.DashStyle = DragActive
                ? System.Drawing.Drawing2D.DashStyle.Solid
                : System.Drawing.Drawing2D.DashStyle.Dash;
            Rectangle border = new Rectangle(0, 0, Width - 1, Height - 1);
            using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(border, 8))
                e.Graphics.DrawPath(pen, path);
        }
    }

    private static System.Drawing.Drawing2D.GraphicsPath Rounded(Rectangle rectangle, int radius)
    {
        int diameter = radius * 2;
        System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
        path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();
        return path;
    }
}

internal sealed class ProcessingRequest
{
    internal string Mode;
    internal string Source;
    internal string Master;
    internal string Channel;
    internal int StartPort;
    internal string Operator;
    internal string Business;
    internal string Important;
    internal string TableAction;
    internal string SignatureMatchMode;
    internal int MaxRows;
    internal string OutputDirectory;
    internal string CancelFile;
}

internal sealed class ProcessingResponse
{
    internal ProcessingResponse(string mode, string outputDirectory, PythonResult result)
    {
        Mode = mode;
        OutputDirectory = outputDirectory;
        Result = result;
    }

    internal string Mode { get; private set; }
    internal string OutputDirectory { get; private set; }
    internal PythonResult Result { get; private set; }
}
