using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;

namespace EdgePwaRedirector;

static class Program
{
    internal static readonly string LogDir = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "EdgePwaRedirector");

    static void RunHidden(string exe, string args) =>
        Process.Start(new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true })
               ?.WaitForExit(10000);

    [STAThread]
    static void Main()
    {
        try
        {
            var self = Process.GetCurrentProcess();

            // If not launched through Explorer's shell chain, Shell_NotifyIcon silently
            // fails to register the tray icon. Relaunch via Task Scheduler (/it /rl LIMITED),
            // which creates a proper interactive session process regardless of calling context.
            bool launchedFromExplorer = false;
            try
            {
                var parent = Process.GetProcessById((int)GetParentProcessId(self.Handle));
                launchedFromExplorer = parent.ProcessName.Equals("explorer", StringComparison.OrdinalIgnoreCase);
            }
            catch { }

            bool isRelaunch = Environment.GetCommandLineArgs().Skip(1).Contains("--relaunch");
            if (!launchedFromExplorer && !isRelaunch)
            {
                var exePath = self.MainModule!.FileName;
                var user = $"{Environment.UserDomainName}\\{Environment.UserName}";
                const string taskName = "EdgePwaRedirectorLaunch";
                RunHidden("schtasks.exe", $"/create /tn \"{taskName}\" /tr \"\\\"{exePath}\\\" --relaunch\" /sc once /st 00:00 /f /ru \"{user}\" /it /rl LIMITED");
                RunHidden("schtasks.exe", $"/run /tn \"{taskName}\"");
                Thread.Sleep(2000);
                RunHidden("schtasks.exe", $"/delete /tn \"{taskName}\" /f");
                return;
            }

            // Kill any previous instance — same user always has permission, and this lets
            // the installer hand off cleanly without needing to kill from an elevated context.
            bool killedAny = false;
            foreach (var other in Process.GetProcessesByName(self.ProcessName))
            {
                if (other.Id != self.Id)
                    try { other.Kill(); other.WaitForExit(3000); killedAny = true; } catch { }
            }
            if (killedAny) Thread.Sleep(1000);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApp());
        }
        catch (Exception ex)
        {
            try
            {
                System.IO.Directory.CreateDirectory(LogDir);
                System.IO.File.WriteAllText(System.IO.Path.Combine(LogDir, "crash.log"), ex.ToString());
            }
            catch { }
            MessageBox.Show(ex.ToString(), "EdgePwaRedirector crashed");
        }
    }

    [DllImport("ntdll.dll")]
    static extern int NtQueryInformationProcess(IntPtr hProcess, int processInformationClass, ref PROCESS_BASIC_INFORMATION pbi, int size, out int returnLength);

    static uint GetParentProcessId(IntPtr hProcess)
    {
        var pbi = new PROCESS_BASIC_INFORMATION();
        NtQueryInformationProcess(hProcess, 0, ref pbi, Marshal.SizeOf(pbi), out _);
        return (uint)pbi.InheritedFromUniqueProcessId;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct PROCESS_BASIC_INFORMATION
    {
        public IntPtr Reserved1; public IntPtr PebBaseAddress;
        public IntPtr Reserved2_0; public IntPtr Reserved2_1;
        public IntPtr UniqueProcessId; public IntPtr InheritedFromUniqueProcessId;
    }
}

class TrayApp : ApplicationContext
{
    private readonly NotifyIcon _trayIcon;
    private readonly RedirectService _service;
    private readonly TrayMessageWindow _messageWindow;
    private readonly UpdateChecker _updateChecker = new();
    private readonly WindowsFormsSynchronizationContext _syncCtx = new();
    private ToolStripMenuItem _updateItem = null!;
    private ToolStripMenuItem _pauseItem = null!;
    private ToolStripMenuItem _statusItem = null!;

    public TrayApp()
    {
        _service = new RedirectService();

        _trayIcon = new NotifyIcon
        {
            Icon = CreateTrayIcon(),
            Text = "Edge PWA Redirector",
            Visible = true,
            ContextMenuStrip = BuildContextMenu()
        };

        _messageWindow = new TrayMessageWindow(_trayIcon);
        _service.Start();
    }

    private static Icon CreateTrayIcon()
    {
        var icoPath = System.IO.Path.Combine(AppContext.BaseDirectory, "app.ico");
        if (System.IO.File.Exists(icoPath))
            return new Icon(icoPath);

        // Fallback: generate a simple arrow icon programmatically.
        var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.FillRectangle(new SolidBrush(Color.FromArgb(0x46, 0x4E, 0xB8)), 0, 0, 32, 32);
            var pen = new Pen(Color.White, 3f);
            g.DrawLine(pen, 6, 16, 22, 16);
            g.DrawLine(pen, 16, 8, 26, 16);
            g.DrawLine(pen, 16, 24, 26, 16);
        }
        return Icon.FromHandle(bmp.GetHicon());
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionLabel = new ToolStripMenuItem($"v{version?.ToString(3) ?? "?"}") { Enabled = false };

        _updateItem = new ToolStripMenuItem("Check for updates");
        _updateItem.Click += OnUpdateItemClick;

        _pauseItem = new ToolStripMenuItem("Pause Redirection") { CheckOnClick = true };
        _pauseItem.CheckedChanged += (_, _) =>
        {
            _service.Paused = _pauseItem.Checked;
            _trayIcon.Text = _pauseItem.Checked ? "Edge PWA Redirector (Paused)" : "Edge PWA Redirector";
        };

        _statusItem = new ToolStripMenuItem("No redirects yet") { Enabled = false };

        var menu = new ContextMenuStrip();
        menu.Items.Add(versionLabel);
        menu.Items.Add(_updateItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(_pauseItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(_statusItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) =>
        {
            _service.Stop();
            _trayIcon.Visible = false;
            Application.Exit();
        });

        menu.Opening += (_, _) =>
        {
            UpdateStatus();
            RefreshUpdateItem();
            _updateChecker.CheckInBackground(() => _syncCtx.Post(_ => RefreshUpdateItem(), null));
        };

        return menu;
    }

    private void RefreshUpdateItem()
    {
        var (state, version, _) = _updateChecker.GetState();
        (_updateItem.Text, _updateItem.Enabled) = state switch
        {
            UpdateChecker.UpdateState.Checking        => ("Checking for updates...", false),
            UpdateChecker.UpdateState.UpdateAvailable => ($"Update to v{version?.ToString(3)}", true),
            UpdateChecker.UpdateState.Downloading     => ("Downloading...", false),
            _                                         => ("Check for updates", true),
        };
    }

    private async void OnUpdateItemClick(object? sender, EventArgs e)
    {
        var (state, _, _) = _updateChecker.GetState();
        if (state == UpdateChecker.UpdateState.UpdateAvailable)
        {
            _updateItem.Enabled = false;
            _updateItem.Text = "Downloading...";
            bool ok = await _updateChecker.DownloadAndInstall();
            if (!ok) RefreshUpdateItem();
        }
        else if (state != UpdateChecker.UpdateState.Checking && state != UpdateChecker.UpdateState.Downloading)
        {
            _updateChecker.ForceCheck();
            RefreshUpdateItem();
            _updateChecker.CheckInBackground(() => _syncCtx.Post(_ => RefreshUpdateItem(), null));
        }
    }

    private void UpdateStatus()
    {
        var (url, time) = _service.GetLastRedirect();
        if (url == null)
        {
            _statusItem.Text = "No redirects yet";
            return;
        }

        string host;
        try { host = new Uri(url).Host; }
        catch { host = url.Length > 50 ? url[..47] + "..." : url; }

        var ago = DateTime.Now - time;
        string timeStr = ago.TotalSeconds < 60 ? "just now"
            : ago.TotalMinutes < 60 ? $"{(int)ago.TotalMinutes}m ago"
            : time.ToString("h:mm tt");

        _statusItem.Text = $"{host}  ·  {timeStr}";
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _messageWindow.ReleaseHandle();
            _trayIcon.Dispose();
            _service.Stop();
        }
        base.Dispose(disposing);
    }
}

class TrayMessageWindow : NativeWindow
{
    private static readonly uint WM_TASKBARCREATED = RegisterWindowMessage("TaskbarCreated");
    private readonly NotifyIcon _trayIcon;

    public TrayMessageWindow(NotifyIcon trayIcon)
    {
        _trayIcon = trayIcon;
        CreateHandle(new CreateParams());
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_TASKBARCREATED)
        {
            _trayIcon.Visible = false;
            _trayIcon.Visible = true;
        }
        base.WndProc(ref m);
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern uint RegisterWindowMessage(string lpString);
}

class UpdateChecker
{
    public enum UpdateState { Idle, Checking, UpToDate, UpdateAvailable, Downloading }

    private UpdateState _state = UpdateState.Idle;
    private Version? _latestVersion;
    private string? _downloadUrl;
    private DateTime _lastCheck = DateTime.MinValue;
    private readonly object _lock = new();

    // Build produces versions like 1.0.N.0; GitHub tags are v1.0.N — normalize to 3 parts for comparison.
    private static readonly Version CurrentVersion = NormalizeVersion(
        System.Reflection.Assembly.GetExecutingAssembly().GetName().Version ?? new Version(0, 0, 0));

    private static Version NormalizeVersion(Version v) =>
        new(v.Major, v.Minor, Math.Max(0, v.Build));

    public (UpdateState State, Version? Version, string? DownloadUrl) GetState()
    {
        lock (_lock) return (_state, _latestVersion, _downloadUrl);
    }

    public void ForceCheck()
    {
        lock (_lock)
        {
            if (_state is UpdateState.Checking or UpdateState.Downloading) return;
            _lastCheck = DateTime.MinValue;
            if (_state == UpdateState.UpToDate) _state = UpdateState.Idle;
        }
    }

    public void CheckInBackground(Action onComplete)
    {
        lock (_lock)
        {
            if (_state is UpdateState.Checking or UpdateState.Downloading) return;
            if (_state == UpdateState.UpdateAvailable) { onComplete(); return; }
            if ((DateTime.Now - _lastCheck).TotalMinutes < 5) return;
            _state = UpdateState.Checking;
        }

        Task.Run(async () =>
        {
            try
            {
                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.ParseAdd("EdgePwaRedirector");
                client.Timeout = TimeSpan.FromSeconds(10);
                var json = await client.GetStringAsync(
                    "https://api.github.com/repos/micahmo/EdgePwaRedirector/releases/latest");

                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                var tag = root.GetProperty("tag_name").GetString() ?? "";
                var latest = Version.Parse(tag.TrimStart('v'));

                string? url = null;
                if (root.TryGetProperty("assets", out var assets))
                    foreach (var asset in assets.EnumerateArray())
                    {
                        var name = asset.GetProperty("name").GetString() ?? "";
                        if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            url = asset.GetProperty("browser_download_url").GetString();
                            break;
                        }
                    }

                lock (_lock)
                {
                    _lastCheck = DateTime.Now;
                    if (latest > CurrentVersion && url != null)
                    {
                        _state = UpdateState.UpdateAvailable;
                        _latestVersion = latest;
                        _downloadUrl = url;
                    }
                    else
                    {
                        _state = UpdateState.UpToDate;
                    }
                }
            }
            catch
            {
                lock (_lock) { _state = UpdateState.Idle; }
            }

            onComplete();
        });
    }

    public async Task<bool> DownloadAndInstall()
    {
        string? url;
        lock (_lock)
        {
            if (_state != UpdateState.UpdateAvailable) return false;
            url = _downloadUrl;
            _state = UpdateState.Downloading;
        }
        if (url == null) return false;

        try
        {
            var tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "EdgePwaRedirector-Update.exe");
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("EdgePwaRedirector");
            var bytes = await client.GetByteArrayAsync(url);
            await System.IO.File.WriteAllBytesAsync(tmp, bytes);
            Process.Start(new ProcessStartInfo(tmp) { UseShellExecute = true });
            return true;
        }
        catch
        {
            lock (_lock) { _state = UpdateState.UpdateAvailable; }
            return false;
        }
    }
}

class RedirectService
{
    private IntPtr _hook;
    private Thread? _hookThread;
    private WinEventDelegate? _delegate;
    private readonly HashSet<IntPtr> _knownWindows = new();
    private readonly object _lock = new();
    private volatile bool _paused;
    private string? _lastUrl;
    private DateTime _lastTime;

    public bool Paused { get => _paused; set => _paused = value; }

    public (string? Url, DateTime Time) GetLastRedirect()
    {
        lock (_lock) return (_lastUrl, _lastTime);
    }

    private const uint EVENT_OBJECT_SHOW = 0x8002;
    private const uint EVENT_OBJECT_DESTROY = 0x8001;
    private const uint WINEVENT_OUTOFCONTEXT = 0x0000;
    private const int GW_OWNER = 4;
    private const int OBJID_WINDOW = 0;
    private const uint WM_CLOSE = 0x0010;
    private const uint WM_KEYDOWN = 0x0100;
    private const uint WM_KEYUP = 0x0101;

    // Known Edge PWA app IDs and window title fragments.
    // Add new entries here when supporting additional PWAs.
    private static readonly string[] KnownPwaAppIds = [
        "cifhbcnohmdccbgoicgdjpfamggdegmo", // Microsoft Teams
        "faolnafnngnfdaknnbpnkhgohbobgegn", // Outlook (PWA)
    ];

    private static readonly string[] KnownPwaTitleFragments = [
        "Microsoft Teams",
        "Microsoft Outlook",
        "Outlook",
    ];

    // URLs belonging to the PWAs' core UI — close the window without redirecting.
    private static readonly string[] OwnedUrlPatterns = [
        "teams.microsoft.com",
        "outlook.office.com",
        "outlook.office365.com",
        "outlook.live.com",
        "webmail.jci.com",
        "webmail.o365.jci.com",
    ];

    // Auth/login flows that need to stay visible for the user to interact with.
    private static readonly string[] AuthUrlPatterns = [
        "microsoftonline.com",
        "microsoft.com/devicelogin",
    ];

    public void Start()
    {
        EnumWindows((hwnd, _) =>
        {
            if (IsEdgeWindow(hwnd))
                lock (_lock) _knownWindows.Add(hwnd);
            return true;
        }, IntPtr.Zero);

        _hookThread = new Thread(HookThreadProc) { IsBackground = true };
        _hookThread.SetApartmentState(ApartmentState.STA);
        _hookThread.Start();
    }

    public void Stop()
    {
        if (_hook != IntPtr.Zero)
        {
            UnhookWinEvent(_hook);
            _hook = IntPtr.Zero;
        }
    }

    private void HookThreadProc()
    {
        _delegate = WinEventCallback;
        _hook = SetWinEventHook(
            EVENT_OBJECT_DESTROY, EVENT_OBJECT_SHOW,
            IntPtr.Zero, _delegate, 0, 0, WINEVENT_OUTOFCONTEXT);
        Log(_hook != IntPtr.Zero ? "hook-ok" : "hook-FAILED");

        var msg = new MSG();
        while (GetMessage(ref msg, IntPtr.Zero, 0, 0) > 0)
        {
            TranslateMessage(ref msg);
            DispatchMessage(ref msg);
        }
    }

    private void WinEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (idObject != OBJID_WINDOW || hwnd == IntPtr.Zero) return;
        if (!IsEdgeWindow(hwnd)) return;

        if (eventType == EVENT_OBJECT_DESTROY)
        {
            lock (_lock) _knownWindows.Remove(hwnd);
            return;
        }

        bool isNew;
        lock (_lock) isNew = _knownWindows.Add(hwnd);
        if (!isNew) return;

        bool spawned = IsSpawnedByKnownPwa(hwnd);
        Log($"new-edge hwnd={hwnd} spawned={spawned} paused={_paused}");
        if (!spawned) return;

        if (_paused) return;

        ThreadPool.QueueUserWorkItem(_ => RedirectToDefaultBrowser(hwnd));
    }

    private static bool IsEdgeWindow(IntPtr hwnd)
    {
        var sb = new StringBuilder(64);
        GetClassName(hwnd, sb, 64);
        return sb.ToString() == "Chrome_WidgetWin_1";
    }

    private bool IsSpawnedByKnownPwa(IntPtr hwnd)
    {
        // Primary: owner window belongs to a known PWA (installed PWA popup)
        var owner = GetWindow(hwnd, GW_OWNER);
        if (owner != IntPtr.Zero && IsKnownPwaWindow(owner))
            return true;

        if (IsKnownPwaWindow(hwnd)) return false; // Don't redirect the PWA window itself

        // Secondary: known PWA installed with --app-id= is running
        if (IsKnownPwaProcessRunning()) return true;

        // Tertiary: Outlook/Teams may run as a regular browser tab (no --app-id=).
        // In that case, links open new windows in the same Edge process. Check if
        // any known-PWA-titled window shares this process.
        GetWindowThreadProcessId(hwnd, out uint newPid);
        lock (_lock)
        {
            foreach (var known in _knownWindows)
            {
                if (known == hwnd) continue;
                GetWindowThreadProcessId(known, out uint knownPid);
                if (knownPid == newPid && IsKnownPwaWindow(known))
                    return true;
            }
        }

        return false;
    }

    private static bool IsKnownPwaWindow(IntPtr hwnd)
    {
        var sb = new StringBuilder(512);
        GetWindowText(hwnd, sb, 512);
        var title = sb.ToString();
        foreach (var fragment in KnownPwaTitleFragments)
            if (title.Contains(fragment)) return true;
        return false;
    }

    private static bool IsKnownPwaProcessRunning()
    {
        foreach (var p in Process.GetProcessesByName("msedge"))
        {
            try
            {
                using var searcher = new System.Management.ManagementObjectSearcher(
                    $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {p.Id}");
                foreach (System.Management.ManagementObject obj in searcher.Get())
                {
                    var cmd = obj["CommandLine"]?.ToString() ?? "";
                    foreach (var appId in KnownPwaAppIds)
                        if (cmd.Contains(appId)) return true;
                }
            }
            catch { }
        }
        return false;
    }

    private static bool IsTransientUrl(string url) =>
        url.StartsWith("about:") ||
        url.Contains("safelinks") ||
        url.Contains("statics.teams.cdn");

    private static bool IsOwnedUrl(string url)
    {
        foreach (var pattern in OwnedUrlPatterns)
            if (url.Contains(pattern)) return true;
        return false;
    }

    private static bool IsAuthUrl(string url)
    {
        foreach (var pattern in AuthUrlPatterns)
            if (url.Contains(pattern)) return true;
        return false;
    }

    private static void Log(string msg)
    {
        try
        {
            System.IO.Directory.CreateDirectory(Program.LogDir);
            System.IO.File.AppendAllText(
                System.IO.Path.Combine(Program.LogDir, "redirect.log"),
                $"[{DateTime.Now:O}] {msg}\n");
        }
        catch { }
    }

    private void RedirectToDefaultBrowser(IntPtr hwnd)
    {
        if (_paused) return;
        Log($"redirect-start hwnd={hwnd}");

        string? url = null;

        // Poll for the URL with short initial intervals, backing off to 300ms.
        // Trying at 0ms first means common-case redirects are near-instant rather
        // than always waiting a fixed 300ms before the first read attempt.
        for (int i = 0; i < 20; i++)
        {
            int sleepMs = i == 0 ? 0 : i < 6 ? 100 : i < 11 ? 200 : 300;
            if (sleepMs > 0) Thread.Sleep(sleepMs);
            if (!IsWindow(hwnd)) return;
            var candidate = ReadAddressBarUrl(hwnd);
            if (!string.IsNullOrEmpty(candidate) && !IsTransientUrl(candidate))
            {
                url = candidate;
                break;
            }
        }

        if (!IsWindow(hwnd)) return;

        Log($"redirect-url='{url ?? "null"}'");
        if (url != null && IsAuthUrl(url)) return;

        if (!string.IsNullOrEmpty(url) && !IsOwnedUrl(url))
        {
            Log($"redirect url='{url}'");
            lock (_lock) { _lastUrl = url; _lastTime = DateTime.Now; }
            Process.Start(new ProcessStartInfo(FindFirefox(), url) { UseShellExecute = false });
        }

        // Popup windows (GW_OWNER set) are new single-tab windows — our fast URL detection
        // runs within ~100ms of open, before any JavaScript has had time to register a
        // beforeunload handler, so WM_CLOSE is clean and near-instant. If it somehow
        // doesn't close (rare: page loaded very fast and has beforeunload), fall back to
        // NavigateToBlank. Regular windows (existing Edge session, tertiary detection)
        // need NavigateToBlank first to safely clear any beforeunload handlers.
        if (GetWindow(hwnd, GW_OWNER) != IntPtr.Zero)
        {
            PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            for (int i = 0; i < 6 && IsWindow(hwnd); i++) Thread.Sleep(50);
            if (IsWindow(hwnd))
            {
                NavigateToBlank(hwnd);
                Thread.Sleep(200);
                if (IsWindow(hwnd)) PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }
        else
        {
            NavigateToBlank(hwnd);
            Thread.Sleep(200);
            if (IsWindow(hwnd))
            {
                bool closed = CloseTabViaUia(hwnd);
                if (!closed)
                {
                    Log("CloseTabViaUia failed, falling back to WM_CLOSE");
                    PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                }
            }
        }
    }

    private static void NavigateToBlank(IntPtr hwnd)
    {
        try
        {
            var root = AutomationElement.FromHandle(hwnd);
            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.NameProperty, "Address and search bar")
            );
            var bar = root.FindFirst(TreeScope.Descendants, condition);
            if (bar?.GetCurrentPattern(ValuePattern.Pattern) is ValuePattern vp)
            {
                bar.SetFocus();
                vp.SetValue("about:blank");
                // NativeWindowHandle is 0 for Chrome address bar (no real Win32 HWND).
                // Send Enter to the top-level window instead — it dispatches to focused control.
                PostMessage(hwnd, WM_KEYDOWN, (IntPtr)0x0D, IntPtr.Zero);
                PostMessage(hwnd, WM_KEYUP, (IntPtr)0x0D, IntPtr.Zero);
            }
        }
        catch { }
    }

    private static bool CloseTabViaUia(IntPtr hwnd)
    {
        try
        {
            var root = AutomationElement.FromHandle(hwnd);
            // Chrome/Edge tab close button — name varies by version/locale, try both
            foreach (var name in new[] { "Close tab", "Close Tab" })
            {
                var btn = root.FindFirst(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.NameProperty, name)
                ));
                if (btn?.GetCurrentPattern(InvokePattern.Pattern) is InvokePattern ip)
                {
                    ip.Invoke();
                    return true;
                }
            }
        }
        catch { }
        return false;
    }

    private static string FindFirefox()
    {
        string[] candidates = [
            @"C:\Program Files\Mozilla Firefox\firefox.exe",
            @"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
        ];
        foreach (var path in candidates)
            if (System.IO.File.Exists(path)) return path;
        return "firefox";
    }

    private static string? ReadAddressBarUrl(IntPtr hwnd)
    {
        try
        {
            var root = AutomationElement.FromHandle(hwnd);
            var condition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                new PropertyCondition(AutomationElement.NameProperty, "Address and search bar")
            );
            var bar = root.FindFirst(TreeScope.Descendants, condition);
            if (bar?.GetCurrentPattern(ValuePattern.Pattern) is ValuePattern vp)
                return vp.Current.Value;
        }
        catch { }
        return null;
    }

    private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [DllImport("user32.dll")]
    static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc,
        WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);
    [DllImport("user32.dll")] static extern bool UnhookWinEvent(IntPtr hWinEventHook);
    [DllImport("user32.dll")] static extern int GetMessage(ref MSG lpMsg, IntPtr hWnd, uint min, uint max);
    [DllImport("user32.dll")] static extern bool TranslateMessage(ref MSG lpMsg);
    [DllImport("user32.dll")] static extern IntPtr DispatchMessage(ref MSG lpMsg);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll", CharSet = CharSet.Auto)] static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public int x, y;
    }
}
