using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace EdgePwaRedirector;

static class Program
{
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
            var log = System.IO.Path.Combine(AppContext.BaseDirectory, "crash.log");
            System.IO.File.WriteAllText(log, ex.ToString());
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

        using var pngMs = new System.IO.MemoryStream();
        bmp.Save(pngMs, System.Drawing.Imaging.ImageFormat.Png);
        bmp.Dispose();
        var png = pngMs.ToArray();

        using var icoMs = new System.IO.MemoryStream();
        using var w = new System.IO.BinaryWriter(icoMs);
        w.Write((short)0);
        w.Write((short)1);
        w.Write((short)1);
        w.Write((byte)32);
        w.Write((byte)32);
        w.Write((byte)0);
        w.Write((byte)0);
        w.Write((short)1);
        w.Write((short)32);
        w.Write(png.Length);
        w.Write(6 + 16);
        w.Write(png);

        icoMs.Position = 0;
        return new Icon(icoMs);
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
        var versionLabel = new ToolStripMenuItem($"v{version?.ToString(3) ?? "?"}") { Enabled = false };

        var menu = new ContextMenuStrip();
        menu.Items.Add(versionLabel);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("Exit", null, (_, _) =>
        {
            _service.Stop();
            _trayIcon.Visible = false;
            Application.Exit();
        });
        return menu;
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

class RedirectService
{
    private IntPtr _hook;
    private Thread? _hookThread;
    private WinEventDelegate? _delegate;
    private readonly HashSet<IntPtr> _knownWindows = new();
    private readonly object _lock = new();

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

        if (!IsSpawnedByKnownPwa(hwnd)) return;

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

    private static void Log(string msg) =>
        System.IO.File.AppendAllText(
            System.IO.Path.Combine(AppContext.BaseDirectory, "redirect.log"),
            $"[{DateTime.Now:O}] {msg}\n");

    private static void RedirectToDefaultBrowser(IntPtr hwnd)
    {
        string? url = null;

        for (int i = 0; i < 20; i++)
        {
            Thread.Sleep(300);
            if (!IsWindow(hwnd)) { Log($"hwnd={hwnd} gone at poll[{i}]"); return; }
            var candidate = ReadAddressBarUrl(hwnd);
            Log($"hwnd={hwnd} poll[{i}] url='{candidate}'");
            if (!string.IsNullOrEmpty(candidate) && !IsTransientUrl(candidate))
            {
                url = candidate;
                break;
            }
        }

        if (!IsWindow(hwnd)) { Log($"hwnd={hwnd} gone after poll"); return; }

        Log($"hwnd={hwnd} final url='{url}' isAuth={url != null && IsAuthUrl(url)} isOwned={url != null && IsOwnedUrl(url)}");

        if (url != null && IsAuthUrl(url)) { Log("returning: auth url"); return; }

        if (!string.IsNullOrEmpty(url) && !IsOwnedUrl(url))
        {
            var firefox = FindFirefox();
            Log($"opening in firefox: {url}");
            Process.Start(new ProcessStartInfo(firefox, url) { UseShellExecute = false });
        }

        GetWindowThreadProcessId(hwnd, out uint closePid);
        Log($"hwnd={hwnd} pid={closePid} parent={GetParent(hwnd)} navigating to blank");
        NavigateToBlank(hwnd, out bool navOk);
        Log($"hwnd={hwnd} navOk={navOk} isWindow={IsWindow(hwnd)}");
        Thread.Sleep(500);

        bool stillThere = IsWindow(hwnd);
        Log($"hwnd={hwnd} stillThere={stillThere} closing tab");
        if (stillThere)
        {
            bool tabClosed = CloseTabViaUia(hwnd);
            Log($"hwnd={hwnd} tabClosed={tabClosed}");
            if (!tabClosed)
                PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        Thread.Sleep(500);
        Log($"hwnd={hwnd} afterClose isWindow={IsWindow(hwnd)}");
    }

    private static void NavigateToBlank(IntPtr hwnd, out bool ok)
    {
        ok = false;
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
                ok = true;
            }
        }
        catch (Exception ex) { Log($"NavigateToBlank ex: {ex.GetType().Name}: {ex.Message}"); }
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
        catch (Exception ex) { Log($"CloseTabViaUia ex: {ex.GetType().Name}: {ex.Message}"); }
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
    [DllImport("user32.dll")] static extern IntPtr GetParent(IntPtr hWnd);

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
