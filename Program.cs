using System;
using System.Collections.Generic;
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
    [STAThread]
    static void Main()
    {
        try
        {
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

    private static bool IsSpawnedByKnownPwa(IntPtr hwnd)
    {
        // Primary: check if the native owner window belongs to a known PWA
        var owner = GetWindow(hwnd, GW_OWNER);
        if (owner != IntPtr.Zero && IsKnownPwaWindow(owner))
            return true;

        // Fallback: check if a known PWA process is running and this isn't the PWA itself
        if (!IsKnownPwaWindow(hwnd) && IsKnownPwaProcessRunning())
            return true;

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

    private static void RedirectToDefaultBrowser(IntPtr hwnd)
    {
        string? url = null;

        for (int i = 0; i < 20; i++)
        {
            Thread.Sleep(300);
            if (!IsWindow(hwnd)) return;
            var candidate = ReadAddressBarUrl(hwnd);
            if (!string.IsNullOrEmpty(candidate) && !IsTransientUrl(candidate))
            {
                url = candidate;
                break;
            }
        }

        if (!IsWindow(hwnd)) return;

        // Auth popups (sign-in flows) must stay open for the user to interact with.
        if (url != null && IsAuthUrl(url)) return;

        // External URL — open in Firefox then close the Edge window.
        if (!string.IsNullOrEmpty(url) && !IsOwnedUrl(url))
        {
            var firefox = FindFirefox();
            Process.Start(new ProcessStartInfo(firefox, url) { UseShellExecute = false });
        }

        // Navigate to about:blank before closing so Edge's session restore doesn't
        // save and re-open the redirect URL on the next link click.
        NavigateToBlank(hwnd);
        Thread.Sleep(300);

        if (IsWindow(hwnd))
            PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
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
                vp.SetValue("about:blank");
                var barHwnd = new IntPtr(bar.Current.NativeWindowHandle);
                if (barHwnd != IntPtr.Zero)
                {
                    PostMessage(barHwnd, WM_KEYDOWN, (IntPtr)0x0D, IntPtr.Zero);
                    PostMessage(barHwnd, WM_KEYUP, (IntPtr)0x0D, IntPtr.Zero);
                }
            }
        }
        catch { }
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
