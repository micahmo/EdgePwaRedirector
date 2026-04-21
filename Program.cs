using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace TeamsLinkRedirector;

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
            MessageBox.Show(ex.ToString(), "TeamsLinkRedirector crashed");
        }
    }
}

class TrayApp : ApplicationContext
{
    private readonly NotifyIcon _trayIcon;
    private readonly RedirectService _service;

    public TrayApp()
    {
        _service = new RedirectService();

        _trayIcon = new NotifyIcon
        {
            Icon = CreateTrayIcon(),
            Text = "Teams Link Redirector",
            Visible = true,
            ContextMenuStrip = BuildContextMenu()
        };

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

        // Encode bitmap as PNG then wrap in an ICO stream
        using var pngMs = new System.IO.MemoryStream();
        bmp.Save(pngMs, System.Drawing.Imaging.ImageFormat.Png);
        bmp.Dispose();
        var png = pngMs.ToArray();

        using var icoMs = new System.IO.MemoryStream();
        using var w = new System.IO.BinaryWriter(icoMs);
        w.Write((short)0);       // reserved
        w.Write((short)1);       // type: icon
        w.Write((short)1);       // image count
        w.Write((byte)32);       // width
        w.Write((byte)32);       // height
        w.Write((byte)0);        // colour count
        w.Write((byte)0);        // reserved
        w.Write((short)1);       // planes
        w.Write((short)32);      // bit depth
        w.Write(png.Length);     // image data size
        w.Write(6 + 16);         // image data offset
        w.Write(png);

        icoMs.Position = 0;
        return new Icon(icoMs);
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();
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
            _trayIcon.Dispose();
            _service.Stop();
        }
        base.Dispose(disposing);
    }
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

    // Teams PWA app ID as installed in Edge Profile 1
    private const string TeamsAppId = "cifhbcnohmdccbgoicgdjpfamggdegmo";

    public void Start()
    {
        // Snapshot existing Edge windows so we only act on new ones
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

        // EVENT_OBJECT_SHOW — check if this is a new window
        bool isNew;
        lock (_lock) isNew = _knownWindows.Add(hwnd);
        if (!isNew) return;

        if (!IsSpawnedByTeams(hwnd)) return;

        ThreadPool.QueueUserWorkItem(_ => RedirectToFirefox(hwnd));
    }

    private static bool IsEdgeWindow(IntPtr hwnd)
    {
        var sb = new StringBuilder(64);
        GetClassName(hwnd, sb, 64);
        return sb.ToString() == "Chrome_WidgetWin_1";
    }

    private static bool IsSpawnedByTeams(IntPtr hwnd)
    {
        // Primary: check if the native owner window is Teams
        var owner = GetWindow(hwnd, GW_OWNER);
        if (owner != IntPtr.Zero && IsTeamsWindow(owner))
            return true;

        // Fallback: check if any Edge process is running the Teams PWA app
        // and this window is NOT the Teams window itself
        if (!IsTeamsWindow(hwnd) && IsTeamsPwaProcessRunning())
            return true;

        return false;
    }

    private static bool IsTeamsWindow(IntPtr hwnd)
    {
        var sb = new StringBuilder(512);
        GetWindowText(hwnd, sb, 512);
        var title = sb.ToString();
        return title.Contains("Microsoft Teams") || title.StartsWith("Teams");
    }

    private static bool IsTeamsPwaProcessRunning()
    {
        foreach (var p in Process.GetProcessesByName("msedge"))
        {
            try
            {
                // Check if any msedge process has the Teams app ID in its command line
                using var searcher = new System.Management.ManagementObjectSearcher(
                    $"SELECT CommandLine FROM Win32_Process WHERE ProcessId = {p.Id}");
                foreach (System.Management.ManagementObject obj in searcher.Get())
                {
                    var cmd = obj["CommandLine"]?.ToString() ?? "";
                    if (cmd.Contains(TeamsAppId) || cmd.Contains("teams.microsoft.com"))
                        return true;
                }
            }
            catch { }
        }
        return false;
    }

    private static bool IsTransientUrl(string url) =>
        url.StartsWith("about:") ||
        url.Contains("safelinks") ||           // ATP Safe Links intermediate page
        url.Contains("statics.teams.cdn");     // Teams CDN safelinks wrapper

    private static void RedirectToFirefox(IntPtr hwnd)
    {
        string? url = null;

        // Poll up to 6s, waiting past about:blank AND ATP safelinks redirects
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

        if (string.IsNullOrEmpty(url)) return;

        // Don't redirect Teams itself or auth flows
        if (url.Contains("teams.microsoft.com") ||
            url.Contains("microsoftonline.com") ||
            url.Contains("microsoft.com/devicelogin"))
            return;

        // Open in Firefox directly to avoid external-launch safety warnings
        var firefox = FindFirefox();
        Process.Start(new ProcessStartInfo(firefox, url) { UseShellExecute = false });

        // Close the Edge window
        if (IsWindow(hwnd))
            PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
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

    // P/Invoke

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
