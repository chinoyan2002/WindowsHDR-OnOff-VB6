using System.Diagnostics;
using System.Globalization;
using System.Management;
using System.Text;
using System.Text.Json;

namespace AutoHdrMode;

// One transition (power off / power on): wait, then run a shell command.
// Shell runs via cmd.exe with the program directory as working directory,
// so relative tool names (e.g. HDRCmd.exe) resolve next to the exe.
// If elevation is needed, point Shell at a self-elevating wrapper script.
internal sealed record Transition
{
    public int DelaySeconds { get; set; } = 5;
    public string Shell { get; set; } = "";
}

internal sealed record AppConfig
{
    public string Language { get; set; } = "auto";
    public int PollSeconds { get; set; } = 4;
    public int StablePolls { get; set; } = 3;
    public bool AutoEnabled { get; set; } = true;
    public string LogDir { get; set; } = "";
    public Transition PowerOff { get; set; } = new();
    public Transition PowerOn { get; set; } = new();

    public static AppConfig Load(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                var cfg = JsonSerializer.Deserialize<AppConfig>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (cfg is not null) return cfg;
            }
        }
        catch { /* fall through to defaults */ }
        return new AppConfig();
    }

    public AppConfig Resolved(string exeDir) => this with
    {
        LogDir = string.IsNullOrWhiteSpace(LogDir) ? exeDir : Resolve(exeDir, LogDir),
    };

    private static string Resolve(string exeDir, string p) =>
        Path.IsPathRooted(p) ? p : Path.Combine(exeDir, p);
}

internal static class Strings
{
    public static bool Chinese { get; set; }
    private static string T(string en, string zh) => Chinese ? zh : en;

    public static string MenuAuto => T("Auto switch", "自動切換");
    public static string MenuFireOff => T("Run power-off command now", "立即執行斷電命令");
    public static string MenuFireOn => T("Run power-on command now", "立即執行通電命令");
    public static string MenuOpenLog => T("Open log folder", "開啟紀錄資料夾");
    public static string MenuExit => T("Exit", "結束程式");

    public static string TipStarting => T("AutoHdrMode starting", "AutoHdrMode 啟動中");
    public static string TipUnknown => T("AutoHdrMode: monitor state unknown", "AutoHdrMode：螢幕狀態未知");
    public static string TipState(bool on) => T($"AutoHdrMode: physical {(on ? "ON" : "OFF")}", $"AutoHdrMode：實體螢幕{(on ? "開" : "關")}");

    public static string LogStart(int poll, int stable, bool auto) => T($"START poll={poll}s stable={stable} auto={auto}", $"啟動 輪詢{poll}秒 去拖{stable}次 自動切換={(auto ? "開" : "關")}");
    public static string LogAutoToggle(bool on) => T("AUTO " + (on ? "on" : "off"), "自動切換 " + (on ? "開" : "關"));
    public static string LogManualFireOff(int rc) => T($"MANUAL fire power-off rc={rc}", $"手動執行斷電命令 rc={rc}");
    public static string LogManualFireOn(int rc) => T($"MANUAL fire power-on rc={rc}", $"手動執行通電命令 rc={rc}");
    public static string LogInit(bool on, int wmi) => T($"INIT physical={(on ? "ON" : "OFF")} wmi={wmi}", $"初始狀態 實體螢幕={(on ? "開" : "關")} WMI筆數={wmi}");
    public static string LogEvent(bool on, int wmi) => T($"EVENT physical={(on ? "ON" : "OFF")} wmi={wmi}", $"事件 實體螢幕={(on ? "開" : "關")} WMI筆數={wmi}");
    public static string LogPlanOff(int delay) => T($"AUTO power-off shell in {delay}s", $"斷電命令排程{delay}秒後執行");
    public static string LogPlanOn(int delay) => T($"AUTO power-on shell in {delay}s", $"通電命令排程{delay}秒後執行");
    public static string LogFiredOff(int rc, string output) => T($"AUTO power-off done rc={rc} out={output}", $"斷電命令完成 rc={rc} 輸出={output}");
    public static string LogFiredOn(int rc, string output) => T($"AUTO power-on done rc={rc} out={output}", $"通電命令完成 rc={rc} 輸出={output}");
    public static string LogSkipOff => T("AUTO state changed during wait, skip power-off", "排程期間狀態已變回，跳過斷電命令");
    public static string LogSkipOn => T("AUTO state changed during wait, skip power-on", "排程期間狀態已變回，跳過通電命令");
    public static string LogEmptyOff => T("AUTO power-off shell not set, skip", "斷電命令未設定，跳過");
    public static string LogEmptyOn => T("AUTO power-on shell not set, skip", "通電命令未設定，跳過");
    public static string LogAutoErr(string m) => T("AUTO-ERR " + m, "自動：例外 " + m);
    public static string LogWmiErr(string m) => T("WMI-ERR " + m, "WMI查詢失敗 " + m);
    public static string LogPollErr(string m) => T("POLL-ERR " + m, "輪詢例外 " + m);
    public static string LogShellErr(string m) => T("SHELL-ERR " + m, "命令執行失敗 " + m);
    public static string LogOpenLogErr(string m) => T("OPENLOG-ERR " + m, "開啟紀錄資料夾失敗 " + m);
    public static string LogProbe(int wmi, string? err) => T($"PROBE wmi={wmi}" + (err is null ? "" : " err=" + err), $"探測 WMI筆數={wmi}" + (err is null ? "" : " 錯誤=" + err));
    public static string LogAutostartInstalled => T("AUTOSTART installed", "開機自啟 已安裝");
    public static string LogAutostartRemoved => T("AUTOSTART removed", "開機自啟 已移除");
    public static string LogAutostartErr(string m) => T("AUTOSTART-ERR " + m, "開機自啟失敗 " + m);
    public static string LogExit => T("EXIT", "結束");
    public static string MarkNoOutput => T("(no output)", "(無輸出)");
    public static string MarkTimeout => T("timeout", "(逾時)");
    public static string MarkStartNull => T("start-null", "(啟動失敗)");
    public static string MarkEx(string m) => T("ex: " + m, "(例外: " + m + ")");
}

internal static class Logger
{
    private static string _file = Path.Combine(Path.GetTempPath(), "AutoHdrMode.log");
    public static void Init(string dir)
    {
        try
        {
            Directory.CreateDirectory(dir);
            _file = Path.Combine(dir, "autohdrmode.log");
        }
        catch { /* keep temp fallback */ }
    }
    public static void Write(string msg)
    {
        try { File.AppendAllText(_file, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {msg}{Environment.NewLine}"); }
        catch { /* ignore */ }
    }
    public static string FilePath => _file;
}

internal static class Util
{
    public static string Trunc(string s, int max = 200)
    {
        if (s is null) return "";
        s = s.Replace('\r', ' ').Replace('\n', ' ');
        return s.Length > max ? s.Substring(0, max) : s;
    }
}

internal static class WmiProbe
{
    public sealed record Result(int Count, string? Error);

    // Returns active physical monitor count, or Count=-1 with Error text.
    // Virtual displays (remote software virtual screens) do not appear here.
    // NOTE (measured 2026-09-19): with the physical monitor powered off,
    // this query throws ("not supported") instead of returning zero rows.
    public static Result GetActivePhysicalCount()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                @"root\WMI", "SELECT Active FROM WmiMonitorBasicDisplayParams");
            int count = 0;
            foreach (ManagementObject mo in searcher.Get())
            {
                try
                {
                    var active = mo["Active"];
                    if (active is bool b ? b : Convert.ToBoolean(active)) count++;
                }
                catch { /* skip one bad row */ }
            }
            return new Result(count, null);
        }
        catch (Exception ex)
        {
            return new Result(-1, $"{ex.GetType().Name} {ex.Message}");
        }
    }
}

internal static class ShellRunner
{
    // Runs an arbitrary shell command hidden (via cmd.exe, workDir as cwd).
    // Returns process exit code (-1 = failed to run / timeout) plus the
    // first line of combined stdout+stderr for the log.
    public static (int Code, string Output) Run(string shell, string workDir, int timeoutMs = 120000)
    {
        try
        {
            var psi = new ProcessStartInfo("cmd.exe", "/s /c \"" + shell + "\"")
            {
                WorkingDirectory = workDir,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
            };
            using var p = Process.Start(psi);
            if (p is null) return (-1, Strings.MarkStartNull);
            var stdout = new StringBuilder();
            var stderr = new StringBuilder();
            p.OutputDataReceived += (_, e) => { if (e.Data is not null) stdout.AppendLine(e.Data); };
            p.ErrorDataReceived += (_, e) => { if (e.Data is not null) stderr.AppendLine(e.Data); };
            p.BeginOutputReadLine();
            p.BeginErrorReadLine();
            if (!p.WaitForExit(timeoutMs))
            {
                try { p.Kill(); } catch { }
                return (-1, Strings.MarkTimeout);
            }
            p.WaitForExit();
            string first = "";
            foreach (var line in (stdout.ToString() + "\n" + stderr.ToString()).Split('\n'))
            {
                var s = line.Trim();
                if (s.Length > 0) { first = s; break; }
            }
            if (first.Length == 0) first = Strings.MarkNoOutput;
            return (p.ExitCode, Util.Trunc(first, 160));
        }
        catch (Exception ex)
        {
            return (-1, Strings.MarkEx(ex.Message));
        }
    }
}

internal sealed class TrayContext : ApplicationContext
{
    private readonly AppConfig _cfg;
    private readonly string _workDir;
    private readonly NotifyIcon _icon;
    private readonly System.Windows.Forms.Timer _timer;
    private readonly ToolStripMenuItem _autoItem;
    private readonly object _stateLock = new();

    private bool? _lastSeen;     // last stable physical present? (true/false)
    private int _stableCount;    // consecutive polls agreeing with candidate
    private bool? _candidate;    // current candidate state
    private string? _lastErr;    // last logged WMI error (throttle repeats)

    public TrayContext(AppConfig cfg, string workDir)
    {
        _cfg = cfg;
        _workDir = workDir;

        _icon = new NotifyIcon
        {
            Icon = SystemIcons.Application,
            Text = Strings.TipStarting,
            Visible = true,
        };

        var menu = new ContextMenuStrip();
        _autoItem = new ToolStripMenuItem(Strings.MenuAuto, null, (_, _) => ToggleAuto())
        {
            Checked = _cfg.AutoEnabled,
            CheckOnClick = true,
        };
        menu.Items.Add(_autoItem);
        menu.Items.Add(new ToolStripMenuItem(Strings.MenuFireOff, null, (_, _) => ManualFire(false)));
        menu.Items.Add(new ToolStripMenuItem(Strings.MenuFireOn, null, (_, _) => ManualFire(true)));
        menu.Items.Add(new ToolStripMenuItem(Strings.MenuOpenLog, null, (_, _) => OpenLog()));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem(Strings.MenuExit, null, (_, _) => Quit()));
        _icon.ContextMenuStrip = menu;

        _timer = new System.Windows.Forms.Timer
        {
            Interval = Math.Max(1, _cfg.PollSeconds) * 1000,
        };
        _timer.Tick += (_, _) => OnPoll();
        _timer.Start();

        Logger.Write(Strings.LogStart(_cfg.PollSeconds, _cfg.StablePolls, _cfg.AutoEnabled));
        OnPoll();
    }

    private void ToggleAuto()
    {
        _cfg.AutoEnabled = _autoItem.Checked;
        Logger.Write(Strings.LogAutoToggle(_cfg.AutoEnabled));
    }

    private void ManualFire(bool isOn)
    {
        var t = isOn ? _cfg.PowerOn : _cfg.PowerOff;
        Task.Run(() =>
        {
            try
            {
                if (string.IsNullOrWhiteSpace(t.Shell))
                {
                    Logger.Write(isOn ? Strings.LogEmptyOn : Strings.LogEmptyOff);
                    return;
                }
                var (rc, output) = ShellRunner.Run(t.Shell, _workDir);
                Logger.Write(isOn ? Strings.LogManualFireOn(rc) : Strings.LogManualFireOff(rc));
            }
            catch (Exception ex) { Logger.Write(Strings.LogAutoErr(ex.Message)); }
        });
    }

    private void OpenLog()
    {
        try
        {
            Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{Logger.FilePath}\"")
            {
                UseShellExecute = true,
            });
        }
        catch (Exception ex) { Logger.Write(Strings.LogOpenLogErr(ex.Message)); }
    }

    private void Quit()
    {
        _timer.Stop();
        _icon.Visible = false;
        _icon.Dispose();
        Logger.Write(Strings.LogExit);
        Application.Exit();
    }

    private void OnPoll()
    {
        try
        {
            var (count, err) = WmiProbe.GetActivePhysicalCount();
            if (err is not null)
            {
                if (err != _lastErr) { Logger.Write(Strings.LogWmiErr(Util.Trunc(err))); _lastErr = err; }
            }
            else _lastErr = null;

            // count > 0   => physical ON.
            // count == 0, or a persistent WMI error (measured OFF signature),
            //             => physical OFF candidate. Debounce below filters blips.
            bool present = count > 0;

            if (_candidate != present) { _candidate = present; _stableCount = 1; }
            else _stableCount++;

            SetTip(Strings.TipState(present));

            if (_stableCount < _cfg.StablePolls) return;
            bool changed = false;
            lock (_stateLock)
            {
                if (_lastSeen is null)
                {
                    _lastSeen = present;
                    Logger.Write(Strings.LogInit(present, count));
                    return;
                }
                if (_lastSeen != present) { _lastSeen = present; changed = true; }
            }
            if (!changed) return;
            Logger.Write(Strings.LogEvent(present, count));

            if (!_cfg.AutoEnabled) return;
            var t = present ? _cfg.PowerOn : _cfg.PowerOff;
            Task.Run(() => RunTransition(present, t));
        }
        catch (Exception ex)
        {
            Logger.Write(Strings.LogPollErr(ex.Message));
        }
    }

    private void RunTransition(bool isOn, Transition t)
    {
        try
        {
            if (t.DelaySeconds > 0) Thread.Sleep(t.DelaySeconds * 1000);
            bool current;
            lock (_stateLock) { current = _lastSeen ?? !isOn; }
            if (current != isOn)
            {
                Logger.Write(isOn ? Strings.LogSkipOn : Strings.LogSkipOff);
                return;
            }
            if (string.IsNullOrWhiteSpace(t.Shell))
            {
                Logger.Write(isOn ? Strings.LogEmptyOn : Strings.LogEmptyOff);
                return;
            }
            Logger.Write(isOn ? Strings.LogPlanOn(t.DelaySeconds) : Strings.LogPlanOff(t.DelaySeconds));
            var (rc, output) = ShellRunner.Run(t.Shell, _workDir);
            Logger.Write(isOn ? Strings.LogFiredOn(rc, output) : Strings.LogFiredOff(rc, output));
        }
        catch (Exception ex) { Logger.Write(Strings.LogAutoErr(ex.Message)); }
    }

    private void SetTip(string text)
    {
        try
        {
            _icon.Text = text.Length > 63 ? text.Substring(0, 63) : text;
        }
        catch { /* tooltip length limits */ }
    }
}

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        string exeDir = AppContext.BaseDirectory;
        var cfg = AppConfig.Load(Path.Combine(exeDir, "config.json")).Resolved(exeDir);
        Logger.Init(cfg.LogDir);
        Strings.Chinese = cfg.Language?.Trim().ToLowerInvariant() switch
        {
            "zh" or "zh-tw" or "zh-cn" or "zh-hk" or "chs" or "cht" or "chinese" => true,
            "en" or "english" => false,
            _ => CultureInfo.CurrentUICulture.Name.StartsWith("zh", StringComparison.OrdinalIgnoreCase),
        };

        var argset = new HashSet<string>(args, StringComparer.OrdinalIgnoreCase);

        if (argset.Contains("--install-autostart"))
            return Autostart.Install();
        if (argset.Contains("--remove-autostart"))
            return Autostart.Remove();

        if (argset.Contains("--probe"))
        {
            var (count, err) = WmiProbe.GetActivePhysicalCount();
            Logger.Write(Strings.LogProbe(count, err));
            // Exit code: 0 = ON, 1 = OFF, 2 = unknown
            if (count > 0) return 0;
            if (count == 0) return 1;
            return 2;
        }
        if (argset.Contains("--fire-off")) return Fire(cfg, exeDir, false);
        if (argset.Contains("--fire-on")) return Fire(cfg, exeDir, true);

        using var mutex = new Mutex(true, @"Global\AutoHdrMode", out bool createdNew);
        if (!createdNew) return 0; // watcher already running in another instance
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new TrayContext(cfg, exeDir));
        return 0;
    }

    private static int Fire(AppConfig cfg, string workDir, bool isOn)
    {
        var t = isOn ? cfg.PowerOn : cfg.PowerOff;
        if (string.IsNullOrWhiteSpace(t.Shell))
        {
            Logger.Write(isOn ? Strings.LogEmptyOn : Strings.LogEmptyOff);
            return 3;
        }
        try
        {
            var (rc, output) = ShellRunner.Run(t.Shell, workDir);
            Logger.Write(isOn ? Strings.LogFiredOn(rc, output) : Strings.LogFiredOff(rc, output));
            return rc < 0 ? 4 : rc;
        }
        catch (Exception ex)
        {
            Logger.Write(Strings.LogShellErr(ex.Message));
            return 4;
        }
    }
}

internal static class Autostart
{
    private const string ValueName = "AutoHdrMode";
    private static string KeyPath => @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";

    public static int Install()
    {
        try
        {
            string exe = Path.Combine(AppContext.BaseDirectory, "AutoHdrMode.exe");
            using var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(KeyPath);
            key?.SetValue(ValueName, $"\"{exe}\"", Microsoft.Win32.RegistryValueKind.String);
            Logger.Write(Strings.LogAutostartInstalled);
            return 0;
        }
        catch (Exception ex) { Logger.Write(Strings.LogAutostartErr(ex.Message)); return 1; }
    }

    public static int Remove()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(KeyPath, writable: true);
            key?.DeleteValue(ValueName, throwOnMissingValue: false);
            Logger.Write(Strings.LogAutostartRemoved);
            return 0;
        }
        catch (Exception ex) { Logger.Write(Strings.LogAutostartErr(ex.Message)); return 1; }
    }
}
