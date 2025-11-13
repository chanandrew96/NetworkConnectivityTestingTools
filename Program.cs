// Program.cs
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System.Diagnostics;
using System.Text;

namespace ConnTestLauncher;

class Program
{
    private static LauncherSettings Settings = null!;
    private static readonly string ConfigFolder = "configs";
    private static readonly string ConfigExtension = ".config.csv";

    static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);
        Settings = builder.Configuration.GetSection("Launcher").Get<LauncherSettings>() ?? new LauncherSettings();

        Console.OutputEncoding = Encoding.UTF8;
        Console.Title = Settings.ConsoleTitle;

        PrintHeader();
        Directory.CreateDirectory(ConfigFolder);

        await RunLauncherAsync();
        await Task.Delay(1);
    }

    static async Task RunLauncherAsync()
    {
        string? autoCsv = DetectConfigCsv();
        if (autoCsv != null && PromptYesNo($"Detected config: {Path.GetFileName(autoCsv)}\nRun from CSV?", true))
        {
            await RunFromCsvAsync(autoCsv);
            WaitAndExit();
            return;
        }

        PrintInfo("=== Select Test Direction ===");
        bool runOutbound = PromptYesNo("Run Outbound Testing?", true);
        bool runInbound = PromptYesNo("Run Inbound Testing?", true);

        if (!runOutbound && !runInbound)
        {
            PrintWarning("No test selected.");
            WaitAndExit();
            return;
        }

        var testConfigs = new List<TestConfig>();

        if (runOutbound)
        {
            PrintInfo("\n--- Outbound Testing (Remote Hosts) ---");
            testConfigs.AddRange(CollectOutboundConfigs());
        }

        if (runInbound)
        {
            PrintInfo("\n--- Inbound Testing (Local Machine) ---");
            if (!IsRunAsAdministrator())
            {
                PrintError("Inbound testing requires Administrator privileges.");
                PrintInfo("Please run this launcher as Administrator.");
                WaitAndExit();
                return;
            }
            testConfigs.AddRange(CollectInboundConfigs());
        }

        if (testConfigs.Count == 0)
        {
            PrintWarning("No configuration entered.");
            WaitAndExit();
            return;
        }

        if (!ConfirmAllTests(testConfigs)) return;

        if (PromptYesNo("Save configuration to CSV for reuse?", true))
        {
            string fileName = $"ConnTest_Config_{DateTime.Now:yyyyMMdd_HHmmss}{ConfigExtension}";
            string savePath = Path.Combine(ConfigFolder, fileName);
            SaveToCsv(testConfigs, savePath);
            PrintSuccess($"Saved: {savePath}");
        }

        string output = Prompt("Output folder", Settings.DefaultOutputFolder);
        string prefix = Prompt("CSV prefix", Settings.DefaultCsvPrefix);
        bool html = PromptYesNo("Generate HTML report?", Settings.DefaultGenerateHtml);

        try { Directory.CreateDirectory(output); }
        catch (Exception ex) { PrintError($"Create folder failed: {ex.Message}"); WaitAndExit(); return; }

        PrintInfo($"\nRunning {testConfigs.Count} test(s)...\n");
        int success = 0;

        for (int i = 0; i < testConfigs.Count; i++)
        {
            var cfg = testConfigs[i];
            Console.WriteLine($"[{i + 1}/{testConfigs.Count}] {cfg.Direction} → {cfg.TargetInfo}");

            bool ok = await RunPowerShellAsync(cfg, output, $"{prefix}_{cfg.Direction}_{i + 1}", html);
            if (ok) success++;
        }

        PrintSuccess($"Completed: {success}/{testConfigs.Count} succeeded.");
        PrintSuccess($"Reports: {output}");

        if (PromptYesNo("Open folder?", true))
            OpenFolder(output);

        WaitAndExit();
    }

    #region Config Collection
    static List<TestConfig> CollectOutboundConfigs()
    {
        var list = new List<TestConfig>();
        int idx = 1;

        while (true)
        {
            PrintInfo($"--- Outbound Test #{idx} ---");
            string hosts = Prompt("Remote host(s) (comma separated)", Settings.DefaultHosts);
            string tcp = Prompt("TCP ports (e.g. 80,443,3389)", Settings.DefaultTcpPorts, allowEmpty: true);
            string udp = Prompt("UDP ports (e.g. 53,137)", Settings.DefaultUdpPorts, allowEmpty: true);
            bool icmp = PromptYesNo("Test ICMP (Ping)?", Settings.DefaultTestIcmp);

            list.Add(new TestConfig
            {
                Direction = "Outbound",
                ScriptPath = Path.Combine(AppContext.BaseDirectory, Settings.OutboundScriptPath),
                Hosts = hosts,
                TcpPorts = tcp,
                UdpPorts = udp,
                Icmp = icmp,
                TargetInfo = hosts
            });

            if (!PromptYesNo("Add another Outbound test?", false)) break;
            idx++; Console.WriteLine();
        }
        return list;
    }

    static List<TestConfig> CollectInboundConfigs()
    {
        var list = new List<TestConfig>();
        int idx = 1;

        while (true)
        {
            PrintInfo($"--- Inbound Test #{idx} (Local) ---");
            string tcp = Prompt("TCP ports to check (e.g. 3389,80)", "3389", allowEmpty: false);
            string udp = Prompt("UDP ports to check (e.g. 53)", "", allowEmpty: true);
            bool icmp = PromptYesNo("Check ICMP (Ping response)?", true);

            list.Add(new TestConfig
            {
                Direction = "Inbound",
                ScriptPath = Path.Combine(AppContext.BaseDirectory, Settings.InboundScriptPath),
                TcpPorts = tcp,
                UdpPorts = udp,
                Icmp = icmp,
                TargetInfo = "Local Machine"
            });

            if (!PromptYesNo("Add another Inbound test?", false)) break;
            idx++; Console.WriteLine();
        }
        return list;
    }
    #endregion

    #region Summary & CSV
    static bool ConfirmAllTests(List<TestConfig> configs)
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(" TEST SUMMARY ");
        Console.WriteLine(" ==========================================");
        Console.ResetColor();

        foreach (var (cfg, i) in configs.Select((c, i) => (c, i)))
        {
            Console.WriteLine($"[{i + 1}] {cfg.Direction} → {cfg.TargetInfo}");
            if (cfg.Direction == "Outbound")
            {
                Console.WriteLine($"    TCP: {(string.IsNullOrEmpty(cfg.TcpPorts) ? "(none)" : cfg.TcpPorts)}");
                Console.WriteLine($"    UDP: {(string.IsNullOrEmpty(cfg.UdpPorts) ? "(none)" : cfg.UdpPorts)}");
                Console.WriteLine($"    ICMP: {(cfg.Icmp ? "Yes" : "No")}");
            }
            else
            {
                Console.WriteLine($"    TCP: {cfg.TcpPorts}");
                Console.WriteLine($"    UDP: {(string.IsNullOrEmpty(cfg.UdpPorts) ? "(none)" : cfg.UdpPorts)}");
                Console.WriteLine($"    ICMP: {(cfg.Icmp ? "Yes" : "No")}");
            }
        }
        Console.WriteLine(" ==========================================");
        return PromptYesNo("Proceed?", true);
    }
    #region CSV 安全寫入與讀取
    static void SaveToCsv(List<TestConfig> configs, string path)
    {
        var lines = new List<string>
    {
        // 使用 ; 分隔，避開端口 , 衝突
        "Direction;Hosts;TcpPorts;UdpPorts;Icmp"
    };

        foreach (var c in configs)
        {
            string tcp = EscapeCsvField(c.TcpPorts);
            string udp = EscapeCsvField(c.UdpPorts);
            string hosts = c.Direction == "Outbound" ? EscapeCsvField(c.Hosts) : "";

            lines.Add($"{c.Direction};{hosts};{tcp};{udp};{c.Icmp}");
        }

        File.WriteAllLines(path, lines, Encoding.UTF8);
        PrintSuccess($"Config saved: {path}");
    }

    static List<TestConfig> LoadFromCsv(string path)
    {
        var list = new List<TestConfig>();
        if (!File.Exists(path))
        {
            PrintError("CSV file not found.");
            return list;
        }

        var lines = File.ReadAllLines(path, Encoding.UTF8);
        if (lines.Length == 0) return list;

        // 跳過 Header
        foreach (var line in lines.Skip(1))
        {
            if (string.IsNullOrWhiteSpace(line)) continue;

            // 使用 ; 分割
            var parts = line.Split(';', 5);
            if (parts.Length < 5) continue;

            var direction = parts[0].Trim();
            if (!new[] { "Outbound", "Inbound" }.Contains(direction, StringComparer.OrdinalIgnoreCase))
                continue;

            var cfg = new TestConfig
            {
                Direction = direction,
                ScriptPath = direction == "Outbound"
                    ? Path.Combine(AppContext.BaseDirectory, Settings.OutboundScriptPath)
                    : Path.Combine(AppContext.BaseDirectory, Settings.InboundScriptPath),
                TcpPorts = UnescapeCsvField(parts[2]),
                UdpPorts = UnescapeCsvField(parts[3]),
                Icmp = bool.TryParse(parts[4], out bool b) && b,
                TargetInfo = direction == "Outbound" ? UnescapeCsvField(parts[1]) : "Local"
            };

            if (cfg.Direction == "Outbound")
                cfg.Hosts = UnescapeCsvField(parts[1]);

            list.Add(cfg);
        }

        return list;
    }

    // 安全跳脫 CSV 欄位（支援 , " ; \n）
    static string EscapeCsvField(string? input)
    {
        if (string.IsNullOrEmpty(input)) return "";
        input = input.Trim();

        bool needsQuote = input.Contains(',') || input.Contains('"') || input.Contains(';') || input.Contains('\n') || input.Contains('\r');

        if (needsQuote)
        {
            input = input.Replace("\"", "\"\""); // 跳脫雙引號
            return $"\"{input}\"";
        }
        return input;
    }

    static string UnescapeCsvField(string input)
    {
        if (string.IsNullOrEmpty(input)) return "";
        input = input.Trim();

        if (input.StartsWith("\"") && input.EndsWith("\""))
        {
            input = input.Substring(1, input.Length - 2);
            input = input.Replace("\"\"", "\"");
        }

        return input;
    }
    #endregion
    #endregion

    #region PowerShell Execution (精確匹配腳本參數)
    static async Task<bool> RunPowerShellAsync(TestConfig cfg, string output, string prefix, bool html)
    {
        if (!File.Exists(cfg.ScriptPath))
        {
            PrintError($"Script missing: {cfg.ScriptPath}");
            return false;
        }

        var args = cfg.Direction == "Outbound"
            ? BuildOutboundArgs(cfg.ScriptPath, cfg.Hosts, cfg.TcpPorts, cfg.UdpPorts, cfg.Icmp, output, prefix, html)
            : BuildInboundArgs(cfg.ScriptPath, cfg.TcpPorts, cfg.UdpPorts, cfg.Icmp, output, prefix, html);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        try
        {
            using var p = Process.Start(psi);
            if (p == null) return false;

            var outTask = p.StandardOutput.ReadToEndAsync();
            var errTask = p.StandardError.ReadToEndAsync();

            string outputLog = await outTask;
            string errorLog = await errTask;

            p.WaitForExit();

            if (!string.IsNullOrEmpty(outputLog)) Console.WriteLine(outputLog);
            if (!string.IsNullOrEmpty(errorLog))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(errorLog);
                Console.ResetColor();
            }

            return p.ExitCode == 0;
        }
        catch (Exception ex)
        {
            PrintError($"PowerShell failed: {ex.Message}");
            return false;
        }
    }

    static string BuildOutboundArgs(string script, string hosts, string tcp, string udp, bool icmp, string outPath, string prefix, bool html)
    {
        var sb = new StringBuilder();
        sb.Append($"-ExecutionPolicy Bypass -File \"{script}\"");
        sb.Append($" -ComputerName \"{hosts}\"");
        if (!string.IsNullOrEmpty(tcp)) sb.Append($" -TcpList \"{tcp}\"");
        if (!string.IsNullOrEmpty(udp)) sb.Append($" -UdpList \"{udp}\"");
        if (icmp) sb.Append(" -Icmp");
        sb.Append($" -OutputPath \"{outPath}\"");
        sb.Append($" -CsvPrefix \"{prefix}\"");
        if (html) sb.Append(" -HtmlReport");
        return sb.ToString();
    }

    static string BuildInboundArgs(string script, string tcp, string udp, bool icmp, string outPath, string prefix, bool html)
    {
        var sb = new StringBuilder();
        sb.Append($"-ExecutionPolicy Bypass -File \"{script}\"");
        sb.Append($" -TCPPorts \"{tcp}\"");
        if (!string.IsNullOrEmpty(udp)) sb.Append($" -UDPPorts \"{udp}\"");
        if (icmp) sb.Append(" -CheckIcmp");
        sb.Append($" -OutputPath \"{outPath}\"");
        sb.Append($" -CsvPrefix \"{prefix}\"");
        if (html) sb.Append(" -GenerateHtml");
        return sb.ToString();
    }
    #endregion

    #region Utils
    static bool IsRunAsAdministrator()
    {
        try
        {
            var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
        catch { return false; }
    }

    static string? DetectConfigCsv() =>
        Directory.GetFiles(ConfigFolder, $"*{ConfigExtension}").FirstOrDefault();

    static async Task RunFromCsvAsync(string csv)
    {
        PrintInfo($"Loading: {Path.GetFileName(csv)}");
        var configs = LoadFromCsv(csv);
        if (configs.Count == 0) { PrintError("No valid config."); return; }

        if (configs.Any(c => c.Direction == "Inbound") && !IsRunAsAdministrator())
        {
            PrintError("Inbound test found but not running as Admin.");
            return;
        }

        string output = Prompt("Output folder", Settings.DefaultOutputFolder);
        string prefix = Prompt("CSV prefix", Path.GetFileNameWithoutExtension(csv));
        bool html = PromptYesNo("Generate HTML?", Settings.DefaultGenerateHtml);

        try { Directory.CreateDirectory(output); }
        catch (Exception ex) { PrintError($"Folder error: {ex.Message}"); return; }

        int success = 0;
        for (int i = 0; i < configs.Count; i++)
        {
            var cfg = configs[i];
            Console.WriteLine($"[{i + 1}/{configs.Count}] {cfg.Direction} → {cfg.TargetInfo}");
            if (await RunPowerShellAsync(cfg, output, $"{prefix}_{i + 1}", html)) success++;
        }
        PrintSuccess($"Batch done: {success}/{configs.Count}");
    }

    static void PrintHeader()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine(" ==========================================");
        Console.WriteLine(" HK Enterprise – Port Test Launcher v4.0");
        Console.WriteLine(" Inbound + Outbound + CSV + Admin Check");
        Console.WriteLine(" ==========================================");
        Console.ResetColor(); Console.WriteLine();
    }

    static string Prompt(string msg, string def, bool allowEmpty = false)
    {
        while (true)
        {
            Console.Write($" {msg}");
            if (!string.IsNullOrWhiteSpace(def)) Console.Write($" [default: {def}]");
            Console.Write(": ");
            string? input = Console.ReadLine()?.Trim();
            if (string.IsNullOrEmpty(input))
            {
                if (allowEmpty) return "";
                if (!string.IsNullOrWhiteSpace(def)) return def;
                PrintWarning("Required.");
                continue;
            }
            return input;
        }
    }

    static bool PromptYesNo(string msg, bool def)
    {
        while (true)
        {
            Console.Write($" {msg} (y/n)");
            Console.Write(def ? " [y]" : " [n]");
            Console.Write(": ");
            string? input = Console.ReadLine()?.Trim().ToLower();
            if (string.IsNullOrEmpty(input)) return def;
            if (input is "y" or "yes") return true;
            if (input is "n" or "no") return false;
            PrintWarning("Enter y or n.");
        }
    }

    static void OpenFolder(string path)
    {
        try { Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true }); }
        catch { PrintWarning("Cannot open folder."); }
    }

    static void WaitAndExit()
    {
        Console.WriteLine(); Console.Write("Press any key to exit..."); Console.ReadKey(true);
    }

    static void PrintError(string m) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine(m); Console.ResetColor(); }
    static void PrintSuccess(string m) { Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine(m); Console.ResetColor(); }
    static void PrintInfo(string m) { Console.ForegroundColor = ConsoleColor.Cyan; Console.WriteLine(m); Console.ResetColor(); }
    static void PrintWarning(string m) { Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine(m); Console.ResetColor(); }
    #endregion
}

public class TestConfig
{
    public string Direction { get; set; } = "";
    public string ScriptPath { get; set; } = "";
    public string Hosts { get; set; } = "";
    public string TcpPorts { get; set; } = "";
    public string UdpPorts { get; set; } = "";
    public bool Icmp { get; set; }
    public string TargetInfo { get; set; } = "";
}

public class LauncherSettings
{
    public string ConsoleTitle { get; set; } = "HK Port Test Launcher v4.0";
    public string OutboundScriptPath { get; set; } = "scripts\\Test-ConnectionPorts.ps1";
    public string InboundScriptPath { get; set; } = "scripts\\Check-InboundConnectivity.ps1";
    public string DefaultOutputFolder { get; set; } = @"C:\ConnTest_Reports";
    public string DefaultCsvPrefix { get; set; } = "ConnTest";
    public string DefaultHosts { get; set; } = "localhost";
    public string DefaultTcpPorts { get; set; } = "80,443,3389";
    public string DefaultUdpPorts { get; set; } = "53";
    public bool DefaultTestIcmp { get; set; } = true;
    public bool DefaultGenerateHtml { get; set; } = true;
}