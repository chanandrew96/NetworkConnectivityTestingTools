using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System.Diagnostics;
using System.Globalization;
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

        // Ensure the configs folder exist
        Directory.CreateDirectory(ConfigFolder);

        await RunLauncherAsync();
        await Task.Delay(1); // Keep host alive
    }

    static async Task RunLauncherAsync()
    {
        string? autoCsv = DetectConfigCsv();
        if (autoCsv != null && PromptYesNo($"Detected config file: {Path.GetFileName(autoCsv)}\nLoad and run tests from this file?", true))
        {
            await RunFromCsvAsync(autoCsv);
            WaitAndExit();
            return;
        }

        PrintInfo("=== Select Test Direction ===");
        bool runInbound = PromptYesNo("Run Inbound Testing?", true);
        bool runOutbound = PromptYesNo("Run Outbound Testing?", true);

        if (!runInbound && !runOutbound)
        {
            PrintWarning("No test selected. Exiting.");
            WaitAndExit();
            return;
        }

        var testConfigs = new List<TestConfig>();

        if (runOutbound)
        {
            PrintInfo("\n--- Outbound Testing Configuration ---");
            testConfigs.AddRange(CollectTestConfigs("Outbound", Settings.OutboundScriptPath));
        }

        if (runInbound)
        {
            PrintInfo("\n--- Inbound Testing Configuration ---");
            testConfigs.AddRange(CollectTestConfigs("Inbound", Settings.InboundScriptPath));
        }

        if (testConfigs.Count == 0)
        {
            PrintWarning("No configuration entered.");
            WaitAndExit();
            return;
        }

        // Show summary
        if (!ConfirmAllTests(testConfigs)) return;

        // Ask if need to save the csv
        string? savePath = null;
        if (PromptYesNo("Save these configurations to CSV for future use?", true))
        {
            string fileName = $"ConnTest_Config_{DateTime.Now:yyyyMMdd_HHmmss}{ConfigExtension}";
            savePath = Path.Combine(ConfigFolder, fileName);
            SaveToCsv(testConfigs, savePath);
            PrintSuccess($"Configuration saved to: {savePath}");
        }

        // Create Output folder
        string output = Prompt("Global output folder", Settings.DefaultOutputFolder);
        string prefix = Prompt("Global CSV prefix", Settings.DefaultCsvPrefix);
        bool html = PromptYesNo("Generate HTML report?", Settings.DefaultGenerateHtml);

        try { Directory.CreateDirectory(output); }
        catch (Exception ex) { PrintError($"Failed to create output folder: {ex.Message}"); WaitAndExit(); return; }

        // Execute all test(s)
        PrintInfo($"\nStarting {testConfigs.Count} test(s)...\n");
        int successCount = 0;

        for (int i = 0; i < testConfigs.Count; i++)
        {
            var cfg = testConfigs[i];
            Console.WriteLine($"[{i + 1}/{testConfigs.Count}] Running {cfg.Direction} Test: {cfg.Hosts} → Ports: {cfg.TcpPorts}{(string.IsNullOrEmpty(cfg.UdpPorts) ? "" : $", {cfg.UdpPorts}")} {(cfg.Icmp ? "+Ping" : "")}");

            bool success = await RunPowerShellAsync(
                cfg.ScriptPath,
                cfg.Hosts,
                cfg.TcpPorts,
                cfg.UdpPorts,
                cfg.Icmp,
                output,
                $"{prefix}_{cfg.Direction}_{i + 1}",
                html
            );

            if (success) successCount++;
            Console.WriteLine();
        }

        PrintSuccess($"All tests completed! {successCount}/{testConfigs.Count} succeeded.");
        PrintSuccess($"Reports saved to: {output}");

        if (PromptYesNo("Open report folder now?", true))
            OpenFolder(output);

        WaitAndExit();
    }

    static List<TestConfig> CollectTestConfigs(string direction, string scriptPath)
    {
        var configs = new List<TestConfig>();
        int index = 1;

        while (true)
        {
            PrintInfo($"--- {direction} Test #{index} ---");
            string hosts = Prompt($"  Target host(s)", Settings.DefaultHosts);
            string tcp = Prompt($"  TCP ports", Settings.DefaultTcpPorts, allowEmpty: true);
            string udp = Prompt($"  UDP ports", Settings.DefaultUdpPorts, allowEmpty: true);
            bool icmp = PromptYesNo($"  Test Ping (ICMP)?", Settings.DefaultTestIcmp);

            configs.Add(new TestConfig
            {
                Direction = direction,
                ScriptPath = Path.Combine(AppContext.BaseDirectory, scriptPath),
                Hosts = hosts,
                TcpPorts = tcp,
                UdpPorts = udp,
                Icmp = icmp
            });

            if (!PromptYesNo($"Add another {direction} test configuration?", false))
                break;

            index++;
            Console.WriteLine();
        }

        return configs;
    }

    static bool ConfirmAllTests(List<TestConfig> configs)
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine(" ==========================================");
        Console.WriteLine(" Test Summary");
        Console.WriteLine(" ==========================================");
        Console.ResetColor();

        foreach (var (cfg, i) in configs.Select((c, i) => (c, i)))
        {
            Console.WriteLine($"[{i + 1}] {cfg.Direction}: {cfg.Hosts}");
            Console.WriteLine($"    TCP: {(string.IsNullOrEmpty(cfg.TcpPorts) ? "(none)" : cfg.TcpPorts)}");
            Console.WriteLine($"    UDP: {(string.IsNullOrEmpty(cfg.UdpPorts) ? "(none)" : cfg.UdpPorts)}");
            Console.WriteLine($"    Ping: {(cfg.Icmp ? "Yes" : "No")}");
        }
        Console.WriteLine(" ==========================================");
        Console.WriteLine();
        return PromptYesNo("Proceed with all tests?", true);
    }

    static string? DetectConfigCsv()
    {
        var files = Directory.GetFiles(ConfigFolder, $"*{ConfigExtension}");
        return files.Length > 0 ? files[0] : null;
    }

    static async Task RunFromCsvAsync(string csvPath)
    {
        PrintInfo($"Loading configurations from: {Path.GetFileName(csvPath)}");
        var configs = LoadFromCsv(csvPath);

        if (configs.Count == 0)
        {
            PrintError("No valid configuration found in CSV.");
            return;
        }

        string output = Prompt("Output folder", Settings.DefaultOutputFolder);
        string prefix = Prompt("CSV prefix", Path.GetFileNameWithoutExtension(csvPath));
        bool html = PromptYesNo("Generate HTML report?", Settings.DefaultGenerateHtml);

        try { Directory.CreateDirectory(output); }
        catch (Exception ex) { PrintError($"Failed to create folder: {ex.Message}"); return; }

        int successCount = 0;
        for (int i = 0; i < configs.Count; i++)
        {
            var cfg = configs[i];
            Console.WriteLine($"[{i + 1}/{configs.Count}] {cfg.Direction} → {cfg.Hosts}");

            bool success = await RunPowerShellAsync(
                cfg.ScriptPath,
                cfg.Hosts,
                cfg.TcpPorts,
                cfg.UdpPorts,
                cfg.Icmp,
                output,
                $"{prefix}_{i + 1}",
                html
            );
            if (success) successCount++;
        }

        PrintSuccess($"Batch completed: {successCount}/{configs.Count} succeeded.");
    }

    static void SaveToCsv(List<TestConfig> configs, string path)
    {
        var lines = new List<string>
        {
            "Direction,Hosts,TcpPorts,UdpPorts,Icmp"
        };

        foreach (var cfg in configs)
        {
            lines.Add($"{cfg.Direction},{cfg.Hosts},{cfg.TcpPorts},{cfg.UdpPorts},{cfg.Icmp}");
        }

        File.WriteAllLines(path, lines, Encoding.UTF8);
    }

    static List<TestConfig> LoadFromCsv(string path)
    {
        var configs = new List<TestConfig>();
        var lines = File.ReadAllLines(path, Encoding.UTF8).Skip(1); // skip header

        foreach (var line in lines)
        {
            var parts = line.Split(',');
            if (parts.Length < 5) continue;

            configs.Add(new TestConfig
            {
                Direction = parts[0],
                Hosts = parts[1],
                TcpPorts = parts[2],
                UdpPorts = parts[3],
                Icmp = bool.TryParse(parts[4], out bool icmp) && icmp,
                ScriptPath = parts[0].Contains("Inbound", StringComparison.OrdinalIgnoreCase)
                    ? Path.Combine(AppContext.BaseDirectory, Settings.InboundScriptPath)
                    : Path.Combine(AppContext.BaseDirectory, Settings.OutboundScriptPath)
            });
        }

        return configs;
    }

    #region PowerShell Execution
    static async Task<bool> RunPowerShellAsync(string scriptPath, string hosts, string tcp, string udp, bool icmp, string output, string prefix, bool html)
    {
        if (!File.Exists(scriptPath))
        {
            PrintError($"Script not found: {scriptPath}");
            return false;
        }

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = BuildArguments(scriptPath, hosts, tcp, udp, icmp, output, prefix, html),
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        try
        {
            using var process = Process.Start(psi);
            if (process == null) return false;

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            string outputLog = await outputTask;
            string errorLog = await errorTask;

            process.WaitForExit();

            if (!string.IsNullOrEmpty(outputLog)) Console.WriteLine(outputLog);
            if (!string.IsNullOrEmpty(errorLog))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(errorLog);
                Console.ResetColor();
            }

            return process.ExitCode == 0;
        }
        catch (Exception ex)
        {
            PrintError($"PowerShell error: {ex.Message}");
            return false;
        }
    }

    static string BuildArguments(string scriptPath, string hosts, string tcp, string udp, bool icmp, string output, string prefix, bool html)
    {
        var args = new StringBuilder();
        args.Append($"-ExecutionPolicy Bypass -File \"{scriptPath}\"");
        args.Append($" -ComputerName \"{hosts}\"");
        if (!string.IsNullOrEmpty(tcp)) args.Append($" -TcpList \"{tcp}\"");
        if (!string.IsNullOrEmpty(udp)) args.Append($" -UdpList \"{udp}\"");
        if (icmp) args.Append(" -Icmp");
        args.Append($" -OutputPath \"{output}\"");
        args.Append($" -CsvPrefix \"{prefix}\"");
        if (html) args.Append(" -HtmlReport");
        return args.ToString();
    }
    #endregion

    #region UI Helpers
    static void PrintHeader()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine(" ==========================================");
        Console.WriteLine(" HK Enterprise – Port & Firewall Test");
        Console.WriteLine("            Multi-Config Launcher v0.1");
        Console.WriteLine(" ==========================================");
        Console.ResetColor();
        Console.WriteLine();
    }

    static string Prompt(string message, string defaultValue, bool allowEmpty = false)
    {
        while (true)
        {
            Console.Write($" {message}");
            if (!string.IsNullOrWhiteSpace(defaultValue))
                Console.Write($" [default: {defaultValue}]");
            Console.Write(": ");
            string? input = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(input))
            {
                if (allowEmpty) return "";
                if (!string.IsNullOrWhiteSpace(defaultValue)) return defaultValue;
                PrintWarning("This field is required.");
                continue;
            }
            return input;
        }
    }

    static bool PromptYesNo(string message, bool defaultYes)
    {
        while (true)
        {
            Console.Write($" {message} (y/n)");
            Console.Write(defaultYes ? " [default: y]" : " [default: n]");
            Console.Write(": ");
            string? input = Console.ReadLine()?.Trim().ToLower();

            if (string.IsNullOrEmpty(input)) return defaultYes;
            if (input is "y" or "yes") return true;
            if (input is "n" or "no") return false;

            PrintWarning("Please enter 'y' or 'n'.");
        }
    }

    static void OpenFolder(string path)
    {
        try { Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true }); }
        catch { PrintWarning("Could not open folder."); }
    }

    static void WaitAndExit()
    {
        Console.WriteLine();
        Console.Write("Press any key to exit...");
        Console.ReadKey(true);
    }

    static void PrintError(string msg) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine(msg); Console.ResetColor(); }
    static void PrintSuccess(string msg) { Console.ForegroundColor = ConsoleColor.Green; Console.WriteLine(msg); Console.ResetColor(); }
    static void PrintInfo(string msg) { Console.ForegroundColor = ConsoleColor.Cyan; Console.WriteLine(msg); Console.ResetColor(); }
    static void PrintWarning(string msg) { Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine(msg); Console.ResetColor(); }
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
}

public class LauncherSettings
{
    public string ConsoleTitle { get; set; } = "HK Enterprise – Port Test Launcher v0.1";
    public string InboundScriptPath { get; set; } = "scripts\\Test-NetInboundConnection.ps1";
    public string OutboundScriptPath { get; set; } = "scripts\\Test-ConnectionPorts.ps1";
    public string DefaultOutputFolder { get; set; } = @"C:\ConnTest_Reports";
    public string DefaultCsvPrefix { get; set; } = "ConnTest";
    public string DefaultHosts { get; set; } = "localhost";
    public string DefaultTcpPorts { get; set; } = "80,443,3389";
    public string DefaultUdpPorts { get; set; } = "53";
    public bool DefaultTestIcmp { get; set; } = true;
    public bool DefaultGenerateHtml { get; set; } = true;
}