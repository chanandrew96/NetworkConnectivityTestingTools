<#
    Test-ConnectionPorts.ps1
    PowerShell 5.1 compatible
    Multi-host TCP/UDP/ICMP testing
    NEW COLUMNS: Port Opening | Firewall blocking | Service Listening
    HTML FIXED – No if/else inside $()
#>

param(
    [Parameter(Mandatory=$true)]
    [string[]]$ComputerName,

    [string[]]$TcpList,
    [string[]]$UdpList,

    [switch]$Icmp,

    [string]$OutputPath = ".",
    [string]$CsvPrefix = "ConnTest",

    [switch]$HtmlReport
)

# -------------------------------------------------
# Service-to-Port map
# -------------------------------------------------
$ServicePorts = @{
    "HTTP" = 80; "HTTPS" = 443; "FTP" = 21; "SSH" = 22
    "TELNET" = 23; "SMTP" = 25; "SMTPS" = 465; "IMAP" = 143
    "IMAPS" = 993; "POP3" = 110; "POP3S" = 995; "RDP" = 3389
    "MYSQL" = 3306; "MSSQL" = 1433; "DNS" = 53; "NTP" = 123
    "SNMP" = 161; "SYSLOG" = 514
}

# -------------------------------------------------
# Parse port list
# -------------------------------------------------
function Parse-PortList {
    param([string[]]$List)
    $ports = @()
    foreach ($item in $List) {
        $item = $item.Trim().ToUpper()
        if ($ServicePorts.ContainsKey($item)) {
            $ports += $ServicePorts[$item]
        }
        elseif ($item -match "^(\d+)-(\d+)$") {
            $s = [int]$matches[1]; $e = [int]$matches[2]
            if ($s -le $e -and $s -ge 1 -and $e -le 65535) {
                $ports += ($s..$e)
            } else { Write-Warning "Invalid range: $item" }
        }
        elseif ($item -match "^\d+$") {
            $p = [int]$item
            if ($p -ge 1 -and $p -le 65535) { $ports += $p }
            else { Write-Warning "Port out of range: $item" }
        }
        else { Write-Warning "Unrecognized: $item" }
    }
    $ports | Sort-Object -Unique
}

$tcpPorts = if ($TcpList) { Parse-PortList -List $TcpList } else { @() }
$udpPorts = if ($UdpList) { Parse-PortList -List $UdpList } else { @() }

# -------------------------------------------------
# Output folder
# -------------------------------------------------
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$timestamp       = Get-Date -Format "yyyyMMdd_HHmmss"
$summaryHtmlPath = Join-Path $OutputPath "$CsvPrefix`_Summary_Report`_$timestamp.html"

$allHostResults = @()
$summaryBody    = ""

# -------------------------------------------------
# Process each host
# -------------------------------------------------
foreach ($hostName in $ComputerName) {
    $hostName = $hostName.Trim()
    if ([string]::IsNullOrWhiteSpace($hostName)) { continue }

    Write-Host "`nTesting: $hostName" -ForegroundColor Magenta

    $results = @()

    # ---------- ICMP ----------
    if ($Icmp) {
        $pingOk = Test-Connection -ComputerName $hostName -Count 1 -Quiet -ErrorAction SilentlyContinue

        $results += [pscustomobject]@{
            Host               = $hostName
            Protocol           = "ICMP"
            Port               = "-"
            Status             = if ($pingOk) { "Success" } else { "Failed" }
            StatusDetail       = if ($pingOk) { "Port Opened" } else { "Firewall blocked" }
            'Port Opening'     = $pingOk
            'Firewall blocking'= -not $pingOk
            'Service Listening'= $false
        }
    }

    # ---------- TCP ----------
    foreach ($port in $tcpPorts) {
        $tnc    = Test-NetConnection -ComputerName $hostName -Port $port -WarningAction SilentlyContinue -InformationLevel Quiet
        $pingOk = Test-Connection -ComputerName $hostName -Count 1 -Quiet -ErrorAction SilentlyContinue

        $portOpen      = $pingOk -and $tnc
        $fwBlock       = -not $pingOk
        $svcListen     = $pingOk -and -not $tnc

        $results += [pscustomobject]@{
            Host               = $hostName
            Protocol           = "TCP"
            Port               = $port
            Status             = if ($portOpen) { "Success" } else { "Failed" }
            StatusDetail       = if ($portOpen) { "Port Opened" }
                                 elseif ($svcListen) { "Service listening" }
                                 else { "Firewall blocked" }
            'Port Opening'     = $portOpen
            'Firewall blocking'= $fwBlock
            'Service Listening'= $svcListen
        }
    }

    # ---------- UDP ----------
    foreach ($port in $udpPorts) {
        $pingOk = Test-Connection -ComputerName $hostName -Count 1 -Quiet -ErrorAction SilentlyContinue
        $udpSuccess = $false

        try {
            $udp = New-Object System.Net.Sockets.UdpClient
            $udp.Client.SendTimeout    = 1000
            $udp.Client.ReceiveTimeout = 1000
            $udp.Connect($hostName, $port)
            $null = $udp.Send([byte[]]@(0), 1)
            $udpSuccess = $true
        } catch { $udpSuccess = $false }
        finally {
            if ($udp) { try { $udp.Close() } catch {} ; $udp.Dispose() }
        }

        $portOpen  = $pingOk -and $udpSuccess
        $fwBlock   = -not $pingOk
        $svcListen = $pingOk -and -not $udpSuccess

        $results += [pscustomobject]@{
            Host               = $hostName
            Protocol           = "UDP"
            Port               = $port
            Status             = if ($portOpen) { "Success" } else { "Failed" }
            StatusDetail       = if ($portOpen) { "Port Opened" }
                                 elseif ($svcListen) { "Service listening" }
                                 else { "Firewall blocked" }
            'Port Opening'     = $portOpen
            'Firewall blocking'= $fwBlock
            'Service Listening'= $svcListen
        }
    }

    # ---------- Export ----------
    $allHost     = $results
    $successHost = $results | Where-Object { $_.Status -eq "Success" }
    $failedHost  = $results | Where-Object { $_.Status -eq "Failed" }

    $safeHost = $hostName -replace '[\\/:*?"<>|]', '_'
    $base     = "$CsvPrefix`_$safeHost`_$timestamp"

    $allFile     = Join-Path $OutputPath "$base`_all.csv"
    $successFile = Join-Path $OutputPath "$base`_success.csv"
    $failedFile  = Join-Path $OutputPath "$base`_failed.csv"
    $htmlFile    = Join-Path $OutputPath "$base`_report.html"

    # CSV
    $allHost |
        Select-Object Host, Protocol, Port, Status, StatusDetail,
                      'Port Opening', 'Firewall blocking', 'Service Listening' |
        Export-Csv -Path $allFile -Encoding UTF8 -NoTypeInformation
    Write-Host " All: $allFile" -ForegroundColor White

    if ($successHost.Count) {
        $successHost | Select-Object Host, Protocol, Port, Status, StatusDetail,
                                     'Port Opening', 'Firewall blocking', 'Service Listening' |
            Export-Csv -Path $successFile -Encoding UTF8 -NoTypeInformation
        Write-Host " Success: $successFile" -ForegroundColor Green
    }
    if ($failedHost.Count) {
        $failedHost | Select-Object Host, Protocol, Port, Status, StatusDetail,
                                    'Port Opening', 'Firewall blocking', 'Service Listening' |
            Export-Csv -Path $failedFile -Encoding UTF8 -NoTypeInformation
        Write-Host " Failed: $failedFile" -ForegroundColor Red
    }

    # ---------- HTML per host ----------
    if ($HtmlReport) {
        $hostHtml = @"
<h2>$hostName</h2>
<div class="summary">
    <strong>Total:</strong> $($allHost.Count) |
    <strong style='color:green'>Success: $($successHost.Count)</strong> |
    <strong style='color:red'>Failed: $($failedHost.Count)</strong>
</div>
<table>
<tr><th>Protocol</th><th>Port</th><th>Status</th><th>Detail</th>
    <th>Port Opening</th><th>Firewall blocking</th><th>Service Listening</th></tr>
"@
        foreach ($r in $allHost) {
            $rowClass = if ($r.Status -eq "Success") { "success" } else { "failed" }

            # Pre-calculate colors
            $poColor = if ($r.'Port Opening')     { "#27ae60" } else { "#95a5a6" }
            $fbColor = if ($r.'Firewall blocking'){ "#c0392b" } else { "#95a5a6" }
            $slColor = if ($r.'Service Listening'){ "#e67e22" } else { "#95a5a6" }

            $hostHtml += @"
<tr class='$rowClass'>
    <td>$($r.Protocol)</td>
    <td>$($r.Port)</td>
    <td>$($r.Status)</td>
    <td>$($r.StatusDetail)</td>
    <td style='color:$poColor'>$($r.'Port Opening')</td>
    <td style='color:$fbColor'>$($r.'Firewall blocking')</td>
    <td style='color:$slColor'>$($r.'Service Listening')</td>
</tr>
"@
        }
        $hostHtml += "</table><hr>"

        $fullHostHtml = @"
<!DOCTYPE html>
<html><head><title>$hostName Report</title>
<style>
    body{font-family:Segoe UI,sans-serif;margin:20px;background:#f4f4f4}
    h1{color:#2c3e50}
    h2{color:#2980b9}
    .summary{background:#fff;padding:12px;border-radius:6px;margin:10px 0;box-shadow:0 1px 5px rgba(0,0,0,.1)}
    table{width:100%;border-collapse:collapse;margin-top:10px}
    th{background:#3498db;color:#fff;padding:10px}
    td{padding:8px;border-bottom:1px solid #ddd}
    .success{background:#d5f5e3}
    .failed{background:#fadbd8}
</style></head><body>
<h1>Connectivity Test – $hostName</h1>
$hostHtml
<div style='color:#7f8c8d;font-size:.9em'>Generated: $(Get-Date)</div>
</body></html>
"@
        $fullHostHtml | Out-File -FilePath $htmlFile -Encoding UTF8
        Write-Host " HTML: $htmlFile" -ForegroundColor Cyan
        $summaryBody += $hostHtml
    }

    $allHostResults += $allHost

    Write-Host " Summary: $($allHost.Count) tests | Success: $($successHost.Count) | Failed: $($failedHost.Count)" -ForegroundColor Yellow
}

# -------------------------------------------------
# Combined Summary HTML
# -------------------------------------------------
if ($HtmlReport -and $allHostResults.Count -gt 0) {
    $totalTests   = $allHostResults.Count
    $totalSuccess = ($allHostResults | Where-Object { $_.Status -eq "Success" }).Count
    $totalFailed  = $totalTests - $totalSuccess

    $summaryHtml = @"
<!DOCTYPE html>
<html><head><title>Multi-Host Connectivity Summary</title>
<style>
    body{font-family:Segoe UI,sans-serif;margin:20px;background:#f9f9f9}
    h1{color:#2c3e50;text-align:center}
    .summary{background:#fff;padding:15px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1);text-align:center;margin-bottom:30px}
    .host-section{margin-bottom:40px}
    h2{color:#2980b9;border-bottom:2px solid #3498db;padding-bottom:5px}
    table{width:100%;border-collapse:collapse;margin-top:10px}
    th{background:#3498db;color:#fff;padding:10px}
    td{padding:8px;border-bottom:1px solid #ddd}
    .success{background:#d5f5e3}
    .failed{background:#fadbd8}
    .footer{margin-top:50px;color:#7f8c8d;font-size:.9em;text-align:center}
</style></head><body>
<h1>Multi-Host Connectivity Test Summary</h1>
<div class="summary">
    <strong>Generated:</strong> $(Get-Date) <br>
    <strong>Total Hosts:</strong> $($ComputerName.Count) <br>
    <strong>Total Tests:</strong> $totalTests |
    <strong style='color:green'>Success: $totalSuccess</strong> |
    <strong style='color:red'>Failed: $totalFailed</strong>
</div>
<div class="host-section">
$summaryBody
</div>
<div class="footer">Generated by Test-ConnectionPorts.ps1</div>
</body></html>
"@
    $summaryHtml | Out-File -FilePath $summaryHtmlPath -Encoding UTF8
    Write-Host "`nCombined Summary HTML: $summaryHtmlPath" -ForegroundColor Cyan
}

Write-Host "`nAll done! Results in: $OutputPath" -ForegroundColor Green