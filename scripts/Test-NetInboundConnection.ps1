#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Accurately checks local inbound TCP/UDP ports and ICMP.
    Uses proper firewall port filter lookup.
    Groups by local IP. Outputs CSV + HTML + Summary.

.PARAMETER TCPPorts
    Comma-separated TCP ports.

.PARAMETER UDPPorts
    Comma-separated UDP ports.

.PARAMETER OutputPath
    Report directory.

.PARAMETER CsvPrefix
    File prefix.

.PARAMETER GenerateHtml
    Generate HTML.

.PARAMETER CheckIcmp
    Include ICMP.

.EXAMPLE
    .\Check-InboundConnectivity.ps1 -TCPPorts "135,445" -UDPPorts "137" -GenerateHtml
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TCPPorts,

    [Parameter(Mandatory=$false)]
    [string]$UDPPorts = "",

    [string]$OutputPath = (Get-Location).Path,

    [string]$CsvPrefix = "Inbound_Check",

    [switch]$GenerateHtml,

    [switch]$CheckIcmp
)

# === Convert ports ===
function Convert-ToIntArray {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    return ($s -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
}

$TCPPortsArray = Convert-ToIntArray $TCPPorts
$UDPPortsArray = Convert-ToIntArray $UDPPorts

if ($TCPPortsArray.Count -eq 0) {
    Write-Error "No valid TCP ports."
    exit 1
}

# === Output setup ===
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$CsvFile  = Join-Path $OutputPath "${CsvPrefix}_${Timestamp}.csv"
$HtmlFile = Join-Path $OutputPath "${CsvPrefix}_${Timestamp}.html"

# === Global Results ===
$Global:Results = @()

# === Get Local IPv4 Addresses ===
$LocalIPs = @()
try {
    $LocalIPs = [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() `
        | Where-Object { $_.OperationalStatus -eq 'Up' -and $_.NetworkInterfaceType -ne 'Loopback' } `
        | ForEach-Object { $_.GetIPProperties().UnicastAddresses } `
        | Where-Object { $_.Address.AddressFamily -eq 'InterNetwork' } `
        | ForEach-Object { $_.Address.IPAddressToString } `
        | Sort-Object -Unique
}
catch { $LocalIPs = @("0.0.0.0") }
if ($LocalIPs.Count -eq 0) { $LocalIPs = @("0.0.0.0") }

Write-Host "`nLocal IP(s): $($LocalIPs -join ', ')" -ForegroundColor Cyan
Write-Host "TCP: $($TCPPortsArray -join ', ')" -ForegroundColor Gray
if ($UDPPortsArray) { Write-Host "UDP: $($UDPPortsArray -join ', ')" -ForegroundColor Gray }

# === Add Result ===
function Add-Result {
    param($IP, $Protocol, $Port, $Listening, $FirewallAllowed, $Status, $Details)
    $Global:Results += [PSCustomObject]@{
        IPAddress      = $IP
        Protocol       = $Protocol
        Port           = $Port
        Listening      = $Listening
        FirewallAllow  = $FirewallAllowed
        Status         = $Status
        Details        = $Details
    }
}

# === CORRECT FIREWALL CHECK FUNCTION ===
function Test-FirewallPortAllowed {
    param($Protocol, $Port)
    $rules = Get-NetFirewallRule -Direction Inbound -Enabled True -ErrorAction SilentlyContinue |
             Where-Object { $_.Enabled -eq $true }

    foreach ($rule in $rules) {
        try {
            $portFilter = $rule | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
            if ($portFilter.Protocol -eq $Protocol) {
                $localPorts = $portFilter.LocalPort
                if ($localPorts -contains $Port -or $localPorts -contains "Any" -or $localPorts -contains "$Port") {
                    return $true
                }
            }
        }
        catch { continue }
    }
    return $false
}

# === TCP CHECK ===
foreach ($ip in $LocalIPs) {
    foreach ($port in $TCPPortsArray) {
        $listening = $false
        $conn = Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue
        if ($conn) {
            $bound = $conn.LocalAddress | Where-Object { $_ -eq $ip -or $_ -eq "0.0.0.0" -or $_ -eq "::" }
            $listening = $null -ne $bound
        }

        $fwAllowed = Test-FirewallPortAllowed "TCP" $port

        if ($listening -and $fwAllowed) {
            Add-Result $ip "TCP" $port "Yes" "Yes" "Allowed" "Listening + Firewall allows"
        }
        elseif ($listening) {
            Add-Result $ip "TCP" $port "Yes" "No" "Warning" "Listening, blocked by firewall"
        }
        else {
            Add-Result $ip "TCP" $port "No" "No" "Blocked" "Not listening"
        }
    }
}

# === UDP CHECK ===
foreach ($ip in $LocalIPs) {
    foreach ($port in $UDPPortsArray) {
        $listening = $false
        $endpoint = Get-NetUDPEndpoint -LocalPort $port -ErrorAction SilentlyContinue
        if ($endpoint) {
            $bound = $endpoint.LocalAddress | Where-Object { $_ -eq $ip -or $_ -eq "0.0.0.0" -or $_ -eq "::" }
            $listening = $null -ne $bound
        }

        $fwAllowed = Test-FirewallPortAllowed "UDP" $port

        if ($listening -and $fwAllowed) {
            Add-Result $ip "UDP" $port "Yes" "Yes" "Allowed" "Listening + Firewall allows"
        }
        elseif ($listening) {
            Add-Result $ip "UDP" $port "Yes" "No" "Warning" "Listening, blocked by firewall"
        }
        else {
            Add-Result $ip "UDP" $port "No" "No" "Blocked" "Not listening"
        }
    }
}

# === ICMP CHECK ===
if ($CheckIcmp) {
    foreach ($ip in $LocalIPs) {
        $icmp4 = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing (Echo Request - ICMPv4-In)" -ErrorAction SilentlyContinue |
                 Where-Object { $_.Enabled -eq $True }
        $allow4 = $null -ne $icmp4
        Add-Result $ip "ICMP" "Ping (IPv4)" "-" $(if($allow4){"Yes"}else{"No"}) $(if($allow4){"Allowed"}else{"Blocked"}) "ICMPv4"

        $icmp6 = Get-NetFirewallRule -DisplayGroup "File and Printer Sharing (Echo Request - ICMPv6-In)" -ErrorAction SilentlyContinue |
                 Where-Object { $_.Enabled -eq $True }
        $allow6 = $null -ne $icmp6
        Add-Result $ip "ICMP" "Ping (IPv6)" "-" $(if($allow6){"Yes"}else{"No"}) $(if($allow6){"Allowed"}else{"Blocked"}) "ICMPv6"
    }
}

# === CONSOLE SUMMARY ===
Write-Host "`n=== RESULTS BY IP ===" -ForegroundColor Cyan
$Global:Results | Group-Object IPAddress | ForEach-Object {
    Write-Host "`n[IP: $($_.Name)]" -ForegroundColor Green
    $_.Group | Format-Table Protocol, Port, Listening, FirewallAllow, Status, Details -AutoSize
}

# === SUCCESS / FAIL SUMMARY ===
$Success = $Global:Results | Where-Object { $_.Status -eq "Allowed" -and $_.Protocol -ne "ICMP" } | ForEach-Object { "$($_.Protocol)/$($_.Port)" } | Sort-Object -Unique
$Failed  = $Global:Results | Where-Object { $_.Status -ne "Allowed" -and $_.Protocol -ne "ICMP" } | ForEach-Object { "$($_.Protocol)/$($_.Port)" } | Sort-Object -Unique

Write-Host "`n=== PORT STATUS ===" -ForegroundColor Yellow
if ($Success) { Write-Host "OPEN: " -NoNewline -ForegroundColor Green; Write-Host ($Success -join ", ") }
else { Write-Host "No ports open." -ForegroundColor Red }

if ($Failed) { Write-Host "BLOCKED: " -NoNewline -ForegroundColor Red; Write-Host ($Failed -join ", ") }

# === EXPORT CSV ===
try {
    $Global:Results | Export-Csv -Path $CsvFile -Encoding UTF8 -NoTypeInformation
    Write-Host "`nCSV: $CsvFile" -ForegroundColor Green
}
catch { Write-Warning "CSV failed: $($_.Exception.Message)" }

# === EXPORT HTML ===
if ($GenerateHtml) {
    $HtmlHeader = @"
<!DOCTYPE html><html><head><title>Inbound Check</title>
<style>
    body {font-family:Segoe UI,sans-serif;margin:20px;background:#f7f9fc;}
    h1,h2 {color:#2c3e50;}
    table {width:100%;border-collapse:collapse;margin:20px 0;}
    th,td {padding:12px;border-bottom:1px solid #ddd;text-align:left;}
    th {background:#3498db;color:white;}
    tr:nth-child(even) {background:#f2f2f2;}
    .allowed {background:#d5f4e6;font-weight:bold;}
    .warning {background:#fef9e7;color:#e67e22;font-weight:bold;}
    .blocked {background:#fadbd8;color:#c0392b;font-weight:bold;}
</style></head><body>
<h1>Local Inbound Report</h1>
<p><strong>Time:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') HKT</p>
<p><strong>Host:</strong> $env:COMPUTERNAME</p>
<p><strong>IPs:</strong> $($LocalIPs -join ', ')</p>
<p><strong>Open:</strong> $(if($Success){$Success -join ', '}else{'None'})</p>
<p><strong>Blocked:</strong> $(if($Failed){$Failed -join ', '}else{'None'})</p>
"@

    $HtmlBody = ""
    $Global:Results | Group-Object IPAddress | ForEach-Object {
        $HtmlBody += "<h2>IP: $($_.Name)</h2><table><tr><th>Proto</th><th>Port</th><th>Listen</th><th>FW</th><th>Status</th><th>Details</th></tr>"
        foreach ($r in $_.Group) {
            $cls = if ($r.Status -eq "Allowed") { "allowed" } elseif ($r.Status -eq "Warning") { "warning" } else { "blocked" }
            $HtmlBody += "<tr><td>$($r.Protocol)</td><td>$($r.Port)</td><td>$($r.Listening)</td><td>$($r.FirewallAllow)</td><td class='$cls'>$($r.Status)</td><td>$($r.Details)</td></tr>"
        }
        $HtmlBody += "</table>"
    }

    $HtmlContent = $HtmlHeader + $HtmlBody + "</body></html>"

    try {
        $HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8
        Write-Host "HTML: $HtmlFile" -ForegroundColor Green
    }
    catch { Write-Warning "HTML failed." }
}

Write-Host "`nDone at $(Get-Date -Format 'HH:mm:ss')!`n" -ForegroundColor Cyan