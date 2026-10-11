<#
.SYNOPSIS
    Generates a self-contained HTML network health report for a Windows laptop.

.DESCRIPTION
    Sections:
      1. Laptop identity (hostname, IP, default gateway, DNS)
      2. Connectivity tests (ping: gateway + configurable hosts)
      3. DNS resolution test
      4. TCP port connectivity test
      5. Network interface status and packet statistics
      6. Active TCP connection summary by state

    Overall status: HEALTHY / WARNING / CRITICAL.
    Exit code: 0 = HEALTHY, 1 = WARNING, 2 = CRITICAL (handy for scheduled tasks / RMM tools).

.EXAMPLE
    .\New-NetworkHealthReport.ps1 -CompanyName "Contoso Ltd" -LogoPath C:\Branding\logo.png -OpenReport

.NOTES
    Requires Windows PowerShell 5.1 or PowerShell 7+ on Windows 10/11 or Server 2016+.
    Administrator rights are not required.
    This file is intentionally ASCII-only so Windows PowerShell 5.1 reads it correctly.
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $env:USERPROFILE ("NetworkHealth_{0}_{1:yyyyMMdd_HHmmss}.html" -f $env:COMPUTERNAME, (Get-Date))),
    [string]$CompanyName = 'Your Company Name',
    [string]$ReportTitle = 'Network Health Report',
    [string]$LogoPath,
    [string[]]$PingTargets = @('8.8.8.8', '1.1.1.1', 'google.com', '10.100.20.130', '10.100.20.131'),
    [string[]]$DnsDomains = @('google.com', 'microsoft.com', 'cloudflare.com'),
    [string]$PortTarget = '10.100.20.11',
    [int[]]$Ports = @(22, 80, 443, 3306),
    [int]$PingCount = 4,
    [int]$PingTimeoutMs = 2000,
    [int]$LatencyWarnMs = 100,
    [int]$DnsWarnMs = 500,
    [int]$TcpTimeoutMs = 3000,
    [switch]$OpenReport
)

$script:StartTime = Get-Date
$script:Results = New-Object System.Collections.ArrayList

#region ---------- Helper functions ----------

function Add-Result {
    param(
        [string]$Section,
        [string]$Item,
        [ValidateSet('OK', 'WARN', 'FAIL', 'INFO')][string]$Level,
        [bool]$Critical = $false,
        [string]$Detail = ''
    )
    [void]$script:Results.Add([pscustomobject]@{
            Section  = $Section
            Item     = $Item
            Level    = $Level
            Critical = $Critical
            Detail   = $Detail
        })
}

function Enc {
    param($Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Get-Badge {
    param([string]$Level, [string]$Text)
    if (-not $Text) { $Text = $Level }
    $cls = $Level.ToLower()
    return "<span class=`"badge $cls`">$(Enc $Text)</span>"
}

function Get-SectionLevel {
    param([string]$Section)
    $items = @($script:Results | Where-Object { $_.Section -eq $Section })
    if (@($items | Where-Object { $_.Level -eq 'FAIL' }).Count -gt 0) { return 'FAIL' }
    if (@($items | Where-Object { $_.Level -eq 'WARN' }).Count -gt 0) { return 'WARN' }
    if ($items.Count -gt 0) { return 'OK' }
    return 'INFO'
}

function New-HtmlTable {
    param(
        [object[]]$Rows,
        [string[]]$Columns,
        [string]$BadgeColumn = 'Health',
        [string[]]$RawColumns = @(),
        [string]$Empty = 'No data available.'
    )
    $Rows = @($Rows)
    if ($Rows.Count -eq 0) { return "<p class=`"empty`">$(Enc $Empty)</p>" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="table-wrap"><table><thead><tr>')
    foreach ($c in $Columns) { [void]$sb.AppendLine("<th>$(Enc $c)</th>") }
    [void]$sb.AppendLine('</tr></thead><tbody>')
    foreach ($r in $Rows) {
        [void]$sb.AppendLine('<tr>')
        foreach ($c in $Columns) {
            $v = $r.$c
            if ($c -eq $BadgeColumn) { $cell = Get-Badge -Level ([string]$v) }
            elseif ($RawColumns -contains $c) { $cell = [string]$v }
            else { $cell = Enc $v }
            [void]$sb.AppendLine("<td>$cell</td>")
        }
        [void]$sb.AppendLine('</tr>')
    }
    [void]$sb.AppendLine('</tbody></table></div>')
    return $sb.ToString()
}

function New-KvGrid {
    param([object[]]$Items)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="kv-grid">')
    foreach ($i in $Items) {
        if ([string]::IsNullOrWhiteSpace([string]$i.Value)) { $val = '<span class="muted">Not found</span>' }
        else { $val = Enc $i.Value }
        $badge = ''
        if ($i.Level) { $badge = ' ' + (Get-Badge -Level $i.Level) }
        [void]$sb.AppendLine("<div class=`"kv`"><div class=`"k`">$(Enc $i.Label)</div><div class=`"v`">$val$badge</div></div>")
    }
    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function New-Section {
    param([int]$Number, [string]$Title, [string]$Subtitle, [string]$Body, [string]$Level = 'INFO')
    $sub = ''
    if ($Subtitle) { $sub = "<p class=`"sub`">$(Enc $Subtitle)</p>" }
    return @"
<section class="card">
  <div class="card-head"><h2><span class="num">$Number</span>$(Enc $Title)</h2>$(Get-Badge -Level $Level)</div>
  $sub
  $Body
</section>
"@
}

function Test-HostPing {
    param([string]$Target, [int]$Count, [int]$TimeoutMs)
    $pinger = New-Object System.Net.NetworkInformation.Ping
    $times = New-Object 'System.Collections.Generic.List[long]'
    $resolved = ''
    $note = ''
    for ($i = 0; $i -lt $Count; $i++) {
        try {
            $reply = $pinger.Send($Target, $TimeoutMs)
            if ($reply.Status -eq [System.Net.NetworkInformation.IPStatus]::Success) {
                $times.Add([long]$reply.RoundtripTime)
                $resolved = $reply.Address.ToString()
            }
            else { $note = [string]$reply.Status }
        }
        catch {
            $inner = $_.Exception
            if ($inner.InnerException) { $inner = $inner.InnerException }
            $note = $inner.Message
        }
        if ($i -lt ($Count - 1)) { Start-Sleep -Milliseconds 250 }
    }
    $recv = $times.Count
    $loss = [math]::Round(100 * ($Count - $recv) / $Count, 0)
    $min = $null; $avg = $null; $max = $null
    if ($recv -gt 0) {
        $m = $times | Measure-Object -Minimum -Maximum -Average
        $min = [int]$m.Minimum; $max = [int]$m.Maximum; $avg = [int][math]::Round($m.Average, 0)
        $note = ''
    }
    [pscustomobject]@{
        Sent = $Count; Received = $recv; LossPct = $loss
        Min = $min; Avg = $avg; Max = $max; Resolved = $resolved; Note = $note
    }
}

function Test-TcpPort {
    param([string]$Target, [int]$Port, [int]$TimeoutMs)
    $client = New-Object System.Net.Sockets.TcpClient
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $state = 'ERROR'
    try {
        $iar = $client.BeginConnect($Target, $Port, $null, $null)
        if ($iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            try { $client.EndConnect($iar); $state = 'OPEN' }
            catch { $state = 'CLOSED' }
        }
        else { $state = 'TIMEOUT' }
    }
    catch { $state = 'ERROR' }
    finally { $sw.Stop(); $client.Close() }
    [pscustomobject]@{ State = $state; Ms = [int]$sw.ElapsedMilliseconds }
}

#endregion

#region ---------- 1. Laptop identity ----------
Write-Host '[1/6] Collecting laptop identity...' -ForegroundColor Cyan

$hostname = $env:COMPUTERNAME
$gateway = $null; $primaryIp = $null; $dnsServers = @(); $ifAlias = $null; $mac = $null

$defRoute = Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue |
    Where-Object { $_.NextHop -and $_.NextHop -ne '0.0.0.0' } |
    Sort-Object {
        $im = (Get-NetIPInterface -InterfaceIndex $_.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).InterfaceMetric
        $_.RouteMetric + $im
    } | Select-Object -First 1

if ($defRoute) {
    $gateway = $defRoute.NextHop
    $ifAlias = $defRoute.InterfaceAlias
    $ipObj = Get-NetIPAddress -InterfaceIndex $defRoute.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($ipObj) { $primaryIp = '{0}/{1}' -f $ipObj.IPAddress, $ipObj.PrefixLength }
    $dnsServers = @((Get-DnsClientServerAddress -InterfaceIndex $defRoute.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses)
    $mac = (Get-NetAdapter -InterfaceIndex $defRoute.InterfaceIndex -ErrorAction SilentlyContinue).MacAddress
}

$osCaption = ''; $domain = ''
try { $osCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).Caption } catch { }
try { $domain = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).Domain } catch { }

$dnsText = ($dnsServers -join ', ')

if ($primaryIp) { $ipLvl = 'OK' } else { $ipLvl = 'FAIL' }
if ($gateway) { $gwLvl = 'OK' } else { $gwLvl = 'FAIL' }
if ($dnsText) { $dnsLvl = 'OK' } else { $dnsLvl = 'FAIL' }

Add-Result 'Identity' 'IP address' $ipLvl ($ipLvl -eq 'FAIL') $(if ($ipLvl -eq 'FAIL') { 'No IPv4 address on the default-route interface' } else { $primaryIp })
Add-Result 'Identity' 'Default gateway' $gwLvl ($gwLvl -eq 'FAIL') $(if ($gwLvl -eq 'FAIL') { 'No default gateway configured' } else { $gateway })
Add-Result 'Identity' 'DNS servers' $dnsLvl ($dnsLvl -eq 'FAIL') $(if ($dnsLvl -eq 'FAIL') { 'No DNS servers configured' } else { $dnsText })

$identityHtml = New-KvGrid -Items @(
    @{ Label = 'Hostname'; Value = $hostname },
    @{ Label = 'IP Address'; Value = $primaryIp; Level = $ipLvl },
    @{ Label = 'Default Gateway'; Value = $gateway; Level = $gwLvl },
    @{ Label = 'DNS Servers'; Value = $dnsText; Level = $dnsLvl },
    @{ Label = 'Active Interface'; Value = $ifAlias },
    @{ Label = 'MAC Address'; Value = $mac },
    @{ Label = 'Domain / Workgroup'; Value = $domain },
    @{ Label = 'Operating System'; Value = $osCaption },
    @{ Label = 'Logged-on User'; Value = "$env:USERDOMAIN\$env:USERNAME" }
)
#endregion

#region ---------- 2. Connectivity (ping) ----------
Write-Host '[2/6] Running ping tests...' -ForegroundColor Cyan

$pingList = @()
if ($gateway) { $pingList += [pscustomobject]@{ Target = $gateway; Role = 'Default Gateway'; Critical = $true } }
foreach ($t in $PingTargets) {
    if ($t -match '^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[01])\.)') { $role = 'Internal' } else { $role = 'Internet' }
    $pingList += [pscustomobject]@{ Target = $t; Role = $role; Critical = $false }
}

$pingRows = @()
$pingFailCount = 0
foreach ($p in $pingList) {
    $r = Test-HostPing -Target $p.Target -Count $PingCount -TimeoutMs $PingTimeoutMs
    if ($r.Received -eq 0) { $lvl = 'FAIL'; $pingFailCount++ }
    elseif ($r.LossPct -gt 0 -or $r.Avg -gt $LatencyWarnMs) { $lvl = 'WARN' }
    else { $lvl = 'OK' }

    if ($r.Received -gt 0) {
        $latency = '{0} / {1} / {2}' -f $r.Min, $r.Avg, $r.Max
        $detail = 'loss {0}%, avg {1} ms' -f $r.LossPct, $r.Avg
    }
    else { $latency = '-'; $detail = 'No reply. ' + $r.Note }

    Add-Result 'Connectivity' ("Ping {0} ({1})" -f $p.Target, $p.Role) $lvl $p.Critical $detail

    $pingRows += [pscustomobject]@{
        'Target'               = $p.Target
        'Role'                 = $p.Role
        'Resolved IP'          = $r.Resolved
        'Sent'                 = $r.Sent
        'Recv'                 = $r.Received
        'Loss %'               = $r.LossPct
        'Min / Avg / Max (ms)' = $latency
        'Notes'                = $r.Note
        'Health'               = $lvl
    }
}
if (-not $gateway) {
    $pingRows = @([pscustomobject]@{
            'Target' = '(none)'; 'Role' = 'Default Gateway'; 'Resolved IP' = ''; 'Sent' = 0; 'Recv' = 0; 'Loss %' = 100
            'Min / Avg / Max (ms)' = '-'; 'Notes' = 'No default gateway found'; 'Health' = 'FAIL'
        }) + $pingRows
}
if ($pingList.Count -gt 0 -and $pingFailCount -eq $pingList.Count) {
    Add-Result 'Connectivity' 'All ping targets' 'FAIL' $true 'Every ping target failed - no network connectivity'
}

$pingHtml = New-HtmlTable -Rows $pingRows -Columns @('Target', 'Role', 'Resolved IP', 'Sent', 'Recv', 'Loss %', 'Min / Avg / Max (ms)', 'Notes', 'Health')
#endregion

#region ---------- 3. DNS resolution ----------
Write-Host '[3/6] Testing DNS resolution...' -ForegroundColor Cyan

$dnsRows = @()
$dnsFail = 0
foreach ($d in $DnsDomains) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $addresses = ''; $note = ''
    try {
        $res = @(Resolve-DnsName -Name $d -Type A -DnsOnly -NoHostsFile -ErrorAction Stop | Where-Object { $_.Type -eq 'A' })
        $sw.Stop()
        if ($res.Count -gt 0) { $addresses = ($res.IPAddress -join ', ') } else { $note = 'No A records returned' }
    }
    catch {
        $sw.Stop()
        $note = $_.Exception.Message
    }
    $ms = [int]$sw.ElapsedMilliseconds
    if (-not $addresses) { $lvl = 'FAIL'; $dnsFail++; $detail = "Resolution failed: $note" }
    elseif ($ms -gt $DnsWarnMs) { $lvl = 'WARN'; $detail = "Slow resolution ($ms ms)" }
    else { $lvl = 'OK'; $detail = "Resolved in $ms ms" }

    Add-Result 'DNS' "Resolve $d" $lvl $false $detail
    $dnsRows += [pscustomobject]@{
        'Domain'             = $d
        'Resolved Addresses' = $addresses
        'Response (ms)'      = $ms
        'Notes'              = $note
        'Health'             = $lvl
    }
}
if ($DnsDomains.Count -gt 0 -and $dnsFail -eq $DnsDomains.Count) {
    Add-Result 'DNS' 'All DNS lookups' 'FAIL' $true 'DNS resolution is completely failing'
}
$dnsHtml = New-HtmlTable -Rows $dnsRows -Columns @('Domain', 'Resolved Addresses', 'Response (ms)', 'Notes', 'Health')
#endregion

#region ---------- 4. Port connectivity ----------
Write-Host "[4/6] Testing TCP ports on $PortTarget..." -ForegroundColor Cyan

$svcNames = @{ 21 = 'FTP'; 22 = 'SSH'; 23 = 'Telnet'; 25 = 'SMTP'; 53 = 'DNS'; 80 = 'HTTP'; 110 = 'POP3'; 143 = 'IMAP'
    389 = 'LDAP'; 443 = 'HTTPS'; 445 = 'SMB'; 1433 = 'MS SQL'; 3306 = 'MySQL/MariaDB'; 3389 = 'RDP'; 5432 = 'PostgreSQL'; 8080 = 'HTTP-Alt' }

$portRows = @()
foreach ($port in $Ports) {
    $t = Test-TcpPort -Target $PortTarget -Port $port -TimeoutMs $TcpTimeoutMs
    if ($t.State -eq 'OPEN') { $lvl = 'OK' } else { $lvl = 'FAIL' }
    if ($svcNames.ContainsKey($port)) { $svc = $svcNames[$port] } else { $svc = '-' }
    Add-Result 'Ports' ("{0}:{1} ({2})" -f $PortTarget, $port, $svc) $lvl $false ("{0} in {1} ms" -f $t.State, $t.Ms)
    $portRows += [pscustomobject]@{
        'Port'          = $port
        'Service'       = $svc
        'Result'        = $t.State
        'Response (ms)' = $t.Ms
        'Health'        = $lvl
    }
}
$portHtml = New-HtmlTable -Rows $portRows -Columns @('Port', 'Service', 'Result', 'Response (ms)', 'Health')
#endregion

#region ---------- 5. Network interfaces ----------
Write-Host '[5/6] Collecting interface status and statistics...' -ForegroundColor Cyan

$stats = @{}
Get-NetAdapterStatistics -ErrorAction SilentlyContinue | ForEach-Object { $stats[$_.Name] = $_ }
$allIps = @(Get-NetIPAddress -ErrorAction SilentlyContinue)

$ifRows = @()
foreach ($a in @(Get-NetAdapter -ErrorAction SilentlyContinue | Sort-Object Status, Name)) {
    $mine = @($allIps | Where-Object { $_.InterfaceIndex -eq $a.ifIndex })
    $v4 = @($mine | Where-Object { $_.AddressFamily -eq 'IPv4' } | ForEach-Object { Enc ('{0}/{1}' -f $_.IPAddress, $_.PrefixLength) })
    $v6 = @($mine | Where-Object { $_.AddressFamily -eq 'IPv6' -and $_.IPAddress -notlike 'fe80*' } | ForEach-Object { Enc $_.IPAddress })
    $ipCell = (($v4 + $v6) -join '<br>')
    if (-not $ipCell) { $ipCell = '<span class="muted">-</span>' }

    $s = $stats[$a.Name]
    $sent = 0; $recv = 0; $errs = 0; $disc = 0
    if ($s) {
        $sent = [int64]$s.SentUnicastPackets
        $recv = [int64]$s.ReceivedUnicastPackets
        $errs = [int64]$s.ReceivedPacketErrors + [int64]$s.OutboundPacketErrors
        $disc = [int64]$s.ReceivedDiscardedPackets + [int64]$s.OutboundDiscardedPackets
    }

    if ($a.Status -eq 'Up') {
        $total = $sent + $recv
        $badPct = 0
        if ($total -gt 0) { $badPct = 100 * ($errs + $disc) / $total }
        if ($badPct -gt 1) { $lvl = 'WARN' } else { $lvl = 'OK' }
        Add-Result 'Interfaces' $a.Name $lvl $false ('{0:N2}% errors/discards ({1} errors, {2} discards)' -f $badPct, $errs, $disc)
    }
    else { $lvl = 'INFO' }

    $ifRows += [pscustomobject]@{
        'Interface'    = $a.Name
        'Description'  = $a.InterfaceDescription
        'State'        = [string]$a.Status
        'Speed'        = $a.LinkSpeed
        'MAC'          = $a.MacAddress
        'IP Addresses' = $ipCell
        'Pkts Sent'    = '{0:N0}' -f $sent
        'Pkts Recv'    = '{0:N0}' -f $recv
        'Errors'       = '{0:N0}' -f $errs
        'Discards'     = '{0:N0}' -f $disc
        'Health'       = $lvl
    }
}
$ifHtml = New-HtmlTable -Rows $ifRows -RawColumns @('IP Addresses') `
    -Columns @('Interface', 'Description', 'State', 'Speed', 'MAC', 'IP Addresses', 'Pkts Sent', 'Pkts Recv', 'Errors', 'Discards', 'Health') `
    -Empty 'No network adapters found.'
#endregion

#region ---------- 6. TCP connection summary ----------
Write-Host '[6/6] Summarising TCP connections...' -ForegroundColor Cyan

$stateNames = @{ Closed = 'CLOSED'; Listen = 'LISTEN'; SynSent = 'SYN_SENT'; SynReceived = 'SYN_RECEIVED'; Established = 'ESTABLISHED'
    FinWait1 = 'FIN_WAIT_1'; FinWait2 = 'FIN_WAIT_2'; CloseWait = 'CLOSE_WAIT'; Closing = 'CLOSING'; LastAck = 'LAST_ACK'
    TimeWait = 'TIME_WAIT'; DeleteTCB = 'DELETE_TCB'; Bound = 'BOUND' }
$stateLimits = @{ CloseWait = 50; TimeWait = 1000; SynSent = 20 }

$conns = @(Get-NetTCPConnection -ErrorAction SilentlyContinue)
$totalConns = $conns.Count
$groups = @($conns | Group-Object State | Sort-Object Count -Descending)
$counts = @{}
foreach ($g in $groups) { $counts[[string]$g.Name] = $g.Count }

$tcpRows = @()
foreach ($g in $groups) {
    $name = [string]$g.Name
    if ($stateNames.ContainsKey($name)) { $label = $stateNames[$name] } else { $label = $name.ToUpper() }
    $pct = 0
    if ($totalConns -gt 0) { $pct = [math]::Round(100 * $g.Count / $totalConns, 1) }
    if ($name -eq 'Listen' -or $name -eq 'Bound') { $lvl = 'INFO' }
    elseif ($stateLimits.ContainsKey($name) -and $g.Count -gt $stateLimits[$name]) { $lvl = 'WARN' }
    else { $lvl = 'OK' }
    $bar = '<div class="bar"><span style="width:{0}%"></span></div>' -f ([math]::Min(100, [math]::Max(1, $pct)))
    $tcpRows += [pscustomobject]@{
        'State'  = $label
        'Count'  = '{0:N0}' -f $g.Count
        'Share'  = '{0}%' -f $pct
        'Chart'  = $bar
        'Health' = $lvl
    }
}
foreach ($k in $stateLimits.Keys) {
    $c = 0
    if ($counts.ContainsKey($k)) { $c = $counts[$k] }
    if ($c -gt $stateLimits[$k]) { $lvl = 'WARN' } else { $lvl = 'OK' }
    Add-Result 'TCP' $stateNames[$k] $lvl $false ("{0} connections (threshold {1})" -f $c, $stateLimits[$k])
}
$tcpHtml = New-HtmlTable -Rows $tcpRows -Columns @('State', 'Count', 'Share', 'Chart', 'Health') -RawColumns @('Chart') -Empty 'No TCP connections found.'
$tcpHtml += "<p class=`"sub`">Total TCP connections: <strong>{0:N0}</strong>. Thresholds: CLOSE_WAIT &gt; {1}, TIME_WAIT &gt; {2}, SYN_SENT &gt; {3}.</p>" -f $totalConns, $stateLimits['CloseWait'], $stateLimits['TimeWait'], $stateLimits['SynSent']
#endregion

#region ---------- Overall status ----------
$critFails = @($script:Results | Where-Object { $_.Level -eq 'FAIL' -and $_.Critical })
$fails = @($script:Results | Where-Object { $_.Level -eq 'FAIL' })
$warns = @($script:Results | Where-Object { $_.Level -eq 'WARN' })
$oks = @($script:Results | Where-Object { $_.Level -eq 'OK' })

if ($critFails.Count -gt 0) { $overall = 'CRITICAL'; $exitCode = 2 }
elseif ($fails.Count -gt 0 -or $warns.Count -gt 0) { $overall = 'WARNING'; $exitCode = 1 }
else { $overall = 'HEALTHY'; $exitCode = 0 }

$overallMsg = switch ($overall) {
    'HEALTHY' { 'All checks passed.' }
    'WARNING' { 'Some checks need attention. Core connectivity is intact.' }
    'CRITICAL' { 'One or more critical checks failed. Immediate attention required.' }
}

$findingRows = @()
$ordered = @($script:Results | Where-Object { $_.Level -eq 'FAIL' -or $_.Level -eq 'WARN' } |
        Sort-Object @{ Expression = { if ($_.Level -eq 'FAIL' -and $_.Critical) { 0 } elseif ($_.Level -eq 'FAIL') { 1 } else { 2 } } })
foreach ($f in $ordered) {
    $sev = ''
    if ($f.Critical -and $f.Level -eq 'FAIL') { $sev = 'Critical' }
    $findingRows += [pscustomobject]@{ 'Section' = $f.Section; 'Check' = $f.Item; 'Detail' = $f.Detail; 'Severity' = $sev; 'Health' = $f.Level }
}
if ($findingRows.Count -gt 0) {
    $findingsBody = New-HtmlTable -Rows $findingRows -Columns @('Section', 'Check', 'Detail', 'Severity', 'Health')
}
else { $findingsBody = '<p class="empty">No issues detected. All checks passed.</p>' }
#endregion

#region ---------- Build HTML ----------
$css = @'
:root{--bg:#f3f5f9;--card:#fff;--text:#1d2733;--muted:#6b7785;--line:#e3e8ef;--accent:#1f6feb}
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);font-size:14px;line-height:1.45}
.header{background:linear-gradient(135deg,#0b2545,#13315c 60%,#1f4f8f);color:#fff;padding:20px 32px;display:flex;align-items:center;justify-content:space-between;gap:20px;flex-wrap:wrap}
.header-left{display:flex;align-items:center;gap:22px;flex-wrap:wrap}
.logo-placeholder{width:180px;height:56px;border:2px dashed rgba(255,255,255,.5);border-radius:6px;display:flex;flex-direction:column;align-items:center;justify-content:center;font-weight:600;letter-spacing:1px;font-size:13px;color:rgba(255,255,255,.85)}
.logo-placeholder small{font-weight:400;opacity:.7;letter-spacing:0}
.logo{max-height:56px;max-width:220px}
.header h1{margin:0;font-size:22px;font-weight:600}
.header .company{font-size:13px;opacity:.85;margin-top:2px}
.header .meta{text-align:right;font-size:12px;opacity:.9}
.container{max-width:1200px;margin:0 auto;padding:24px 20px 40px}
.overall{border-radius:10px;padding:22px 28px;color:#fff;display:flex;align-items:center;justify-content:space-between;gap:24px;flex-wrap:wrap;box-shadow:0 2px 8px rgba(0,0,0,.12);margin-bottom:20px}
.overall.healthy{background:linear-gradient(135deg,#1e8e3e,#2fb457)}
.overall.warning{background:linear-gradient(135deg,#e69500,#f9ab00);color:#2b1d00}
.overall.critical{background:linear-gradient(135deg,#c5221f,#e0443e)}
.overall .label{font-size:12px;letter-spacing:2px;text-transform:uppercase;opacity:.85}
.overall .state{font-size:40px;font-weight:700;letter-spacing:1px;line-height:1.1}
.overall .msg{margin-top:4px;opacity:.95}
.counters{display:flex;gap:12px}
.counter{background:rgba(255,255,255,.22);border-radius:8px;padding:10px 18px;text-align:center;min-width:86px}
.counter b{display:block;font-size:26px;line-height:1.2}
.counter span{font-size:11px;text-transform:uppercase;letter-spacing:1px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px 24px;margin-bottom:18px;box-shadow:0 1px 3px rgba(16,24,40,.05)}
.card-head{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:6px}
.card h2{margin:0;font-size:17px;font-weight:600;display:flex;align-items:center;gap:10px}
.num{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:50%;background:#0b2545;color:#fff;font-size:13px;font-weight:600}
.sub{color:var(--muted);margin:2px 0 12px;font-size:13px}
.kv-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-top:12px}
.kv{border:1px solid var(--line);border-radius:8px;padding:10px 14px;background:#fafbfd}
.kv .k{font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:3px}
.kv .v{font-size:15px;font-weight:600;word-break:break-word}
.muted{color:var(--muted);font-weight:400}
.table-wrap{overflow-x:auto;margin-top:8px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:#eef2f7;text-align:left;padding:9px 12px;font-weight:600;color:#33415c;border-bottom:2px solid var(--line);white-space:nowrap}
td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
tbody tr:nth-child(even){background:#fafbfd}
tbody tr:hover{background:#f1f6ff}
.badge{display:inline-block;padding:2px 11px;border-radius:999px;font-size:11px;font-weight:700;letter-spacing:.5px;border:1px solid transparent;white-space:nowrap}
.badge.ok{background:#e6f4ea;color:#1e8e3e;border-color:#b7e1c1}
.badge.warn{background:#fff4d6;color:#b06000;border-color:#f3d98b}
.badge.fail{background:#fce8e6;color:#c5221f;border-color:#f4b9b5}
.badge.info{background:#eceff3;color:#5f6b7a;border-color:#d5dbe3}
.bar{background:#e8edf3;border-radius:4px;height:8px;width:160px;overflow:hidden}
.bar span{display:block;height:100%;background:var(--accent)}
.empty{color:var(--muted);padding:8px 0}
.legend{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-top:8px;color:var(--muted);font-size:12px}
.footer{color:var(--muted);font-size:12px;text-align:center;margin-top:24px}
@media print{body{background:#fff}.card,.overall{box-shadow:none;break-inside:avoid}.header,.overall{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
'@

$logoHtml = '<div class="logo-placeholder">YOUR LOGO<small>180 x 56 px</small></div>'
if ($LogoPath -and (Test-Path -LiteralPath $LogoPath)) {
    $ext = [System.IO.Path]::GetExtension($LogoPath).TrimStart('.').ToLower()
    $mime = switch ($ext) {
        'png' { 'image/png' }
        'jpg' { 'image/jpeg' }
        'jpeg' { 'image/jpeg' }
        'gif' { 'image/gif' }
        'svg' { 'image/svg+xml' }
        default { 'image/png' }
    }
    $b64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($LogoPath))
    $logoHtml = "<img class=`"logo`" alt=`"Company logo`" src=`"data:$mime;base64,$b64`">"
}

$elapsed = [math]::Round(((Get-Date) - $script:StartTime).TotalSeconds, 1)
$generated = Get-Date -Format 'dddd, dd MMMM yyyy HH:mm:ss'
$overallClass = $overall.ToLower()

$sec1 = New-Section 1 'Laptop Identity' 'Identity of this machine and its primary network configuration (interface holding the default route).' $identityHtml (Get-SectionLevel 'Identity')
$sec2 = New-Section 2 'Connectivity Tests' ("ICMP ping, {0} packets per target, {1} ms timeout. Latency warning above {2} ms." -f $PingCount, $PingTimeoutMs, $LatencyWarnMs) $pingHtml (Get-SectionLevel 'Connectivity')
$sec3 = New-Section 3 'DNS Resolution' ("A-record lookups using this laptop's configured resolver(s): {0}. Slow-response threshold {1} ms." -f $dnsText, $DnsWarnMs) $dnsHtml (Get-SectionLevel 'DNS')
$sec4 = New-Section 4 'Port Connectivity' ("TCP connect tests to {0}, {1} ms timeout." -f $PortTarget, $TcpTimeoutMs) $portHtml (Get-SectionLevel 'Ports')
$sec5 = New-Section 5 'Network Interface Status' 'All network adapters, including virtual and disconnected ones. Disconnected adapters are informational only.' $ifHtml (Get-SectionLevel 'Interfaces')
$sec6 = New-Section 6 'Active TCP Connections' 'Summary of TCP connections on this laptop grouped by state.' $tcpHtml (Get-SectionLevel 'TCP')

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$(Enc $ReportTitle) - $(Enc $hostname)</title>
<style>
$css
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    $logoHtml
    <div>
      <h1>$(Enc $ReportTitle)</h1>
      <div class="company">$(Enc $CompanyName)</div>
    </div>
  </div>
  <div class="meta">
    <div><strong>$(Enc $hostname)</strong></div>
    <div>$(Enc $generated)</div>
  </div>
</div>

<div class="container">

  <div class="overall $overallClass">
    <div>
      <div class="label">Overall Network Health</div>
      <div class="state">$overall</div>
      <div class="msg">$(Enc $overallMsg)</div>
    </div>
    <div class="counters">
      <div class="counter"><b>$($oks.Count)</b><span>OK</span></div>
      <div class="counter"><b>$($warns.Count)</b><span>Warnings</span></div>
      <div class="counter"><b>$($fails.Count)</b><span>Failures</span></div>
    </div>
  </div>

  <section class="card">
    <div class="card-head"><h2>Key Findings</h2></div>
    $findingsBody
  </section>

  $sec1
  $sec2
  $sec3
  $sec4
  $sec5
  $sec6

  <section class="card">
    <div class="card-head"><h2>How to read this report</h2></div>
    <div class="legend">
      <span>$(Get-Badge -Level 'OK') check passed</span>
      <span>$(Get-Badge -Level 'WARN') degraded / needs attention</span>
      <span>$(Get-Badge -Level 'FAIL') check failed</span>
      <span>$(Get-Badge -Level 'INFO') informational only</span>
    </div>
    <p class="sub">CRITICAL = no IP/gateway/DNS configured, gateway unreachable, all pings failing, or all DNS lookups failing. WARNING = any other failed or degraded check. HEALTHY = all checks passed.</p>
  </section>

  <div class="footer">Generated by New-NetworkHealthReport.ps1 on $(Enc $hostname) in $elapsed seconds.</div>
</div>
</body>
</html>
"@
#endregion

#region ---------- Save ----------
$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
[System.IO.File]::WriteAllText($OutputPath, $html, (New-Object System.Text.UTF8Encoding($false)))

$color = switch ($overall) { 'HEALTHY' { 'Green' } 'WARNING' { 'Yellow' } 'CRITICAL' { 'Red' } }
Write-Host ''
Write-Host "Overall status : $overall" -ForegroundColor $color
Write-Host "Report saved   : $OutputPath" -ForegroundColor Cyan

if ($OpenReport) { Start-Process $OutputPath }
exit $exitCode
#endregion
