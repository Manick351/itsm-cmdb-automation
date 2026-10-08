<#
.SYNOPSIS
    Finds what is left in DNS and Active Directory after servers were retired in the CMDB, and says what to do with it.

.DESCRIPTION
    Read-only. Sources:
      - GLPI: retired servers (name, IP addresses, IPs mentioned in the description) and active names;
      - DNS: all A, PTR and CNAME records of the AD-integrated zones, read directly over LDAP (dnsNode objects
        in DomainDnsZones / ForestDnsZones) - one pass, no per-name queries;
      - network: ping and a TCP connect to -TcpPort for every address involved (runspace pool);
      - AD: the computer object (enabled, last logon).

    Each retired name gets a verdict:
      Clean      nothing left;
      Fix        records to delete (A, PTR, CNAME) - nobody uses the address;
      Reuse      the address now belongs to another host: delete the A record, replace the PTR;
      Manual     something still looks alive: the address answers without another owner, the A record is
                 dynamic and was refreshed recently, the AD object is enabled and logged on recently;
      Skip       the name is active again in the CMDB (server re-created).

    Output: objects (pipe to Out-GridView / Export-Csv) and an XLSX report.

.PARAMETER Name
    Check only these retired names.

.PARAMETER TcpPort
    Ports tried when ping fails. Default: 22, 135, 445, 3389, 5985.

.PARAMETER FreshDays
    A dynamic A record refreshed within this many days blocks automatic cleanup. Default: 30.

.PARAMETER AdLogonDays
    An enabled AD computer that logged on within this many days blocks automatic cleanup. Default: 60.

.PARAMETER SkipAd
    Do not check AD computer objects.

.PARAMETER ThrottleLimit
    Parallel probes. Default: 32.

.PARAMETER OutputPath
    Folder for the XLSX report. Default: Settings.ReportPath.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Get-RetiredDnsResidue.ps1 | Where-Object Verdict -ne 'Clean' | Out-GridView
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$Name,
    [int[]]$TcpPort = @(22, 135, 445, 3389, 5985),
    [ValidateRange(1, 365)][int]$FreshDays = 30,
    [ValidateRange(1, 3650)][int]$AdLogonDays = 60,
    [switch]$SkipAd,
    [ValidateRange(1, 128)][int]$ThrottleLimit = 32,
    [string]$OutputPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Get-RetiredDnsResidue'
$ipRx = '\b(?:\d{1,3}\.){3}\d{1,3}\b'

function Get-ShortName([string]$n) {
    $s = "$n".Trim().TrimEnd('.')
    if (-not $s -or $s -match "^$ipRx$") { return '' }
    return ($s -split '\.')[0].ToLower()
}
function Add-To([hashtable]$h, [string]$k, $v) {
    if (-not $h.ContainsKey($k)) { $h[$k] = New-Object System.Collections.ArrayList }
    [void]$h[$k].Add($v)
}
function Get-From([hashtable]$h, [string]$k) { if ($h.ContainsKey($k)) { return @($h[$k]) } else { return @() } }

# --- GLPI -------------------------------------------------------------------------------------------------
$glpi = Connect-Glpi -Config $config
try { $all = @(Get-CmdbServer -Session $glpi -Config $config) } finally { Disconnect-Glpi $glpi }
$retiredStatus = Get-CmdbStatusValue -Config $config -Kind Retired
$activeNames = @{}
foreach ($s in ($all | Where-Object { $_.Status -ne $retiredStatus })) { $n = Get-ShortName $s.Name; if ($n) { $activeNames[$n] = $s.Status } }
$glpiIpOwner = @{}
foreach ($s in ($all | Where-Object { $_.Status -ne $retiredStatus })) {
    foreach ($ip in [regex]::Matches("$($s.Ip)", $ipRx)) { Add-To $glpiIpOwner $ip.Value (Get-ShortName $s.Name) }
}
$targets = [ordered]@{}
foreach ($s in ($all | Where-Object { $_.Status -eq $retiredStatus })) {
    $n = Get-ShortName $s.Name
    if (-not $n -or ($Name -and $Name -notcontains $n)) { continue }
    if (-not $targets.Contains($n)) { $targets[$n] = [PSCustomObject]@{ Name = $n; IPs = @(); Extra = @(); RetireDate = $s.RetireDate; Owner = $s.Owner; Description = $s.Description; Records = 0 } }
    $t = $targets[$n]
    $t.Records++
    $t.IPs = @(@($t.IPs) + @([regex]::Matches("$($s.Ip)", $ipRx) | ForEach-Object Value) | Select-Object -Unique)
    $t.Extra = @(@($t.Extra) + @([regex]::Matches("$($s.Description)", $ipRx) | ForEach-Object Value) | Where-Object { $t.IPs -notcontains $_ } | Select-Object -Unique)
}
Write-Log "GLPI: $($targets.Count) retired name(s), $($activeNames.Count) active"

# --- DNS over LDAP ----------------------------------------------------------------------------------------
function Get-DnsLabelName([byte[]]$b, [int]$o) {
    # DNS_COUNT_NAME: length, label count, then length-prefixed labels
    $cnt = $b[$o + 1]; $p = $o + 2; $labels = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $cnt -and $p -lt $b.Length; $i++) {
        $l = $b[$p]; if ($l -eq 0) { break }
        [void]$labels.Add([Text.Encoding]::UTF8.GetString($b, $p + 1, $l)); $p += $l + 1
    }
    return ($labels -join '.')
}
$aByName = @{}; $aByIp = @{}; $ptrByIp = @{}; $ptrByTarget = @{}; $cnByTarget = @{}
$domainDn = ([ADSI]'LDAP://RootDSE').defaultNamingContext
if (-not $domainDn) { throw 'Cannot read RootDSE: run on a domain member.' }
$epoch = New-Object DateTime 1601, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
foreach ($part in 'DomainDnsZones', 'ForestDnsZones') {
    $ds = New-Object DirectoryServices.DirectorySearcher
    $ds.SearchRoot = New-Object DirectoryServices.DirectoryEntry("LDAP://DC=$part,$domainDn")
    $ds.Filter = '(objectClass=dnsNode)'
    $ds.PageSize = 1000
    [void]$ds.PropertiesToLoad.AddRange([string[]]@('name', 'dnsrecord', 'dnstombstoned', 'distinguishedname'))
    try { $res = $ds.FindAll() } catch { Write-Log "DNS partition $part not readable: $($_.Exception.Message)" -Level Warn; continue }
    foreach ($r in $res) {
        if ($r.Properties['dnstombstoned'].Count -and $r.Properties['dnstombstoned'][0]) { continue }
        if ("$($r.Properties['distinguishedname'][0])" -notmatch '^DC=(?:[^,\\]|\\.)*,DC=((?:[^,\\]|\\.)*),CN=MicrosoftDNS') { continue }
        $zone = $Matches[1] -replace '\\(.)', '$1'
        if ($zone -in 'RootDNSServers', '..TrustAnchors' -or $zone -like '_msdcs.*') { continue }
        $node = "$($r.Properties['name'][0])"
        if (-not $node -or $node -eq '@') { continue }
        $isRev = $zone -like '*.in-addr.arpa'
        foreach ($b in $r.Properties['dnsrecord']) {
            if ($b.Length -lt 26) { continue }
            $type = [BitConverter]::ToUInt16($b, 2)
            $ts = [BitConverter]::ToUInt32($b, 20)           # hours since 1601; 0 = static record
            $upd = if ($ts) { $epoch.AddHours($ts) } else { $null }
            if ($type -eq 1 -and -not $isRev -and $b.Length -ge 28) {
                $rec = [PSCustomObject]@{ Name = $node.ToLower(); Fqdn = "$node.$zone"; IP = "$($b[24]).$($b[25]).$($b[26]).$($b[27])"; Static = (-not $ts); Updated = $upd }
                Add-To $aByName $rec.Name $rec; Add-To $aByIp $rec.IP $rec
            }
            elseif ($type -eq 5 -and -not $isRev) {
                $target = Get-DnsLabelName $b 24
                Add-To $cnByTarget (Get-ShortName $target) ([PSCustomObject]@{ Alias = "$node.$zone"; Target = $target })
            }
            elseif ($type -eq 12 -and $isRev) {
                $labels = @(("$node." + ($zone -replace '\.in-addr\.arpa$', '')) -split '\.')
                if ($labels.Count -ne 4 -or @($labels | Where-Object { $_ -notmatch '^\d{1,3}$' }).Count) { continue }
                [Array]::Reverse($labels)
                $target = (Get-DnsLabelName $b 24).TrimEnd('.')
                $rec = [PSCustomObject]@{ IP = $labels -join '.'; Target = $target; Short = Get-ShortName $target }
                Add-To $ptrByIp $rec.IP $rec; Add-To $ptrByTarget $rec.Short $rec
            }
        }
    }
    $res.Dispose(); $ds.Dispose()
}
Write-Log "DNS: $($aByName.Count) A names, $($ptrByIp.Count) PTR addresses, $($cnByTarget.Count) CNAME targets"

# --- probes (parallel) ------------------------------------------------------------------------------------
$probeBlock = {
    param([string]$Ip, [int[]]$Ports)
    $alive = $false; $ttl = $null
    $pg = New-Object System.Net.NetworkInformation.Ping
    try {
        for ($i = 0; $i -lt 2 -and -not $alive; $i++) {
            try { $rp = $pg.Send($Ip, 800); if ($rp.Status -eq 'Success') { $alive = $true; if ($rp.Options) { $ttl = $rp.Options.Ttl } } } catch { }
        }
    }
    finally { $pg.Dispose() }
    $open = @()
    if (-not $alive) {
        foreach ($p in $Ports) {
            $c = New-Object System.Net.Sockets.TcpClient
            try { $ar = $c.BeginConnect($Ip, $p, $null, $null); if ($ar.AsyncWaitHandle.WaitOne(400) -and $c.Connected) { $open += $p } } catch { } finally { $c.Close() }
        }
    }
    [PSCustomObject]@{ IP = $Ip; Alive = ($alive -or $open.Count -gt 0); Ping = $alive; Ttl = $ttl; Ports = $open }
}
$ips = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($t in $targets.Values) {
    foreach ($a in (Get-From $aByName $t.Name)) { [void]$ips.Add($a.IP) }
    foreach ($p in (Get-From $ptrByTarget $t.Name)) { [void]$ips.Add($p.IP) }
    foreach ($ip in @($t.IPs) + @($t.Extra)) { if ($ip) { [void]$ips.Add($ip) } }
}
$probe = @{}
if ($ips.Count) {
    Write-Log "Probing $($ips.Count) address(es)..."
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit); $pool.Open()
    $jobs = foreach ($ip in $ips) {
        $ps = [PowerShell]::Create().AddScript($probeBlock).AddArgument($ip).AddArgument($TcpPort)
        $ps.RunspacePool = $pool
        @{ PS = $ps; Handle = $ps.BeginInvoke() }
    }
    foreach ($j in $jobs) { foreach ($r in $j.PS.EndInvoke($j.Handle)) { $probe[$r.IP] = $r }; $j.PS.Dispose() }
    $pool.Close(); $pool.Dispose()
}

function Get-AdComputer([string]$n) {
    $ds = New-Object DirectoryServices.DirectorySearcher
    $ds.Filter = "(&(objectCategory=computer)(sAMAccountName=$n`$))"
    [void]$ds.PropertiesToLoad.AddRange([string[]]@('useraccountcontrol', 'lastlogontimestamp', 'distinguishedname'))
    $f = $ds.FindOne(); $ds.Dispose()
    if (-not $f) { return $null }
    $ll = if ($f.Properties['lastlogontimestamp'].Count) { [DateTime]::FromFileTimeUtc([int64]$f.Properties['lastlogontimestamp'][0]) } else { $null }
    return [PSCustomObject]@{ Disabled = [bool]([int]$f.Properties['useraccountcontrol'][0] -band 2); LastLogon = $ll; Dn = "$($f.Properties['distinguishedname'][0])" }
}

# --- verdicts ---------------------------------------------------------------------------------------------
$now = [DateTime]::UtcNow
$results = foreach ($t in $targets.Values) {
    $A = Get-From $aByName $t.Name; $P = Get-From $ptrByTarget $t.Name; $C = Get-From $cnByTarget $t.Name
    if ($activeNames.ContainsKey($t.Name)) {
        [PSCustomObject]@{ Name = $t.Name; Verdict = 'Skip'; Actions = ''; Blockers = "name is active again in the CMDB ($($activeNames[$t.Name]))"; A = $A.Count; PTR = $P.Count; CNAME = $C.Count; AD = ''; Probe = ''; RetireDate = $t.RetireDate; Owner = $t.Owner }
        continue
    }
    $actions = New-Object System.Collections.Generic.List[string]
    $block = New-Object System.Collections.Generic.List[string]
    $reuse = $false
    $dnsOwners = { param($ip) @(Get-From $aByIp $ip | Where-Object { $_.Name -ne $t.Name } | ForEach-Object Name | Select-Object -Unique) }
    $glpiOwners = { param($ip) @(Get-From $glpiIpOwner $ip | Where-Object { $_ -ne $t.Name } | Select-Object -Unique) }
    foreach ($a in $A) {
        $alive = $probe[$a.IP] -and $probe[$a.IP].Alive
        $dOwn = & $dnsOwners $a.IP; $gOwn = & $glpiOwners $a.IP
        if (-not $a.Static -and $a.Updated -and ($now - $a.Updated).TotalDays -le $FreshDays) { $block.Add("A $($a.Fqdn) is dynamic and was refreshed $($a.Updated.ToString('yyyy-MM-dd')): the host still registers") }
        if ($alive) {
            if ($dOwn) { $reuse = $true }
            elseif ($gOwn) { $block.Add("$($a.IP) answers; the CMDB assigns it to $($gOwn -join ', ') but DNS does not") }
            else { $block.Add("$($a.IP) answers and has no other owner: the server may still run") }
        }
        $actions.Add("delete A $($a.Fqdn) -> $($a.IP)")
    }
    foreach ($p in $P) {
        $dOwn = & $dnsOwners $p.IP
        $otherPtr = @(Get-From $ptrByIp $p.IP | Where-Object { $_.Short -ne $t.Name })
        if (@($A | Where-Object IP -eq $p.IP).Count -eq 0 -and $probe[$p.IP] -and $probe[$p.IP].Alive -and -not $dOwn -and -not (& $glpiOwners $p.IP)) {
            $block.Add("$($p.IP) (PTR -> $($p.Target)) answers and has no other owner")
        }
        if ($otherPtr) { $actions.Add("delete PTR $($p.IP) -> $($p.Target) (other PTR exists)") }
        elseif ($dOwn) { $actions.Add("replace PTR $($p.IP) -> $($p.Target) with $($dOwn[0])") }
        else { $actions.Add("delete PTR $($p.IP) -> $($p.Target)") }
    }
    foreach ($c in $C) {
        if ($reuse) { $block.Add("CNAME $($c.Alias) now points to a live address and may be in use") }
        $actions.Add("delete CNAME $($c.Alias) -> $($c.Target)")
    }
    $adText = ''
    if (-not $SkipAd) {
        try {
            $ad = Get-AdComputer $t.Name
            if ($ad) {
                $adText = $(if ($ad.Disabled) { 'disabled' } else { 'ENABLED' }) + $(if ($ad.LastLogon) { ", logon $($ad.LastLogon.ToString('yyyy-MM-dd'))" })
                if (-not $ad.Disabled -and $ad.LastLogon -and ($now - $ad.LastLogon).TotalDays -le $AdLogonDays) { $block.Add('AD computer is enabled and logged on recently') }
                $actions.Add("remove AD computer $($ad.Dn)")
            }
            else { $adText = 'none' }
        }
        catch { $adText = "error: $($_.Exception.Message)" }
    }
    $probeText = (@($probe.Values | Where-Object { @($A.IP) + @($P.IP) + @($t.IPs) -contains $_.IP }) | ForEach-Object {
            if ($_.Ping) { "$($_.IP) ping" } elseif ($_.Ports) { "$($_.IP) tcp $($_.Ports -join ',')" } else { "$($_.IP) silent" } }) -join '; '
    $verdict = if ($block.Count) { 'Manual' } elseif (-not $actions.Count) { 'Clean' } elseif ($reuse) { 'Reuse' } else { 'Fix' }
    [PSCustomObject]@{
        Name = $t.Name; Verdict = $verdict; Actions = $actions -join "`n"; Blockers = $block -join "`n"
        A = $A.Count; PTR = $P.Count; CNAME = $C.Count; AD = $adText; Probe = $probeText; RetireDate = $t.RetireDate; Owner = $t.Owner
        __Highlight = ($verdict -eq 'Manual')
    }
}
$results = @($results | Sort-Object @{ e = { @('Manual', 'Reuse', 'Fix', 'Skip', 'Clean').IndexOf($_.Verdict) } }, Name)
$file = Join-Path $OutputPath ('retired-dns-residue_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
Export-Xlsx -Path $file -Sheets @(New-XlsxSheet -Name 'Retired servers' -Rows $results -Widths @{ Actions = 60; Blockers = 60; Probe = 30; AD = 26 })
Write-Log ("Verdicts: " + (($results | Group-Object Verdict | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', ') + "; report $file") -Level Ok
$results | Select-Object * -ExcludeProperty __Highlight
