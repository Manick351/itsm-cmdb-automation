<#
.SYNOPSIS
    Read-only audit of how hosts enter and leave Zabbix: configuration snapshot plus a list of hygiene problems.

.DESCRIPTION
    Collects into a timestamped folder (JSON + CSV):
      user and role of the API account, host groups with host counts, templates, autoregistration and network
      discovery rules/actions, proxies, autoregistration events of the last -AutoregDays days, housekeeping
      settings, and all hosts flattened (interfaces, groups, templates, tags, disable marker).

    Then checks every host against Zabbix.Audit in the config and writes one CSV per finding:
      DisabledWithoutMarker  disabled, but no "Disabled: <date>" line in the description (who and why?);
      DisabledStale          disabled longer than StaleDisabledDays (delete or re-enable);
      NoTemplate             nothing is monitored;
      NoInterface            nothing can be monitored;
      NameNotMatching        technical name outside NamePattern;
      MissingGroup           enabled host not in one of RequiredGroups;
      VisibleNameMismatch    visible name differs from the technical name;
      AgentUnavailable       enabled host whose agent interface is unavailable.

    Works with Zabbix 5.0 - 7.x (group property and proxy field names are chosen by API version).

.PARAMETER OutputPath
    Base folder; a subfolder per run is created. Default: Settings.ReportPath\zabbix-audit.

.PARAMETER AutoregDays
    Look-back for autoregistration events. Default: 90.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Test-ZabbixHostLifecycle.ps1
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputPath,
    [ValidateRange(1, 365)][int]$AutoregDays = 90,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Zabbix.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$audit = if ($config.Zabbix.Audit) { $config.Zabbix.Audit } else { @{} }
$staleDays = if ($audit.StaleDisabledDays) { [int]$audit.StaleDisabledDays } else { 30 }
if (-not $OutputPath) { $OutputPath = Join-Path $config.Settings.ReportPath 'zabbix-audit' }
$runDir = Join-Path $OutputPath ('{0:yyyyMMdd_HHmmss}' -f (Get-Date))
New-Item -ItemType Directory -Path $runDir -Force | Out-Null
$null = Start-ToolkitLog -LogPath $runDir -Name 'audit'

function Save-Json($Object, [string]$Name) { if ($null -ne $Object) { $Object | ConvertTo-Json -Depth 20 | Set-Content (Join-Path $runDir "$Name.json") -Encoding UTF8 } }
function Save-Csv($Object, [string]$Name) { if ($Object) { @($Object) | Export-Csv (Join-Path $runDir "$Name.csv") -NoTypeInformation -Encoding UTF8 } }
function Get-Safe([scriptblock]$Call, [string]$What) {
    try { & $Call } catch { Write-Log "  $What unavailable: $($_.Exception.Message)" -Level Warn; $null }
}

$zbx = Connect-Zabbix -Config $config
$summary = [ordered]@{ ApiVersion = $zbx.Version; Collected = (Get-Date).ToString('s') }
try {
    $grp = Get-ZabbixGroupParam $zbx
    $proxyProp = if ($zbx.VersionNumber -ge 700) { 'proxyid' } else { 'proxy_hostid' }
    Write-Log "Zabbix $($zbx.Version), output $runDir"

    Write-Log '[1/9] API account'
    $me = Get-Safe { Invoke-ZabbixApi -Session $zbx -Method 'user.checkAuthentication' -Params $(if ($zbx.Bearer) { @{ token = $zbx.Token } } else { @{ sessionid = $zbx.Token } }) } 'user.checkAuthentication'
    Save-Json $me '01_account'

    Write-Log '[2/9] Host groups'
    $groups = Invoke-ZabbixApi -Session $zbx -Method 'hostgroup.get' -Params @{ output = @('groupid', 'name'); selectHosts = 'count' }
    Save-Csv ($groups | Select-Object groupid, name, hosts | Sort-Object name) '02_hostgroups'
    $summary.HostGroups = @($groups).Count
    foreach ($g in @($audit.RequiredGroups)) { if ($g -and -not ($groups | Where-Object name -eq $g)) { Write-Log "  required group '$g' does not exist" -Level Warn } }

    Write-Log '[3/9] Templates'
    $templates = Invoke-ZabbixApi -Session $zbx -Method 'template.get' -Params @{ output = @('templateid', 'host', 'name') }
    Save-Csv ($templates | Sort-Object name) '03_templates'
    $summary.Templates = @($templates).Count

    Write-Log '[4/9] Autoregistration and discovery'
    $arActions = Invoke-ZabbixApi -Session $zbx -Method 'action.get' -Params @{ output = 'extend'; filter = @{ eventsource = 2 }; selectFilter = 'extend'; selectOperations = 'extend' }
    $dActions = Invoke-ZabbixApi -Session $zbx -Method 'action.get' -Params @{ output = 'extend'; filter = @{ eventsource = 1 }; selectFilter = 'extend'; selectOperations = 'extend' }
    $drules = Invoke-ZabbixApi -Session $zbx -Method 'drule.get' -Params @{ output = 'extend'; selectDChecks = 'extend' }
    Save-Json $arActions '04_actions_autoregistration'; Save-Json $dActions '04_actions_discovery'; Save-Json $drules '04_discovery_rules'
    $summary.AutoregActionsEnabled = @($arActions | Where-Object status -eq '0').Count
    $summary.DiscoveryRulesEnabled = @($drules | Where-Object status -eq '0').Count

    Write-Log '[5/9] Proxies'
    $proxies = Get-Safe { Invoke-ZabbixApi -Session $zbx -Method 'proxy.get' -Params @{ output = 'extend'; selectHosts = 'count' } } 'proxy.get'
    Save-Json $proxies '05_proxies'
    $summary.Proxies = @($proxies).Count

    Write-Log '[6/9] Autoregistration settings'
    Save-Json (Get-Safe { Invoke-ZabbixApi -Session $zbx -Method 'autoregistration.get' -Params @{ output = 'extend' } } 'autoregistration.get') '06_autoregistration'

    Write-Log "[7/9] Autoregistration events ($AutoregDays days)"
    $since = [int64](([DateTimeOffset](Get-Date).AddDays(-$AutoregDays)).ToUnixTimeSeconds())
    $events = Get-Safe { Invoke-ZabbixApi -Session $zbx -Method 'event.get' -Params @{ output = 'extend'; source = 2; object = 3; time_from = $since; sortfield = @('clock'); sortorder = 'DESC'; limit = 1000 } } 'event.get'
    Save-Json $events '07_autoreg_events'
    $summary.AutoregEvents = @($events).Count

    Write-Log '[8/9] Housekeeping'
    Save-Json (Get-Safe { Invoke-ZabbixApi -Session $zbx -Method 'housekeeping.get' -Params @{ output = 'extend' } } 'housekeeping.get') '08_housekeeping'

    Write-Log '[9/9] Hosts'
    $hostParams = @{
        output = @('hostid', 'host', 'name', 'status', 'description', 'flags', $proxyProp)
        selectInterfaces = @('type', 'main', 'useip', 'ip', 'dns', 'port', 'available')
        selectParentTemplates = @('templateid', 'name'); selectTags = 'extend'
    }
    $hostParams[$grp.Select] = @('groupid', 'name')
    $hosts = Invoke-ZabbixApi -Session $zbx -Method 'host.get' -Params $hostParams
    $flat = foreach ($h in $hosts) {
        $if = @($h.interfaces) | Where-Object { $_.main -eq '1' } | Select-Object -First 1
        $marker = @(([string]$h.description) -split "`n" | Where-Object { $_ -match 'Disabled:\s*(\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})' }) | Select-Object -Last 1
        $markerDate = $null
        if ($marker -and $marker -match '(\d{4}-\d{2}-\d{2}|\d{2}\.\d{2}\.\d{4})') {
            $markerDate = [datetime]::ParseExact($Matches[1], [string[]]@('yyyy-MM-dd', 'dd.MM.yyyy'), [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None)
        }
        # Zabbix 7 uses "available" on the interface, older agents report it the same way (0 unknown, 1 up, 2 down)
        [PSCustomObject]@{
            hostid = $h.hostid; host = $h.host; visible_name = $h.name
            status = if ($h.status -eq '0') { 'Enabled' } else { 'Disabled' }
            discovered = ($h.flags -eq '4')
            proxy = $h.$proxyProp
            iface_type = $if.type; iface_ip = $if.ip; iface_dns = $if.dns; iface_available = $if.available
            iface_count = @($h.interfaces).Count
            groups = (@($h.($grp.Property)) | ForEach-Object name) -join '; '
            template_count = @($h.parentTemplates).Count
            templates = (@($h.parentTemplates) | ForEach-Object name) -join '; '
            tags = (@($h.tags) | ForEach-Object { "$($_.tag)=$($_.value)" }) -join '; '
            disabled_marker = $marker
            days_disabled = if ($markerDate) { [int]((Get-Date) - $markerDate).TotalDays } else { $null }
            description = ([string]$h.description) -replace "`r?`n", ' | '
        }
    }
    $flat = @($flat | Sort-Object host)
    Save-Csv $flat '09_hosts'
    $summary.Hosts = $flat.Count

    $enabled = @($flat | Where-Object status -eq 'Enabled')
    $disabled = @($flat | Where-Object status -eq 'Disabled')
    $findings = [ordered]@{
        DisabledWithoutMarker = @($disabled | Where-Object { -not $_.disabled_marker })
        DisabledStale         = @($disabled | Where-Object { $_.days_disabled -gt $staleDays })
        NoTemplate            = @($flat | Where-Object { $_.template_count -eq 0 })
        NoInterface           = @($flat | Where-Object { $_.iface_count -eq 0 })
        NameNotMatching       = @(if ($audit.NamePattern) { $flat | Where-Object { $_.host -notmatch $audit.NamePattern } })
        MissingGroup          = @(foreach ($h in $enabled) { $miss = @($audit.RequiredGroups | Where-Object { $_ -and ("; $($h.groups);" -notlike "*; $_;*") }); if ($miss) { $h | Select-Object *, @{ n = 'missing_groups'; e = { $miss -join ', ' } } } })
        VisibleNameMismatch   = @($flat | Where-Object { $_.host -ne $_.visible_name })
        AgentUnavailable      = @($enabled | Where-Object { $_.iface_type -eq '1' -and $_.iface_available -eq '2' })
    }
    Write-Log ''
    foreach ($k in $findings.Keys) {
        $summary[$k] = $findings[$k].Count
        Save-Csv $findings[$k] "A_$k"
        Write-Log ('{0,-24} {1}' -f $k, $findings[$k].Count) -Level $(if ($findings[$k].Count) { 'Warn' } else { 'Ok' })
    }
    $summary | ConvertTo-Json | Set-Content (Join-Path $runDir '00_summary.json') -Encoding UTF8
    Write-Log "Done: $runDir" -Level Ok
    [PSCustomObject]$summary
}
finally { Disconnect-Zabbix $zbx }
