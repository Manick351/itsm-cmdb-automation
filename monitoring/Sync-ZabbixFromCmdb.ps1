<#
.SYNOPSIS
    Creates Zabbix hosts for active CMDB servers that are not monitored yet.

.DESCRIPTION
    Compares active servers in the GLPI CMDB with Zabbix hosts (by technical name, case-insensitive). For each
    missing server the first matching rule of Zabbix.HostRules decides how it is monitored:

      @{ When = @{ Os = 'Windows' }; Groups = @('Windows servers'); Templates = @('Windows by Zabbix agent'); Tags = @{ os_family = 'Windows' } }

    When uses logical CMDB fields with regex values (see lib\Glpi.ps1). Zabbix.CommonGroups are added to every
    host. Groups and templates are referenced by NAME and resolved through the API, so the config is portable
    between Zabbix instances. The agent interface uses the CMDB IP and FQDN (port Zabbix.AgentPort, 10050).

    Servers matching Zabbix.SkipPattern (appliances, hosts monitored by other means) and servers without an IP
    are reported, not created. Without -Execute nothing is created.

.PARAMETER Name
    Only these servers.

.PARAMETER Execute
    Create the hosts.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Sync-ZabbixFromCmdb.ps1

.EXAMPLE
    .\Sync-ZabbixFromCmdb.ps1 -Name srv-app-07 -Execute
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string[]]$Name,
    [switch]$Execute,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')
. (Join-Path $PSScriptRoot '..\lib\Zabbix.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$z = $config.Zabbix
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Sync-ZabbixFromCmdb'
if (-not $Execute) { Write-Log 'Dry run: no hosts are created (use -Execute).' -Level Warn }

$glpi = Connect-Glpi -Config $config
try {
    $active = Get-CmdbStatusValue -Config $config -Kind Active
    $servers = @(Get-CmdbServer -Session $glpi -Config $config | Where-Object { $_.Status -eq $active })
}
finally { Disconnect-Glpi $glpi }
if ($Name) { $servers = @($servers | Where-Object { $Name -contains $_.Name }) }

$zbx = Connect-Zabbix -Config $config
$results = New-Object System.Collections.Generic.List[object]
try {
    $existing = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($h in (Invoke-ZabbixApi -Session $zbx -Method 'host.get' -Params @{ output = @('host') })) { [void]$existing.Add($h.host) }
    $missing = @($servers | Where-Object { -not $existing.Contains([string]$_.Name) })
    Write-Log "CMDB active: $($servers.Count); not in Zabbix: $($missing.Count)"

    $groupIds = @{}; $templateIds = @{}
    function Get-GroupId([string]$n) {
        if (-not $groupIds.ContainsKey($n)) {
            $g = @(Invoke-ZabbixApi -Session $zbx -Method 'hostgroup.get' -Params @{ output = @('groupid'); filter = @{ name = @($n) } })
            $groupIds[$n] = if ($g) { $g[0].groupid } else { $null }
        }
        return $groupIds[$n]
    }
    function Get-TemplateId([string]$n) {
        if (-not $templateIds.ContainsKey($n)) {
            $t = @(Invoke-ZabbixApi -Session $zbx -Method 'template.get' -Params @{ output = @('templateid'); filter = @{ name = @($n) } })
            if (-not $t) { $t = @(Invoke-ZabbixApi -Session $zbx -Method 'template.get' -Params @{ output = @('templateid'); filter = @{ host = @($n) } }) }
            $templateIds[$n] = if ($t) { $t[0].templateid } else { $null }
        }
        return $templateIds[$n]
    }

    foreach ($s in $missing) {
        $row = [ordered]@{ Server = $s.Name; Ip = $s.Ip; Rule = ''; Result = ''; Detail = '' }
        try {
            if ($z.SkipPattern -and $s.Name -match $z.SkipPattern) { $row.Result = 'Skipped'; $row.Detail = 'matches Zabbix.SkipPattern'; continue }
            $ip = [regex]::Match([string]$s.Ip, '\b(?:\d{1,3}\.){3}\d{1,3}\b').Value
            if (-not $ip) { $row.Result = 'Skipped'; $row.Detail = 'no IP address in the CMDB'; continue }
            $record = [ordered]@{}
            foreach ($p in $s.PSObject.Properties) { $record[$p.Name] = $p.Value }
            $rule = @($z.HostRules) | Where-Object { $_ -and (Test-RuleCondition -When $(if ($_.When) { $_.When } else { @{} }) -Record $record) } | Select-Object -First 1
            if (-not $rule) { $row.Result = 'NoRule'; $row.Detail = 'no Zabbix.HostRules entry matches'; continue }
            $row.Rule = if ($rule.Name) { $rule.Name } else { ($rule.Templates -join ', ') }
            $groupNames = @(@($rule.Groups) + @($z.CommonGroups) | Where-Object { $_ } | Select-Object -Unique)
            $groups = foreach ($g in $groupNames) { $id = Get-GroupId $g; if (-not $id) { throw "host group '$g' not found" }; @{ groupid = $id } }
            $templates = foreach ($t in @($rule.Templates | Where-Object { $_ })) { $id = Get-TemplateId $t; if (-not $id) { throw "template '$t' not found" }; @{ templateid = $id } }
            $tags = foreach ($k in @($rule.Tags.Keys)) { @{ tag = $k; value = [string]$rule.Tags[$k] } }
            $dns = if ($s.Fqdn) { [string]$s.Fqdn } else { '' }
            $params = @{
                host = $s.Name
                description = (@("Created from CMDB $((Get-Date).ToString('yyyy-MM-dd'))", [string]$s.Description) | Where-Object { $_ }) -join "`n"
                interfaces = @(@{ type = 1; main = 1; useip = 1; ip = $ip; dns = $dns; port = [string]$(if ($z.AgentPort) { $z.AgentPort } else { 10050 }) })
                groups = @($groups); templates = @($templates); tags = @($tags)
            }
            $row.Detail = "groups: $($groupNames -join ', '); templates: $(@($rule.Templates) -join ', ')"
            if (-not $Execute) { $row.Result = 'WouldCreate'; continue }
            $r = Invoke-ZabbixApi -Session $zbx -Method 'host.create' -Params $params
            $row.Result = 'Created'; $row.Detail = "hostid $($r.hostids[0]); $($row.Detail)"
        }
        catch { $row.Result = 'Error'; $row.Detail = $_.Exception.Message }
        finally {
            $results.Add([PSCustomObject]$row)
            Write-Log ("{0}: {1} {2}" -f $row.Server, $row.Result, $row.Detail) -Level $(if ($row.Result -eq 'Error') { 'Error' } elseif ($row.Result -in 'Created', 'WouldCreate') { 'Info' } else { 'Warn' })
        }
    }
}
finally { Disconnect-Zabbix $zbx }
$results
if (@($results | Where-Object { $_.Result -eq 'Error' })) { exit 1 }
