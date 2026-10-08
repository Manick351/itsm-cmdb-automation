<#
.SYNOPSIS
    Closes the loop between the service desk and the CMDB: closed deployment tasks create server records,
    closed decommission tasks retire them (and optionally disable the host in Zabbix).

.DESCRIPTION
    Reads tasks closed in the last -Days from the ITSM adapter (lib\Itsm.ps1) and:

      Commission  tasks whose title matches Lifecycle.CommissionTitle -> cmdb\Add-GlpiServer.ps1 in batch
                  mode with the parent request's form fields (same validation, rules and dry run);
      Decommission tasks whose title matches Lifecycle.DecommissionTitle -> the server named in the field
                  Lifecycle.DecommissionNameField is set to the Retired status with the retirement date (the
                  task's closing date) and "Decommission request: <number>" appended to the comment;
                  with -DisableMonitoring the Zabbix host is disabled and marked "Disabled: <date> <request>"
                  in its description (monitoring\Test-ZabbixHostLifecycle.ps1 reads this marker).

    Request numbers in Lifecycle.IgnoreRequests are skipped (test requests, known exceptions). Without
    -Execute nothing is changed anywhere.

.PARAMETER Stage
    Commission, Decommission or both (default).

.PARAMETER Days
    How far back to look for closed tasks. Default: 30.

.PARAMETER DisableMonitoring
    Also disable decommissioned hosts in Zabbix.

.PARAMETER Execute
    Apply the changes.

.PARAMETER SendMail
    Mail a summary when something was done or failed.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Invoke-CmdbLifecycle.ps1

.EXAMPLE
    .\Invoke-CmdbLifecycle.ps1 -Stage Decommission -DisableMonitoring -Execute -SendMail
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Commission', 'Decommission')][string[]]$Stage = @('Commission', 'Decommission'),
    [ValidateRange(1, 365)][int]$Days = 30,
    [switch]$DisableMonitoring,
    [switch]$Execute,
    [switch]$SendMail,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')
. (Join-Path $PSScriptRoot '..\lib\Itsm.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$lc = $config.Lifecycle
$log = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Invoke-CmdbLifecycle'
if (-not $Execute) { Write-Log 'Dry run: nothing is changed (use -Execute).' -Level Warn }
$since = (Get-Date).Date.AddDays(-$Days)
$ignore = @($lc.IgnoreRequests)
$teams = if ($lc.Teams) { @($lc.Teams) } else { @('') }
$closedStates = if ($lc.ClosedStates) { @($lc.ClosedStates) } else { @('closed') }

$itsm = Connect-Itsm -Config $config
$results = New-Object System.Collections.Generic.List[object]

if ($Stage -contains 'Commission') {
    # several deployment tasks of one request describe the same server: keep one per request
    $tasks = @(Get-ItsmTicket -Session $itsm -State $closedStates -Team $teams -TitlePattern $lc.CommissionTitle -ClosedAfter $since |
        Where-Object { $ignore -notcontains $_.RequestNumber } | Group-Object RequestNumber | ForEach-Object { $_.Group[0] })
    Write-Log "Commission: $($tasks.Count) request(s) with closed deployment tasks since $($since.ToString('yyyy-MM-dd'))"
    if ($tasks) {
        $batch = foreach ($t in $tasks) { [PSCustomObject]@{ Fields = $t.Fields; Ticket = $t.RequestNumber } }
        $add = @(& (Join-Path $PSScriptRoot 'Add-GlpiServer.ps1') -Batch $batch -Execute:$Execute -ConfigPath $ConfigPath)
        foreach ($r in $add) {
            $results.Add([PSCustomObject]@{ Stage = 'Commission'; Request = $r.Ticket; Server = $r.Name; Result = $r.Status; Detail = (@($r.Message) + @($r.Warnings) | Where-Object { $_ }) -join '; ' })
        }
    }
}

if ($Stage -contains 'Decommission') {
    $tasks = @(Get-ItsmTicket -Session $itsm -State $closedStates -Team $teams -TitlePattern $lc.DecommissionTitle -ClosedAfter $since |
        Where-Object { $ignore -notcontains $_.RequestNumber })
    Write-Log "Decommission: $($tasks.Count) closed task(s) since $($since.ToString('yyyy-MM-dd'))"
    $glpi = $null; $zbx = $null
    try {
        if ($tasks) {
            $glpi = Connect-Glpi -Config $config
            $servers = @(Get-CmdbServer -Session $glpi -Config $config)
            $retired = Get-CmdbStatusValue -Config $config -Kind Retired
            $protected = @($lc.ProtectedServers)
            if ($DisableMonitoring) { . (Join-Path $PSScriptRoot '..\lib\Zabbix.ps1'); $zbx = Connect-Zabbix -Config $config }
        }
        foreach ($t in $tasks) {
            $nameField = if ($lc.DecommissionNameField) { $lc.DecommissionNameField } else { 'Server name' }
            $name = ([string]$t.Fields[$nameField]).Trim()
            $row = [ordered]@{ Stage = 'Decommission'; Request = $t.RequestNumber; Server = $name; Result = ''; Detail = '' }
            try {
                if (-not $name -or $name -match '[\s.]') { throw "field '$nameField' does not hold a short server name: '$name'" }
                if ($protected -contains $name) { $row.Result = 'Skipped'; $row.Detail = 'listed in Lifecycle.ProtectedServers'; continue }
                $s = @($servers | Where-Object { $_.Name -eq $name })
                if (-not $s) { $row.Result = 'NotInCmdb'; $row.Detail = 'no CMDB record with this name'; continue }
                $live = @($s | Where-Object { $_.Status -ne $retired })
                $actions = New-Object System.Collections.Generic.List[string]
                if ($live) {
                    $date = $(if ($t.Closed) { $t.Closed } else { Get-Date }).ToString('yyyy-MM-dd')
                    $comment = (@(([string]$live[0].Comment).Trim(), "Decommission request: $($t.RequestNumber)") | Where-Object { $_ }) -join "`n"
                    $actions.Add("CMDB: $($live[0].Status) -> $retired, retired $date")
                    if ($Execute) {
                        $in = ConvertTo-CmdbInput -Session $glpi -Config $config -Values ([ordered]@{ Status = $retired; RetireDate = $date; Comment = $comment })
                        Set-CmdbServer -Session $glpi -Config $config -Server $live[0] -CmdbInput $in
                    }
                    if ($live.Count -gt 1) { $actions.Add("WARNING: $($live.Count) active records, only the first was retired") }
                }
                if ($zbx) {
                    $h = @(Invoke-ZabbixApi -Session $zbx -Method 'host.get' -Params @{ output = @('hostid', 'host', 'status', 'description'); filter = @{ host = @($name) } })
                    if ($h -and $h[0].status -eq '0') {
                        $actions.Add('Zabbix: host disabled')
                        if ($Execute) {
                            $desc = (@(([string]$h[0].description).Trim(), "Disabled: $((Get-Date).ToString('yyyy-MM-dd')) $($t.RequestNumber)") | Where-Object { $_ }) -join "`n"
                            $null = Invoke-ZabbixApi -Session $zbx -Method 'host.update' -Params @{ hostid = $h[0].hostid; status = 1; description = $desc }
                        }
                    }
                    elseif (-not $h) { $actions.Add('Zabbix: no host') }
                }
                if (-not $actions.Count) { $row.Result = 'AlreadyDone'; continue }
                $row.Result = if ($Execute) { 'Retired' } else { 'WouldRetire' }
                $row.Detail = $actions -join '; '
            }
            catch { $row.Result = 'Error'; $row.Detail = Get-RestErrorText $_ }
            finally {
                $results.Add([PSCustomObject]$row)
                Write-Log ("{0} {1}: {2} {3}" -f $row.Request, $row.Server, $row.Result, $row.Detail) -Level $(if ($row.Result -eq 'Error') { 'Error' } elseif ($row.Result -in 'NotInCmdb', 'Skipped') { 'Warn' } else { 'Info' })
            }
        }
    }
    finally {
        Disconnect-Glpi $glpi
        if ($zbx) { Disconnect-Zabbix $zbx }
    }
}

$results | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
$notable = @($results | Where-Object { $_.Result -notin 'AlreadyExists', 'AlreadyDone' })
if ($SendMail -and $notable) {
    $errors = @($notable | Where-Object Result -eq 'Error').Count
    $subject = 'CMDB lifecycle: ' + (($notable | Group-Object Result | ForEach-Object { "$($_.Name) $($_.Count)" }) -join ', ')
    $body = '<p>Changes from closed service desk tasks' + $(if (-not $Execute) { ' (dry run)' }) + ':</p>' + (ConvertTo-HtmlTable -Rows $notable)
    Send-ToolkitMail -Config $config -Subject $subject -Body $body -BodyAsHtml -Attachment $(if ($errors) { $log })
}
$results
