<#
.SYNOPSIS
    Workload analytics for a support team: what the queue consists of, how it changes month by month, who does what.

.DESCRIPTION
    Built to answer "which work can be automated or handed to another team": closed tasks and requests of a team
    are grouped by service and by task title, with share of the total, examples and assignees.

    Data comes from the ITSM adapter (lib\Itsm.ps1) or from a JSON export (-InputPath). Fetched data is cached
    per day in Settings.StatePath, so re-running the analysis does not load the service desk again;
    -Offline uses the newest cache only.

    For a task the "effective service" is the service of its parent request (tasks themselves usually have
    none), which keeps tasks and requests comparable.

    Workbook sheets:
      By service   count, share, tasks / requests per service;
      Detailed     per (kind, service, title): count, share, top assignees, last seen, sample numbers;
      By month     totals per month (trend);
      By assignee  per person: count and top services;
      Raw          one row per ticket.

.PARAMETER InputPath
    JSON file with tickets instead of the service desk (for example made by examples\New-SyntheticTickets.ps1).

.PARAMETER Since
    Only tickets created on or after this date.

.PARAMETER Offline
    Use the newest cache, do not call the service desk.

.PARAMETER OutputPath
    Folder for the workbook. Default: Settings.ReportPath.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    ..\examples\New-SyntheticTickets.ps1 -Count 5000 -Path .\tickets.json; .\Get-TicketWorkload.ps1 -InputPath .\tickets.json

.EXAMPLE
    .\Get-TicketWorkload.ps1 -Since 2026-01-01
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InputPath,
    [datetime]$Since,
    [switch]$Offline,
    [string]$OutputPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Itsm.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
foreach ($d in $OutputPath, $config.Settings.StatePath) { if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null } }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Get-TicketWorkload'
$wl = if ($config.Workload) { $config.Workload } else { @{} }

# --- load ---------------------------------------------------------------------------------------------------
if ($InputPath) {
    $raw = @(Get-Content -LiteralPath $InputPath -Raw -Encoding UTF8 | ConvertFrom-Json)
    $tickets = @(foreach ($r in $raw) { ConvertTo-ItsmTicket $r $config.Itsm.FieldMap })
    Write-Log "Loaded $($tickets.Count) ticket(s) from $InputPath"
}
else {
    $cacheFile = Join-Path $config.Settings.StatePath ('workload_{0:yyyyMMdd}.json' -f (Get-Date))
    $newest = Get-ChildItem -LiteralPath $config.Settings.StatePath -Filter 'workload_*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 1
    if ($Offline -or (Test-Path -LiteralPath $cacheFile)) {
        if (-not $newest) { throw 'No cache yet: run once without -Offline.' }
        $raw = @(Get-Content -LiteralPath $newest.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
        Write-Log "Cache $($newest.Name): $($raw.Count) ticket(s)"
    }
    else {
        $itsm = Connect-Itsm -Config $config
        $states = if ($wl.States) { @($wl.States) } else { @('closed', 'resolved') }
        $teams = if ($wl.Teams) { @($wl.Teams) } else { @('') }
        $raw = @(Get-ItsmTicket -Session $itsm -State $states -Team $teams | ForEach-Object Raw)
        ConvertTo-Json -InputObject $raw -Depth 20 | Set-Content -LiteralPath $cacheFile -Encoding UTF8
        Write-Log "Fetched $($raw.Count) ticket(s), cached to $cacheFile"
    }
    $tickets = @(foreach ($r in $raw) { ConvertTo-ItsmTicket $r $config.Itsm.FieldMap })
}
if ($PSBoundParameters.ContainsKey('Since')) { $tickets = @($tickets | Where-Object { $_.Created -and $_.Created -ge $Since }) }
if (-not $tickets) { throw 'No tickets to analyse.' }

# Parent requests are in the same data set (requests and tasks of the team); others keep their own service.
$requests = @{}
foreach ($t in $tickets) { if (-not $t.RequestNumber -and $t.Number) { $requests[$t.Number] = $t } }

$rows = foreach ($t in $tickets) {
    $kind = if ($t.RequestNumber) { 'Task' } else { 'Request' }
    $parent = if ($kind -eq 'Task') { $requests[$t.RequestNumber] } else { $null }
    $service = if ($t.Service) { $t.Service } elseif ($parent) { $parent.Service } else { '(unknown)' }
    [PSCustomObject]@{
        Kind = $kind; Number = $t.Number; Title = $t.Title; Service = $service; Request = $t.RequestNumber
        State = $t.State; Assignee = $(if ($t.Assignee) { $t.Assignee } else { '(none)' }); Created = $t.Created
        Month = if ($t.Created) { $t.Created.ToString('yyyy-MM') } else { '' }
    }
}
$rows = @($rows | Sort-Object Created -Descending)
$total = $rows.Count
$pct = { param($n) [Math]::Round($n * 100.0 / $total, 1) }

$byService = foreach ($g in ($rows | Group-Object Service | Sort-Object Count -Descending)) {
    [PSCustomObject]@{ Service = $g.Name; Count = $g.Count; 'Share %' = & $pct $g.Count
        Tasks = @($g.Group | Where-Object Kind -eq 'Task').Count; Requests = @($g.Group | Where-Object Kind -eq 'Request').Count; Direction = '' }
}
$detailed = foreach ($g in ($rows | Group-Object Kind, Service, Title | Sort-Object Count -Descending)) {
    $f = $g.Group[0]
    $top = ($g.Group | Group-Object Assignee | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
    [PSCustomObject]@{ Kind = $f.Kind; Service = $f.Service; Title = $f.Title; Count = $g.Count; 'Share %' = & $pct $g.Count
        'Top assignees' = $top; 'Last seen' = ($g.Group | Sort-Object Created -Descending | Select-Object -First 1).Created
        Samples = (@($g.Group | Select-Object -First 3 | ForEach-Object Number) -join ', '); Direction = '' }
}
$byMonth = foreach ($g in ($rows | Where-Object Month | Group-Object Month | Sort-Object Name)) {
    [PSCustomObject]@{ Month = $g.Name; Total = $g.Count; Tasks = @($g.Group | Where-Object Kind -eq 'Task').Count; Requests = @($g.Group | Where-Object Kind -eq 'Request').Count }
}
$byAssignee = foreach ($g in ($rows | Group-Object Assignee | Sort-Object Count -Descending)) {
    [PSCustomObject]@{ Assignee = $g.Name; Count = $g.Count; 'Share %' = & $pct $g.Count
        'Top services' = ($g.Group | Group-Object Service | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ' }
}

$file = Join-Path $OutputPath ('ticket-workload_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
Export-Xlsx -Path $file -Sheets @(
    (New-XlsxSheet -Name 'By service' -Rows @($byService) -Widths @{ Service = 36; Direction = 20 })
    (New-XlsxSheet -Name 'Detailed' -Rows @($detailed) -Widths @{ Service = 30; Title = 40; 'Top assignees' = 40; Samples = 30; 'Last seen' = 17 })
    (New-XlsxSheet -Name 'By month' -Rows @($byMonth))
    (New-XlsxSheet -Name 'By assignee' -Rows @($byAssignee) -Widths @{ 'Top services' = 60 })
    (New-XlsxSheet -Name 'Raw' -Rows $rows -Widths @{ Title = 40; Service = 30; Created = 17 })
)
$oldest = ($rows | Where-Object Created | Select-Object -Last 1).Created
Write-Log ("{0} ticket(s) since {1:yyyy-MM-dd}, {2} service(s), {3} distinct task types; report {4}" -f $total, $oldest, @($byService).Count, @($detailed).Count, $file) -Level Ok
$file
