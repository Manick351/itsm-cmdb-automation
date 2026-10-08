<#
.SYNOPSIS
    Generates a synthetic service desk export (requests with tasks) for demos and tests.

.DESCRIPTION
    No real data: services, titles and people are invented, volumes follow simple weights with a mild seasonal
    trend. The output matches the default field map of lib\Itsm.ps1 (id, number, title, service, state, team,
    assignee, created, closed, requestNumber, fields) and can be used with:

      tickets\Get-TicketWorkload.ps1 -InputPath <file>
      Itsm.Provider = 'File' for the ticket router and the CMDB lifecycle scripts (use -Open to get open tasks)

.PARAMETER Count
    Number of requests. Each gets 1-3 tasks. Default: 2000.

.PARAMETER Months
    Time span ending today. Default: 12.

.PARAMETER Open
    Make the newest 5% of tasks open and unassigned (for the router demo).

.PARAMETER Seed
    Random seed for reproducible output. Default: 42.

.PARAMETER Path
    Output JSON file.

.EXAMPLE
    .\New-SyntheticTickets.ps1 -Count 3000 -Path ..\state\tickets.json
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateRange(1, 100000)][int]$Count = 2000,
    [ValidateRange(1, 60)][int]$Months = 12,
    [switch]$Open,
    [int]$Seed = 42,
    [Parameter(Mandatory = $true)][string]$Path
)

$rnd = New-Object System.Random($Seed)
$people = 'alice', 'bob', 'carol', 'dave', 'erin', 'frank'
$catalog = @(
    @{ Service = 'Server provisioning'; Weight = 14; Tasks = @('Server Deployment', 'Add server to CMDB', 'Adding server to the backup scope', 'Add server to monitoring') }
    @{ Service = 'Server decommission'; Weight = 6; Tasks = @('Final backup', 'Shutdown of server', 'Removing server from backup', 'Disabling server in AD') }
    @{ Service = 'Access management'; Weight = 22; Tasks = @('Grant RDP access', 'Grant admin access', 'Revoke access') }
    @{ Service = 'Backup and restore'; Weight = 10; Tasks = @('Restore files', 'Restore VM', 'Change backup policy') }
    @{ Service = 'Snapshots'; Weight = 9; Tasks = @('Create Snapshot', 'Delete Snapshot') }
    @{ Service = 'Server changes'; Weight = 15; Tasks = @('Extend disk', 'Change CPU/RAM', 'Change criticality', 'Change business administrator') }
    @{ Service = 'Incidents'; Weight = 18; Tasks = @('Investigate alert', 'Service restart', 'Performance issue') }
    @{ Service = 'Certificates'; Weight = 6; Tasks = @('Renew certificate', 'Issue certificate') }
)
$weights = @($catalog | ForEach-Object { $_.Weight })
$sum = ($weights | Measure-Object -Sum).Sum
function Get-RandomService {
    $x = $rnd.Next(0, $sum); $acc = 0
    foreach ($c in $catalog) { $acc += $c.Weight; if ($x -lt $acc) { return $c } }
    return $catalog[-1]
}

$end = Get-Date
$start = $end.AddMonths(-$Months)
$spanDays = ($end - $start).TotalDays
$out = New-Object System.Collections.Generic.List[object]
$taskNo = 50000
for ($i = 1; $i -le $Count; $i++) {
    # later months slightly busier: skew the day towards the end
    $day = $spanDays * [Math]::Pow($rnd.NextDouble(), 0.8)
    $created = $start.AddDays($day).AddMinutes($rnd.Next(0, 600))
    $svc = Get-RandomService
    $reqNo = 'REQ-{0:D6}' -f (10000 + $i)
    $assignee = $people[$rnd.Next(0, $people.Count)]
    $closed = $created.AddHours($rnd.Next(2, 120))
    $server = 'srv-{0}-{1:D2}' -f @('app', 'db', 'web', 'file')[$rnd.Next(0, 4)], $rnd.Next(1, 60)
    $fields = [ordered]@{ 'Server name' = $server; 'IP' = "10.0.{0}.{1}" -f $rnd.Next(1, 30), $rnd.Next(2, 250); 'Operating System' = @('Windows Server 2022 Standard', 'Ubuntu Linux')[$rnd.Next(0, 2)]
        'Used for' = @('Production', 'Test', 'Development')[$rnd.Next(0, 3)]; 'Impact' = @('Low', 'Medium', 'High')[$rnd.Next(0, 3)]; 'Windows account' = "user$($rnd.Next(100, 999))" }
    $out.Add([PSCustomObject]@{ id = [guid]::NewGuid().ToString(); number = $reqNo; title = "$($svc.Service) request"; service = $svc.Service; state = 'closed'; team = 'server-team'
            assignee = $assignee; created = $created.ToString('s'); closed = $closed.ToString('s'); requestNumber = $null; fields = $fields })
    $n = $rnd.Next(1, 4)
    for ($k = 0; $k -lt $n; $k++) {
        $taskNo++
        $title = $svc.Tasks[$rnd.Next(0, $svc.Tasks.Count)]
        $tc = $created.AddMinutes(5 + 30 * $k)
        $out.Add([PSCustomObject]@{ id = [guid]::NewGuid().ToString(); number = "TSK-$taskNo"; title = $title; service = $null; state = 'closed'; team = 'server-team'
                assignee = $people[$rnd.Next(0, $people.Count)]; created = $tc.ToString('s'); closed = $tc.AddHours($rnd.Next(1, 72)).ToString('s'); requestNumber = $reqNo; fields = $fields })
    }
}
if ($Open) {
    $tasks = @($out | Where-Object { $_.requestNumber } | Sort-Object created -Descending)
    foreach ($t in ($tasks | Select-Object -First ([Math]::Max(1, [int]($tasks.Count * 0.05))))) { $t.state = 'registered'; $t.assignee = $null; $t.closed = $null }
}
ConvertTo-Json -InputObject $out.ToArray() -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8
Write-Host "Wrote $($out.Count) tickets ($Count requests) to $Path"
