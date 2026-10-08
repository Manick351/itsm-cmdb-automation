<#
.SYNOPSIS
    Routes open service desk tasks: automates the ones a script can do, assigns the rest to people in turn.

.DESCRIPTION
    A pattern for "L2 queue automation", independent of the service desk product (lib\Itsm.ps1):

      1. read unassigned tasks of Router.Teams in Router.States;
      2. for each task find the first route in Router.Routes whose When matches (regex per ticket property:
         Title, Service, Number, or Field:<form field name>);
      3. run the route's handler from tickets\handlers\<Handler>.ps1 with the route's Params;
      4. tasks without a route go to Router.Default (round-robin by default).

    Built-in handlers:
      RoundRobin    next person from Router.Assignees, skipping Router.ExcludedAssignees and people absent
                    today (Router.AbsencePath CSV: Name,From,To); the position is kept in a state file with an
                    exclusive lock, so parallel runs do not hand out the same person twice;
      Assign        a fixed person or group (Params.To);
      AddToAdGroup  adds the account from a form field to an AD group built from a template, comments, closes;
                    falls back to round-robin with a note on any problem;
      AddToCmdb     cmdb\Add-GlpiServer.ps1 with the request fields; closes on success, returns the task to the
                    request owner with the reason otherwise;
      Close         closes with a fixed note (informational tasks).

    A handler is a script returning @{ Result = '...'; Note = '...' }; see tickets\handlers\README.md.
    Without -Execute every decision is logged and nothing is written to the service desk or AD.

.PARAMETER Execute
    Perform the actions.

.PARAMETER MaxTickets
    Upper limit of tasks per run. Default: Router.MaxTickets or 100.

.PARAMETER SendMail
    Mail the list of handled tasks with their result.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Invoke-TicketRouter.ps1

.EXAMPLE
    .\Invoke-TicketRouter.ps1 -Execute -SendMail     # every 5 minutes as a scheduled task
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$Execute,
    [ValidateRange(1, 1000)][int]$MaxTickets,
    [switch]$SendMail,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Itsm.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$router = $config.Router
if (-not $router.Assignees) { throw 'Router.Assignees is empty.' }
if (-not $MaxTickets) { $MaxTickets = if ($router.MaxTickets) { [int]$router.MaxTickets } else { 100 } }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Invoke-TicketRouter'
if (-not $Execute) { Write-Log 'Dry run: decisions are logged, nothing is changed (use -Execute).' -Level Warn }
if (-not (Test-Path -LiteralPath $config.Settings.StatePath)) { New-Item -ItemType Directory -Path $config.Settings.StatePath -Force | Out-Null }

#region Shared helpers for handlers
# Handlers run as child scripts, so these helpers use uniquely named router variables ($RouterItsm,
# $RouterAbsent, $RouterDryRun) found through the call stack; $script: would point at the handler's scope.
$RouterDryRun = @{ Last = $null }

function Get-AbsentToday {
    if (-not $router.AbsencePath) { return @() }
    $path = $router.AbsencePath
    if (-not [IO.Path]::IsPathRooted($path)) { $path = Join-Path (Split-Path -Parent $PSScriptRoot) $path }
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    $today = (Get-Date).Date
    @(Import-Csv -LiteralPath $path | Where-Object { $_.Name -and [datetime]$_.From -le $today -and [datetime]$_.To -ge $today } | ForEach-Object Name)
}

function Get-NextAssignee {
    <#
        Round-robin over the available people. The last position lives in state\rotation.json; the file is
        opened exclusively (retry for up to 30 s) so that two router runs never read the same position.
    #>
    param([string[]]$Pool = $router.Assignees)
    $available = @($Pool | Where-Object { $router.ExcludedAssignees -notcontains $_ -and $RouterAbsent -notcontains $_ })
    if (-not $available) { throw 'nobody is available (all excluded or absent)' }
    $file = Join-Path $config.Settings.StatePath 'rotation.json'
    if (-not $Execute) {
        # dry run: rotate in memory, starting from the saved position, and never write the state file
        if ($null -eq $RouterDryRun.Last) { $RouterDryRun.Last = if (Test-Path -LiteralPath $file) { (Get-Content -LiteralPath $file -Raw | ConvertFrom-Json).Last } else { '' } }
        $i = [array]::IndexOf($Pool, $RouterDryRun.Last)
        for ($k = 1; $k -le $Pool.Count; $k++) { $cand = $Pool[($i + $k) % $Pool.Count]; if ($available -contains $cand) { $RouterDryRun.Last = $cand; return $cand } }
    }
    $deadline = (Get-Date).AddSeconds(30)
    while ($true) {
        try { $fs = [IO.File]::Open($file, 'OpenOrCreate', 'ReadWrite', 'None'); break }
        catch [IO.IOException] { if ((Get-Date) -gt $deadline) { throw "rotation file is locked: $file" }; Start-Sleep -Milliseconds 300 }
    }
    try {
        $reader = New-Object IO.StreamReader($fs)
        $text = $reader.ReadToEnd()
        $last = if ($text) { ($text | ConvertFrom-Json).Last } else { '' }
        $i = [array]::IndexOf($Pool, $last)
        # walk the full pool from the last person, so that the order is stable when someone is absent
        $next = $null
        for ($k = 1; $k -le $Pool.Count; $k++) {
            $cand = $Pool[($i + $k) % $Pool.Count]
            if ($available -contains $cand) { $next = $cand; break }
        }
        $fs.SetLength(0)
        $writer = New-Object IO.StreamWriter($fs)
        $writer.Write((@{ Last = $next; Time = (Get-Date).ToString('s') } | ConvertTo-Json -Compress))
        $writer.Flush()
        return $next
    }
    finally { $fs.Dispose() }
}

function Set-TicketAssignee($Ticket, [string]$Assignee, [string]$Note) {
    Write-Log "  $($Ticket.Number): assign to $Assignee$(if ($Note) { " ($Note)" })"
    if (-not $Execute) { return }
    Invoke-ItsmAction -Session $RouterItsm -Action Assign -Ticket $Ticket -Values @{ assignee = $Assignee }
    if ($Note) { Invoke-ItsmAction -Session $RouterItsm -Action Comment -Ticket $Ticket -Values @{ text = $Note } }
}

function Add-TicketComment($Ticket, [string]$Text) {
    Write-Log "  $($Ticket.Number): comment '$Text'"
    if ($Execute) { Invoke-ItsmAction -Session $RouterItsm -Action Comment -Ticket $Ticket -Values @{ text = $Text } }
}

function Close-Ticket($Ticket, [string]$Note) {
    Write-Log "  $($Ticket.Number): close '$Note'"
    if ($Execute) { Invoke-ItsmAction -Session $RouterItsm -Action Close -Ticket $Ticket -Values @{ text = $Note } }
}

function Expand-RouteValue([string]$Template, $Ticket) {
    # {Field:Name} -> form field, {Number}/{Title}/{Service} -> ticket property
    return [regex]::Replace($Template, '\{(Field:)?([^}]+)\}', {
            param($m)
            if ($m.Groups[1].Success) { [string]$Ticket.Fields[$m.Groups[2].Value] } else { [string]$Ticket.($m.Groups[2].Value) }
        }).Trim()
}

function Test-Route($Route, $Ticket) {
    foreach ($k in @($Route.When.Keys)) {
        $v = if ($k -like 'Field:*') { [string]$Ticket.Fields[$k.Substring(6)] } else { [string]$Ticket.$k }
        if ($v -notmatch $Route.When[$k]) { return $false }
    }
    return $true
}

#endregion

$RouterAbsent = Get-AbsentToday
if ($RouterAbsent) { Write-Log "Absent today: $($RouterAbsent -join ', ')" }
$RouterItsm = Connect-Itsm -Config $config
$states = if ($router.States) { @($router.States) } else { @('') }
$teams = if ($router.Teams) { @($router.Teams) } else { @('') }
$tickets = @(Get-ItsmTicket -Session $RouterItsm -State $states -Team $teams -UnassignedOnly | Sort-Object Created | Select-Object -First $MaxTickets)
Write-Log "Unassigned tasks: $($tickets.Count)"

$handlerDir = Join-Path $PSScriptRoot 'handlers'
$context = @{ Config = $config; ConfigPath = $ConfigPath; Execute = [bool]$Execute; RepoRoot = (Split-Path -Parent $PSScriptRoot) }
$results = New-Object System.Collections.Generic.List[object]
foreach ($t in $tickets) {
    $route = @($router.Routes) | Where-Object { $_ -and (Test-Route $_ $t) } | Select-Object -First 1
    if (-not $route) { $route = if ($router.Default) { $router.Default } else { @{ Name = 'default'; Handler = 'RoundRobin' } } }
    $routeName = if ($route.Name) { $route.Name } else { $route.Handler }
    Write-Log "$($t.Number) [$($t.Service)] '$($t.Title)' -> $routeName"
    $row = [ordered]@{ Ticket = $t.Number; Title = $t.Title; Route = $routeName; Result = ''; Note = '' }
    try {
        $script = Join-Path $handlerDir "$($route.Handler).ps1"
        if (-not (Test-Path -LiteralPath $script)) { throw "handler '$($route.Handler)' not found in $handlerDir" }
        $params = if ($route.Params) { $route.Params } else { @{} }
        $r = & $script -Ticket $t -Params $params -Context $context
        $row.Result = $r.Result; $row.Note = $r.Note
    }
    catch {
        $row.Result = 'Error'; $row.Note = $_.Exception.Message
        Write-Log "  $($t.Number): $($row.Note)" -Level Error
        try { Set-TicketAssignee $t (Get-NextAssignee) "Automation failed: $($row.Note)" } catch { Write-Log "  fallback assignment failed: $($_.Exception.Message)" -Level Error }
    }
    $results.Add([PSCustomObject]$row)
}

if ($results.Count) { $results | Format-Table -AutoSize | Out-String -Width 220 | Write-Host }
if ($SendMail -and $results.Count) {
    $mode = if ($Execute) { '' } else { ' (dry run)' }
    Send-ToolkitMail -Config $config -Subject "Ticket router: $($results.Count) task(s)$mode" -BodyAsHtml `
        -Body ('<p>Tasks handled by the router' + $mode + ':</p>' + (ConvertTo-HtmlTable -Rows $results.ToArray()))
}
$results
