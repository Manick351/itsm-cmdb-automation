<#
.SYNOPSIS
    Mails the list of temporary servers whose planned retirement date has passed or is near.

.DESCRIPTION
    Temporary servers (tests, migrations, pilots) get a planned retirement date in the CMDB when they are
    ordered; this report makes sure somebody acts on it. Active servers with PlannedRetireDate are grouped:

      Overdue   the date has passed;
      Due       within -DaysAhead days;
      Later     further away (listed for completeness).

    Hardware is left out: records with physical-asset markers (any of the logical fields listed in
    TemporaryReport.HardwareFields, for example Serial or SupportEnd, or IsVirtual = 0) carry a planned
    replacement date, not a temporary-server end date.

    No mail is sent when nothing is overdue or due, unless -SendAlways. -PreviewPath writes the HTML instead.

.PARAMETER DaysAhead
    Horizon for "due". Default: 14.

.PARAMETER To
    Recipients. Default: TemporaryReport.To or Mail.To.

.PARAMETER SendAlways
    Send even when nothing is overdue or due.

.PARAMETER IncludeHardware
    Add the excluded hardware records as a separate table.

.PARAMETER PreviewPath
    Write the HTML body to this file instead of sending.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Send-TemporaryServerReport.ps1 -PreviewPath .\temp-servers.html

.EXAMPLE
    .\Send-TemporaryServerReport.ps1 -DaysAhead 21
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateRange(1, 365)][int]$DaysAhead = 14,
    [string[]]$To,
    [switch]$SendAlways,
    [switch]$IncludeHardware,
    [string]$PreviewPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Send-TemporaryServerReport'
$tr = if ($config.TemporaryReport) { $config.TemporaryReport } else { @{} }
if (-not $To) { $To = if ($tr.To) { @($tr.To) } else { @($config.Mail.To) } }
$hwFields = if ($tr.HardwareFields) { @($tr.HardwareFields) } else { @('Serial', 'SupportEnd') }

function Test-Value($v) { $s = "$v".Trim(); return ($s -and $s -ne '0' -and $s -notmatch '^0000') }

$glpi = Connect-Glpi -Config $config
try {
    $active = Get-CmdbStatusValue -Config $config -Kind Active
    $servers = @(Get-CmdbServer -Session $glpi -Config $config | Where-Object { $_.Status -eq $active -and (Test-Value $_.PlannedRetireDate) })
}
finally { Disconnect-Glpi $glpi }

$today = (Get-Date).Date
$rows = foreach ($s in $servers) {
    $date = $null
    try { $date = [datetime]"$($s.PlannedRetireDate)" } catch { Write-Log "$($s.Name): cannot parse '$($s.PlannedRetireDate)'" -Level Warn; continue }
    $hw = @($hwFields | Where-Object { Test-Value $s.$_ })
    if ($s.PSObject.Properties['IsVirtual'] -and "$($s.IsVirtual)" -eq '0') { $hw += 'IsVirtual=0' }
    $days = ($date.Date - $today).Days
    [PSCustomObject]@{
        Kind = if ($hw) { 'Hardware' } else { 'Temporary' }
        Server = $s.Name; Ip = $s.Ip; Location = $s.Location
        Owner = if (Test-Value $s.Owner) { $s.Owner } else { 'NOT SET' }
        Date = $date; DaysLeft = $days
        Bucket = if ($days -lt 0) { 'Overdue' } elseif ($days -le $DaysAhead) { 'Due' } else { 'Later' }
    }
}
$rows = @($rows | Sort-Object Date)
$temp = @($rows | Where-Object Kind -eq 'Temporary')
$hw = @($rows | Where-Object Kind -eq 'Hardware')
$overdue = @($temp | Where-Object Bucket -eq 'Overdue')
$due = @($temp | Where-Object Bucket -eq 'Due')
$later = @($temp | Where-Object Bucket -eq 'Later')
Write-Log "Temporary servers: $($temp.Count) (overdue $($overdue.Count), due $($due.Count), later $($later.Count)); hardware excluded: $($hw.Count)"

if (-not $overdue -and -not $due -and -not $SendAlways -and -not $PreviewPath) { Write-Log 'Nothing overdue or due, no mail.' -Level Ok; return }

function ConvertTo-ServerTableHtml($Items, [string]$Color) {
    $view = foreach ($r in $Items) {
        [PSCustomObject]@{
            Server = $r.Server; IP = $r.Ip; Location = $r.Location; Owner = $r.Owner; 'Planned retirement' = $r.Date.ToString('yyyy-MM-dd')
            'Days left' = if ($r.DaysLeft -lt 0) { "overdue by $(-$r.DaysLeft)" } else { $r.DaysLeft }
        }
    }
    (ConvertTo-HtmlTable -Rows @($view)) -replace '<td style="border:1px solid #ddd;padding:4px 8px">(overdue by \d+)</td>', "<td style=""border:1px solid #ddd;padding:4px 8px;color:$Color"">`$1</td>"
}
$body = "<div style='font-family:Segoe UI,Arial;font-size:13px'><h3>Temporary servers - planned retirement</h3><p style='color:#666'>CMDB, $(Get-Date -Format 'yyyy-MM-dd')</p>"
if ($overdue) { $body += "<h4 style='color:#c00'>Overdue: $($overdue.Count)</h4>" + (ConvertTo-ServerTableHtml $overdue '#c00') }
if ($due) { $body += "<h4 style='color:#b8860b'>Due within $DaysAhead days: $($due.Count)</h4>" + (ConvertTo-ServerTableHtml $due '#b8860b') }
if (-not $overdue -and -not $due) { $body += "<p style='color:#090'>Nothing overdue or due.</p>" }
if ($later) { $body += "<h4 style='color:#555'>Later: $($later.Count)</h4>" + (ConvertTo-ServerTableHtml $later '#555') }
if ($IncludeHardware -and $hw) { $body += "<h4 style='color:#888'>Excluded as hardware (planned replacement): $($hw.Count)</h4>" + (ConvertTo-ServerTableHtml $hw '#888') }
$body += "<p style='color:#888;font-size:11px'>Records with hardware markers ($($hwFields -join ', '), IsVirtual = 0) are excluded: $($hw.Count).</p></div>"
$parts = @()
if ($overdue) { $parts += "overdue $($overdue.Count)" }
if ($due) { $parts += "due $($due.Count)" }
if (-not $parts) { $parts = @('no changes') }
$subject = "Temporary servers - $($parts -join ', ') ($(Get-Date -Format 'yyyy-MM-dd'))"

if ($PreviewPath) {
    [IO.File]::WriteAllText($PreviewPath, $body, (New-Object Text.UTF8Encoding($false)))
    Write-Log "Preview: $PreviewPath ($subject)" -Level Ok
}
else { Send-ToolkitMail -Config $config -To $To -Subject $subject -Body $body -BodyAsHtml }
