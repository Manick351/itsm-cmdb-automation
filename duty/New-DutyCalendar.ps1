<#
.SYNOPSIS
    Generates next month's on-call (duty) calendar as an XLSX template, optionally uploads it and mails the link.

.DESCRIPTION
    One sheet per month, ready to be filled in by the team:

      - per person two rows: "Duty" (1 = on call that day) and "Overtime" (hours worked; an input prompt
        reminds to add a comment with start/end time and what was done);
      - non-working days (weekends and public holidays) are coloured; the patch weekend (the weekend after the
        N-th Tuesday, -PatchWeekendAfterTuesday) gets its own colour so maintenance duty is visible;
      - a Total column with SUM formulas, and conditional formatting that turns a day red when more than one
        person is on duty;
      - a legend and the overtime rules (Duty.Notes) below the table.

    The workbook is written as Office Open XML directly: no Excel, no EPPlus, no ImportExcel.

    Holidays (-HolidaySource):
      IsDayOff  isdayoff.ru production calendar (RU, includes transferred working days and short days)
      Nager     date.nager.at public holidays for -CountryCode, plus weekends
      File      Duty.HolidaysPath: a text file with one yyyy-MM-dd per line, plus weekends
      None      weekends only

.PARAMETER Month
    Any date in the target month. Default: the first day of next month.

.PARAMETER PeoplePath
    CSV with Name, Phone, Role (one row per person). Default: Duty.PeoplePath.

.PARAMETER HolidaySource
    IsDayOff, Nager, File or None. Default: Duty.HolidaySource or None.

.PARAMETER CountryCode
    ISO country code for Nager. Default: Duty.CountryCode.

.PARAMETER PatchWeekendAfterTuesday
    Highlight the weekend after this Tuesday of the month (2 = after Patch Tuesday). 0 disables. Default: 2.

.PARAMETER OutputPath
    Folder for the file. Default: Settings.ReportPath.

.PARAMETER Publish
    Upload to Duty.WebDavUrl (Nextcloud / ownCloud / any WebDAV) as <yyyy.MM>.xlsx with Duty.WebDavCredential.

.PARAMETER SendMail
    Mail Duty.To (or Mail.To) a link (Duty.ShareLink) or, without a link, the file as attachment.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\New-DutyCalendar.ps1 -PeoplePath ..\examples\duty-people.example.csv -HolidaySource Nager -CountryCode DE

.EXAMPLE
    .\New-DutyCalendar.ps1 -Publish -SendMail     # monthly scheduled task on the 20th
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [datetime]$Month = (Get-Date -Day 1).Date.AddMonths(1),
    [string]$PeoplePath,
    [ValidateSet('IsDayOff', 'Nager', 'File', 'None')][string]$HolidaySource,
    [string]$CountryCode,
    [ValidateRange(0, 4)][int]$PatchWeekendAfterTuesday = 2,
    [string]$OutputPath,
    [switch]$Publish,
    [switch]$SendMail,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$duty = $config.Duty
if (-not $PeoplePath) { $PeoplePath = $duty.PeoplePath }
if (-not $PeoplePath) { throw 'No people list: use -PeoplePath or Duty.PeoplePath.' }
if (-not [IO.Path]::IsPathRooted($PeoplePath)) { $PeoplePath = Join-Path (Split-Path -Parent $PSScriptRoot) $PeoplePath }
if (-not $HolidaySource) { $HolidaySource = if ($duty.HolidaySource) { $duty.HolidaySource } else { 'None' } }
if (-not $CountryCode) { $CountryCode = $duty.CountryCode }
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'New-DutyCalendar'

$people = @(Import-Csv -LiteralPath $PeoplePath -Encoding UTF8 | Where-Object { $_.Name })
if (-not $people) { throw "No people in $PeoplePath." }
$year = $Month.Year; $mon = $Month.Month
$days = [DateTime]::DaysInMonth($year, $mon)
$inv = [Globalization.CultureInfo]::InvariantCulture

# --- non-working days ---------------------------------------------------------------------------------------
$offDays = New-Object 'System.Collections.Generic.HashSet[int]'
$addWeekends = { for ($d = 1; $d -le $days; $d++) { if ((Get-Date -Year $year -Month $mon -Day $d).DayOfWeek -in 'Saturday', 'Sunday') { [void]$offDays.Add($d) } } }
switch ($HolidaySource) {
    'IsDayOff' {
        $resp = Invoke-ToolkitRest -Uri "https://isdayoff.ru/api/getdata?year=$year&month=$('{0:D2}' -f $mon)"
        $codes = ([string]$resp.Content).Trim()
        if ($codes.Length -ne $days) { throw "isdayoff.ru returned '$codes' for $days days." }
        for ($d = 1; $d -le $days; $d++) { if ($codes[$d - 1] -eq '1') { [void]$offDays.Add($d) } }
    }
    'Nager' {
        if (-not $CountryCode) { throw 'Nager needs -CountryCode (for example DE, GB, US).' }
        & $addWeekends
        $resp = Invoke-ToolkitRest -Uri "https://date.nager.at/api/v3/PublicHolidays/$year/$CountryCode"
        foreach ($h in ($resp.Content | ConvertFrom-Json)) {
            $dt = [datetime]::ParseExact($h.date, 'yyyy-MM-dd', $inv)
            if ($dt.Month -eq $mon -and ($null -eq $h.global -or $h.global)) { [void]$offDays.Add($dt.Day) }
        }
    }
    'File' {
        & $addWeekends
        foreach ($line in (Get-Content -LiteralPath $duty.HolidaysPath)) {
            if ($line -match '^\s*(\d{4}-\d{2}-\d{2})') { $dt = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd', $inv); if ($dt.Year -eq $year -and $dt.Month -eq $mon) { [void]$offDays.Add($dt.Day) } }
        }
    }
    default { & $addWeekends }
}
$patchDays = New-Object 'System.Collections.Generic.HashSet[int]'
if ($PatchWeekendAfterTuesday) {
    $tuesdays = @(1..$days | Where-Object { (Get-Date -Year $year -Month $mon -Day $_).DayOfWeek -eq 'Tuesday' })
    if ($tuesdays.Count -ge $PatchWeekendAfterTuesday) {
        $t = $tuesdays[$PatchWeekendAfterTuesday - 1]
        foreach ($d in ($t + 4), ($t + 5)) { if ($d -le $days) { [void]$patchDays.Add($d) } }
    }
}
Write-Log ("{0:yyyy-MM}: {1} people, {2} non-working day(s) ({3}), patch weekend: {4}" -f $Month, $people.Count, $offDays.Count, $HolidaySource, $(if ($patchDays.Count) { ($patchDays -join ', ') } else { 'none' }))

# --- sheet XML ----------------------------------------------------------------------------------------------
# Styles (cellXfs): 0 default, 1 header, 2 body, 3 off-day body, 4 patch body, 5 duty label, 6 overtime label,
# 7 off-day header, 8 patch header, 9 total, 10 note, 11 note header
$S = @{ Header = 1; Body = 2; Off = 3; Patch = 4; Duty = 5; Over = 6; OffHead = 7; PatchHead = 8; Total = 9; Note = 10; NoteHead = 11 }
$firstDayCol = 6                                   # F
$lastDayCol = $firstDayCol + $days - 1
$totalCol = $lastDayCol + 1
$lastRow = 3 + 2 * $people.Count
$col = { param([int]$c) ConvertTo-XlsxColumnName $c }
$cells = @{}                                       # "row" -> list of cell xml, keyed by column
function Set-Cell([int]$Row, [int]$Col, $Value, [int]$Style, [string]$Formula) {
    if (-not $cells.ContainsKey($Row)) { $cells[$Row] = @{} }
    $ref = (ConvertTo-XlsxColumnName $Col) + $Row
    $xml = if ($Formula) { "<c r=""$ref"" s=""$Style""><f>$Formula</f></c>" }
           elseif ($null -eq $Value -or "$Value" -eq '') { "<c r=""$ref"" s=""$Style""/>" }
           elseif ($Value -is [int] -or $Value -is [double]) { "<c r=""$ref"" s=""$Style""><v>$(([double]$Value).ToString($inv))</v></c>" }
           else { "<c r=""$ref"" s=""$Style"" t=""inlineStr""><is><t xml:space=""preserve"">$(ConvertTo-XlsxText $Value)</t></is></c>" }
    $cells[$Row][$Col] = $xml
}
$merges = New-Object System.Collections.Generic.List[string]
function Add-Merge([int]$R1, [int]$C1, [int]$R2, [int]$C2, $Value, [int]$Style) {
    for ($r = $R1; $r -le $R2; $r++) { for ($c = $C1; $c -le $C2; $c++) { Set-Cell $r $c $(if ($r -eq $R1 -and $c -eq $C1) { $Value } else { $null }) $Style } }
    $merges.Add("$(& $col $C1)$R1`:$(& $col $C2)$R2")
}

$headers = '#', 'Name', 'Phone', 'Role', 'Type'
for ($i = 0; $i -lt 5; $i++) { Add-Merge 1 ($i + 1) 3 ($i + 1) $headers[$i] $S.Header }
Add-Merge 1 $firstDayCol 1 $lastDayCol ($Month.ToString('MMMM yyyy', $inv)) $S.Header
Add-Merge 1 $totalCol 3 $totalCol 'Total' $S.Total
for ($d = 1; $d -le $days; $d++) {
    $c = $firstDayCol + $d - 1
    $date = Get-Date -Year $year -Month $mon -Day $d
    $hs = if ($patchDays.Contains($d)) { $S.PatchHead } elseif ($offDays.Contains($d)) { $S.OffHead } else { $S.Header }
    $bs = if ($patchDays.Contains($d)) { $S.Patch } elseif ($offDays.Contains($d)) { $S.Off } else { $S.Body }
    Set-Cell 2 $c ($date.DayOfWeek.ToString().Substring(0, 2)) $hs
    Set-Cell 3 $c $d $hs
    for ($r = 4; $r -le $lastRow; $r++) { Set-Cell $r $c $null $bs }
}
$row = 4
$n = 1
foreach ($p in $people) {
    Add-Merge $row 1 ($row + 1) 1 $n $S.Body
    Add-Merge $row 2 ($row + 1) 2 $p.Name $S.Body
    Add-Merge $row 3 ($row + 1) 3 $p.Phone $S.Body
    Add-Merge $row 4 ($row + 1) 4 $p.Role $S.Body
    Set-Cell $row 5 'Duty' $S.Duty
    Set-Cell ($row + 1) 5 'Overtime' $S.Over
    foreach ($r in $row, ($row + 1)) { Set-Cell $r $totalCol $null $S.Total "SUM($(& $col $firstDayCol)$r`:$(& $col $lastDayCol)$r)" }
    $row += 2; $n++
}

# legend and notes
$legendRow = $lastRow + 2
Set-Cell $legendRow $firstDayCol 1 $S.Duty;  Add-Merge $legendRow ($firstDayCol + 1) $legendRow ($firstDayCol + 6) 'On duty' 0
Set-Cell ($legendRow + 1) $firstDayCol $null $S.Off;  Add-Merge ($legendRow + 1) ($firstDayCol + 1) ($legendRow + 1) ($firstDayCol + 6) 'Non-working day' 0
if ($patchDays.Count) { Set-Cell ($legendRow + 2) $firstDayCol $null $S.Patch; Add-Merge ($legendRow + 2) ($firstDayCol + 1) ($legendRow + 2) ($firstDayCol + 6) 'Patch weekend' 0 }
$notes = if ($duty.Notes) { @($duty.Notes) } else { @(
        'Overtime: enter the hours worked on a duty day in the "Overtime" row.'
        'Add a comment to the cell with the start and end time of every piece of work and what was done.'
        'Example: 20:00-21:30 failover test; 22:30-23:00 authentication incident.'
        'Fill in the calendar before the 28th of the previous month.') }
$noteRow = $legendRow
Add-Merge $noteRow 1 $noteRow 5 'Rules' $S.NoteHead
foreach ($t in $notes) { $noteRow++; Add-Merge $noteRow 1 $noteRow 5 $t $S.Note }

# --- assemble -----------------------------------------------------------------------------------------------
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
[void]$sb.Append('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>')
[void]$sb.Append("<sheetViews><sheetView workbookViewId=""0"" tabSelected=""1""><pane xSplit=""5"" ySplit=""3"" topLeftCell=""F4"" activePane=""bottomRight"" state=""frozen""/></sheetView></sheetViews>")
[void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
foreach ($w in @(@(1, 4), @(2, 24), @(3, 16), @(4, 14), @(5, 10))) { [void]$sb.Append("<col min=""$($w[0])"" max=""$($w[0])"" width=""$($w[1])"" customWidth=""1""/>") }
[void]$sb.Append("<col min=""$firstDayCol"" max=""$lastDayCol"" width=""4.2"" customWidth=""1""/><col min=""$totalCol"" max=""$totalCol"" width=""7"" customWidth=""1""/></cols><sheetData>")
foreach ($r in ($cells.Keys | Sort-Object)) {
    $ht = if ($r -ge 4 -and $r -le $lastRow) { ' ht="22" customHeight="1"' } elseif ($r -gt $legendRow) { ' ht="30" customHeight="1"' } else { '' }
    [void]$sb.Append("<row r=""$r""$ht>")
    foreach ($c in ($cells[$r].Keys | Sort-Object)) { [void]$sb.Append($cells[$r][$c]) }
    [void]$sb.Append('</row>')
}
[void]$sb.Append('</sheetData>')
[void]$sb.Append("<mergeCells count=""$($merges.Count)"">" + (($merges | ForEach-Object { "<mergeCell ref=""$_""/>" }) -join '') + '</mergeCells>')
# more than one person on duty on a day -> red. One rule over all duty rows; the formula is relative to F4.
$dutyRows = @(for ($r = 4; $r -le $lastRow; $r += 2) { $r })
$sqref = ($dutyRows | ForEach-Object { "$(& $col $firstDayCol)$_`:$(& $col $lastDayCol)$_" }) -join ' '
$sumExpr = ($dutyRows | ForEach-Object { "$(& $col $firstDayCol)`$$_" }) -join '+'
[void]$sb.Append("<conditionalFormatting sqref=""$sqref""><cfRule type=""expression"" dxfId=""0"" priority=""1""><formula>($sumExpr)&gt;1</formula></cfRule></conditionalFormatting>")
$overRows = ($dutyRows | ForEach-Object { "$(& $col $firstDayCol)$($_ + 1):$(& $col $lastDayCol)$($_ + 1)" }) -join ' '
[void]$sb.Append("<dataValidations count=""1""><dataValidation allowBlank=""1"" showInputMessage=""1"" promptTitle=""Add a comment"" prompt=""Enter hours, then add a cell comment with start/end time and what was done."" sqref=""$overRows""/></dataValidations>")
[void]$sb.Append('<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/></worksheet>')
$sheetXml = $sb.ToString()

$fills = 'D9D9D9', 'FFC000', '3399FF', '00B050', 'FFFF00', 'FFF2CC', 'C00000'
$fillXml = '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    (($fills | ForEach-Object { "<fill><patternFill patternType=""solid""><fgColor rgb=""FF$_""/><bgColor indexed=""64""/></patternFill></fill>" }) -join '')
$border = '<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right><top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>'
$center = '<alignment horizontal="center" vertical="center" wrapText="1"/>'
$xf = { param($font, $fill, $align = $center) "<xf numFmtId=""0"" fontId=""$font"" fillId=""$fill"" borderId=""1"" xfId=""0"" applyFont=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"">$align</xf>" }
$cellXfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    (& $xf 1 2) + (& $xf 0 0) + (& $xf 0 3) + (& $xf 0 4) + (& $xf 0 5) + (& $xf 0 6) + (& $xf 1 3) + (& $xf 1 4) + (& $xf 1 2) +
    (& $xf 0 7 '<alignment horizontal="left" vertical="center" wrapText="1" indent="1"/>') + (& $xf 2 8)
$styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3"><font><sz val="10"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="10"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    "<fills count=""$($fills.Count + 2)"">$fillXml</fills><borders count=""2""><border><left/><right/><top/><bottom/><diagonal/></border>$border</borders>" +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    "<cellXfs count=""12"">$cellXfs</cellXfs><cellStyles count=""1""><cellStyle name=""Normal"" xfId=""0"" builtinId=""0""/></cellStyles>" +
    '<dxfs count="1"><dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs></styleSheet>'

$sheetName = $Month.ToString('yyyy-MM')
$head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
$parts = [ordered]@{
    '[Content_Types].xml'        = $head + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'
    '_rels/.rels'                = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    'xl/workbook.xml'            = $head + "<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets><sheet name=""$sheetName"" sheetId=""1"" r:id=""rId1""/></sheets><calcPr calcId=""191029"" fullCalcOnLoad=""1""/></workbook>"
    'xl/_rels/workbook.xml.rels' = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    'xl/styles.xml'              = $styles
    'xl/worksheets/sheet1.xml'   = $sheetXml
}
$file = Join-Path $OutputPath ('duty-calendar_{0:yyyy.MM}.xlsx' -f $Month)
Add-Type -AssemblyName System.IO.Compression
if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force }
$stream = [IO.File]::Open($file, [IO.FileMode]::CreateNew)
$zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($name in $parts.Keys) {
        $w = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
        try { $w.Write($parts[$name]) } finally { $w.Dispose() }
    }
}
finally { $zip.Dispose(); $stream.Dispose() }
Write-Log "Calendar: $file" -Level Ok

# --- publish and notify -------------------------------------------------------------------------------------
$link = $duty.ShareLink
if ($Publish) {
    if (-not $duty.WebDavUrl) { throw 'Duty.WebDavUrl is not set.' }
    $cred = Get-ToolkitCredential -Name $duty.WebDavCredential -CredentialPath $config.Settings.CredentialPath
    $auth = @{ Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$($cred.UserName):$($cred.GetNetworkCredential().Password)")) }
    $folder = $duty.WebDavUrl.TrimEnd('/')
    # MKCOL is not a standard Invoke-WebRequest method in Windows PowerShell 5.1; 405 = folder exists
    try {
        $mk = [Net.WebRequest]::Create($folder); $mk.Method = 'MKCOL'; $mk.Headers.Add('Authorization', $auth.Authorization)
        $mk.GetResponse().Dispose()
    }
    catch { }
    $remote = "$folder/$('{0:yyyy.MM}' -f $Month).xlsx"
    $null = Invoke-ToolkitRest -Uri $remote -Method Put -Headers $auth -Body ([IO.File]::ReadAllBytes($file)) -ContentType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    Write-Log "Uploaded to $remote" -Level Ok
}
if ($SendMail) {
    $to = if ($duty.To) { @($duty.To) } else { @($config.Mail.To) }
    $deadline = (Get-Date -Year $Month.Year -Month $Month.Month -Day 1).AddMonths(-1).AddDays(27)
    $body = "<p>Hello,</p><p>The duty calendar for <b>$($Month.ToString('MMMM yyyy', $inv))</b> is ready" +
        $(if ($link) { ": <a href=""$link"">open it here</a>." } else { ' (attached).' }) +
        "</p><p>Please fill it in by $($deadline.ToString('d MMMM', $inv)).</p>"
    Send-ToolkitMail -Config $config -To $to -Subject "Duty calendar $($Month.ToString('yyyy-MM'))" -Body $body -BodyAsHtml -Attachment $(if (-not $link) { $file })
}
$file
