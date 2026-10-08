<#
.SYNOPSIS
    Keeps CMDB server records in line with the hypervisor inventory: compares fields, updates the drift, lists gaps.

.DESCRIPTION
    Reads an inventory CSV (for example the "VMs" sheet of vmware-cloud-toolkit's Get-VMInventory saved as CSV,
    or any export with one row per VM) and every active server in GLPI. For each server:

      - the inventory row is matched by name (VM names like "srv-app-01 (old)" are matched by the first token);
        duplicates are narrowed down to powered-on rows, then to rows with a location;
      - Inventory.FieldMap columns are compared with the CMDB values (dropdowns by name); differences become
        updates. A multi-address IP cell is never written (reported instead).

    Report (XLSX): "Changes" (field, CMDB value, inventory value, applied), "Not in inventory" (active CMDB
    servers without a VM: physical, other platforms or forgotten decommissions) and "Not in CMDB" (VMs without
    a record). Without -Execute nothing is written to GLPI.

.PARAMETER InventoryPath
    Inventory CSV. Default: Inventory.Path from the config.

.PARAMETER Execute
    Write the changes to GLPI.

.PARAMETER OutputPath
    Folder for the report. Default: Settings.ReportPath.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Sync-GlpiInventory.ps1 -InventoryPath .\vm-inventory.csv

.EXAMPLE
    .\Sync-GlpiInventory.ps1 -Execute
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$InventoryPath,
    [switch]$Execute,
    [string]$OutputPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')
. (Join-Path $PSScriptRoot '..\lib\Xlsx.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$inv = $config.Inventory
if (-not $InventoryPath) { $InventoryPath = $inv.Path }
if (-not $InventoryPath) { throw 'No inventory file: use -InventoryPath or Inventory.Path.' }
if (-not $OutputPath) { $OutputPath = $config.Settings.ReportPath }
if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Sync-GlpiInventory'
if (-not $Execute) { Write-Log 'Dry run: differences are reported, GLPI is not changed (use -Execute).' -Level Warn }

$nameCol = if ($inv.NameColumn) { $inv.NameColumn } else { 'Name' }
$rows = @(Import-Csv -LiteralPath $InventoryPath -Encoding UTF8 | Where-Object { $_.$nameCol })
if ($inv.ExcludePattern) { $rows = @($rows | Where-Object { $_.$nameCol -notmatch $inv.ExcludePattern }) }
$byName = @{}
foreach ($r in $rows) {
    $key = (($r.$nameCol -split '[\s(]')[0]).ToLower()
    if (-not $byName.ContainsKey($key)) { $byName[$key] = New-Object System.Collections.Generic.List[object] }
    $byName[$key].Add($r)
}
Write-Log "Inventory: $($rows.Count) row(s), $($byName.Count) name(s)"

function Get-InventoryValue($Row, $Spec) {
    if ($Spec -is [string]) { return ([string]$Row.$Spec).Trim() }
    $v = ([string]$Row.($Spec.Column)).Trim()
    if ($v -and $Spec.Multiply) { $v = [string][int64]([double]$v * $Spec.Multiply) }
    if ($v -and $Spec.Map) { foreach ($k in $Spec.Map.Keys) { if ($v -match $k) { $v = $Spec.Map[$k]; break } } }
    return $v
}

$changes = New-Object System.Collections.Generic.List[object]
$notInInventory = New-Object System.Collections.Generic.List[object]
$glpi = $null
try {
    $glpi = Connect-Glpi -Config $config
    $active = Get-CmdbStatusValue -Config $config -Kind Active
    $servers = @(Get-CmdbServer -Session $glpi -Config $config | Where-Object { $_.Status -eq $active })
    Write-Log "CMDB: $($servers.Count) active server(s)"
    $updated = 0

    foreach ($s in $servers) {
        $cands = $byName[([string]$s.Name).ToLower()]
        if (-not $cands) {
            $notInInventory.Add([PSCustomObject]@{ Name = $s.Name; Ip = $s.Ip; Location = $s.Location; Environment = $s.Environment; Description = $s.Description })
            continue
        }
        $pick = $cands.ToArray()
        if ($pick.Count -gt 1 -and $inv.PowerColumn) { $on = @($pick | Where-Object { $_.($inv.PowerColumn) -match '^(on|poweredon|powered_on)$' }); if ($on) { $pick = $on } }
        if ($pick.Count -gt 1 -and $inv.LocationColumn) { $loc = @($pick | Where-Object { $_.($inv.LocationColumn) }); if ($loc) { $pick = $loc } }
        if ($pick.Count -gt 1) { Write-Log "$($s.Name): $($pick.Count) inventory rows after filtering, using the first" -Level Warn }
        $row = $pick[0]

        $diff = [ordered]@{}
        foreach ($logical in $inv.FieldMap.Keys) {
            $new = Get-InventoryValue $row $inv.FieldMap[$logical]
            if (-not $new) { continue }
            $old = ([string]$s.$logical).Trim()
            if ($old -eq $new) { continue }
            $note = ''
            if ($logical -eq 'Ip' -and $new -match '[\s,;]') { $note = 'several addresses, not written' }
            $changes.Add([PSCustomObject]@{ Server = $s.Name; Field = $logical; CMDB = $old; Inventory = $new; Applied = $false; Note = $note })
            if (-not $note) { $diff[$logical] = $new }
        }
        if (-not $diff.Count) { continue }
        Write-Log ("$($s.Name): " + (($diff.Keys | ForEach-Object { "$_ '$($s.$_)' -> '$($diff[$_])'" }) -join '; '))
        if (-not $Execute) { continue }
        try {
            $cmdbInput = ConvertTo-CmdbInput -Session $glpi -Config $config -Values $diff
            foreach ($w in $cmdbInput.Warnings) { Write-Log "$($s.Name): $w" -Level Warn }
            Set-CmdbServer -Session $glpi -Config $config -Server $s -CmdbInput $cmdbInput
            foreach ($c in $changes) { if ($c.Server -eq $s.Name -and $diff.Contains($c.Field)) { $c.Applied = $true } }
            $updated++
        }
        catch { Write-Log "$($s.Name): update failed: $(Get-RestErrorText $_)" -Level Error }
    }

    $cmdbNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($s in $servers) { [void]$cmdbNames.Add([string]$s.Name) }
    $notInCmdb = foreach ($k in ($byName.Keys | Sort-Object)) {
        if ($cmdbNames.Contains($k)) { continue }
        $r = $byName[$k][0]
        $o = [ordered]@{ Name = $r.$nameCol }
        foreach ($logical in $inv.FieldMap.Keys) { $o[$logical] = Get-InventoryValue $r $inv.FieldMap[$logical] }
        [PSCustomObject]$o
    }
    $notInCmdb = @($notInCmdb)

    $file = Join-Path $OutputPath ('cmdb-inventory-sync_{0:yyyyMMdd_HHmm}.xlsx' -f (Get-Date))
    Export-Xlsx -Path $file -Sheets @(
        (New-XlsxSheet -Name 'Changes' -Rows $changes.ToArray() -Widths @{ Server = 22; CMDB = 30; Inventory = 30; Note = 30 })
        (New-XlsxSheet -Name 'Not in inventory' -Rows $notInInventory.ToArray() -Widths @{ Description = 40 })
        (New-XlsxSheet -Name 'Not in CMDB' -Rows $notInCmdb)
    )
    Write-Log ("Differences: {0} in {1} server(s), updated {2}; not in inventory {3}; not in CMDB {4}" -f $changes.Count,
        @($changes | Select-Object -ExpandProperty Server -Unique).Count, $updated, $notInInventory.Count, $notInCmdb.Count) -Level Ok
    Write-Log "Report: $file" -Level Ok
    $file
}
finally { Disconnect-Glpi $glpi }
