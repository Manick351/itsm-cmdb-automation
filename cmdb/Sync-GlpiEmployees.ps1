<#
.SYNOPSIS
    Adds Active Directory users to a GLPI object used as "owner" / "administrator" dropdown of CMDB records.

.DESCRIPTION
    Many GLPI CMDB designs keep business and service owners in a separate object (for example a Generic Object
    "Person") instead of GLPI users, so that people without a GLPI login can be referenced. This script keeps
    that object filled from AD:

      - enabled users under Employees.SearchBase (ADSI, no ActiveDirectory module) with a display name;
      - the key (Employees.KeyField, by default the SID) decides whether a user already exists;
      - Employees.Map builds the GLPI fields from AD attributes with {attribute} placeholders;
      - with -UpdateChanged, existing entries whose mapped fields differ (renamed users) are updated.

    Without -Execute only the planned additions and updates are listed.

.PARAMETER Execute
    Write to GLPI.

.PARAMETER UpdateChanged
    Also update existing entries whose mapped values changed.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.EXAMPLE
    .\Sync-GlpiEmployees.ps1

.EXAMPLE
    .\Sync-GlpiEmployees.ps1 -Execute -UpdateChanged
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [switch]$Execute,
    [switch]$UpdateChanged,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$e = $config.Employees
if (-not $e -or -not $e.Itemtype -or -not $e.Map) { throw 'Employees.Itemtype and Employees.Map must be configured.' }
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Sync-GlpiEmployees'
if (-not $Execute) { Write-Log 'Dry run: GLPI is not changed (use -Execute).' -Level Warn }
$keyField = if ($e.KeyField) { $e.KeyField } else { 'contact' }

# --- AD users ---------------------------------------------------------------------------------------------
$root = if ($e.SearchBase) { New-Object DirectoryServices.DirectoryEntry("LDAP://$($e.SearchBase)") } else { New-Object DirectoryServices.DirectoryEntry }
$searcher = New-Object DirectoryServices.DirectorySearcher($root)
$searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2))(displayName=*)$($e.LdapFilter))"
$searcher.PageSize = 1000
$attrs = @([regex]::Matches(($e.Map.Values -join ' '), '\{(\w+)\}') | ForEach-Object { $_.Groups[1].Value.ToLower() } | Select-Object -Unique)
foreach ($a in $attrs) { [void]$searcher.PropertiesToLoad.Add($a) }
$users = foreach ($r in $searcher.FindAll()) {
    $vals = @{}
    foreach ($a in $attrs) {
        $raw = if ($r.Properties[$a].Count) { $r.Properties[$a][0] } else { '' }
        if ($a -eq 'objectsid' -and $raw -is [byte[]]) { $raw = (New-Object Security.Principal.SecurityIdentifier($raw, 0)).Value }
        $vals[$a] = [string]$raw
    }
    $fields = @{}
    foreach ($k in $e.Map.Keys) { $fields[$k] = [regex]::Replace([string]$e.Map[$k], '\{(\w+)\}', { param($m) $vals[$m.Groups[1].Value.ToLower()] }).Trim() }
    if ($fields[$keyField]) { [PSCustomObject]@{ Key = $fields[$keyField]; Fields = $fields } }
}
$users = @($users)
Write-Log "AD: $($users.Count) enabled user(s)"

# --- GLPI -------------------------------------------------------------------------------------------------
$glpi = $null
$added = 0; $changed = 0; $failed = 0
try {
    $glpi = Connect-Glpi -Config $config
    $existing = @{}
    foreach ($x in (Get-GlpiItem -Session $glpi -Itemtype $e.Itemtype)) { if ($x.$keyField) { $existing[[string]$x.$keyField] = $x } }
    Write-Log "GLPI $($e.Itemtype): $($existing.Count) entr(ies)"
    foreach ($u in $users) {
        $cur = $existing[$u.Key]
        try {
            if (-not $cur) {
                Write-Log "add: $($u.Fields.name)"
                if ($Execute) { $null = Invoke-GlpiApi -Session $glpi -Path $e.Itemtype -Method Post -InputObject $u.Fields }
                $added++
                continue
            }
            if (-not $UpdateChanged) { continue }
            $diff = @{}
            foreach ($k in $u.Fields.Keys) { if ([string]$cur.$k -ne $u.Fields[$k]) { $diff[$k] = $u.Fields[$k] } }
            if (-not $diff.Count) { continue }
            Write-Log "update $($cur.id): $(($diff.Keys | ForEach-Object { "$_ '$($cur.$_)' -> '$($diff[$_])'" }) -join '; ')"
            if ($Execute) { $null = Invoke-GlpiApi -Session $glpi -Path "$($e.Itemtype)/$($cur.id)" -Method Put -InputObject $diff }
            $changed++
        }
        catch { $failed++; Write-Log "$($u.Fields.name): $(Get-RestErrorText $_)" -Level Error }
    }
}
finally { Disconnect-Glpi $glpi }
$verb = if ($Execute) { '' } else { ' (dry run)' }
Write-Log "Added $added, updated $changed, failed $failed$verb" -Level $(if ($failed) { 'Warn' } else { 'Ok' })
if ($failed) { exit 1 }
