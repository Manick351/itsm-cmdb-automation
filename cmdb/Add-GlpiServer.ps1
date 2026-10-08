<#
.SYNOPSIS
    Creates a server record in the GLPI CMDB from a provisioning request (form fields), with validation and rules.

.DESCRIPTION
    The request is a set of form fields as the service desk delivers them (JSON file, hashtable from the ticket
    router, CSV row). The config translates it in four steps:

      1. Request.FieldMap      request field -> logical field (Name, Ip, Os, Environment, Owner, ...);
      2. Request.ValueMap      per logical field, request wording -> CMDB wording (for example OS names);
         Request.Disks         data disk fields "Data Disk #1 Volume (GB)" ... -> one Disks text;
         Request.Defaults      fixed values (Status, IsVirtual, ...);
      3. Rules                 declarative rules that derive values (backup policy from environment and
                               criticality, support group from location, ...);
      4. Glpi.Fields           logical field -> GLPI field / dropdown (see lib\Glpi.ps1).

    Checks before writing: name present and not an FQDN/IP, Request.NameRules.SkipPattern (desktops etc.),
    Request.Required fields present (otherwise Incomplete: nothing is written), no
    active record with the same name, no record retired less than Glpi.RetiredGraceDays ago (a server re-created
    right after decommissioning usually needs a human look). The FQDN is taken from DNS when the name resolves.

    Without -Execute nothing is written: the resolved GLPI input and all warnings are shown.

.PARAMETER RequestPath
    JSON file with one request object (or an array of them).

.PARAMETER Request
    Request fields as a hashtable (used by tickets\Invoke-TicketRouter.ps1).

.PARAMETER Batch
    Several requests at once: objects with Fields (hashtable) and Ticket (request number). Used by
    Invoke-CmdbLifecycle.ps1 so that GLPI is read once.

.PARAMETER TicketNumber
    Request number, written to the comment ("Commission request: <number>").

.PARAMETER Execute
    Create the record.

.PARAMETER ConfigPath
    Configuration (.psd1). Default: ..\config\config.psd1.

.OUTPUTS
    One object per request: Name, Status (Created, WouldCreate, AlreadyExists, GracePeriod, Skipped, Incomplete, Error),
    AssetId, Message, Warnings.

.EXAMPLE
    .\Add-GlpiServer.ps1 -RequestPath ..\examples\server-request.example.json

.EXAMPLE
    .\Add-GlpiServer.ps1 -RequestPath .\req-1042.json -TicketNumber REQ-1042 -Execute
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'File')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'File')][string]$RequestPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'Object')][System.Collections.IDictionary]$Request,
    [Parameter(Mandatory = $true, ParameterSetName = 'Batch')][object[]]$Batch,
    [string]$TicketNumber,
    [switch]$Execute,
    [string]$ConfigPath = (Join-Path $PSScriptRoot '..\config\config.psd1')
)

. (Join-Path $PSScriptRoot '..\lib\Common.ps1')
. (Join-Path $PSScriptRoot '..\lib\Glpi.ps1')

$config = Import-ToolkitConfig -Path $ConfigPath
$null = Start-ToolkitLog -LogPath $config.Settings.LogPath -Name 'Add-GlpiServer'
$req = $config.Request
$graceDays = if ($null -ne $config.Glpi.RetiredGraceDays) { [int]$config.Glpi.RetiredGraceDays } else { 30 }

$requests = switch ($PSCmdlet.ParameterSetName) {
    'Object' { @{ Fields = $Request; Ticket = $TicketNumber } }
    'Batch' { $Batch }
    default {
        foreach ($r in @(Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json)) {
            $h = [ordered]@{}
            foreach ($p in $r.PSObject.Properties) { $h[$p.Name] = $p.Value }
            @{ Fields = $h; Ticket = $TicketNumber }
        }
    }
}

function ConvertFrom-ServerRequest([System.Collections.IDictionary]$Fields) {
    $rec = [ordered]@{}
    if ($req.FieldMap) {
        foreach ($logical in $req.FieldMap.Keys) {
            $src = $req.FieldMap[$logical]
            $v = if ($Fields.Contains($src)) { $Fields[$src] } else { $null }
            if ($v -is [string]) { $v = $v.Trim() }
            if ($null -ne $v -and "$v" -ne '') { $rec[$logical] = $v }
        }
    }
    if ($req.ValueMap) {
        foreach ($logical in $req.ValueMap.Keys) {
            if (-not $rec.Contains($logical)) { continue }
            $map = $req.ValueMap[$logical]
            foreach ($k in $map.Keys) { if ([string]$rec[$logical] -match $k) { $rec[$logical] = $map[$k]; break } }
        }
    }
    if ($req.Bool) {
        foreach ($logical in $req.Bool) {
            if ($rec.Contains($logical)) { $rec[$logical] = [int]([string]$rec[$logical] -match '^(yes|true|1|y)$') }
        }
    }
    if ($req.Disks) {
        $lines = New-Object System.Collections.Generic.List[string]
        $sys = $req.Disks.SystemDiskGB
        if ($sys) {
            $size = if ([string]$rec.Os -match 'Windows' -and $sys.Windows) { $sys.Windows } else { $sys.Default }
            if ($size) { $lines.Add("Disk 0: $size GB (system)") }
        }
        for ($n = 1; $n -le 26; $n++) {
            $sizeKey = $req.Disks.SizePattern -replace '\{n\}', $n
            if (-not $Fields.Contains($sizeKey) -or -not "$($Fields[$sizeKey])".Trim()) { break }
            $type = if ($req.Disks.TypePattern) { "$($Fields[($req.Disks.TypePattern -replace '\{n\}', $n)])".Trim() } else { '' }
            $lines.Add(("Disk {0}: {1} GB{2}" -f $n, "$($Fields[$sizeKey])".Trim(), $(if ($type) { " ($type)" } else { '' })))
        }
        if ($lines.Count) { $rec['Disks'] = $lines -join "`n" }
    }
    if ($rec.Contains('RamGB')) { $rec['RamMB'] = [int]([double]$rec.RamGB * 1024); $rec.Remove('RamGB') }
    foreach ($k in @($req.Defaults.Keys)) { if (-not $rec.Contains($k)) { $rec[$k] = $req.Defaults[$k] } }
    return $rec
}

$glpi = $null
try {
    $glpi = Connect-Glpi -Config $config
    $existing = @(Get-CmdbServer -Session $glpi -Config $config)
    $active = Get-CmdbStatusValue -Config $config -Kind Active
    $retired = Get-CmdbStatusValue -Config $config -Kind Retired

    foreach ($item in @($requests)) {
        $rec = ConvertFrom-ServerRequest $item.Fields
        $ticket = [string]$item.Ticket
        $name = [string]$rec.Name
        $result = [ordered]@{ Name = $name; Ticket = $ticket; Status = ''; AssetId = $null; Message = ''; Warnings = @() }
        $warnings = New-Object System.Collections.Generic.List[string]
        try {
            if (-not $name) { throw 'server name is empty' }
            if ($name -match '\.' -or $name -match '\s') { throw "'$name' is not a short host name (FQDN, IP or spaces)" }
            if ($req.NameRules.SkipPattern -and $name -match $req.NameRules.SkipPattern) {
                $result.Status = 'Skipped'; $result.Message = "name matches Request.NameRules.SkipPattern"; continue
            }
            $same = @($existing | Where-Object { $_.Name -eq $name })
            $live = @($same | Where-Object { $_.Status -eq $active })
            if ($live) { $result.Status = 'AlreadyExists'; $result.AssetId = $live[0].AssetId; $result.Message = "active record $($live[0].AssetId)"; continue }
            $recent = @($same | Where-Object { $_.Status -eq $retired -and $_.RetireDate -and ((Get-Date) - [datetime]$_.RetireDate).TotalDays -le $graceDays })
            if ($recent) {
                $result.Status = 'GracePeriod'
                $result.Message = "retired on $(([datetime]$recent[0].RetireDate).ToString('yyyy-MM-dd')), less than $graceDays days ago; create manually if intended"
                continue
            }

            $dns = Resolve-HostAddress $name
            if ($dns -and $dns.Fqdn -like "$name.*") {
                $rec['Fqdn'] = $dns.Fqdn
                if (-not $rec.Contains('Ip')) { $rec['Ip'] = $dns.Ip }
                elseif ($rec.Ip -ne $dns.Ip) { $warnings.Add("DNS resolves $name to $($dns.Ip), request says $($rec.Ip)") }
            }
            elseif ($req.DefaultDomain) { $rec['Fqdn'] = "$name.$($req.DefaultDomain)" }
            $rec['Status'] = $active
            $rec['CommissionDate'] = (Get-Date).ToString('yyyy-MM-dd')
            if ($ticket) { $rec['Comment'] = "Commission request: $ticket" }

            $applied = @(Invoke-CmdbRule -Rules $config.Rules -Record $rec)
            if ($applied) { Write-Log "${name}: rules applied: $($applied -join ', ')" -Level Debug }
            $missing = @($req.Required | Where-Object { $_ -and -not $rec.Contains($_) })
            if ($missing) {
                $result.Status = 'Incomplete'; $result.Message = "required field(s) empty: $($missing -join ', ')"
                continue
            }

            $cmdbInput = ConvertTo-CmdbInput -Session $glpi -Config $config -Values $rec
            foreach ($w in $cmdbInput.Warnings) { $warnings.Add($w) }
            $unmapped = @($rec.Keys | Where-Object { -not $config.Glpi.Fields.ContainsKey($_) })
            if ($unmapped) { $warnings.Add("not mapped to GLPI (ignored): $($unmapped -join ', ')") }

            Write-Log "${name}: $($cmdbInput.Asset.Count) asset field(s), $($cmdbInput.Container.Count) CMDB field(s), $($warnings.Count) warning(s)"
            foreach ($k in $rec.Keys) { Write-Log ('    {0,-16} {1}' -f $k, (([string]$rec[$k]) -replace "`n", ' | ')) -Level Debug }
            if (-not $Execute) { $result.Status = 'WouldCreate'; $result.Message = 'dry run, use -Execute'; continue }

            $ids = New-CmdbServer -Session $glpi -Config $config -CmdbInput $cmdbInput
            $result.Status = 'Created'; $result.AssetId = $ids.AssetId
            $existing += [PSCustomObject]@{ Name = $name; Status = $active; AssetId = $ids.AssetId; RetireDate = $null }
            $result.Message = "asset $($ids.AssetId)" + $(if ($ids.ContainerId) { ", CMDB record $($ids.ContainerId)" } else { '' })
            Write-Log "${name}: created ($($result.Message))" -Level Ok
        }
        catch {
            $result.Status = 'Error'; $result.Message = Get-RestErrorText $_
            Write-Log "${name}: $($result.Message)" -Level Error
        }
        finally {
            foreach ($w in $warnings) { Write-Log "${name}: $w" -Level Warn }
            $result.Warnings = $warnings.ToArray()
            if ($result.Status -in 'Skipped', 'AlreadyExists', 'GracePeriod', 'Incomplete') { Write-Log "${name}: $($result.Status) - $($result.Message)" -Level Warn }
            [PSCustomObject]$result
        }
    }
}
finally { Disconnect-Glpi $glpi }
