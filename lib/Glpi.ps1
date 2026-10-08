# GLPI REST API client and a configurable CMDB layer on top of it.
# Requires lib\Common.ps1. Windows PowerShell 5.1 and PowerShell 7 compatible.
#
# GLPI is a construction kit: every installation models servers differently (core Computer fields, the Fields
# plugin, Generic Objects). Scripts therefore never use GLPI field names directly. They work with LOGICAL
# fields (Name, Status, Ip, Environment, ...) mapped in the config:
#
#   Glpi.Fields.Environment = @{ On = 'Container'; Field = 'plugin_fields_environmentfielddropdowns_id'; Dropdown = 'PluginFieldsEnvironmentfieldDropdown' }
#
#   On        Asset (the AssetItemtype, usually Computer) or Container (a Fields plugin container, ContainerItemtype)
#   Field     GLPI field name
#   Dropdown  itemtype of the dropdown, when the field stores an id (names are resolved to ids on write)
#   Multi     $true for Fields plugin multiple dropdowns (stored as "[id1,id2]")

#region Session

function Connect-Glpi {
    <#
    .SYNOPSIS
        Opens a GLPI API session (initSession). AuthMode UserToken (personal API token) or Basic (login/password).
    #>
    param([Parameter(Mandatory = $true)][hashtable]$Config)
    $g = $Config.Glpi
    if (-not $g.Url) { throw 'Glpi.Url is not set.' }
    $cred = Get-ToolkitCredential -Name $g.Credential -CredentialPath $Config.Settings.CredentialPath
    $headers = @{ Accept = 'application/json' }
    if ($g.AppToken) {
        $headers['App-Token'] = (Get-ToolkitCredential -Name $g.AppToken -CredentialPath $Config.Settings.CredentialPath).GetNetworkCredential().Password
    }
    $auth = @{} + $headers
    if ($g.AuthMode -eq 'Basic') {
        $pair = '{0}:{1}' -f $cred.UserName, $cred.GetNetworkCredential().Password
        $auth.Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    }
    else { $auth.Authorization = 'user_token ' + $cred.GetNetworkCredential().Password }
    $skip = [bool]$Config.Settings.SkipCertificateCheck
    $base = $g.Url.TrimEnd('/') -replace '/apirest\.php$', ''
    $resp = Invoke-ToolkitRest -Uri "$base/apirest.php/initSession" -Headers $auth -SkipCertificateCheck:$skip
    $token = ($resp.Content | ConvertFrom-Json).session_token
    if (-not $token) { throw 'GLPI did not return a session token.' }
    $headers['Session-Token'] = $token
    return [PSCustomObject]@{ Base = "$base/apirest.php"; Headers = $headers; SkipCert = $skip; Dropdowns = @{} }
}

function Disconnect-Glpi {
    param([Parameter(Position = 0)]$Session)
    if (-not $Session) { return }
    try { Invoke-ToolkitRest -Uri "$($Session.Base)/killSession" -Headers $Session.Headers -SkipCertificateCheck:$Session.SkipCert -Retries 0 | Out-Null } catch { }
}

function Invoke-GlpiApi {
    <#
    .SYNOPSIS
        Calls the GLPI API. -Input is wrapped as { "input": ... } for POST/PUT. Returns parsed JSON.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Method = 'Get',
        $InputObject
    )
    $request = @{ Uri = "$($Session.Base)/$($Path.TrimStart('/'))"; Method = $Method; Headers = $Session.Headers; SkipCertificateCheck = $Session.SkipCert }
    if ($null -ne $InputObject) {
        $request.Body = [Text.Encoding]::UTF8.GetBytes((@{ input = $InputObject } | ConvertTo-Json -Depth 10))
        $request.ContentType = 'application/json'
    }
    $resp = Invoke-ToolkitRest @request
    $content = if ($resp.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($resp.Content) } else { [string]$resp.Content }
    if (-not $content) { return $null }
    return $content | ConvertFrom-Json
}

function Get-GlpiItem {
    <#
    .SYNOPSIS
        Reads all items of an itemtype page by page (range=start-end). -Expand returns dropdown names instead of ids.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Itemtype,
        [switch]$Expand,
        [int]$PageSize = 500,
        [string]$Query
    )
    $start = 0
    do {
        $path = "$Itemtype`?range=$start-$($start + $PageSize - 1)"
        if ($Expand) { $path += '&expand_dropdowns=true' }
        if ($Query) { $path += "&$Query" }
        try { $page = @(Invoke-GlpiApi -Session $Session -Path $path) }
        catch {
            # GLPI answers 400 ERROR_RANGE_EXCEED_TOTAL when start is past the end
            if ($start -gt 0 -and "$($_.ErrorDetails.Message)$($_.Exception.Message)" -match 'RANGE|400') { break }
            throw
        }
        foreach ($i in $page) { $i }
        $start += $PageSize
    } while ($page.Count -eq $PageSize)
}

function Resolve-GlpiDropdown {
    <#
    .SYNOPSIS
        Dropdown name -> id (cached per session). Matches name, then completename ("Parent > Child"). Returns $null if unknown.
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][string]$Itemtype, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if (-not $Session.Dropdowns.ContainsKey($Itemtype)) {
        $map = @{}
        foreach ($d in (Get-GlpiItem -Session $Session -Itemtype $Itemtype -PageSize 1000)) {
            foreach ($key in @($d.name, $d.completename)) {
                if ($key -and -not $map.ContainsKey(([string]$key).ToLower())) { $map[([string]$key).ToLower()] = $d.id }
            }
        }
        $Session.Dropdowns[$Itemtype] = $map
    }
    $hit = $Session.Dropdowns[$Itemtype][$Name.Trim().ToLower()]
    if ($null -eq $hit) { return $null }
    return $hit
}

#endregion

#region CMDB layer

function Get-CmdbFieldMap {
    param([hashtable]$Config)
    $f = $Config.Glpi.Fields
    if (-not $f -or -not $f.Name -or -not $f.Status) { throw 'Glpi.Fields must define at least Name and Status.' }
    return $f
}

function Get-CmdbServer {
    <#
    .SYNOPSIS
        Reads all servers as objects with logical properties (AssetId, ContainerId, Name, Status, ...).
    .DESCRIPTION
        Asset items and, if configured, Fields plugin container items are read with expanded dropdowns and joined
        by items_id. Only assets with a container record are returned when a container is configured and
        Glpi.ContainerRequired is $true (the container usually marks "this computer is a managed server").
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][hashtable]$Config)
    $g = $Config.Glpi
    $fields = Get-CmdbFieldMap $Config
    $asset = if ($g.AssetItemtype) { $g.AssetItemtype } else { 'Computer' }
    $assets = @{}
    foreach ($a in (Get-GlpiItem -Session $Session -Itemtype $asset -Expand)) { $assets[[string]$a.id] = $a }
    $containers = @{}
    if ($g.ContainerItemtype) {
        foreach ($c in (Get-GlpiItem -Session $Session -Itemtype $g.ContainerItemtype -Expand)) {
            if (-not $c.itemtype -or $c.itemtype -eq $asset) { $containers[[string]$c.items_id] = $c }
        }
    }
    foreach ($id in $assets.Keys) {
        $a = $assets[$id]
        $c = $containers[$id]
        if ($g.ContainerItemtype -and $g.ContainerRequired -and -not $c) { continue }
        $o = [ordered]@{ AssetId = [int]$id; ContainerId = $(if ($c) { [int]$c.id } else { $null }) }
        foreach ($name in $fields.Keys) {
            $def = $fields[$name]
            $src = if ($def.On -eq 'Container') { $c } else { $a }
            $v = if ($src) { $src.($def.Field) } else { $null }
            if ($v -is [string] -and $v -eq '0' -and $def.Dropdown) { $v = $null }
            if ($v -is [int] -and $v -eq 0 -and $def.Dropdown) { $v = $null }
            $o[$name] = $v
        }
        [PSCustomObject]$o
    }
}

function ConvertTo-CmdbInput {
    <#
    .SYNOPSIS
        Logical values -> @{ Asset = @{field=value}; Container = @{...} } with dropdown names resolved to ids.
    .OUTPUTS
        Hashtable with Asset, Container and Warnings.
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][hashtable]$Config, [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Values)
    $fields = Get-CmdbFieldMap $Config
    $out = @{ Asset = @{}; Container = @{}; Warnings = New-Object System.Collections.Generic.List[string] }
    foreach ($name in $Values.Keys) {
        if (-not $fields.ContainsKey($name)) { continue }
        $def = $fields[$name]
        $value = $Values[$name]
        if ($null -eq $value -or "$value" -eq '') { continue }
        if ($def.Dropdown) {
            $names = if ($def.Multi) { @($value) } else { @($value)[0] }
            $ids = foreach ($n in @($names)) {
                $id = Resolve-GlpiDropdown -Session $Session -Itemtype $def.Dropdown -Name ([string]$n)
                if ($null -eq $id) { $out.Warnings.Add("$name '$n' not found in $($def.Dropdown)") } else { $id }
            }
            $ids = @($ids)
            if (-not $ids) { continue }
            $value = if ($def.Multi) { '[' + ($ids -join ',') + ']' } else { $ids[0] }
        }
        $target = if ($def.On -eq 'Container') { $out.Container } else { $out.Asset }
        $target[$def.Field] = $value
    }
    return $out
}

function New-CmdbServer {
    <#
    .SYNOPSIS
        Creates the asset (and the container record) for a server. Returns @{ AssetId; ContainerId }.
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][hashtable]$Config, [Parameter(Mandatory = $true)][hashtable]$CmdbInput)
    $g = $Config.Glpi
    $asset = if ($g.AssetItemtype) { $g.AssetItemtype } else { 'Computer' }
    $created = Invoke-GlpiApi -Session $Session -Path $asset -Method Post -InputObject $CmdbInput.Asset
    $assetId = [int]$created.id
    if (-not $assetId) { throw "GLPI did not return an id for the new $asset." }
    $containerId = $null
    if ($g.ContainerItemtype) {
        $body = @{ items_id = $assetId; itemtype = $asset } + $CmdbInput.Container
        $c = Invoke-GlpiApi -Session $Session -Path $g.ContainerItemtype -Method Post -InputObject $body
        $containerId = [int]$c.id
    }
    return @{ AssetId = $assetId; ContainerId = $containerId }
}

function Set-CmdbServer {
    <#
    .SYNOPSIS
        Updates logical fields of an existing server (PUT on the asset and/or the container record).
    #>
    param([Parameter(Mandatory = $true)]$Session, [Parameter(Mandatory = $true)][hashtable]$Config, [Parameter(Mandatory = $true)]$Server, [Parameter(Mandatory = $true)][hashtable]$CmdbInput)
    $g = $Config.Glpi
    $asset = if ($g.AssetItemtype) { $g.AssetItemtype } else { 'Computer' }
    if ($CmdbInput.Asset.Count) { $null = Invoke-GlpiApi -Session $Session -Path "$asset/$($Server.AssetId)" -Method Put -InputObject $CmdbInput.Asset }
    if ($CmdbInput.Container.Count) {
        if (-not $g.ContainerItemtype) { throw 'Container fields given but Glpi.ContainerItemtype is not set.' }
        if ($Server.ContainerId) {
            $null = Invoke-GlpiApi -Session $Session -Path "$($g.ContainerItemtype)/$($Server.ContainerId)" -Method Put -InputObject $CmdbInput.Container
        }
        else {
            $body = @{ items_id = $Server.AssetId; itemtype = $asset } + $CmdbInput.Container
            $null = Invoke-GlpiApi -Session $Session -Path $g.ContainerItemtype -Method Post -InputObject $body
        }
    }
}

function Get-CmdbStatusValue {
    # Display name of a lifecycle status (Glpi.StatusValues.Active / Retired).
    param([hashtable]$Config, [ValidateSet('Active', 'Retired')][string]$Kind)
    $v = $Config.Glpi.StatusValues
    if ($v -and $v[$Kind]) { return $v[$Kind] }
    if ($Kind -eq 'Active') { return 'In use' } else { return 'Retired' }
}

#endregion

#region Rules

function Test-RuleCondition {
    # Every key of When is a logical field; the value is a regex matched against the record (all must match).
    param([System.Collections.IDictionary]$When, [System.Collections.IDictionary]$Record)
    foreach ($k in $When.Keys) {
        $v = if ($Record.Contains($k)) { [string]$Record[$k] } else { '' }
        if ($v -notmatch $When[$k]) { return $false }
    }
    return $true
}

function Invoke-CmdbRule {
    <#
    .SYNOPSIS
        Applies declarative rules to a record: @{ When = @{ Field = 'regex' }; Set = @{ Field = 'value' } }.
        Rules run in order; later rules override earlier ones. Values "{Field}" are replaced with the record value.
    .OUTPUTS
        Names of the rules that matched.
    #>
    param([object[]]$Rules, [System.Collections.IDictionary]$Record)
    foreach ($r in @($Rules | Where-Object { $_ })) {
        $when = if ($r.When) { $r.When } else { @{} }
        if (-not (Test-RuleCondition -When $when -Record $Record)) { continue }
        foreach ($k in $r.Set.Keys) {
            $val = $r.Set[$k]
            if ($val -is [string]) {
                $val = [regex]::Replace($val, '\{(\w+)\}', { param($m) [string]$Record[$m.Groups[1].Value] })
            }
            $Record[$k] = $val
        }
        if ($r.Name) { $r.Name } else { ($r.Set.Keys -join ',') }
    }
}

#endregion
