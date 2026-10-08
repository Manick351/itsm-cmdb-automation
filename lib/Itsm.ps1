# Generic ITSM adapter. Requires lib\Common.ps1.
#
# Service desks differ in everything but the idea: tasks with a number, a title, a service, a state, an assignee
# and the form fields of the parent request. Scripts use normalized ticket objects:
#
#   Id, Number, Title, Service, State, Team, Assignee, Created, Closed, RequestNumber, Fields (hashtable), Raw
#
# Providers (config Itsm.Provider):
#   Rest  any REST API described by URL templates and a field map (Itsm.Endpoints, Itsm.FieldMap)
#   File  a JSON file with tickets (demo, tests, offline analysis); write operations go to a journal file

function Connect-Itsm {
    param([Parameter(Mandatory = $true)][hashtable]$Config)
    $i = $Config.Itsm
    $provider = if ($i.Provider) { $i.Provider } else { 'Rest' }
    $s = [PSCustomObject]@{ Provider = $provider; Config = $i; Headers = @{ Accept = 'application/json' }; SkipCert = [bool]$Config.Settings.SkipCertificateCheck; Journal = $null; Data = $null }
    if ($provider -eq 'File') {
        $path = $i.File
        if (-not [IO.Path]::IsPathRooted($path)) { $path = Join-Path (Split-Path -Parent $PSScriptRoot) $path }
        $s.Data = @(Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)
        if (-not (Test-Path -LiteralPath $Config.Settings.StatePath)) { New-Item -ItemType Directory -Path $Config.Settings.StatePath -Force | Out-Null }
        $s.Journal = Join-Path $Config.Settings.StatePath 'itsm-journal.jsonl'
        return $s
    }
    if (-not $i.BaseUrl) { throw 'Itsm.BaseUrl is not set.' }
    $cred = Get-ToolkitCredential -Name $i.Credential -CredentialPath $Config.Settings.CredentialPath
    $header = if ($i.AuthHeader) { $i.AuthHeader } else { 'Authorization' }
    $format = if ($i.AuthFormat) { $i.AuthFormat } else { 'Bearer {0}' }
    $s.Headers[$header] = $format -f $cred.GetNetworkCredential().Password
    return $s
}

function Expand-ItsmTemplate {
    # Replaces {name} placeholders with URL-encoded values (or raw values for bodies).
    param([string]$Template, [hashtable]$Values, [switch]$Raw)
    return [regex]::Replace($Template, '\{(\w+)\}', {
            param($m)
            $v = [string]$Values[$m.Groups[1].Value]
            if ($Raw) { $v } else { [Uri]::EscapeDataString($v) }
        })
}

function Get-PathValue {
    # Reads a dotted path ("request.service.title") from an object.
    param($Object, [string]$Path)
    if (-not $Path) { return $null }
    $cur = $Object
    foreach ($p in $Path -split '\.') { if ($null -eq $cur) { return $null }; $cur = $cur.$p }
    return $cur
}

function ConvertTo-ItsmTicket {
    param($Raw, [hashtable]$FieldMap)
    $map = @{ Id = 'id'; Number = 'number'; Title = 'title'; Service = 'service'; State = 'state'; Team = 'team'; Assignee = 'assignee'
        Created = 'created'; Closed = 'closed'; RequestNumber = 'requestNumber'; Fields = 'fields' }
    if ($FieldMap) { foreach ($k in $FieldMap.Keys) { $map[$k] = $FieldMap[$k] } }
    $fields = @{}
    $f = Get-PathValue $Raw $map.Fields
    if ($f) {
        if ($f -is [System.Collections.IDictionary]) { foreach ($k in $f.Keys) { $fields[$k] = $f[$k] } }
        elseif ($f -is [array]) { foreach ($x in $f) { if ($x.name) { $fields[[string]$x.name] = $x.value } } }   # [{name, value}] style
        else { foreach ($p in $f.PSObject.Properties) { $fields[$p.Name] = $p.Value } }
    }
    $dt = { param($v) if ($v) { try { [datetime]$v } catch { $null } } }
    [PSCustomObject]@{
        Id            = [string](Get-PathValue $Raw $map.Id)
        Number        = [string](Get-PathValue $Raw $map.Number)
        Title         = [string](Get-PathValue $Raw $map.Title)
        Service       = [string](Get-PathValue $Raw $map.Service)
        State         = [string](Get-PathValue $Raw $map.State)
        Team          = [string](Get-PathValue $Raw $map.Team)
        Assignee      = [string](Get-PathValue $Raw $map.Assignee)
        Created       = & $dt (Get-PathValue $Raw $map.Created)
        Closed        = & $dt (Get-PathValue $Raw $map.Closed)
        RequestNumber = [string](Get-PathValue $Raw $map.RequestNumber)
        Fields        = $fields
        Raw           = $Raw
    }
}

function Get-ItsmTicket {
    <#
    .SYNOPSIS
        Returns normalized tickets filtered by state, team, title regex and closing date.
    .DESCRIPTION
        Rest: Itsm.Endpoints.FindTasks is called for every State x Team combination with {state}, {team},
        {offset}, {limit} placeholders and paged until a short page. The response is either an array or an
        object with an "items" property. Filtering by title/date is local.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [string[]]$State = @(''),
        [string[]]$Team = @(''),
        [string]$TitlePattern,
        [datetime]$ClosedAfter,
        [switch]$UnassignedOnly
    )
    $all = if ($Session.Provider -eq 'File') {
        foreach ($r in $Session.Data) { ConvertTo-ItsmTicket $r $Session.Config.FieldMap }
    }
    else {
        $limit = if ($Session.Config.PageSize) { [int]$Session.Config.PageSize } else { 200 }
        foreach ($st in $State) {
            foreach ($tm in $Team) {
                $offset = 0
                do {
                    $path = Expand-ItsmTemplate $Session.Config.Endpoints.FindTasks @{ state = $st; team = $tm; offset = $offset; limit = $limit }
                    $resp = Invoke-ToolkitRest -Uri "$($Session.Config.BaseUrl.TrimEnd('/'))/$path" -Headers $Session.Headers -SkipCertificateCheck:$Session.SkipCert
                    $data = $resp.Content | ConvertFrom-Json
                    $items = if ($data.PSObject.Properties['items']) { @($data.items) } else { @($data) }
                    foreach ($r in $items) { ConvertTo-ItsmTicket $r $Session.Config.FieldMap }
                    $offset += $limit
                } while ($items.Count -eq $limit)
            }
        }
    }
    foreach ($t in $all) {
        if ($State -and $State[0] -and $t.State -notin $State) { continue }
        if ($Team -and $Team[0] -and $t.Team -notin $Team) { continue }
        if ($TitlePattern -and $t.Title -notmatch $TitlePattern) { continue }
        if ($PSBoundParameters.ContainsKey('ClosedAfter') -and (-not $t.Closed -or $t.Closed -lt $ClosedAfter)) { continue }
        if ($UnassignedOnly -and $t.Assignee) { continue }
        $t
    }
}

function Invoke-ItsmAction {
    <#
    .SYNOPSIS
        Performs a write action (Assign, Comment, Close) described by Itsm.Endpoints.<Action> = @{ Method; Path; Body }.
        File provider: the action is appended to the journal instead.
    #>
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][ValidateSet('Assign', 'Comment', 'Close')][string]$Action,
        [Parameter(Mandatory = $true)]$Ticket,
        [hashtable]$Values = @{}
    )
    $vals = @{ id = $Ticket.Id; number = $Ticket.Number }
    foreach ($k in $Values.Keys) { $vals[$k] = $Values[$k] }
    if ($Session.Provider -eq 'File') {
        $entry = [ordered]@{ time = (Get-Date).ToString('s'); action = $Action; ticket = $Ticket.Number }
        foreach ($k in $Values.Keys) { $entry[$k] = $Values[$k] }
        Add-Content -LiteralPath $Session.Journal -Value ($entry | ConvertTo-Json -Compress) -Encoding UTF8
        return
    }
    $ep = $Session.Config.Endpoints[$Action]
    if (-not $ep) { throw "Itsm.Endpoints.$Action is not configured." }
    $body = $null
    if ($ep.Body) {
        $b = @{}
        foreach ($k in $ep.Body.Keys) { $b[$k] = Expand-ItsmTemplate ([string]$ep.Body[$k]) $vals -Raw }
        $body = [Text.Encoding]::UTF8.GetBytes(($b | ConvertTo-Json -Depth 5))
    }
    $method = if ($ep.Method) { $ep.Method } else { 'Post' }
    $uri = "$($Session.Config.BaseUrl.TrimEnd('/'))/$(Expand-ItsmTemplate $ep.Path $vals)"
    $null = Invoke-ToolkitRest -Uri $uri -Method $method -Headers $Session.Headers -Body $body -ContentType 'application/json' -SkipCertificateCheck:$Session.SkipCert
}
