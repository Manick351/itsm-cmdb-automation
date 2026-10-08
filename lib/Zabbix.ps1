# Zabbix JSON-RPC client. Requires lib\Common.ps1.
#
# Authentication (config Zabbix.Credential): if the credential user name is "token", the password part is an API
# token (Zabbix 5.4+); otherwise user.login is called with user/password. Zabbix 6.4+ receives the token in the
# Authorization: Bearer header, older versions in the "auth" property of the request.

function Connect-Zabbix {
    param([Parameter(Mandatory = $true)][hashtable]$Config)
    $z = $Config.Zabbix
    if (-not $z.Url) { throw 'Zabbix.Url is not set.' }
    $api = $z.Url.TrimEnd('/') -replace '/api_jsonrpc\.php$', ''
    $session = [PSCustomObject]@{
        Api = "$api/api_jsonrpc.php"; Token = $null; Bearer = $false; Version = $null; VersionNumber = 0
        SkipCert = [bool]$Config.Settings.SkipCertificateCheck; Id = 0; LoggedIn = $false
    }
    $session.Version = Invoke-ZabbixApi -Session $session -Method 'apiinfo.version' -Params @() -NoAuth
    $p = $session.Version.Split('.')
    $session.VersionNumber = [int]$p[0] * 100 + [int]$p[1]
    $session.Bearer = $session.VersionNumber -ge 604
    $cred = Get-ToolkitCredential -Name $z.Credential -CredentialPath $Config.Settings.CredentialPath
    if ($cred.UserName -eq 'token') {
        $session.Token = $cred.GetNetworkCredential().Password
    }
    else {
        $login = if ($session.VersionNumber -ge 504) { @{ username = $cred.UserName; password = $cred.GetNetworkCredential().Password } }
                 else { @{ user = $cred.UserName; password = $cred.GetNetworkCredential().Password } }
        $session.Token = Invoke-ZabbixApi -Session $session -Method 'user.login' -Params $login -NoAuth
        $session.LoggedIn = $true
    }
    return $session
}

function Disconnect-Zabbix {
    param([Parameter(Position = 0)]$Session)
    if ($Session -and $Session.LoggedIn) { try { $null = Invoke-ZabbixApi -Session $Session -Method 'user.logout' -Params @{} } catch { } }
}

function Invoke-ZabbixApi {
    param(
        [Parameter(Mandatory = $true)]$Session,
        [Parameter(Mandatory = $true)][string]$Method,
        $Params = @{},
        [switch]$NoAuth
    )
    $Session.Id++
    $body = [ordered]@{ jsonrpc = '2.0'; method = $Method; params = $Params; id = $Session.Id }
    $headers = @{}
    if (-not $NoAuth) {
        if ($Session.Bearer) { $headers.Authorization = "Bearer $($Session.Token)" } else { $body.auth = $Session.Token }
    }
    $json = $body | ConvertTo-Json -Depth 20 -Compress
    $resp = Invoke-ToolkitRest -Uri $Session.Api -Method Post -Headers $headers -ContentType 'application/json-rpc' `
        -Body ([Text.Encoding]::UTF8.GetBytes($json)) -SkipCertificateCheck:$Session.SkipCert
    $content = if ($resp.Content -is [byte[]]) { [Text.Encoding]::UTF8.GetString($resp.Content) } else { [string]$resp.Content }
    $r = $content | ConvertFrom-Json
    if ($r.error) { throw "Zabbix $Method failed: $($r.error.message) $($r.error.data)" }
    return $r.result
}

function Get-ZabbixGroupParam {
    # Zabbix 6.2 renamed selectGroups/groups to selectHostGroups/hostgroups.
    param($Session)
    if ($Session.VersionNumber -ge 602) { return @{ Select = 'selectHostGroups'; Property = 'hostgroups' } }
    return @{ Select = 'selectGroups'; Property = 'groups' }
}
