<#
    Grants access by adding an account to an AD group, then comments and closes the task.

    Params:
      AccountField  form field with the sAMAccountName (DOMAIN\ prefix is stripped)
      Group         group name, may use placeholders: 'RDP-{Field:Server name}', 'APP-{Service}-Users'
      AllowedGroups regex the resulting group name must match (guard against typos in requests)
      Note          closing note, placeholders allowed

    Uses ADSI (no ActiveDirectory module). Any problem - empty account, group not found, group outside
    AllowedGroups - hands the task to a person with the reason instead of failing silently.
#>
param($Ticket, [hashtable]$Params, [hashtable]$Context)
$account = ([string]$Ticket.Fields[$Params.AccountField]).Trim() -replace '^.*\\', ''
$group = Expand-RouteValue $Params.Group $Ticket
$problem = $null
if (-not $account) { $problem = "form field '$($Params.AccountField)' is empty" }
elseif (-not $group -or $group -match '\{|\}|^\W|\W$') { $problem = "group name '$group' could not be built from the request" }
elseif ($Params.AllowedGroups -and $group -notmatch $Params.AllowedGroups) { $problem = "group '$group' is outside AllowedGroups" }

function Find-AdObject([string]$Filter) {
    $ds = New-Object DirectoryServices.DirectorySearcher
    $ds.Filter = $Filter
    $r = $ds.FindOne(); $ds.Dispose()
    return $r
}
if (-not $problem) {
    # RFC 4515 escaping of filter values
    $esc = { param($s) (([string]$s).ToCharArray() | ForEach-Object { if ('\*()'.Contains([string]$_) -or [int]$_ -eq 0) { '\{0:x2}' -f [int]$_ } else { [string]$_ } }) -join '' }
    $u = Find-AdObject "(&(objectCategory=person)(sAMAccountName=$(& $esc $account)))"
    $g = Find-AdObject "(&(objectCategory=group)(sAMAccountName=$(& $esc $group)))"
    if (-not $u) { $problem = "account '$account' not found in AD" }
    elseif (-not $g) { $problem = "group '$group' not found in AD" }
}
if ($problem) {
    Set-TicketAssignee $Ticket (Get-NextAssignee) "Not automated: $problem"
    return @{ Result = 'Handed over'; Note = $problem }
}

$already = @($g.Properties['member']) -contains [string]$u.Properties['distinguishedname'][0]
if (-not $already) {
    Write-Log "  $($Ticket.Number): add $account to $group"
    if ($Context.Execute) {
        $entry = $g.GetDirectoryEntry()
        $entry.Properties['member'].Add([string]$u.Properties['distinguishedname'][0]) | Out-Null
        $entry.CommitChanges()
    }
}
$note = if ($Params.Note) { Expand-RouteValue $Params.Note $Ticket } else { "Account $account added to group $group" }
if ($already) { $note = "Account $account is already a member of $group" }
Add-TicketComment $Ticket $note
Close-Ticket $Ticket 'Access granted'
@{ Result = 'Closed'; Note = $note }
