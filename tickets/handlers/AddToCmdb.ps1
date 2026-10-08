<#
    Creates the CMDB record for a new server from the request form (cmdb\Add-GlpiServer.ps1) and closes the task.

    Params:
      ReturnTo  'Requester' (default: the parent request's assignee, form field Params.RequesterField) or a person
    Created / AlreadyExists -> close; GracePeriod / Skipped / Error -> hand back with the reason.
#>
param($Ticket, [hashtable]$Params, [hashtable]$Context)
$add = Join-Path $Context.RepoRoot 'cmdb\Add-GlpiServer.ps1'
$r = & $add -Request $Ticket.Fields -TicketNumber $Ticket.RequestNumber -Execute:$Context.Execute -ConfigPath $Context.ConfigPath | Select-Object -Last 1
switch ($r.Status) {
    { $_ -in 'Created', 'AlreadyExists', 'WouldCreate' } {
        $note = "CMDB: $($r.Status) $($r.Name) $($r.Message)".Trim()
        if ($r.Warnings) { $note += "`nWarnings: " + ($r.Warnings -join '; ') }
        Close-Ticket $Ticket $note
        return @{ Result = 'Closed'; Note = $note }
    }
    default {
        $to = if ($Params.ReturnTo -and $Params.ReturnTo -ne 'Requester') { $Params.ReturnTo }
              elseif ($Params.RequesterField -and $Ticket.Fields[$Params.RequesterField]) { [string]$Ticket.Fields[$Params.RequesterField] }
              else { Get-NextAssignee }
        Set-TicketAssignee $Ticket $to "CMDB automation: $($r.Status) - $($r.Message)"
        return @{ Result = 'Handed over'; Note = "$($r.Status): $($r.Message)" }
    }
}
