# Assigns the task to a fixed person or group: Params.To (string or list; the first one not absent wins).
param($Ticket, [hashtable]$Params, [hashtable]$Context)
$candidates = @($Params.To | Where-Object { $_ })
if (-not $candidates) { throw 'Assign handler needs Params.To' }
$who = @($candidates | Where-Object { $RouterAbsent -notcontains $_ })[0]
if (-not $who) { $who = $candidates[0] }
Set-TicketAssignee $Ticket $who $Params.Note
@{ Result = 'Assigned'; Note = $who }
