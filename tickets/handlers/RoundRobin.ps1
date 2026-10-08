# Assigns the task to the next available person (Router.Assignees, or Params.Pool for a sub-team).
param($Ticket, [hashtable]$Params, [hashtable]$Context)
$pool = if ($Params.Pool) { [string[]]$Params.Pool } else { [string[]]$Context.Config.Router.Assignees }
$who = Get-NextAssignee -Pool $pool
Set-TicketAssignee $Ticket $who $Params.Note
@{ Result = 'Assigned'; Note = $who }
