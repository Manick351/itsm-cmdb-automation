# Closes informational tasks with a fixed note (Params.Note).
param($Ticket, [hashtable]$Params, [hashtable]$Context)
$note = if ($Params.Note) { Expand-RouteValue $Params.Note $Ticket } else { 'Closed automatically' }
Close-Ticket $Ticket $note
@{ Result = 'Closed'; Note = $note }
