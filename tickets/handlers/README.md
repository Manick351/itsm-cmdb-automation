# Ticket router handlers

A handler is a script `<Name>.ps1` in this folder, referenced from `Router.Routes[].Handler` in the config.

```powershell
param($Ticket, [hashtable]$Params, [hashtable]$Context)
# $Ticket   normalized ticket: Id, Number, Title, Service, State, Assignee, RequestNumber, Fields (form fields)
# $Params   the route's Params hashtable from the config
# $Context  Config, ConfigPath, Execute (bool), RepoRoot
#
# Helpers provided by Invoke-TicketRouter.ps1 (they respect the dry run):
#   Get-NextAssignee [-Pool <names>]      next available person (round-robin, absence aware)
#   Set-TicketAssignee $Ticket <who> [note]
#   Add-TicketComment  $Ticket <text>
#   Close-Ticket       $Ticket <note>
#   Expand-RouteValue  <template> $Ticket  '{Field:Server name}', '{Number}', '{Service}'
#   Write-Log <text> [-Level Info|Warn|Error|Ok]
@{ Result = 'Closed'; Note = 'what was done' }
```

Rules of thumb:

- Never fail silently. If the request is incomplete or something unexpected happens, hand the task to a person
  with the reason (`Set-TicketAssignee $Ticket (Get-NextAssignee) "Not automated: ..."`).
- Check `$Context.Execute` before changing anything outside the service desk (AD, CMDB, hypervisors).
- Validate everything taken from a form field (allowed patterns, existence) before using it.
