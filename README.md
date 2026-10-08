# itsm-cmdb-automation

PowerShell automation around a GLPI CMDB, Zabbix and a service desk: server records created and retired from
tickets, CMDB kept in line with the hypervisor inventory, monitoring coverage, DNS leftovers of retired servers,
a rule-based ticket router, workload analytics and an on-call calendar.

> **🇷🇺 Кратко.** Автоматизация вокруг CMDB на GLPI, Zabbix и сервис-деска. Записи о серверах создаются и
> списываются по закрытым заявкам, CMDB сверяется с инвентарём гипервизоров, хосты без мониторинга заводятся в
> Zabbix, остатки списанных серверов ищутся в DNS и AD. Плюс маршрутизатор заявок (таблица правил + обработчики),
> аналитика нагрузки по заявкам и шаблон графика дежурств. Поля GLPI, правила и API сервис-деска описываются
> в конфиге, поэтому под другую CMDB или ITSM меняется конфиг, а не код. Всё, что меняет данные, по умолчанию
> работает как dry-run.

## Contents

| Area | Script | What it does | Changes anything? |
|---|---|---|---|
| CMDB | [cmdb/Add-GlpiServer.ps1](cmdb/Add-GlpiServer.ps1) | Provisioning request (form fields) -> CMDB record: field mapping, value mapping, declarative rules, duplicate / recently-retired / incomplete checks | Dry run unless `-Execute` |
| | [cmdb/Invoke-CmdbLifecycle.ps1](cmdb/Invoke-CmdbLifecycle.ps1) | Closed deployment tasks create records, closed decommission tasks retire them and disable the Zabbix host | Dry run unless `-Execute` |
| | [cmdb/Sync-GlpiInventory.ps1](cmdb/Sync-GlpiInventory.ps1) | CMDB vs hypervisor inventory: field drift (updated with `-Execute`), servers without a VM, VMs without a record | Dry run unless `-Execute` |
| | [cmdb/Sync-GlpiEmployees.ps1](cmdb/Sync-GlpiEmployees.ps1) | AD users -> GLPI object used for owners / administrators | Dry run unless `-Execute` |
| | [cmdb/Send-TemporaryServerReport.ps1](cmdb/Send-TemporaryServerReport.ps1) | Temporary servers past or near their planned retirement date, hardware filtered out | Mail only |
| | [cmdb/Get-RetiredDnsResidue.ps1](cmdb/Get-RetiredDnsResidue.ps1) | Retired servers vs AD-integrated DNS (A / PTR / CNAME over LDAP), ping/TCP, AD objects: verdict and exact cleanup actions | No |
| Monitoring | [monitoring/Sync-ZabbixFromCmdb.ps1](monitoring/Sync-ZabbixFromCmdb.ps1) | Active CMDB servers not in Zabbix -> hosts with groups / templates / tags chosen by rules | Dry run unless `-Execute` |
| | [monitoring/Test-ZabbixHostLifecycle.ps1](monitoring/Test-ZabbixHostLifecycle.ps1) | Zabbix configuration snapshot and host hygiene findings (disabled without reason, no template, missing group, ...) | No |
| Tickets | [tickets/Invoke-TicketRouter.ps1](tickets/Invoke-TicketRouter.ps1) | Route table + [handlers](tickets/handlers): automate what a script can do, round-robin the rest (absence aware) | Dry run unless `-Execute` |
| | [tickets/Get-TicketWorkload.ps1](tickets/Get-TicketWorkload.ps1) | What the queue consists of: by service, task type, month, assignee | No |
| Duty | [duty/New-DutyCalendar.ps1](duty/New-DutyCalendar.ps1) | Monthly on-call calendar XLSX (holidays, patch weekend, conflict highlighting, totals), WebDAV upload, mail | Writes a file |

Libraries in [lib/](lib): `Glpi.ps1` (REST client and a configurable CMDB layer with rules), `Zabbix.ps1` (JSON-RPC,
5.0 - 7.x), `Itsm.ps1` (service desk adapter: generic REST or JSON file), `Common.ps1` (config, credentials,
logging, mail, DNS), `Xlsx.ps1` (XLSX writer without Excel or ImportExcel).

## Design

**GLPI is a construction kit.** Every installation models servers differently: core Computer fields, the
Fields plugin, Generic Objects. The scripts never use GLPI field names. They work with logical fields (`Name`,
`Status`, `Ip`, `Environment`, `Owner`, `RetireDate`, ...) mapped in the config:

```powershell
Fields = @{
    Status      = @{ On = 'Asset'; Field = 'states_id'; Dropdown = 'State' }
    Environment = @{ On = 'Container'; Field = 'plugin_fields_environmentfielddropdowns_id'; Dropdown = 'PluginFieldsEnvironmentfieldDropdown' }
    Ip          = @{ On = 'Container'; Field = 'ipaddressfield' }
}
```

Dropdown names are resolved to ids on write (cached per run). The field map in
[examples/config.example.psd1](examples/config.example.psd1) is an example and has to be adapted to the
target GLPI.

**Rules instead of code.** Derived values are declarative and applied in order:

```powershell
Rules = @(
    @{ Name = 'backup-prod'; When = @{ Environment = '^Production' }; Set = @{ BackupPolicy = 'daily-30d' } }
    @{ Name = 'backup-prod-critical'; When = @{ Environment = '^Production'; Criticality = '^(High|Critical)$' }; Set = @{ BackupPolicy = 'hourly-30d-offsite' } }
)
```

The same `When` syntax selects Zabbix groups and templates (`Zabbix.HostRules`) and ticket routes
(`Router.Routes`).

**Any service desk.** `lib/Itsm.ps1` normalizes tickets (`Number`, `Title`, `Service`, `State`, `Assignee`,
`RequestNumber`, `Fields`). The `Rest` provider is described by URL templates and a field map. The `File`
provider reads a JSON export and writes actions to a journal, which is handy for demos and tests.

## Requirements

- Windows PowerShell 5.1 or PowerShell 7.
- GLPI 9.5 / 10.x with the REST API enabled (user API token or login, App-Token if configured).
- Zabbix 5.0 - 7.x (API token or service user).
- `Get-RetiredDnsResidue`, `Sync-GlpiEmployees` and the `AddToAdGroup` handler use ADSI and need a domain-joined
  Windows host. Everything else also runs on PowerShell 7 for Linux / macOS.

## Setup

```powershell
Copy-Item examples\config.example.psd1 config\config.psd1        # config\ is git-ignored
notepad config\config.psd1                                         # GLPI field map, rules, Zabbix, service desk

Set-Secret -Name glpi-api -Secret (Get-Credential api-user)        # password = GLPI personal API token
# or a DPAPI file for the same user and machine:
Get-Credential api-user | Export-Clixml config\credentials\glpi-api.xml
```

Every script has comment-based help: `Get-Help .\cmdb\Add-GlpiServer.ps1 -Full`.

## Examples

```powershell
# what would be created from a request, then create it
.\cmdb\Add-GlpiServer.ps1 -RequestPath .\examples\server-request.example.json
.\cmdb\Add-GlpiServer.ps1 -RequestPath .\req-1042.json -TicketNumber REQ-1042 -Execute

# nightly: tickets -> CMDB, inventory -> CMDB, CMDB -> Zabbix
.\cmdb\Invoke-CmdbLifecycle.ps1 -DisableMonitoring -Execute -SendMail
.\cmdb\Sync-GlpiInventory.ps1 -InventoryPath .\vm-inventory.csv -Execute
.\monitoring\Sync-ZabbixFromCmdb.ps1 -Execute

# every 5 minutes: the ticket router
.\tickets\Invoke-TicketRouter.ps1 -Execute

# DNS cleanup list for retired servers
.\cmdb\Get-RetiredDnsResidue.ps1 | Where-Object Verdict -ne 'Clean' | Out-GridView

# demo of the analytics on synthetic data
.\examples\New-SyntheticTickets.ps1 -Count 3000 -Path .\state\tickets.json
.\tickets\Get-TicketWorkload.ps1 -InputPath .\state\tickets.json
```

The hypervisor inventory for `Sync-GlpiInventory.ps1` can come from
[vmware-cloud-toolkit](https://github.com/Manick351/vmware-cloud-toolkit) (`Get-VMInventory.ps1`) or any CSV with
one row per VM. For scheduling, see `Register-ScriptTask.ps1` in
[windows-server-ops](https://github.com/Manick351/windows-server-ops).

`Add-GlpiServer.ps1` on a batch of requests (illustrative):

```text
Name                Status        AssetId Message
----                ------        ------- -------
srv-app-07          Created           100 asset 100, CMDB record 200
srv-db-02           AlreadyExists       1 active record 1
srv-recent-05       GracePeriod           retired on 2026-10-01, less than 30 days ago; create manually if intended
srv-r-00            Incomplete            required field(s) empty: Ip, Os, Environment, Owner
ws-0042             Skipped               name matches Request.NameRules.SkipPattern
```

`Invoke-TicketRouter.ps1` (illustrative):

```text
Ticket    Title                           Route        Result      Note
------    -----                           -----        ------      ----
TSK-99100 Add server to CMDB              CMDB record  Closed      CMDB: Created srv-new-20 asset 102
TSK-99101 Add server to CMDB              CMDB record  Handed over GracePeriod: retired on 2026-10-01 ...
TSK-99102 Information: maintenance window FYI tasks    Closed      Acknowledged (TSK-99102)
TSK-50049 Grant admin access              round-robin  Assigned    alice
TSK-50050 Grant admin access              round-robin  Assigned    bob
```

## Safety

- Everything that writes to GLPI, Zabbix, AD or the service desk is a dry run unless `-Execute` is given; the
  router's dry run even simulates the round-robin without touching its state file.
- Automation never fails silently: an incomplete or unexpected request is handed to a person with the reason.
- A server re-created within `Glpi.RetiredGraceDays` of its retirement is not created automatically.
- Form values used for AD changes are validated (`AllowedGroups` pattern, account and group must exist).
- Credentials come from SecretManagement or DPAPI files, never from the config or the command line.

## Tested

Offline against mock GLPI, Zabbix and service desk APIs: record creation and all refusal paths, inventory sync,
temporary server report, lifecycle (commission and decommission with Zabbix disable), Zabbix sync and audit,
the router with both providers and all handlers except `AddToAdGroup`, workload analytics on 9 000 synthetic
tickets, and the duty calendar rendered in LibreOffice. The ADSI-based parts (`Get-RetiredDnsResidue`,
`Sync-GlpiEmployees`, `AddToAdGroup`) need a domain and were not covered.

## Limitations

- The GLPI layer handles one asset itemtype plus one Fields plugin container; other relations (network ports,
  software, Item_OperatingSystem) are not written.
- `Get-RetiredDnsResidue.ps1` reads AD-integrated zones only (no BIND / cloud DNS), IPv4 only.
- The ticket router processes tasks one after another; it is meant to run every few minutes.
- `New-DutyCalendar.ps1` builds one sheet per run (one month).

## License

[MIT](LICENSE)
