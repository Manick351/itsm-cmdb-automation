# itsm-cmdb-automation configuration example. Copy to config\config.psd1 (git-ignored) and adjust.
# Every "Credential" value is the NAME of a credential (SecretManagement secret, or
# config\credentials\<name>.xml made with  Get-Credential | Export-Clixml ...). API tokens: put the token
# into the password; for Zabbix use the user name "token" to mark an API token.
@{
    Settings = @{
        LogPath              = 'C:\Ops\itsm-cmdb-automation\logs'
        ReportPath           = 'C:\Ops\itsm-cmdb-automation\out'
        StatePath            = 'C:\Ops\itsm-cmdb-automation\state'     # router rotation, ticket cache, journal
        SkipCertificateCheck = $false
    }

    Mail = @{
        SmtpServer = 'smtp.contoso.local'; Port = 587; UseSsl = $true
        From       = 'cmdb-automation@contoso.local'
        To         = @('it-ops@contoso.local')
        Credential = 'smtp-relay'
    }

    # ---------------------------------------------------------------------------------------------------------
    # GLPI. The field map is an EXAMPLE: GLPI is a construction kit and every installation models servers
    # differently. Map the logical names used by the scripts to your fields. On = Asset (core Computer
    # fields) or Container (a Fields plugin container); Dropdown = itemtype of the dropdown.
    Glpi = @{
        Url               = 'https://glpi.contoso.local'
        Credential        = 'glpi-api'          # password = personal API token (AuthMode UserToken)
        AuthMode          = 'UserToken'         # or Basic (user/password)
        AppToken          = 'glpi-app-token'    # API client App-Token, if required by the GLPI setup
        AssetItemtype     = 'Computer'
        ContainerItemtype = 'PluginFieldsComputerserverinfo'
        ContainerRequired = $true               # only computers with a "server info" record are servers
        RetiredGraceDays  = 30
        StatusValues      = @{ Active = 'In production'; Retired = 'Retired' }
        Fields            = @{
            Name              = @{ On = 'Asset'; Field = 'name' }
            Status            = @{ On = 'Asset'; Field = 'states_id'; Dropdown = 'State' }
            Location          = @{ On = 'Asset'; Field = 'locations_id'; Dropdown = 'Location' }
            TechGroup         = @{ On = 'Asset'; Field = 'groups_id_tech'; Dropdown = 'Group' }
            Serial            = @{ On = 'Asset'; Field = 'serial' }
            Comment           = @{ On = 'Asset'; Field = 'comment' }
            Fqdn              = @{ On = 'Container'; Field = 'fqdnfield' }
            Ip                = @{ On = 'Container'; Field = 'ipaddressfield' }
            Os                = @{ On = 'Container'; Field = 'plugin_fields_osfielddropdowns_id'; Dropdown = 'PluginFieldsOsfieldDropdown' }
            Environment       = @{ On = 'Container'; Field = 'plugin_fields_environmentfielddropdowns_id'; Dropdown = 'PluginFieldsEnvironmentfieldDropdown' }
            Criticality       = @{ On = 'Container'; Field = 'plugin_fields_criticalityfielddropdowns_id'; Dropdown = 'PluginFieldsCriticalityfieldDropdown' }
            BackupPolicy      = @{ On = 'Container'; Field = 'plugin_fields_backuppolicyfielddropdowns_id'; Dropdown = 'PluginFieldsBackuppolicyfieldDropdown' }
            Owner             = @{ On = 'Container'; Field = 'users_id_ownerfield'; Dropdown = 'User' }
            Cpu               = @{ On = 'Container'; Field = 'cpucountfield' }
            RamMB             = @{ On = 'Container'; Field = 'rammbfield' }
            Disks             = @{ On = 'Container'; Field = 'disksfield' }
            Description       = @{ On = 'Container'; Field = 'descriptionfield' }
            IsVirtual         = @{ On = 'Container'; Field = 'isvirtualfield' }
            CommissionDate    = @{ On = 'Container'; Field = 'commissiondatefield' }
            RetireDate        = @{ On = 'Container'; Field = 'retireddatefield' }
            PlannedRetireDate = @{ On = 'Container'; Field = 'plannedretirementdatefield' }
            SupportEnd        = @{ On = 'Container'; Field = 'supportenddatefield' }
        }
    }

    # Provisioning request (form) -> logical fields. Add-GlpiServer, Invoke-CmdbLifecycle, router AddToCmdb.
    Request = @{
        FieldMap      = @{
            Name = 'Server name'; Ip = 'IP'; Os = 'Operating System'; Environment = 'Used for'; Criticality = 'Impact'
            Owner = 'Business owner'; Location = 'Location'; Cpu = 'CPU (cores)'; RamGB = 'RAM (GB)'
            Description = 'Server description'; PlannedRetireDate = 'Retirement date (temporary servers)'
        }
        ValueMap      = @{
            Os = @{ '^Windows Server 2022' = 'Windows Server 2022'; '^Ubuntu' = 'Linux - Ubuntu'; '^(Red Hat|RHEL)' = 'Linux - RHEL' }
        }
        Disks         = @{ SizePattern = 'Data disk #{n} size (GB)'; TypePattern = 'Data disk #{n} type'; SystemDiskGB = @{ Windows = 80; Default = 40 } }
        Defaults      = @{ IsVirtual = 1; TechGroup = 'Server team' }
        Required      = @('Ip', 'Os', 'Environment', 'Owner')
        NameRules     = @{ SkipPattern = '^ws-' }               # workstation names are not servers
        DefaultDomain = 'contoso.local'
    }

    # Declarative rules, applied in order (later rules override earlier ones). When: logical field -> regex.
    Rules = @(
        @{ Name = 'backup-test'; When = @{ Environment = '^(Test|Development)' }; Set = @{ BackupPolicy = 'daily-7d' } }
        @{ Name = 'backup-prod'; When = @{ Environment = '^Production' }; Set = @{ BackupPolicy = 'daily-30d' } }
        @{ Name = 'backup-prod-critical'; When = @{ Environment = '^Production'; Criticality = '^(High|Critical)$' }; Set = @{ BackupPolicy = 'hourly-30d-offsite' } }
        @{ Name = 'siteb-group'; When = @{ Location = 'SiteB' }; Set = @{ TechGroup = 'Server team SiteB' } }
    )

    # Owners object filled from AD (Sync-GlpiEmployees). Example: a Generic Object "Person".
    Employees = @{
        Itemtype   = 'PluginGenericobjectPerson'
        SearchBase = 'OU=Users,DC=contoso,DC=local'
        KeyField   = 'contact'
        Map        = @{ name = '{displayName} ({sAMAccountName})'; serial = '{sAMAccountName}'; otherserial = '{mail}'; contact = '{objectSid}' }
    }

    # Hypervisor inventory CSV -> CMDB fields (Sync-GlpiInventory).
    Inventory = @{
        Path           = 'C:\Ops\vmware-cloud-toolkit\out\vm-inventory.csv'
        NameColumn     = 'Name'
        PowerColumn    = 'Power'
        LocationColumn = 'Location'
        ExcludePattern = '(clone|restore|_old)$'
        FieldMap       = @{
            Ip = 'IP address'; Cpu = 'vCPU'; Location = 'Location'
            RamMB = @{ Column = 'RAM GB'; Multiply = 1024 }
            Os = @{ Column = 'OS'; Map = @{ 'Windows Server 2022' = 'Windows Server 2022'; 'Ubuntu' = 'Linux - Ubuntu' } }
        }
    }

    TemporaryReport = @{ To = @('server-team@contoso.local'); HardwareFields = @('Serial', 'SupportEnd') }

    # ---------------------------------------------------------------------------------------------------------
    Zabbix = @{
        Url          = 'https://zabbix.contoso.local/zabbix'
        Credential   = 'zabbix-api'               # user "token" + API token, or a service user
        AgentPort    = 10050
        SkipPattern  = '^(fw|lb)-'
        CommonGroups = @('Servers')
        HostRules    = @(
            @{ Name = 'windows'; When = @{ Os = 'Windows' }; Groups = @('Windows servers'); Templates = @('Windows by Zabbix agent'); Tags = @{ os_family = 'Windows' } }
            @{ Name = 'linux'; When = @{ Os = 'Linux' }; Groups = @('Linux servers'); Templates = @('Linux by Zabbix agent'); Tags = @{ os_family = 'Linux' } }
        )
        Audit        = @{ RequiredGroups = @('Servers'); NamePattern = '^[a-z]{2,5}-[a-z0-9-]+$'; StaleDisabledDays = 30 }
    }

    # ---------------------------------------------------------------------------------------------------------
    # Service desk adapter. Rest: describe your API with URL templates and a field map. File: JSON tickets.
    Itsm = @{
        Provider   = 'Rest'
        BaseUrl    = 'https://servicedesk.contoso.local/api/v1'
        Credential = 'itsm-api'
        AuthHeader = 'Authorization'; AuthFormat = 'Bearer {0}'
        PageSize   = 200
        Endpoints  = @{
            FindTasks = 'tasks?team={team}&state={state}&offset={offset}&limit={limit}'
            Assign    = @{ Method = 'Patch'; Path = 'tasks/{id}'; Body = @{ assignee = '{assignee}' } }
            Comment   = @{ Method = 'Post'; Path = 'tasks/{id}/comments'; Body = @{ text = '{text}' } }
            Close     = @{ Method = 'Patch'; Path = 'tasks/{id}'; Body = @{ state = 'resolved'; resolution = '{text}' } }
        }
        FieldMap   = @{ Id = 'id'; Number = 'number'; Title = 'title'; Service = 'request.service'; State = 'state'; Team = 'team'
            Assignee = 'assignee'; Created = 'created'; Closed = 'closed'; RequestNumber = 'request.number'; Fields = 'request.fields' }
        # Provider = 'File'; File = 'state\tickets.json'; FieldMap = @{}    # demo with examples\New-SyntheticTickets.ps1
    }

    Lifecycle = @{
        Teams                 = @('server-team')
        ClosedStates          = @('closed', 'resolved')
        CommissionTitle       = '^Server Deployment$'
        DecommissionTitle     = '^Shutdown of server$'
        DecommissionNameField = 'Server name'
        IgnoreRequests        = @('REQ-000001')        # test requests, known exceptions
        ProtectedServers      = @()                   # never retired automatically
    }

    Router = @{
        Teams             = @('server-team')
        States            = @('registered', 'inprogress')
        MaxTickets        = 100
        Assignees         = @('alice', 'bob', 'carol', 'dave')
        ExcludedAssignees = @()                      # for example a person on a project this month
        AbsencePath       = 'state\absence.csv'      # Name,From,To
        Routes            = @(
            @{ Name = 'CMDB record'; When = @{ Title = '^Add server to CMDB$' }; Handler = 'AddToCmdb'; Params = @{ ReturnTo = 'Requester'; RequesterField = 'Requester' } }
            @{ Name = 'RDP access'; When = @{ Title = '^Grant RDP access$'; 'Field:Server name' = '^srv-' }; Handler = 'AddToAdGroup'
                Params = @{ AccountField = 'Windows account'; Group = 'RDP-{Field:Server name}'; AllowedGroups = '^RDP-srv-[a-z0-9-]+$' } }
            @{ Name = 'Backups'; When = @{ Title = 'backup' }; Handler = 'Assign'; Params = @{ To = @('erin', 'frank') } }
            @{ Name = 'FYI tasks'; When = @{ Title = '^Information:' }; Handler = 'Close'; Params = @{ Note = 'Acknowledged ({Number})' } }
        )
        Default           = @{ Name = 'round-robin'; Handler = 'RoundRobin' }
    }

    Workload = @{ Teams = @('server-team'); States = @('closed', 'resolved') }

    Duty = @{
        PeoplePath       = 'examples\duty-people.example.csv'
        HolidaySource    = 'Nager'                   # IsDayOff (RU), Nager, File, None
        CountryCode      = 'DE'
        WebDavUrl        = 'https://cloud.contoso.local/remote.php/dav/files/svc-duty/Duty'
        WebDavCredential = 'webdav-duty'
        ShareLink        = 'https://cloud.contoso.local/apps/files/?dir=/Duty'
        To               = @('server-team@contoso.local')
    }
}
