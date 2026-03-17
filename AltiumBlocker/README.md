## Altium Designer Blocker

Selective network blocking for Altium Designer that preserves Component Search and login while blocking telemetry and licensing endpoints.

### What it does

- **Firewall**: Creates block rules for discovered Altium `.exe` files (inbound + outbound) except allowlisted components.
- **Hosts**: Adds a selective block list (telemetry/licensing/update domains only).
- **Services**: Detects Altium-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\AltiumBlocker"
.\AltiumDesignerBlocker.ps1
```

### Operation modes

- **Block**: Apply selective blocking rules.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files\Altium\AD*\`
- `C:\Program Files (x86)\Altium\AD*\`
- `C:\ProgramData\Altium\`
- `%LOCALAPPDATA%\Altium\`
- `%APPDATA%\Altium\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `AltiumBlocker_Logs/AltiumBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `AltiumBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `AltiumBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to validate that required Altium online features remain reachable.

### Tool-specific notes

- **Preserved online features**: Component Search, Manufacturer Part Search, Octopart/PartQuest/AltiumLive access, authentication/login.
- **Selective hosts blocking**: Only telemetry/licensing/update domains are blocked. Component and login endpoints are intentionally preserved.
- **Allowlist-based filtering**: Files matching allowlisted keywords are not blocked.

### Uninstall

Run the script and select **Unblock** mode.

