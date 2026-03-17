## MultiSim Blocker

Blocks NI MultiSim internet access via Windows Firewall rules (exe-only) and optionally modifies the hosts file and blocks IP ranges.

### What it does

- **Firewall**: Creates block rules for discovered NI/MultiSim `.exe` files (inbound + outbound).
- **Hosts**: Adds NI domain entries to the Windows hosts file.
- **IP ranges**: Creates IP range block rules for known NI subnets.
- **Services**: Detects NI-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\MultiSimBlocker"
.\MultiSimBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall + hosts + IP blocks.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files (x86)\National Instruments\`
- `C:\Program Files\National Instruments\`
- `C:\ProgramData\National Instruments\`
- `%LOCALAPPDATA%\National Instruments\`
- `%APPDATA%\National Instruments\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `MultiSimBlocker_Logs/MultiSimBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `MultiSimBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `MultiSimBlocker_Logs/MultiSimBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to confirm what will be affected.
- This tool changes **Firewall**, the **hosts file**, and may add **IP range rules**.

### Tool-specific notes

- Targets NI licensing/update/telemetry domains and common NI install locations.

### Uninstall

Run the script and select **Unblock** mode.
