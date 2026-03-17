## Proteus Blocker

Blocks Proteus Design Suite internet access via Windows Firewall rules (exe-only) and optionally modifies the hosts file and blocks IP ranges.

### What it does

- **Firewall**: Creates block rules for discovered Proteus `.exe` files (inbound + outbound).
- **Hosts**: Adds Labcenter/Proteus domain entries to the Windows hosts file.
- **IP ranges**: Creates IP range block rules for known Labcenter subnets.
- **Services**: Detects Proteus-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\ProteusBlocker"
.\ProteusBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall + hosts + IP blocks.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files\Labcenter Electronics\Proteus*\`
- `C:\Program Files (x86)\Labcenter Electronics\Proteus*\`
- `C:\Proteus\`
- `C:\ProgramData\Labcenter\`
- `%LOCALAPPDATA%\Labcenter\`
- `%APPDATA%\Labcenter\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `ProteusBlocker_Logs/ProteusBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `ProteusBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `ProteusBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to confirm what will be affected.
- This tool changes **Firewall**, the **hosts file**, and may add **IP range rules**.

### Tool-specific notes

- Multi-version support via `Proteus*` scan folders.

### Uninstall

Run the script and select **Unblock** mode.

