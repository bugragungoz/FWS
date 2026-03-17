## PLECS Blocker

PowerShell script that blocks PLECS executables via Windows Firewall (inbound and outbound).

### What it does

- **Firewall**: Creates block rules for discovered PLECS-related `.exe` files (inbound + outbound).
- **Outputs**: Generates logs, backups, and a report under the script directory.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\PlecsBlocker"
.\PLECSBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall rules (strict full block).
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool.
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- Strict full block (no exceptions): inbound + outbound for discovered `.exe` files.
- Generates logs, backups, and a report under the script directory during execution.

### Outputs

- **Logs**: `PLECSBlocker_Logs/PLECSBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `PLECSBlocker_Backups/FirewallRules_*.json`
- **Reports**: `PLECSBlocker_Logs/PLECSBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to validate scan results.

### Tool-specific notes

- This tool is a strict full block with no exceptions.

### Uninstall

Run the script and select **Unblock** mode.

