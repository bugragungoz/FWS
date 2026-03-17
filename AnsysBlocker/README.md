## ANSYS Blocker

Blocks ANSYS network traffic via Windows Firewall while preserving localhost license access.

### What it does

- **Firewall**: Creates block rules for discovered ANSYS-related `.exe` files (inbound + outbound).
- **Localhost exception**: Allows only `127.0.0.1` for local license verification.
- **Outputs**: Generates logs, backups, and a report under the script directory.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\AnsysBlocker"
.\ANSYSBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall rules with localhost exception.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool.
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

The tool scans common ANSYS installation roots and also discovers keyword-matching directories under standard Windows locations.

### Outputs

- **Logs**: `ANSYSBlocker_Logs/ANSYSBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `ANSYSBlocker_Backups/FirewallRules_*.json`
- **Reports**: `ANSYSBlocker_Logs/ANSYSBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to validate scan results.

### Tool-specific notes

- Uses a remote-address strategy that blocks all IPs except `127.0.0.1` (default license: `1055@localhost`).

### Uninstall

Run the script and select **Unblock** mode.

