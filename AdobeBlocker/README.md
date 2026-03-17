## Adobe Blocker

Blocks Adobe application internet access via Windows Firewall rules (exe-only) and optionally modifies the hosts file and blocks IP ranges.

### What it does

- **Firewall**: Creates block rules for discovered Adobe-related `.exe` files (inbound + outbound).
- **Hosts**: Adds domain entries to the Windows hosts file (activation/licensing/telemetry).
- **IP ranges**: Creates IP range block rules for known Adobe subnets.
- **Services**: Detects Adobe-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\AdobeBlocker"
.\AdobeBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall + hosts + IP blocks.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files\Adobe\`
- `C:\Program Files (x86)\Adobe\`
- `C:\Program Files\Common Files\Adobe\`
- `C:\Program Files (x86)\Common Files\Adobe\`
- `C:\ProgramData\Adobe\`
- `%LOCALAPPDATA%\Adobe\`
- `%APPDATA%\Adobe\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `AdobeBlocker_Logs/AdobeBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `AdobeBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `AdobeBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to confirm what will be affected.
- This tool changes **Firewall**, the **hosts file**, and adds **IP range rules**.

### Tool-specific notes

- Hosts entries are written as `0.0.0.0 <domain>` and the hosts file is backed up as `hosts.backup_<SessionId>`.
- IP blocks are created as separate inbound/outbound rules per range.

### Uninstall

Run the script and select **Unblock** mode.

