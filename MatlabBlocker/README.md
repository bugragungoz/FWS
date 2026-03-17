## MATLAB Blocker

Blocks MATLAB/MathWorks internet access via Windows Firewall rules (exe-only) and optionally modifies the hosts file and blocks IP ranges.

### What it does

- **Firewall**: Creates block rules for discovered MATLAB `.exe` files (inbound + outbound).
- **Hosts**: Adds MathWorks domain entries to the Windows hosts file.
- **IP ranges**: Creates IP range block rules for known MathWorks subnets.
- **Services**: Detects MATLAB/MathWorks-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\MatlabBlocker"
.\MatlabBlocker.ps1
```

### Operation modes

- **Block**: Apply firewall + hosts + IP blocks.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files\MATLAB\R*\`
- `C:\Program Files (x86)\MATLAB\R*\`
- `C:\ProgramData\MathWorks\`
- `%LOCALAPPDATA%\MathWorks\`
- `%APPDATA%\MathWorks\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `MatlabBlocker_Logs/MatlabBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `MatlabBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `MatlabBlocker_Report_*.txt`

### Safety notes

- Run **Dry Run** first to confirm what will be affected.
- This tool changes **Firewall**, the **hosts file**, and may add **IP range rules**.

### Tool-specific notes

- Multi-version support via `R*` scan folders (e.g., R2020a+).

### Uninstall

Run the script and select **Unblock** mode.

