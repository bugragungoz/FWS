## SketchUp Blocker

Full network isolation for SketchUp using Windows Firewall rules (exe-only) plus domain blocking via hosts file.

### What it does

- **Firewall**: Creates block rules for discovered SketchUp `.exe` files (inbound + outbound).
- **Hosts**: Adds SketchUp/Trimble domain entries to the Windows hosts file.
- **Services**: Detects SketchUp-related Windows services and reports them.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\SketchUpBlocker"
.\SketchUpBlocker.ps1
```

### Operation modes

- **Block**: Apply full isolation rules.
- **Dry Run**: Analyze and report without making changes.
- **Unblock**: Remove firewall rules created by this tool and restore hosts file from backup (if available).
- **Rollback**: Restore from the most recent backup created by the tool.

### Scan locations

- `C:\Program Files\SketchUp\` (all versions)
- `C:\Program Files (x86)\SketchUp\`
- `C:\ProgramData\SketchUp\`
- `%LOCALAPPDATA%\SketchUp\`
- `%APPDATA%\SketchUp\`
- Custom path prompt if nothing is found

### Outputs

- **Logs**: `SketchUpBlocker_Logs/SketchUpBlocker_YYYYMMDD_HHMMSS.log`
- **Backups**: `SketchUpBlocker_Backups/FirewallRules_*.xml`
- **Reports**: `SketchUpBlocker_Report_*.txt`

### Safety notes

- This tool is intended for full offline use cases. Run **Dry Run** first if you are unsure.
- This tool changes **Firewall** and the **hosts file**.

### Tool-specific notes

- Domain blocking targets SketchUp/Trimble related endpoints (license, telemetry, updates, warehouses, assets).
- The goal is full isolation; online services (extensions/warehouses/cloud) are expected to stop working.

### Uninstall

Run the script and select **Unblock** mode.
