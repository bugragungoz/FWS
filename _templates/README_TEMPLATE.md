## <ToolName>

<One-line summary. Keep it precise and technical.>

### What it does

- **Firewall**: Creates Windows Firewall block rules for discovered `.exe` files (inbound/outbound depending on tool).
- **Optional**: Hosts file changes (domain blocking).
- **Optional**: IP range blocking.
- **Optional**: Service detection/reporting.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\<ToolFolder>"
.\<ToolScript>.ps1
```

### Operation modes

- **Block**: Apply blocking actions.
- **Dry Run**: Analyze and report without making system changes.
- **Unblock**: Remove rules created by the tool and restore related changes (if applicable).
- **Rollback**: Restore from backups created by the tool (if applicable).

### Scan locations

List default scan roots and any version-aware scanning logic (wildcards, multi-version folders).

### Outputs

- **Logs**: `<Tool>_Logs/`
- **Backups**: `<Tool>_Backups/`
- **Reports**: `<Tool>_Report_*.txt` (or under logs folder, depending on tool)

### Safety notes

- Run in **Dry Run** first to validate scan results.
- Review what the tool modifies (Firewall / Hosts / IP rules) before applying.

### Tool-specific notes

Document exceptions, preserved online features, localhost license exceptions, allowlists, etc.

### Uninstall

Run the tool and select **Unblock** mode.

