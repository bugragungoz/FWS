## Application Internet Blocker

General-purpose tool to block internet access for a specific application using Windows Firewall rules.

### What it does

- **Firewall**: Creates block rules for selected targets (recommended: `.exe` only).
- **Rule management**: List/remove rules created by the tool.
- **Optional**: System Restore Point prompt (if enabled on the system drive).
- **Logging**: Writes actions to a log file.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\AppInternetBlocker"
.\AppIntBlockerEnhanced.ps1
```

### Operation modes

This tool provides an interactive menu to:

- Block internet access for an application
- Manage/remove rules created by the tool
- Open Windows Firewall Advanced Security UI

### Scan locations

User-provided application directory (the tool scans within that directory).

### Outputs

- **Logs**: `AppBlocker.log` (in the current working directory)

### Safety notes

- Prefer blocking **executables** only.
- Consider creating a System Restore Point before applying changes.

### Tool-specific notes

- This tool is general-purpose and does not target any specific vendor.

### Uninstall

Use the tool's rule management menu to remove rules created by the tool.