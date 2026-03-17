## PSIM Blocker

Blocks Altair PSIM internet access via Windows Firewall rules (exe-only).

### What it does

- **Firewall**: Creates block rules for discovered PSIM-related `.exe` files (inbound + outbound).
- **Rule lifecycle**: Removes previously managed rules (by prefix) and recreates them on each run.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

### Usage

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
cd "C:\path\to\FWS\PsimBlocker"
.\PSIMBlocker.ps1
```

### Operation modes

This tool runs as a single workflow (no interactive modes):

- Discover install directories and relevant `.exe` files
- Remove previously managed rules
- Create inbound + outbound block rules

### Scan locations

The tool discovers Altair/PSIM directories using known paths plus keyword-based discovery under common Windows roots.

### Outputs

Firewall rules are created under the group `Altair PSIM Privacy Block`.

### Safety notes

- This tool recreates managed rules on each run; review rule prefix before executing.

### Tool-specific notes

- Uses stable rule naming with a short SHA-256 based hash per executable path.

### Uninstall

Remove rules by prefix:

```powershell
Get-NetFirewallRule -DisplayName "Altair_PSIM_Privacy_Block*" | Remove-NetFirewallRule
```

