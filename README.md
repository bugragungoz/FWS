## FWS (Firewall Scripts)

Windows Firewall automation toolset (exe-only rule targets). One tool per folder.

### Layout

- `<ToolName>/` -> `<ToolScript>.ps1` + `README.md`
- `_templates/` -> `README_TEMPLATE.md`, `SCRIPT_TEMPLATE.ps1`

### Tools

- `AdobeBlocker`, `MatlabBlocker`, `ProteusBlocker`, `MultiSimBlocker`: Firewall + hosts + IP ranges
- `AltiumBlocker`: Selective blocking (preserves Component Search/login)
- `SketchUpBlocker`: Full isolation (Firewall + hosts)
- `AnsysBlocker`, `CadenceBlocker`: Localhost license exception profiles
- `PlecsBlocker`: Strict firewall-only full block
- `PsimBlocker`: Firewall-only workflow with stable rule names
- `AppInternetBlocker`: General-purpose interactive app blocker

### Requirements

- Windows 10/11, PowerShell 5.1+, Administrator privileges

