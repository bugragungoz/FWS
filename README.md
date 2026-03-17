## FWS (Firewall Scripts)

Consolidated Windows Firewall automation scripts (exe-only rule targets) organized as one tool per folder.

### Structure

- `FWS/<ToolName>/`
  - `<ToolScript>.ps1`
  - `README.md`

### Templates

- `FWS/_templates/README_TEMPLATE.md`
- `FWS/_templates/SCRIPT_TEMPLATE.ps1`

### Tools

- `AdobeBlocker`: Firewall + hosts + IP range blocking for Adobe.
- `AltiumBlocker`: Selective blocking while preserving Component Search and login.
- `MatlabBlocker`: Firewall + hosts + IP range blocking for MATLAB/MathWorks.
- `ProteusBlocker`: Firewall + hosts + IP range blocking for Proteus/Labcenter.
- `SketchUpBlocker`: Full isolation profile (Firewall + hosts).
- `MultiSimBlocker`: Firewall + hosts + IP range blocking for NI/MultiSim.
- `AnsysBlocker`: Firewall blocking with localhost license exception.
- `CadenceBlocker`: WAN blocking with localhost license exception (outbound rules).
- `PlecsBlocker`: Strict firewall-only full block.
- `PsimBlocker`: Firewall-only workflow with stable hashed rule names.
- `AppInternetBlocker`: General-purpose interactive app blocker.

### Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges

