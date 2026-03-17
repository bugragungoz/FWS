#Requires -RunAsAdministrator

<#
.SYNOPSIS
    <ToolName> - Firewall automation tool (exe-only)

.DESCRIPTION
    <Short description of what is blocked and what is preserved.>

.NOTES
    Name:           <ToolName>
    Author:         Bugra
    Version:        1.0.0
    Created:        2026

.LEGAL DISCLAIMER
    This tool is provided for LEGAL USE ONLY. Users are solely responsible for compliance
    with software licenses and local laws. The author accepts no liability for misuse.
#>

$ErrorActionPreference = 'Stop'

$script:Config = @{
    LogDirectory    = "$PSScriptRoot\<Tool>_Logs"
    BackupDirectory = "$PSScriptRoot\<Tool>_Backups"
    LogFile         = ''
    BackupFile      = ''
    ReportFile      = ''
    SessionID       = (Get-Date -Format 'yyyyMMdd_HHmmss')
    DryRun          = $false
    RulePrefix      = '<RulePrefix>'
    RuleGroup       = '<RuleGroup>'
}

$script:Statistics = @{
    TotalFilesScanned    = 0
    FirewallRulesCreated = 0
    ExecutionStartTime   = Get-Date
    ExecutionEndTime     = $null
    BlockedFilesList     = @()
}

$script:BlockedFilesHashSet = @{}

function Show-Banner {
    Write-Host ''
    Write-Host "  <ToolName> | Session: $($script:Config.SessionID)" -ForegroundColor Cyan
    Write-Host "  Rule Group: $($script:Config.RuleGroup)" -ForegroundColor Gray
    Write-Host ''
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    if (-not $script:Config.LogFile) { return }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $script:Config.LogFile -Value "[$timestamp] [$Level] $Message" -ErrorAction SilentlyContinue
}

function Initialize-Environment {
    if (-not (Test-Path $script:Config.LogDirectory)) {
        New-Item -ItemType Directory -Path $script:Config.LogDirectory -Force | Out-Null
    }
    if (-not (Test-Path $script:Config.BackupDirectory)) {
        New-Item -ItemType Directory -Path $script:Config.BackupDirectory -Force | Out-Null
    }

    $script:Config.LogFile = Join-Path $script:Config.LogDirectory "<Tool>_$($script:Config.SessionID).log"
    $script:Config.BackupFile = Join-Path $script:Config.BackupDirectory "FirewallRules_$($script:Config.SessionID).json"
    $script:Config.ReportFile = Join-Path $script:Config.LogDirectory "<Tool>_Report_$($script:Config.SessionID).txt"
}

function Test-Prerequisites {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
    if (-not $isAdmin) { throw 'Administrator privileges are required.' }
}

function Show-MainMenu {
    Write-Host ''
    Write-Host '  [1] BLOCK MODE' -ForegroundColor Green
    Write-Host '  [2] DRY RUN MODE' -ForegroundColor Yellow
    Write-Host '  [3] UNBLOCK MODE' -ForegroundColor Red
    Write-Host '  [4] ROLLBACK MODE' -ForegroundColor Magenta
    Write-Host '  [5] EXIT' -ForegroundColor Gray
    Write-Host ''
    return (Read-Host '  Select operation mode [1-5]')
}

function Get-UserConsent {
    param([Parameter(Mandatory = $true)][string]$Mode)
    return $true
}

function Get-ScanDirectories {
    # Return unique directories to scan.
    return @()
}

function Get-TargetExecutables {
    param([Parameter(Mandatory = $true)][string[]]$Directories)

    $all = @()
    foreach ($dir in $Directories) {
        if (-not (Test-Path -LiteralPath $dir -PathType Container)) { continue }
        try {
            $all += Get-ChildItem -LiteralPath $dir -Filter '*.exe' -Recurse -File -ErrorAction SilentlyContinue
        } catch {}
    }
    return $all | Select-Object -Unique -Property FullName | ForEach-Object { Get-Item -LiteralPath $_.FullName }
}

function New-BlockRule {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$ProgramPath,
        [Parameter(Mandatory = $true)][ValidateSet('Inbound', 'Outbound')][string]$Direction
    )

    if ($script:Config.DryRun) { $script:Statistics.FirewallRulesCreated++; return }

    $ruleDisplayName = "$($script:Config.RulePrefix) - $DisplayName ($Direction)"
    if (Get-NetFirewallRule -DisplayName $ruleDisplayName -ErrorAction SilentlyContinue) { return }

    New-NetFirewallRule `
        -DisplayName $ruleDisplayName `
        -Group $script:Config.RuleGroup `
        -Direction $Direction `
        -Program $ProgramPath `
        -Action Block `
        -Enabled True | Out-Null

    $script:Statistics.FirewallRulesCreated++
}

function Invoke-BlockOrDryRun {
    param([Parameter(Mandatory = $true)][bool]$DryRun)

    $script:Config.DryRun = $DryRun
    $directories = @(Get-ScanDirectories | Select-Object -Unique)
    $files = @(Get-TargetExecutables -Directories $directories)

    if ($files.Count -eq 0) {
        Write-Host '  [WARNING] No .exe files found to process.' -ForegroundColor Yellow
        return
    }

    $fileArray = @($files)
    for ($i = 0; $i -lt $fileArray.Count; $i++) {
        $file = $fileArray[$i]
        $percentage = [math]::Round((($i + 1) / $fileArray.Count) * 100)
        Write-Progress -Activity 'Processing Files' -Status "$percentage% Complete" -PercentComplete $percentage -CurrentOperation $file.Name

        $script:Statistics.TotalFilesScanned++
        if ($script:BlockedFilesHashSet.ContainsKey($file.FullName)) { continue }
        $script:BlockedFilesHashSet[$file.FullName] = $true

        $displayName = "$($file.BaseName) - $($file.Extension)"
        New-BlockRule -DisplayName $displayName -ProgramPath $file.FullName -Direction Outbound
        New-BlockRule -DisplayName $displayName -ProgramPath $file.FullName -Direction Inbound
        $script:Statistics.BlockedFilesList += $file.FullName
    }

    Write-Progress -Activity 'Processing Files' -Completed
}

function Invoke-UnblockMode {
    $rules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }
    if (-not $rules) { return }
    $ruleCount = $rules.Count

    $counter = 0
    foreach ($rule in $rules) {
        $counter++
        $percentage = [math]::Round(($counter / $ruleCount) * 100)
        Write-Progress -Activity 'Removing Firewall Rules' -Status "$percentage% Complete" -PercentComplete $percentage
        Remove-NetFirewallRule -Name $rule.Name
    }

    Write-Progress -Activity 'Removing Firewall Rules' -Completed
}

function Generate-Report {
    $script:Statistics.ExecutionEndTime = Get-Date
    $duration = $script:Statistics.ExecutionEndTime - $script:Statistics.ExecutionStartTime

    $report = @"
================================================================================
<TOOL> EXECUTION REPORT
================================================================================

Session ID:       $($script:Config.SessionID)
Duration:         $($duration.ToString('hh\:mm\:ss'))
Mode:             $(if ($script:Config.DryRun) { 'DRY RUN' } else { 'LIVE' })

Files Scanned:    $($script:Statistics.TotalFilesScanned)
Firewall Rules:   $($script:Statistics.FirewallRulesCreated)

================================================================================
"@

    $report | Out-File -FilePath $script:Config.ReportFile -Encoding UTF8
    Write-Host "  [OK] Report saved: $($script:Config.ReportFile)" -ForegroundColor Green
}

try {
    Show-Banner
    Test-Prerequisites
    Initialize-Environment

    $mode = Show-MainMenu
    switch ($mode) {
        '1' { if (Get-UserConsent -Mode 'BLOCK') { Invoke-BlockOrDryRun -DryRun:$false; Generate-Report } }
        '2' { if (Get-UserConsent -Mode 'DRY RUN') { Invoke-BlockOrDryRun -DryRun:$true; Generate-Report } }
        '3' { if (Get-UserConsent -Mode 'UNBLOCK') { Invoke-UnblockMode } }
        '4' { Write-Host 'Rollback mode template placeholder.' -ForegroundColor Yellow }
        '5' { exit 0 }
        default { throw 'Invalid selection.' }
    }
}
catch {
    Write-Host "CRITICAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

