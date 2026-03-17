#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Plexim PLECS Internet Access Blocker - Strict Full Block

.DESCRIPTION
    Blocks all inbound and outbound network traffic for PLECS executables
    via Windows Firewall. No exceptions. Use with local license file only.

.NOTES
    Name:           PLECS Internet Access Blocker
    Author:         Bugra
    Version:        1.0.0
    Created:        2025

.LEGAL DISCLAIMER
    This tool is provided for LEGAL USE ONLY. The author accepts NO RESPONSIBILITY
    for any misuse, damage, or legal consequences arising from the use of this script.

    - Users are SOLELY RESPONSIBLE for ensuring compliance with software licenses
    - Users are SOLELY RESPONSIBLE for compliance with local laws and regulations
    - This script is intended for network security and testing purposes only
    - Always backup your system and create restore points before execution
    - The author disclaims all warranties, express or implied

    BY USING THIS SCRIPT, YOU ACKNOWLEDGE AND ACCEPT FULL RESPONSIBILITY FOR YOUR ACTIONS.
#>

$ErrorActionPreference = "Stop"

$script:Config = @{
    LogDirectory    = "$PSScriptRoot\PLECSBlocker_Logs"
    BackupDirectory = "$PSScriptRoot\PLECSBlocker_Backups"
    LogFile         = ""
    BackupFile      = ""
    ReportFile      = ""
    SessionID       = (Get-Date -Format "yyyyMMdd_HHmmss")
    DryRun          = $false
    RulePrefix      = "App Internet Blocker - PLECS"
    RuleGroup       = "PLECS Internet Blocker"
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
    Write-Host ""
    Write-Host "  PLECS Internet Blocker v1.0.0 | Session: $($script:Config.SessionID)" -ForegroundColor Cyan
    Write-Host "  Rule Group: $($script:Config.RuleGroup)" -ForegroundColor Gray
    Write-Host "  [*] Strict full block - Inbound + Outbound (exe only)" -ForegroundColor Yellow
    Write-Host ""
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS', 'DEBUG')]
        [string]$Level = 'INFO'
    )

    if ($script:Config.LogFile) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$Level] $Message"
        Add-Content -Path $script:Config.LogFile -Value $logEntry -ErrorAction SilentlyContinue
    }
}

function Initialize-Environment {
    try {
        if (-not (Test-Path $script:Config.LogDirectory)) {
            New-Item -ItemType Directory -Path $script:Config.LogDirectory -Force | Out-Null
        }

        if (-not (Test-Path $script:Config.BackupDirectory)) {
            New-Item -ItemType Directory -Path $script:Config.BackupDirectory -Force | Out-Null
        }

        $script:Config.LogFile = Join-Path $script:Config.LogDirectory "PLECSBlocker_$($script:Config.SessionID).log"
        $script:Config.BackupFile = Join-Path $script:Config.BackupDirectory "FirewallRules_$($script:Config.SessionID).json"
        $script:Config.ReportFile = Join-Path $script:Config.LogDirectory "PLECSBlocker_Report_$($script:Config.SessionID).txt"

        Write-Log "Environment initialized successfully" -Level SUCCESS
        Write-Host "  [OK] Log directory: $($script:Config.LogDirectory)" -ForegroundColor Green
        Write-Host "  [OK] Backup directory: $($script:Config.BackupDirectory)" -ForegroundColor Green

        return $true
    }
    catch {
        Write-Host "  [ERROR] Failed to initialize environment: $_" -ForegroundColor Red
        return $false
    }
}

function Test-Prerequisites {
    Write-Host "[Step 1] Checking prerequisites..." -ForegroundColor Cyan

    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "  [ERROR] This script requires Administrator privileges!" -ForegroundColor Red
        Write-Log "Script execution failed: Not running as Administrator" -Level ERROR
        return $false
    }

    Write-Host "  [OK] Running as Administrator" -ForegroundColor Green
    Write-Log "Prerequisites check passed" -Level SUCCESS
    return $true
}

function Prompt-SystemRestorePoint {
    $response = Read-Host "  Create system restore point first? (y/n, default n)"
    if ($response -match '^(yes|y)$') {
        Write-Host "  Create restore point via: Win + R -> sysdm.cpl -> System Protection -> Create" -ForegroundColor Gray
        $cont = Read-Host "  Continue anyway? (y/n)"
        if ($cont -notmatch '^(yes|y)$') {
            Write-Host "  [CANCELLED]" -ForegroundColor Yellow
            return $false
        }
    }
    return $true
}

function Backup-FirewallRules {
    Write-Host ""
    Write-Host "[Step 2] Backing up existing firewall rules..." -ForegroundColor Cyan

    try {
        $existingRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }

        if ($existingRules) {
            $backupEntries = @()
            foreach ($rule in $existingRules) {
                $appFilter = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                $backupEntries += @{
                    DisplayName = $rule.DisplayName
                    Direction   = $rule.Direction
                    Action      = $rule.Action
                    Enabled     = $rule.Enabled
                    Program     = if ($appFilter) { $appFilter.Program } else { $null }
                }
            }
            $backupEntries | ConvertTo-Json -Depth 5 | Out-File -FilePath $script:Config.BackupFile -Encoding UTF8

            Write-Host "  [OK] Backed up $($existingRules.Count) existing rules to:" -ForegroundColor Green
            Write-Host "       $($script:Config.BackupFile)" -ForegroundColor Gray
            Write-Log "Backed up $($existingRules.Count) existing firewall rules" -Level SUCCESS
        }
        else {
            Write-Host "  [INFO] No existing PLECS blocker rules found" -ForegroundColor Cyan
            Write-Log "No existing rules to backup" -Level INFO
        }

        return $true
    }
    catch {
        Write-Host "  [WARNING] Could not backup firewall rules: $_" -ForegroundColor Yellow
        Write-Log "Firewall backup failed: $_" -Level WARNING
        return $true
    }
}

function Get-PLECSScanDirectories {
    $directories = @()

    $primaryPaths = @(
        "C:\Program Files\Plexim",
        "C:\Program Files (x86)\Plexim",
        "$env:LOCALAPPDATA\Plexim",
        "C:\ProgramData\Plexim"
    )

    foreach ($path in $primaryPaths) {
        if (Test-Path $path) {
            $directories += $path
        }
    }

    $searchKeywords = @("Plexim", "PLECS")
    $basePaths = @(
        "C:\Program Files",
        "C:\Program Files (x86)",
        "C:\ProgramData",
        $env:LOCALAPPDATA,
        $env:APPDATA
    )

    foreach ($basePath in $basePaths) {
        if (-not (Test-Path $basePath)) { continue }

        try {
            $subdirs = Get-ChildItem -Path $basePath -Directory -Recurse -ErrorAction SilentlyContinue
            foreach ($dir in $subdirs) {
                foreach ($keyword in $searchKeywords) {
                    if ($dir.Name -like "*$keyword*") {
                        $directories += $dir.FullName
                        break
                    }
                }
            }
        }
        catch {
            Write-Log "Failed to scan $basePath : $_" -Level WARNING
        }
    }

    return $directories | Select-Object -Unique
}

function Get-PLECSFiles {
    param([string[]]$Directories)

    $allFiles = @()
    $extensions = @("*.exe")

    foreach ($dir in $Directories) {
        if (-not (Test-Path -LiteralPath $dir)) { continue }

        try {
            foreach ($ext in $extensions) {
                $files = Get-ChildItem -LiteralPath $dir -Filter $ext -Recurse -File -ErrorAction SilentlyContinue
                $allFiles += $files
            }
        }
        catch {
            Write-Log "Failed to scan $dir : $_" -Level WARNING
        }
    }

    return $allFiles | Select-Object -Unique -Property FullName | ForEach-Object { Get-Item -LiteralPath $_.FullName }
}

function New-PLECSFirewallRule {
    param(
        [string]$DisplayName,
        [string]$FilePath,
        [ValidateSet('Inbound', 'Outbound')]
        [string]$Direction
    )

    if ($script:Config.DryRun) {
        $script:Statistics.FirewallRulesCreated++
        return $true
    }

    try {
        $ruleDisplayName = "$($script:Config.RulePrefix) - $DisplayName ($Direction)"
        if (Get-NetFirewallRule -DisplayName $ruleDisplayName -ErrorAction SilentlyContinue) {
            return $true
        }

        New-NetFirewallRule `
            -DisplayName $ruleDisplayName `
            -Group $script:Config.RuleGroup `
            -Direction $Direction `
            -Program $FilePath `
            -Action Block `
            -Enabled True `
            -ErrorAction Stop | Out-Null

        $script:Statistics.FirewallRulesCreated++
        return $true
    }
    catch {
        Write-Log "Failed to create rule for $FilePath ($Direction): $_" -Level ERROR
        return $false
    }
}

function Remove-PLECSFirewallRules {
    $rules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }
    if ($rules) {
        $rules | Remove-NetFirewallRule
        return $rules.Count
    }
    return 0
}

function Show-MainMenu {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "                              OPERATION MODE                                    " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] BLOCK MODE      - Full block Inbound + Outbound (exe only)" -ForegroundColor Green
    Write-Host "  [2] DRY RUN MODE    - Analyze and report without making changes" -ForegroundColor Yellow
    Write-Host "  [3] UNBLOCK MODE    - Remove all blocking rules" -ForegroundColor Red
    Write-Host "  [4] ROLLBACK MODE   - Restore from backup" -ForegroundColor Magenta
    Write-Host "  [5] EXIT            - Exit script" -ForegroundColor Gray
    Write-Host "  [6] DISCLAIMER & HELP - View legal info and documentation" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""

    $choice = Read-Host "  Select operation mode [1-6]"
    return $choice
}

function Show-BlockModeDisclaimer {
    Write-Host ""
    Write-Host "  BLOCK MODE: Strict full block - Inbound and Outbound for all exe." -ForegroundColor Yellow
    Write-Host "  No exceptions. Use with local license file only." -ForegroundColor Gray
    Write-Host ""
}

function Show-DryRunDisclaimer {
    Write-Host ""
    Write-Host "  DRY RUN: Scan and report only, no changes." -ForegroundColor Yellow
    Write-Host ""
}

function Show-UnblockDisclaimer {
    Write-Host ""
    Write-Host "  UNBLOCK: Remove all PLECS blocker rules, restore full internet access." -ForegroundColor Yellow
    Write-Host ""
}

function Show-RollbackDisclaimer {
    Write-Host ""
    Write-Host "  ROLLBACK: Restore from most recent backup." -ForegroundColor Yellow
    Write-Host ""
}

function Show-DisclaimerAndHelp {
    Clear-Host
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "             PLECS BLOCKER - DISCLAIMER & HELP DOCUMENTATION                     " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "LEGAL DISCLAIMER:" -ForegroundColor Red
    Write-Host "  This script is for LEGAL, EDUCATIONAL, and TESTING purposes only." -ForegroundColor White
    Write-Host "  Author: Bugra | Version: 1.0.0" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "WHAT THIS SCRIPT DOES:" -ForegroundColor Green
    Write-Host "  - Inbound + Outbound firewall rules for Plexim PLECS exe" -ForegroundColor White
    Write-Host "  - Full block - no exceptions" -ForegroundColor White
    Write-Host "  - Intended for use with local license file (no network license)" -ForegroundColor White
    Write-Host ""
    Write-Host "MANUAL CLEANUP - Remove all PLECS blocker rules:" -ForegroundColor Yellow
    Write-Host "  Get-NetFirewallRule -DisplayName 'App Internet Blocker - PLECS*' | Remove-NetFirewallRule" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Press any key to return to main menu..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Get-UserConsent {
    param([string]$Mode)
    return $true
}

function Remove-DuplicateRules {
    Write-Host ""
    Write-Host "[Step 3] Checking for duplicate rules..." -ForegroundColor Cyan

    try {
        $existingRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }

        if ($existingRules) {
            Write-Host "  [INFO] Found $($existingRules.Count) existing rules" -ForegroundColor Yellow
            Write-Host "  [1] Keep existing and add new  [2] Remove all and start fresh  [3] Cancel" -ForegroundColor Cyan
            $choice = Read-Host "  Select option [1-3]"

            switch ($choice) {
                "1" { Write-Host "  [OK] Keeping existing rules" -ForegroundColor Green; return $true }
                "2" {
                    $existingRules | Remove-NetFirewallRule
                    Write-Host "  [OK] Existing rules removed" -ForegroundColor Green
                    return $true
                }
                "3" { Write-Host "  [CANCELLED] Operation cancelled" -ForegroundColor Yellow; return $false }
                default { Write-Host "  [ERROR] Invalid choice" -ForegroundColor Red; return $false }
            }
        }
        else {
            Write-Host "  [OK] No duplicate rules found" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to check duplicates: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-UnblockMode {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Red
    Write-Host "                              UNBLOCK MODE                                      " -ForegroundColor Red
    Write-Host "================================================================================" -ForegroundColor Red

    Show-UnblockDisclaimer
    if (-not (Get-UserConsent -Mode "UNBLOCK")) { return }

    Write-Host ""
    Write-Host "[Step 1] Removing firewall rules..." -ForegroundColor Cyan

    try {
        $removed = Remove-PLECSFirewallRules
        if ($removed -gt 0) {
            Write-Host "  [OK] Removed $removed firewall rules" -ForegroundColor Green
            Write-Log "Removed $removed firewall rules" -Level SUCCESS
        }
        else {
            Write-Host "  [INFO] No rules found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to remove rules: $_" -ForegroundColor Red
        Write-Log "Rule removal failed: $_" -Level ERROR
    }

    Write-Host ""
    Write-Host "  PLECS internet access has been restored." -ForegroundColor Green
    Write-Host ""
}

function Invoke-RollbackMode {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Magenta
    Write-Host "                             ROLLBACK MODE                                      " -ForegroundColor Magenta
    Write-Host "================================================================================" -ForegroundColor Magenta

    Show-RollbackDisclaimer
    if (-not (Get-UserConsent -Mode "ROLLBACK")) { return }

    $backupFiles = Get-ChildItem -Path $script:Config.BackupDirectory -Filter "FirewallRules_*.json" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending

    if (-not $backupFiles) {
        Write-Host "  [ERROR] No backup files found. Use UNBLOCK MODE instead." -ForegroundColor Red
        return
    }

    $latestBackup = $backupFiles[0]
    $removed = Remove-PLECSFirewallRules
    if ($removed -gt 0) {
        Write-Host "  [INFO] Removed $removed current rules" -ForegroundColor Cyan
    }

    try {
        $backupData = Get-Content -Path $latestBackup.FullName -Raw | ConvertFrom-Json
        $rulesToRestore = if ($backupData -is [System.Array]) { $backupData } else { @($backupData) }
        $restoredCount = 0

        foreach ($entry in $rulesToRestore) {
            try {
                $params = @{
                    DisplayName = $entry.DisplayName
                    Group       = $script:Config.RuleGroup
                    Direction   = $entry.Direction
                    Action      = $entry.Action
                    Enabled     = $entry.Enabled
                    ErrorAction = 'Stop'
                }
                if ($entry.Program) {
                    $params['Program'] = $entry.Program
                }
                New-NetFirewallRule @params | Out-Null
                $restoredCount++
            }
            catch {
                Write-Log "Failed to restore rule $($entry.DisplayName): $_" -Level WARNING
            }
        }

        Write-Host "  [OK] Restored $restoredCount firewall rules from backup" -ForegroundColor Green
        Write-Log "Rollback: Restored $restoredCount rules" -Level SUCCESS
    }
    catch {
        Write-Host "  [ERROR] Backup restore failed: $_" -ForegroundColor Red
        Write-Log "Rollback failed: $_" -Level ERROR
    }

    Write-Host ""
}

function Process-PLECSDirectories {
    Write-Host ""
    Write-Host "[Step 4] Scanning for Plexim/PLECS directories..." -ForegroundColor Cyan

    $directories = Get-PLECSScanDirectories

    if ($directories.Count -eq 0) {
        Write-Host "  [WARNING] No Plexim/PLECS directories found in standard locations" -ForegroundColor Yellow
        Write-Host "  [INFO] Specify custom directory? (yes/no): " -ForegroundColor Cyan -NoNewline
        $response = Read-Host
        if ($response -match '^(yes|y)$') {
            Write-Host "  Enter path: " -ForegroundColor Cyan -NoNewline
            $customPath = Read-Host
            if (Test-Path -LiteralPath $customPath) {
                $directories = @($customPath)
            }
            else {
                Write-Host "  [ERROR] Invalid path" -ForegroundColor Red
                return 0
            }
        }
        else {
            return 0
        }
    }

    Write-Host "  [INFO] Found $($directories.Count) directory/directories to scan" -ForegroundColor Cyan
    foreach ($dir in $directories) {
        Write-Host "    - $dir" -ForegroundColor Gray
    }

    $files = Get-PLECSFiles -Directories $directories
    if ($files.Count -eq 0) {
        Write-Host "  [WARNING] No .exe files found in scanned directories" -ForegroundColor Yellow
        return 0
    }

    Write-Host ""
    Write-Host "  [INFO] Found $($files.Count) file(s) to process (exe only)" -ForegroundColor Cyan

    $processedCount = 0
    $fileArray = @($files)
    for ($i = 0; $i -lt $fileArray.Count; $i++) {
        $file = $fileArray[$i]
        $percentage = [math]::Round((($i + 1) / $fileArray.Count) * 100)
        Write-Progress -Activity "Processing Files" -Status "$percentage% Complete" -PercentComplete $percentage -CurrentOperation $file.Name
        $script:Statistics.TotalFilesScanned++

        if (-not $script:BlockedFilesHashSet.ContainsKey($file.FullName)) {
            $script:BlockedFilesHashSet[$file.FullName] = $true
            $displayName = "$($file.BaseName) - $($file.Extension)"

            $outboundOk = New-PLECSFirewallRule -DisplayName $displayName -FilePath $file.FullName -Direction Outbound
            $inboundOk = New-PLECSFirewallRule -DisplayName $displayName -FilePath $file.FullName -Direction Inbound

            if ($outboundOk -and $inboundOk) {
                $processedCount++
                $script:Statistics.BlockedFilesList += $file.FullName

                if ($script:Config.DryRun) {
                    Write-Host "    [DRY RUN] Would block: $($file.Name)" -ForegroundColor DarkGray
                }
                else {
                    Write-Host "    [OK] Blocked: $($file.Name)" -ForegroundColor Green
                }
            }
        }
    }

    Write-Progress -Activity "Processing Files" -Completed
    Write-Host ""
    Write-Host "  [OK] Processed $processedCount file(s)" -ForegroundColor Green
    Write-Log "Processed $processedCount files from $($directories.Count) directories" -Level SUCCESS
    return $processedCount
}

function Generate-Report {
    Write-Host ""
    Write-Host "[Step 5] Generating report..." -ForegroundColor Cyan

    $script:Statistics.ExecutionEndTime = Get-Date
    $duration = $script:Statistics.ExecutionEndTime - $script:Statistics.ExecutionStartTime

    $report = @"
================================================================================
                    PLECS BLOCKER EXECUTION REPORT
================================================================================

Session ID:       $($script:Config.SessionID)
Start Time:       $($script:Statistics.ExecutionStartTime)
End Time:         $($script:Statistics.ExecutionEndTime)
Duration:         $($duration.ToString("hh\:mm\:ss"))
Mode:             $(if ($script:Config.DryRun) { "DRY RUN" } else { "LIVE" })

Files Scanned:    $($script:Statistics.TotalFilesScanned)
Firewall Rules:   $($script:Statistics.FirewallRulesCreated)

Rule Group:       $($script:Config.RuleGroup)
Block Mode:       Full (Inbound + Outbound, exe only)

================================================================================
"@

    try {
        $report | Out-File -FilePath $script:Config.ReportFile -Encoding UTF8
        Write-Host "  [OK] Report saved to: $($script:Config.ReportFile)" -ForegroundColor Green
        Write-Host $report -ForegroundColor White
    }
    catch {
        Write-Host "  [ERROR] Report failed: $_" -ForegroundColor Red
    }
}

try {
    Show-Banner

    if (-not (Test-Prerequisites)) { exit 1 }
    if (-not (Initialize-Environment)) { exit 1 }

    Write-Log "Script execution started" -Level INFO

    $mode = Show-MainMenu

    switch ($mode) {
        "1" {
            $script:Config.DryRun = $false
            Show-BlockModeDisclaimer
            if (-not (Get-UserConsent -Mode "BLOCK")) { break }
            if (-not (Prompt-SystemRestorePoint)) { break }
            if (-not (Backup-FirewallRules)) { break }
            if (-not (Remove-DuplicateRules)) { break }

            Process-PLECSDirectories | Out-Null
            Generate-Report

            Write-Host ""
            Write-Host "  PLECS strict internet block applied." -ForegroundColor Green
            Write-Host ""
        }
        "2" {
            $script:Config.DryRun = $true
            Show-DryRunDisclaimer
            if (-not (Get-UserConsent -Mode "DRY RUN")) { break }

            Process-PLECSDirectories | Out-Null
            Generate-Report

            Write-Host ""
            Write-Host "  DRY RUN complete. No changes were made." -ForegroundColor Green
            Write-Host ""
        }
        "3" { Invoke-UnblockMode }
        "4" { Invoke-RollbackMode }
        "5" {
            Write-Log "Script exited by user" -Level INFO
            exit 0
        }
        "6" {
            Show-DisclaimerAndHelp
            & $PSCommandPath
        }
        default {
            Write-Host "  [ERROR] Invalid selection" -ForegroundColor Red
            exit 1
        }
    }

    Write-Log "Script execution completed" -Level SUCCESS
}
catch {
    Write-Host ""
    Write-Host "  CRITICAL ERROR: $_" -ForegroundColor Red
    Write-Log "Critical error: $_" -Level ERROR
    exit 1
}
finally {
    Write-Host ""
    Write-Host "  Finished at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
    if ($script:Config.LogFile) { Write-Host "  Log: $($script:Config.LogFile)" -ForegroundColor Gray }
    Write-Host ""
}
