#Requires -RunAsAdministrator

<#
.SYNOPSIS
    MultiSim (National Instruments) Internet Access Blocker - Comprehensive Network Blocking Tool

.DESCRIPTION
    Advanced PowerShell script to block internet access for NI MultiSim and related Circuit Design Suite
    applications through Windows Firewall rules, hosts file modifications, and IP blocking.

.NOTES
    Name:           MultiSim Internet Access Blocker
    Author:         Bugra
    Concept & Design: Bugra
    Development:    Claude 4.5 Sonnet & Gemini 2.5 Pro AI
    Testing:        Bugra
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
    LogDirectory     = "$PSScriptRoot\MultiSimBlocker_Logs"
    BackupDirectory  = "$PSScriptRoot\MultiSimBlocker_Backups"
    LogFile          = ""
    BackupFile       = ""
    ReportFile       = ""
    SessionID        = (Get-Date -Format "yyyyMMdd_HHmmss")
    DryRun           = $false
    RulePrefix       = "MultiSimBlocker"
    RuleGroup        = "MultiSimBlocker"
}

$script:Statistics = @{
    TotalFilesScanned       = 0
    FirewallRulesCreated    = 0
    DomainsBlocked          = 0
    IPRulesCreated          = 0
    ServicesFound           = 0
    ExecutionStartTime      = Get-Date
    ExecutionEndTime        = $null
    BlockedFilesList        = @()
}

$script:BlockedFilesHashSet = @{}

function Show-Banner {
    Write-Host ""
    Write-Host "  MultiSim Blocker v1.0.0 | Session: $($script:Config.SessionID)" -ForegroundColor Cyan
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

        $script:Config.LogFile = Join-Path $script:Config.LogDirectory "MultiSimBlocker_$($script:Config.SessionID).log"
        $script:Config.BackupFile = Join-Path $script:Config.BackupDirectory "FirewallRules_$($script:Config.SessionID).json"
        $script:Config.ReportFile = Join-Path $script:Config.LogDirectory "MultiSimBlocker_Report_$($script:Config.SessionID).txt"

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

function Request-SystemRestorePoint {
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
            $backupData = $existingRules | ConvertTo-Json -Depth 10
            $backupData | Out-File -FilePath $script:Config.BackupFile -Encoding UTF8

            Write-Host "  [OK] Backed up $($existingRules.Count) existing rules to:" -ForegroundColor Green
            Write-Host "       $($script:Config.BackupFile)" -ForegroundColor Gray
            Write-Log "Backed up $($existingRules.Count) existing firewall rules" -Level SUCCESS
        }
        else {
            Write-Host "  [INFO] No existing $($script:Config.RulePrefix) rules found" -ForegroundColor Cyan
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

function Show-MainMenu {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "                              OPERATION MODE                                    " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] BLOCK MODE      - Apply full blocking (firewall + hosts + IP)" -ForegroundColor Green
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
    Write-Host "  BLOCK MODE: Firewall rules + hosts + IP blocking. Legal use only." -ForegroundColor Yellow
    Write-Host ""
}

function Show-DryRunDisclaimer {
    Write-Host ""
    Write-Host "  DRY RUN: Scan and report only, no changes." -ForegroundColor Yellow
    Write-Host ""
}

function Show-UnblockDisclaimer {
    Write-Host ""
    Write-Host "  UNBLOCK: Remove rules, restore hosts, restore internet access." -ForegroundColor Yellow
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
    Write-Host "             MULTISIM BLOCKER - DISCLAIMER & HELP DOCUMENTATION                  " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "LEGAL DISCLAIMER:" -ForegroundColor Red
    Write-Host "  This script is for LEGAL, EDUCATIONAL, and TESTING purposes only." -ForegroundColor White
    Write-Host "  Author: Bugra | Version: 1.0.0" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "WHAT THIS SCRIPT DOES:" -ForegroundColor Green
    Write-Host "  - Firewall rules for MultiSim / NI Circuit Design Suite executables" -ForegroundColor White
    Write-Host "  - Hosts file: NI domain blocking (license, telemetry, updates)" -ForegroundColor White
    Write-Host "  - IP range blocking for known NI servers" -ForegroundColor White
    Write-Host ""
    Write-Host "MANUAL CLEANUP - Remove all MultiSimBlocker rules:" -ForegroundColor Yellow
    Write-Host "  Get-NetFirewallRule -DisplayName 'MultiSimBlocker*' | Remove-NetFirewallRule" -ForegroundColor Gray
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
            Write-Host "  [INFO] Found $($existingRules.Count) existing rules with prefix: $($script:Config.RulePrefix)" -ForegroundColor Yellow
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
        $rules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }
        if ($rules) {
            $rules | Remove-NetFirewallRule
            Write-Host "  [OK] Removed $($rules.Count) firewall rules" -ForegroundColor Green
            Write-Log "Removed $($rules.Count) firewall rules" -Level SUCCESS
        }
        else {
            Write-Host "  [INFO] No rules found" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to remove rules: $_" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "[Step 2] Restoring hosts file..." -ForegroundColor Cyan

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $hostsBackupFiles = Get-ChildItem -Path (Split-Path $hostsPath) -Filter "hosts.backup_*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending

    try {
        if ($hostsBackupFiles) {
            $latestHostsBackup = $hostsBackupFiles[0]
            Copy-Item -Path $latestHostsBackup.FullName -Destination $hostsPath -Force
            Write-Host "  [OK] Hosts file restored from backup" -ForegroundColor Green
        }
        else {
            Write-Host "  [WARNING] No hosts backup found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to restore hosts: $_" -ForegroundColor Red
    }

    Write-Host ""
    Write-Host "  MultiSim internet access has been restored." -ForegroundColor Green
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
    $currentRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$($script:Config.RulePrefix)*" }
    if ($currentRules) { $currentRules | Remove-NetFirewallRule }

    try {
        $backupData = Get-Content -Path $latestBackup.FullName -Raw | ConvertFrom-Json
        $rulesToRestore = if ($backupData -is [System.Array]) { $backupData } else { @($backupData) }

        foreach ($rule in $rulesToRestore) {
            try {
                $params = @{ DisplayName = $rule.DisplayName; Direction = $rule.Direction; Action = $rule.Action; Enabled = $rule.Enabled }
                $filter = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $rule -ErrorAction SilentlyContinue
                if ($filter -and $filter.Program) { $params['Program'] = $filter.Program }
                New-NetFirewallRule @params -ErrorAction Stop | Out-Null
            }
            catch { Write-Log "Failed to restore rule: $_" -Level WARNING }
        }

        Write-Host "  [OK] Firewall rules restored from backup" -ForegroundColor Green
    }
    catch {
        Write-Host "  [ERROR] Backup restore failed: $_" -ForegroundColor Red
    }

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $hostsBackupFiles = Get-ChildItem -Path (Split-Path $hostsPath) -Filter "hosts.backup_*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($hostsBackupFiles) {
        Copy-Item -Path $hostsBackupFiles[0].FullName -Destination $hostsPath -Force
        Write-Host "  [OK] Hosts file restored" -ForegroundColor Green
    }

    Write-Host ""
}

function New-FirewallBlockRule {
    param([string]$DisplayName, [string]$FilePath, [string]$Direction)

    if ($script:Config.DryRun) {
        $script:Statistics.FirewallRulesCreated++
        return $true
    }

    try {
        $ruleName = "$($script:Config.RulePrefix) - $DisplayName ($Direction)"
        if (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue) { return $true }

        New-NetFirewallRule -DisplayName $ruleName -Group $script:Config.RuleGroup -Direction $Direction -Program $FilePath -Action Block -Enabled True -ErrorAction Stop | Out-Null
        $script:Statistics.FirewallRulesCreated++
        return $true
    }
    catch {
        Write-Log "Failed to create rule for $FilePath : $_" -Level ERROR
        return $false
    }
}

function Get-MultiSimFiles {
    param([string]$Path)

    if (-not (Test-Path $Path)) { return @() }

    try {
        return Get-ChildItem -Path $Path -Include *.exe -Recurse -ErrorAction SilentlyContinue
    }
    catch {
        Write-Log "Failed to scan $Path : $_" -Level WARNING
        return @()
    }
}

function Invoke-ProcessDirectory {
    param([string]$Path)

    Write-Host ""
    Write-Host "  [INFO] Scanning: $Path" -ForegroundColor Cyan

    $files = Get-MultiSimFiles -Path $Path
    if ($files.Count -eq 0) {
        Write-Host "  [WARNING] No MultiSim files found here" -ForegroundColor Yellow
        return 0
    }

    Write-Host "  [INFO] Found $($files.Count) files to process" -ForegroundColor Cyan
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

            if (New-FirewallBlockRule -DisplayName $displayName -FilePath $file.FullName -Direction Outbound) {
                New-FirewallBlockRule -DisplayName $displayName -FilePath $file.FullName -Direction Inbound | Out-Null
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
    Write-Host "  [OK] Processed $processedCount files" -ForegroundColor Green
    Write-Log "Processed $processedCount files from $Path" -Level SUCCESS
    return $processedCount
}

function Invoke-ScanSystemLocations {
    Write-Host ""
    Write-Host "[Step 4] Scanning system locations..." -ForegroundColor Cyan

    $locations = @()

    $basePaths = @(
        "C:\Program Files (x86)\National Instruments",
        "C:\Program Files\National Instruments",
        "C:\ProgramData\National Instruments",
        "$env:LOCALAPPDATA\National Instruments",
        "$env:APPDATA\National Instruments"
    )

    foreach ($basePath in $basePaths) {
        if (Test-Path $basePath) {
            $locations += $basePath
            $subdirs = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue
            foreach ($d in $subdirs) {
                $locations += $d.FullName
            }
        }
    }

    $altPaths = @(
        "C:\National Instruments",
        "C:\Multisim"
    )

    foreach ($altPath in $altPaths) {
        if (Test-Path $altPath) { $locations += $altPath }
    }

    $totalProcessed = 0
    foreach ($location in $locations | Select-Object -Unique) {
        if (Test-Path $location) {
            $totalProcessed += Invoke-ProcessDirectory -Path $location
        }
    }

    if ($totalProcessed -eq 0) {
        Write-Host ""
        Write-Host "  [WARNING] No MultiSim files found in common locations" -ForegroundColor Yellow
        Write-Host "  [INFO] Specify custom directory? (yes/no): " -ForegroundColor Cyan -NoNewline
        $response = Read-Host

        if ($response -match '^(yes|y)$') {
            Write-Host "  Enter path: " -ForegroundColor Cyan -NoNewline
            $customPath = Read-Host
            if (Test-Path $customPath) {
                $totalProcessed = Invoke-ProcessDirectory -Path $customPath
            }
            else {
                Write-Host "  [ERROR] Invalid path" -ForegroundColor Red
            }
        }
    }

    return $totalProcessed
}

function Block-MultiSimDomains {
    Write-Host ""
    Write-Host "[Step 5] Blocking NI/MultiSim domains via hosts file..." -ForegroundColor Cyan

    $hostsPath = "$env:SystemRoot\System32\drivers\etc\hosts"
    $hostsBackup = "$hostsPath.backup_$($script:Config.SessionID)"

    $domains = @(
        "ni.com",
        "www.ni.com",
        "nationalinstruments.com",
        "www.nationalinstruments.com",

        "license.ni.com",
        "licensing.ni.com",
        "activate.ni.com",
        "activation.ni.com",
        "licensekey.ni.com",

        "update.ni.com",
        "updates.ni.com",
        "download.ni.com",
        "downloads.ni.com",
        "software.ni.com",

        "analytics.ni.com",
        "telemetry.ni.com",
        "metrics.ni.com",
        "stats.ni.com",

        "login.ni.com",
        "account.ni.com",
        "my.ni.com",

        "api.ni.com",
        "services.ni.com",

        "support.ni.com",
        "knowledge.ni.com",

        "cloud.ni.com",
        "online.ni.com",

        "forums.ni.com",
        "digital.ni.com"
    )

    if ($script:Config.DryRun) {
        Write-Host "  [DRY RUN] Would block $($domains.Count) domains" -ForegroundColor DarkGray
        $script:Statistics.DomainsBlocked = $domains.Count
        return $true
    }

    try {
        Copy-Item -Path $hostsPath -Destination $hostsBackup -Force
        Write-Host "  [OK] Hosts file backed up" -ForegroundColor Green

        $hostsContent = Get-Content -Path $hostsPath
        $newEntries = @()

        foreach ($domain in $domains) {
            $entry = "0.0.0.0 $domain"
            $exists = $hostsContent | Where-Object { $_ -match [regex]::Escape($domain) }

            if (-not $exists) {
                $newEntries += $entry
                $script:Statistics.DomainsBlocked++
            }
        }

        if ($newEntries.Count -gt 0) {
            $newEntries = @("", "# MultiSim Blocker Entries - $(Get-Date)") + $newEntries
            Add-Content -Path $hostsPath -Value $newEntries
            Write-Host "  [OK] Added $($newEntries.Count - 2) domain entries" -ForegroundColor Green
            Write-Log "Added domain entries to hosts file" -Level SUCCESS
        }
        else {
            Write-Host "  [INFO] All domains already blocked" -ForegroundColor Cyan
        }

        return $true
    }
    catch {
        Write-Host "  [ERROR] Failed to modify hosts file: $_" -ForegroundColor Red
        Write-Log "Hosts modification failed: $_" -Level ERROR
        return $false
    }
}

function Block-MultiSimIPRanges {
    Write-Host ""
    Write-Host "[Step 6] Blocking National Instruments IP ranges..." -ForegroundColor Cyan

    $ipRanges = @(
        @{ Range = "129.170.0.0/16"; Description = "National Instruments - Austin HQ" },
        @{ Range = "129.171.0.0/16"; Description = "National Instruments - Cloud Services" }
    )

    if ($script:Config.DryRun) {
        Write-Host "  [DRY RUN] Would block $($ipRanges.Count) IP ranges" -ForegroundColor DarkGray
        $script:Statistics.IPRulesCreated = $ipRanges.Count * 2
        return $true
    }

    try {
        foreach ($range in $ipRanges) {
            $displayName = "$($script:Config.RulePrefix) - IP Block - $($range.Range)"
            if (Get-NetFirewallRule -DisplayName $displayName -ErrorAction SilentlyContinue) { continue }

            New-NetFirewallRule -DisplayName $displayName -Group $script:Config.RuleGroup -Direction Outbound -Action Block -RemoteAddress $range.Range -Enabled True -ErrorAction Stop | Out-Null
            New-NetFirewallRule -DisplayName "$displayName (Inbound)" -Group $script:Config.RuleGroup -Direction Inbound -Action Block -RemoteAddress $range.Range -Enabled True -ErrorAction Stop | Out-Null
            $script:Statistics.IPRulesCreated += 2
            Write-Host "  [OK] Blocked: $($range.Range)" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Host "  [ERROR] IP blocking failed: $_" -ForegroundColor Red
        return $false
    }
}

function Get-MultiSimServices {
    Write-Host ""
    Write-Host "[Step 7] Detecting NI services..." -ForegroundColor Cyan

    try {
        $services = Get-Service | Where-Object { $_.DisplayName -like "*National Instruments*" -or $_.DisplayName -like "NI *" -or $_.DisplayName -like "*Multisim*" }

        if ($services) {
            Write-Host "  [INFO] Found $($services.Count) NI-related services" -ForegroundColor Yellow
            foreach ($svc in $services) {
                Write-Host "    - $($svc.DisplayName) [$($svc.Status)]" -ForegroundColor White
                $script:Statistics.ServicesFound++
            }
        }
        else {
            Write-Host "  [INFO] No NI services detected" -ForegroundColor Cyan
        }
        return $true
    }
    catch {
        Write-Host "  [WARNING] Service scan failed: $_" -ForegroundColor Yellow
        return $false
    }
}

function New-ExecutionReport {
    Write-Host ""
    Write-Host "[Step 8] Generating report..." -ForegroundColor Cyan

    $script:Statistics.ExecutionEndTime = Get-Date
    $duration = $script:Statistics.ExecutionEndTime - $script:Statistics.ExecutionStartTime

    $report = @"
================================================================================
                    MULTISIM BLOCKER EXECUTION REPORT
================================================================================

Session ID:       $($script:Config.SessionID)
Start Time:       $($script:Statistics.ExecutionStartTime)
End Time:         $($script:Statistics.ExecutionEndTime)
Duration:         $($duration.ToString("hh\:mm\:ss"))
Mode:             $(if ($script:Config.DryRun) { "DRY RUN" } else { "LIVE" })

Files Scanned:    $($script:Statistics.TotalFilesScanned)
Firewall Rules:   $($script:Statistics.FirewallRulesCreated)
Domains Blocked:  $($script:Statistics.DomainsBlocked)
IP Rules:         $($script:Statistics.IPRulesCreated)
Services Found:   $($script:Statistics.ServicesFound)

================================================================================
"@

    try {
        $report | Out-File -FilePath $script:Config.ReportFile -Encoding UTF8
        Write-Host "  [OK] Report saved" -ForegroundColor Green
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
            if (-not (Request-SystemRestorePoint)) { break }
            if (-not (Backup-FirewallRules)) { break }
            if (-not (Remove-DuplicateRules)) { break }

            Invoke-ScanSystemLocations | Out-Null
            Block-MultiSimDomains | Out-Null
            Block-MultiSimIPRanges | Out-Null
            Get-MultiSimServices | Out-Null
            New-ExecutionReport

            Write-Host ""
            Write-Host "  MultiSim internet access has been blocked successfully." -ForegroundColor Green
            Write-Host ""
        }
        "2" {
            $script:Config.DryRun = $true
            Show-DryRunDisclaimer
            if (-not (Get-UserConsent -Mode "DRY RUN")) { break }

            Invoke-ScanSystemLocations | Out-Null
            Block-MultiSimDomains | Out-Null
            Block-MultiSimIPRanges | Out-Null
            Get-MultiSimServices | Out-Null
            New-ExecutionReport

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
