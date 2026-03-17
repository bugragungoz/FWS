#Requires -RunAsAdministrator

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$PrimaryExecutablePath = 'C:\Altair\Altair_PSIM_2025\bin\psim.exe'
$AlternateExecutablePath = 'C:\Altair\Altair_PSIM_2025\PSIM.exe'
$PrimaryInstallPath    = 'C:\Altair\Altair_PSIM_2025'
$PrimaryInboundRule    = 'Altair_PSIM_Privacy_Block_In'
$PrimaryOutboundRule   = 'Altair_PSIM_Privacy_Block_Out'
$RulePrefix            = 'Altair_PSIM_Privacy_Block'
$RuleGroup             = 'Altair PSIM Privacy Block'
$SearchKeywords        = @('altair', 'psim')

function Test-Administrator {
    $isAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )

    if (-not $isAdministrator) {
        throw 'Administrator privileges are required. Please run this script in an elevated PowerShell session.'
    }
}

function Get-ShortTextHash {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())
        $hashBytes = $sha256.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hashBytes) -replace '-', '').Substring(0, 12)
    }
    finally {
        $sha256.Dispose()
    }
}

function Add-DirectoryCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$DirectoryMap
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return
    }

    try {
        $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
        [void]$DirectoryMap.Set_Item($resolvedPath, $true)
    }
    catch {
        [void]$DirectoryMap.Set_Item($Path, $true)
    }
}

function Add-ExecutableCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$ExecutableMap
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return
    }

    $resolvedPath = $null

    try {
        $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    }
    catch {
        $resolvedPath = $Path
    }

    if ($ExecutableMap.ContainsKey($resolvedPath)) {
        return $null
    }

    $ExecutableMap[$resolvedPath] = $true
    return $resolvedPath
}

function Get-AltairPSIMScanDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PrimaryExePath,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryInstallRoot,
        [Parameter(Mandatory = $true)]
        [string[]]$Keywords
    )

    $directoryMap = @{}

    $primaryExeDirectory = Split-Path -Path $PrimaryExePath -Parent
    $primarySuitePath = if ($primaryExeDirectory) { Split-Path -Path $primaryExeDirectory -Parent } else { $null }

    $knownPaths = @(
        $PrimaryInstallRoot,
        $primaryExeDirectory,
        $primarySuitePath,
        'C:\Altair',
        'C:\Program Files\Altair',
        'C:\Program Files\Altair Engineering',
        'C:\Program Files (x86)\Altair',
        'C:\Program Files (x86)\Altair Engineering',
        'C:\ProgramData\Altair',
        "$env:LOCALAPPDATA\Altair",
        "$env:APPDATA\Altair",
        'D:\Altair',
        'D:\Program Files\Altair',
        'D:\Program Files (x86)\Altair'
    )

    foreach ($path in $knownPaths) {
        Add-DirectoryCandidate -Path $path -DirectoryMap $directoryMap
    }

    $discoveryRoots = @(
        'C:\Program Files',
        'C:\Program Files (x86)',
        'C:\ProgramData',
        $env:LOCALAPPDATA,
        $env:APPDATA,
        $PrimaryInstallRoot,
        'C:\Altair',
        'D:\Program Files',
        'D:\Program Files (x86)',
        'D:\ProgramData',
        'D:\Altair'
    ) | Select-Object -Unique

    foreach ($root in $discoveryRoots) {
        if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path -LiteralPath $root -PathType Container)) {
            continue
        }

        try {
            $matchedDirectories = Get-ChildItem -LiteralPath $root -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object {
                $lowerFullName = $_.FullName.ToLowerInvariant()
                foreach ($keyword in $Keywords) {
                    if ($lowerFullName.Contains($keyword)) {
                        return $true
                    }
                }
                return $false
            }

            foreach ($directory in $matchedDirectories) {
                Add-DirectoryCandidate -Path $directory.FullName -DirectoryMap $directoryMap
            }
        }
        catch {
            Write-Verbose "Skipping inaccessible root: $root"
        }
    }

    return $directoryMap.Keys | Sort-Object
}

function Get-AltairPSIMExecutables {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Directories,
        [Parameter(Mandatory = $true)]
        [string[]]$Keywords,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryExePath
    )

    $executableMap = @{}

    if (Test-Path -LiteralPath $PrimaryExePath -PathType Leaf) {
        $addedPrimaryPath = Add-ExecutableCandidate -Path $PrimaryExePath -ExecutableMap $executableMap
        if ($addedPrimaryPath) {
            Write-Host ("[FOUND {0}] {1}" -f $executableMap.Count, $addedPrimaryPath) -ForegroundColor Gray
        }
    }

    $totalDirectories = $Directories.Count
    $directoryIndex = 0

    foreach ($directoryPath in $Directories) {
        $directoryIndex++
        Write-Host ("[SCAN {0}/{1}] Searching directory: {2}" -f $directoryIndex, $totalDirectories, $directoryPath) -ForegroundColor DarkGray

        if (-not (Test-Path -LiteralPath $directoryPath -PathType Container)) {
            Write-Host '  [SKIP] Directory is not accessible.' -ForegroundColor DarkYellow
            continue
        }

        try {
            $executables = Get-ChildItem -LiteralPath $directoryPath -Filter '*.exe' -Recurse -File -ErrorAction SilentlyContinue
            $matchesInDirectory = 0

            foreach ($executable in $executables) {
                $lowerPath = $executable.FullName.ToLowerInvariant()
                $isRelevant = $false

                foreach ($keyword in $Keywords) {
                    if ($lowerPath.Contains($keyword)) {
                        $isRelevant = $true
                        break
                    }
                }

                if ($isRelevant) {
                    $addedPath = Add-ExecutableCandidate -Path $executable.FullName -ExecutableMap $executableMap
                    if ($addedPath) {
                        $matchesInDirectory++
                        Write-Host ("  [FOUND {0}] {1}" -f $executableMap.Count, $addedPath) -ForegroundColor Gray
                    }
                }
            }

            Write-Host ("  [SCAN RESULT] Directory matches: {0} | Running total: {1}" -f $matchesInDirectory, $executableMap.Count) -ForegroundColor DarkCyan
        }
        catch {
            Write-Host ("  [WARNING] Failed to scan directory: {0}" -f $directoryPath) -ForegroundColor Yellow
            Write-Verbose "Skipping inaccessible directory: $directoryPath"
        }
    }

    return $executableMap.Keys | Sort-Object
}

function Remove-ManagedFirewallRules {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayNamePrefix
    )

    $existingRules = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "$DisplayNamePrefix*" }
    if (-not $existingRules) {
        return 0
    }

    $existingRules | Remove-NetFirewallRule
    return $existingRules.Count
}

function New-ProgramBlockRules {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProgramPath,
        [Parameter(Mandatory = $true)]
        [string]$DisplayPrefix,
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryPath,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryInRule,
        [Parameter(Mandatory = $true)]
        [string]$PrimaryOutRule
    )

    $inboundRuleName = $null
    $outboundRuleName = $null

    if (-not [string]::IsNullOrWhiteSpace($PrimaryPath) -and ($ProgramPath -ieq $PrimaryPath)) {
        $inboundRuleName = $PrimaryInRule
        $outboundRuleName = $PrimaryOutRule
    }
    else {
        $ruleHash = Get-ShortTextHash -Text $ProgramPath
        $inboundRuleName = "${DisplayPrefix}_In_${ruleHash}"
        $outboundRuleName = "${DisplayPrefix}_Out_${ruleHash}"
    }

    $commonRuleParams = @{
        Program = $ProgramPath
        Action  = 'Block'
        Profile = 'Any'
        Enabled = 'True'
        Group   = $GroupName
    }

    New-NetFirewallRule @commonRuleParams -DisplayName $inboundRuleName -Direction Inbound | Out-Null
    New-NetFirewallRule @commonRuleParams -DisplayName $outboundRuleName -Direction Outbound | Out-Null

    return [PSCustomObject]@{
        ProgramPath   = $ProgramPath
        InboundRule   = $inboundRuleName
        OutboundRule  = $outboundRuleName
    }
}

function Test-RuleSetActive {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$RuleSet
    )

    foreach ($ruleEntry in $RuleSet) {
        $inboundRule = Get-NetFirewallRule -DisplayName $ruleEntry.InboundRule -ErrorAction SilentlyContinue
        $outboundRule = Get-NetFirewallRule -DisplayName $ruleEntry.OutboundRule -ErrorAction SilentlyContinue

        if (($null -eq $inboundRule) -or ($null -eq $outboundRule) -or ($inboundRule.Enabled -ne 'True') -or ($outboundRule.Enabled -ne 'True')) {
            return $false
        }
    }

    return $true
}

try {
    Test-Administrator

    Write-Host ''
    Write-Host '[1/5] Checking primary install and executable paths...' -ForegroundColor Cyan

    $resolvedPrimaryPath = $null
    $resolvedInstallPath = $null

    if (Test-Path -LiteralPath $PrimaryInstallPath -PathType Container) {
        $resolvedInstallPath = (Resolve-Path -LiteralPath $PrimaryInstallPath).Path
        Write-Host "[OK] Primary install root found: $resolvedInstallPath" -ForegroundColor Green
    }
    else {
        Write-Host "[WARNING] Primary install root not found: $PrimaryInstallPath" -ForegroundColor Yellow
    }

    if (Test-Path -LiteralPath $PrimaryExecutablePath -PathType Leaf) {
        $resolvedPrimaryPath = (Resolve-Path -LiteralPath $PrimaryExecutablePath).Path
        Write-Host "[OK] Primary executable found: $resolvedPrimaryPath" -ForegroundColor Green
    }
    elseif (Test-Path -LiteralPath $AlternateExecutablePath -PathType Leaf) {
        $resolvedPrimaryPath = (Resolve-Path -LiteralPath $AlternateExecutablePath).Path
        Write-Host "[OK] Alternate executable found: $resolvedPrimaryPath" -ForegroundColor Green
    }
    else {
        Write-Host "[WARNING] Primary executable not found: $PrimaryExecutablePath" -ForegroundColor Yellow

        if ($resolvedInstallPath) {
            Write-Host '[INFO] Searching fallback psim.exe under install root...' -ForegroundColor Cyan
            $fallbackPsim = Get-ChildItem -LiteralPath $resolvedInstallPath -Filter 'psim.exe' -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1

            if ($fallbackPsim) {
                $resolvedPrimaryPath = $fallbackPsim.FullName
                Write-Host "[OK] Fallback psim.exe discovered: $resolvedPrimaryPath" -ForegroundColor Green
            }
            else {
                Write-Host '[WARNING] No fallback psim.exe found under install root. Continuing with keyword-based executable discovery.' -ForegroundColor Yellow
            }
        }
        else {
            Write-Host '[INFO] Continuing with Altair/PSIM executable discovery.' -ForegroundColor Yellow
        }
    }

    Write-Host ''
    Write-Host '[2/5] Discovering known Altair/PSIM directories...' -ForegroundColor Cyan
    $scanDirectories = @(Get-AltairPSIMScanDirectories -PrimaryExePath $PrimaryExecutablePath -PrimaryInstallRoot $PrimaryInstallPath -Keywords $SearchKeywords)

    if ($resolvedInstallPath -and ($scanDirectories -notcontains $resolvedInstallPath)) {
        $scanDirectories = @($resolvedInstallPath) + $scanDirectories
        $scanDirectories = $scanDirectories | Select-Object -Unique
        Write-Host "[INFO] Forced primary install root into scan list: $resolvedInstallPath" -ForegroundColor Cyan
    }

    if ($scanDirectories.Count -eq 0) {
        throw 'No Altair/PSIM directories were found in known scan locations.'
    }

    Write-Host "[OK] Directories discovered: $($scanDirectories.Count)" -ForegroundColor Green
    $directoryListIndex = 0
    foreach ($directoryPath in $scanDirectories) {
        $directoryListIndex++
        Write-Host ("  [DIR {0}/{1}] {2}" -f $directoryListIndex, $scanDirectories.Count, $directoryPath) -ForegroundColor Gray
    }

    Write-Host ''
    Write-Host '[3/5] Collecting Altair/PSIM executable files (*.exe only)...' -ForegroundColor Cyan
    $effectivePrimaryExePath = if ($resolvedPrimaryPath) { $resolvedPrimaryPath } else { $PrimaryExecutablePath }
    $targetExecutables = @(Get-AltairPSIMExecutables -Directories $scanDirectories -Keywords $SearchKeywords -PrimaryExePath $effectivePrimaryExePath)

    if ($targetExecutables.Count -eq 0) {
        throw 'No Altair/PSIM executable files were found to block.'
    }

    Write-Host "[OK] Executables found: $($targetExecutables.Count)" -ForegroundColor Green
    Write-Host '[INFO] Final target executable list:' -ForegroundColor Cyan
    $executableListIndex = 0
    foreach ($executablePath in $targetExecutables) {
        $executableListIndex++
        Write-Host ("  [EXE {0}/{1}] {2}" -f $executableListIndex, $targetExecutables.Count, $executablePath) -ForegroundColor Gray
    }

    Write-Host ''
    Write-Host '[4/5] Removing previously managed firewall rules...' -ForegroundColor Cyan
    $removedRuleCount = Remove-ManagedFirewallRules -DisplayNamePrefix $RulePrefix
    Write-Host "[OK] Removed existing managed rules: $removedRuleCount" -ForegroundColor Green

    Write-Host ''
    Write-Host '[5/5] Creating inbound and outbound block rules...' -ForegroundColor Cyan
    $createdRuleEntries = @()
    $processedExecutables = 0
    $createdRulesInRun = 0

    foreach ($programPath in $targetExecutables) {
        $processedExecutables++
        $percentage = [math]::Round(($processedExecutables / $targetExecutables.Count) * 100)
        Write-Progress -Activity "Creating Firewall Rules for PSIM" -Status "$percentage% Complete" -PercentComplete $percentage -CurrentOperation $programPath
        
        Write-Host ("[RULE {0}/{1}] Processing: {2}" -f $processedExecutables, $targetExecutables.Count, $programPath) -ForegroundColor Cyan

        $newRuleEntry = New-ProgramBlockRules `
            -ProgramPath $programPath `
            -DisplayPrefix $RulePrefix `
            -GroupName $RuleGroup `
            -PrimaryPath $resolvedPrimaryPath `
            -PrimaryInRule $PrimaryInboundRule `
            -PrimaryOutRule $PrimaryOutboundRule

        $createdRuleEntries += $newRuleEntry
        $createdRulesInRun += 2

        Write-Host ("  [ADDED] {0}" -f $newRuleEntry.InboundRule) -ForegroundColor Gray
        Write-Host ("  [ADDED] {0}" -f $newRuleEntry.OutboundRule) -ForegroundColor Gray
        Write-Host ("  [PROGRESS] Executables: {0}/{1} | Rules created in run: {2}" -f $processedExecutables, $targetExecutables.Count, $createdRulesInRun) -ForegroundColor Green
    }

    Write-Progress -Activity "Creating Firewall Rules for PSIM" -Completed
    if (-not (Test-RuleSetActive -RuleSet $createdRuleEntries)) {
        throw 'Firewall rules were created but could not be verified as active.'
    }

    $totalRules = $createdRuleEntries.Count * 2

    Write-Host ''
    Write-Host 'Success: Altair/PSIM firewall block rules are active.' -ForegroundColor Green
    Write-Host "Executables blocked: $($createdRuleEntries.Count)" -ForegroundColor Green
    Write-Host "Rules active: $totalRules (Inbound + Outbound)" -ForegroundColor Green
    Write-Host "Rule group: $RuleGroup" -ForegroundColor Green
    Write-Host ''
}
catch {
    Write-Error "Failed to configure firewall blocking for Altair/PSIM. $($_.Exception.Message)"
    exit 1
}