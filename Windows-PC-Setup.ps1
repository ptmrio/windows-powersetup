#Requires -Version 5.1
<#
.SYNOPSIS
    Windows PC Setup Utility - IT Admin tool for cleaning up and configuring new Windows PCs
.DESCRIPTION
    A streamlined GUI-based PowerShell utility for IT administrators to:
    - Remove bloatware (Windows built-in, Lenovo, and other OEM)
    - Configure system settings (Taskbar, Start Menu, Power)
    - Install common applications via winget
.NOTES
    Author: IT Admin Utility
    Version: 2.0
    Supports: Windows 10 and Windows 11
#>

# ============================================================================
# STRICT MODE AND ERROR HANDLING
# ============================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# Global flags
$script:DryRun = $false

# ============================================================================
# INITIALIZATION AND ADMIN CHECK
# ============================================================================

# Check for administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # Relaunch as administrator
    $arguments = "& '" + $MyInvocation.MyCommand.Definition + "'"
    Start-Process powershell -Verb RunAs -ArgumentList $arguments
    exit
}

# Set up logging
$script:LogPath = Join-Path $env:TEMP "PCSetup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $script:LogPath -Append

# Detect Windows version
$OSBuild = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
$OSName = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductName
$IsWindows11 = [int]$OSBuild -ge 22000

Write-Host "Windows PC Setup Utility v2.0 Starting..."
Write-Host "OS: $OSName (Build $OSBuild)"
Write-Host "Windows 11: $IsWindows11"
Write-Host "Log file: $script:LogPath"

# ============================================================================
# LOAD WINDOWS FORMS
# ============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================================
# DEFINE BLOATWARE LISTS
# ============================================================================

# Apps that are ALWAYS protected and never shown for removal
$ProtectedApps = @(
    "Microsoft.WindowsStore",
    "Microsoft.StorePurchaseApp",
    "Microsoft.DesktopAppInstaller",
    "Microsoft.VCLibs*",
    "Microsoft.UI.Xaml*",
    "Microsoft.NET*",
    "Microsoft.WindowsCalculator",
    "Microsoft.WindowsNotepad",
    "Microsoft.Paint",
    "Microsoft.MSPaint",
    "Microsoft.ScreenSketch",
    "Microsoft.Windows.Photos",
    "Microsoft.WindowsCamera",
    "Microsoft.WindowsTerminal",
    "Microsoft.SecHealthUI",
    "Microsoft.AAD.BrokerPlugin",
    "Microsoft.AccountsControl",
    "Microsoft.LockApp",
    "Microsoft.Windows.ShellExperienceHost",
    "Microsoft.Windows.StartMenuExperienceHost",
    "Microsoft.HEIFImageExtension",
    "Microsoft.VP9VideoExtensions",
    "Microsoft.WebMediaExtensions",
    "Microsoft.WebpImageExtension",
    "Microsoft.RawImageExtension",
    "windows.immersivecontrolpanel",
    "Windows.PrintDialog",
    "Microsoft.Windows.ContentDeliveryManager",
    "Microsoft.Windows.Search",
    "Microsoft.Edge",
    "Microsoft.MicrosoftEdge*"
)

# Common bloatware - pre-checked for removal
$CommonBloatware = @(
    # Third-party promotional apps (SAFE to remove)
    @{ Name = "Candy Crush Saga"; PackageName = "king.com.CandyCrushSaga"; PreChecked = $true },
    @{ Name = "Candy Crush Soda Saga"; PackageName = "king.com.CandyCrushSodaSaga"; PreChecked = $true },
    @{ Name = "Candy Crush Friends"; PackageName = "king.com.CandyCrushFriends"; PreChecked = $true },
    @{ Name = "Bubble Witch 3 Saga"; PackageName = "king.com.BubbleWitch3Saga"; PreChecked = $true },
    @{ Name = "TikTok"; PackageName = "BytedancePte.Ltd.TikTok"; PreChecked = $true },
    @{ Name = "Spotify"; PackageName = "SpotifyAB.SpotifyMusic"; PreChecked = $true },
    @{ Name = "Netflix"; PackageName = "4DF9E0F8.Netflix"; PreChecked = $true },
    @{ Name = "Disney+"; PackageName = "Disney.37853FC22B2CE"; PreChecked = $true },
    @{ Name = "Amazon"; PackageName = "Amazon.com.Amazon"; PreChecked = $true },
    @{ Name = "Prime Video"; PackageName = "AmazonVideo.PrimeVideo"; PreChecked = $true },
    @{ Name = "Facebook"; PackageName = "Facebook.Facebook"; PreChecked = $true },
    @{ Name = "Instagram"; PackageName = "Facebook.Instagram"; PreChecked = $true },
    @{ Name = "Twitter/X"; PackageName = "9E2F88E3.Twitter"; PreChecked = $true },
    @{ Name = "LinkedIn"; PackageName = "7EE7776C.LinkedInforWindows"; PreChecked = $true },
    @{ Name = "Duolingo"; PackageName = "D5EA27B7.Duolingo"; PreChecked = $true },
    # Microsoft bloatware (SAFE to remove)
    @{ Name = "Clipchamp"; PackageName = "Clipchamp.Clipchamp"; PreChecked = $true },
    @{ Name = "Microsoft Solitaire"; PackageName = "Microsoft.MicrosoftSolitaireCollection"; PreChecked = $true },
    @{ Name = "Bing News"; PackageName = "Microsoft.BingNews"; PreChecked = $true },
    @{ Name = "Bing Weather"; PackageName = "Microsoft.BingWeather"; PreChecked = $true },
    @{ Name = "Bing Finance"; PackageName = "Microsoft.BingFinance"; PreChecked = $true },
    @{ Name = "Bing Sports"; PackageName = "Microsoft.BingSports"; PreChecked = $true },
    @{ Name = "Get Help"; PackageName = "Microsoft.GetHelp"; PreChecked = $true },
    @{ Name = "Get Started (Tips)"; PackageName = "Microsoft.Getstarted"; PreChecked = $true },
    @{ Name = "Mixed Reality Portal"; PackageName = "Microsoft.MixedReality.Portal"; PreChecked = $true },
    @{ Name = "3D Viewer"; PackageName = "Microsoft.Microsoft3DViewer"; PreChecked = $true },
    @{ Name = "Office Hub"; PackageName = "Microsoft.MicrosoftOfficeHub"; PreChecked = $true },
    @{ Name = "People"; PackageName = "Microsoft.People"; PreChecked = $true },
    @{ Name = "Skype"; PackageName = "Microsoft.SkypeApp"; PreChecked = $true },
    @{ Name = "Groove Music"; PackageName = "Microsoft.ZuneMusic"; PreChecked = $true },
    @{ Name = "Movies & TV"; PackageName = "Microsoft.ZuneVideo"; PreChecked = $true },
    @{ Name = "Feedback Hub"; PackageName = "Microsoft.WindowsFeedbackHub"; PreChecked = $true },
    @{ Name = "Maps"; PackageName = "Microsoft.WindowsMaps"; PreChecked = $true },
    @{ Name = "Power Automate"; PackageName = "Microsoft.PowerAutomateDesktop"; PreChecked = $true },
    @{ Name = "Microsoft To Do"; PackageName = "Microsoft.Todos"; PreChecked = $true },
    # Xbox - CAUTION: Some games may need these
    @{ Name = "Xbox App"; PackageName = "Microsoft.XboxApp"; PreChecked = $true },
    @{ Name = "Xbox Game Bar"; PackageName = "Microsoft.XboxGamingOverlay"; PreChecked = $true },
    # Optional apps - not pre-checked (user may want these)
    @{ Name = "Phone Link (Your Phone)"; PackageName = "Microsoft.YourPhone"; PreChecked = $false },
    @{ Name = "Cortana"; PackageName = "Microsoft.549981C3F5F10"; PreChecked = $false },
    @{ Name = "Copilot"; PackageName = "Microsoft.Copilot"; PreChecked = $false },
    @{ Name = "OneDrive"; PackageName = "Microsoft.OneDrive"; PreChecked = $false },
    @{ Name = "Mail and Calendar"; PackageName = "microsoft.windowscommunicationsapps"; PreChecked = $false },
    @{ Name = "Outlook (New)"; PackageName = "Microsoft.OutlookForWindows"; PreChecked = $false },
    @{ Name = "OneNote"; PackageName = "Microsoft.Office.OneNote"; PreChecked = $false }
)

# Lenovo-specific bloatware
$LenovoBloatware = @(
    @{ Name = "Lenovo Companion"; PackageName = "E046963F.LenovoCompanion"; PreChecked = $true },
    @{ Name = "Lenovo Utility"; PackageName = "E0469640.LenovoUtility"; PreChecked = $true },
    @{ Name = "Lenovo Settings"; PackageName = "LenovoCorporation.LenovoSettings"; PreChecked = $true },
    @{ Name = "Lenovo ID"; PackageName = "LenovoCorporation.LenovoID"; PreChecked = $true },
    @{ Name = "Lenovo Vantage (Enterprise)"; PackageName = "E046963F.LenovoSettingsforEnterprise"; PreChecked = $true }
)

# Win32 bloatware patterns (traditional programs)
$Win32BloatwarePatterns = @(
    @{ Pattern = "^McAfee"; PreChecked = $true; Description = "McAfee Security" },
    @{ Pattern = "^Norton"; PreChecked = $true; Description = "Norton Security" },
    @{ Pattern = "^Avast"; PreChecked = $true; Description = "Avast Antivirus" },
    @{ Pattern = "^AVG "; PreChecked = $true; Description = "AVG Antivirus" },
    @{ Pattern = "WildTangent"; PreChecked = $true; Description = "WildTangent Games" },
    @{ Pattern = "^ExpressVPN"; PreChecked = $true; Description = "ExpressVPN (bundled)" },
    @{ Pattern = "Booking\.com"; PreChecked = $true; Description = "Booking.com" },
    @{ Pattern = "Lenovo App Explorer"; PreChecked = $true; Description = "Lenovo App Explorer" },
    @{ Pattern = "Lenovo Welcome"; PreChecked = $true; Description = "Lenovo Welcome" },
    @{ Pattern = "Lenovo Now"; PreChecked = $true; Description = "Lenovo Now" }
)

# Protected Win32 patterns - never show for removal
$ProtectedWin32Patterns = @(
    "^Microsoft Visual C\+\+",
    "^Microsoft Windows Desktop Runtime",
    "^Microsoft \.NET",
    "^Intel\(R\)",
    "^AMD ",
    "^NVIDIA",
    "^Realtek",
    "^Synaptics"
)

# Apps to install via winget - your specific list
$WingetApps = @(
    @{ Name = "Google Chrome"; WingetId = "Google.Chrome"; Category = "Browser" },
    @{ Name = "Brave Browser"; WingetId = "Brave.Brave"; Category = "Browser" },
    @{ Name = "Adobe Acrobat Reader"; WingetId = "Adobe.Acrobat.Reader.64-bit"; Category = "PDF" },
    @{ Name = "Google Drive"; WingetId = "Google.GoogleDrive"; Category = "Cloud" },
    @{ Name = "Proton Pass"; WingetId = "Proton.ProtonPass"; Category = "Security" },
    @{ Name = "Todoist"; WingetId = "Doist.Todoist"; Category = "Productivity" },
    @{ Name = "PhraseVault"; WingetId = "ptmrio.phrasevault"; Category = "Productivity" },
    @{ Name = "Microsoft PowerToys"; WingetId = "Microsoft.PowerToys"; Category = "Utility" },
    @{ Name = "LocalSend"; WingetId = "LocalSend.LocalSend"; Category = "Utility" },
    @{ Name = "Visual Studio Code"; WingetId = "Microsoft.VisualStudioCode"; Category = "Editor" },
    @{ Name = "7-Zip"; WingetId = "7zip.7zip"; Category = "Archive" },
    @{ Name = "VLC Media Player"; WingetId = "VideoLAN.VLC"; Category = "Media" }
)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = if ($script:DryRun -and $Level -ne "INFO") { "[DRY RUN] " } else { "" }
    $logMessage = "[$timestamp] [$Level] $prefix$Message"

    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "ERROR" { "Red" }
        "WARN" { "Yellow" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
}

function Update-Status {
    param([string]$Message)
    if ($script:StatusLabel) {
        $script:StatusLabel.Text = $Message
        $script:StatusLabel.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }
    Write-Log $Message
}

# ============================================================================
# WIN32 BLOATWARE DETECTION
# ============================================================================

function Test-ProtectedWin32 {
    param([string]$DisplayName)
    foreach ($pattern in $ProtectedWin32Patterns) {
        if ($DisplayName -match $pattern) { return $true }
    }
    return $false
}

function Get-Win32BloatwareMatch {
    param([string]$DisplayName)
    foreach ($bloat in $Win32BloatwarePatterns) {
        if ($DisplayName -match $bloat.Pattern) {
            return $bloat
        }
    }
    return $null
}

function Get-InstalledWin32Programs {
    $programs = @()
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $regPaths) {
        try {
            $items = Get-ItemProperty $path -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.DisplayName -and
                    $_.DisplayName.Trim() -ne "" -and
                    ($_.SystemComponent -ne 1) -and
                    -not (Test-ProtectedWin32 $_.DisplayName)
                }

            foreach ($item in $items) {
                $match = Get-Win32BloatwareMatch $item.DisplayName
                if ($match) {
                    $canUninstall = $false
                    $uninstallCmd = $null

                    if ($item.QuietUninstallString) {
                        $canUninstall = $true
                        $uninstallCmd = $item.QuietUninstallString
                    }
                    elseif ($item.UninstallString -match '\{[0-9A-Fa-f\-]{36}\}') {
                        $canUninstall = $true
                        $guid = $Matches[0]
                        $uninstallCmd = "msiexec.exe /x $guid /qn /norestart"
                    }

                    if ($canUninstall) {
                        $programs += @{
                            DisplayName = $item.DisplayName
                            UninstallCommand = $uninstallCmd
                            PreChecked = $match.PreChecked
                            Category = $match.Description
                        }
                    }
                }
            }
        }
        catch { }
    }

    return $programs | Sort-Object DisplayName -Unique
}

function Uninstall-Win32Program {
    param([string]$DisplayName, [string]$UninstallCommand)

    if ($script:DryRun) {
        Write-Log "Would uninstall Win32: $DisplayName" "INFO"
        return $true
    }

    try {
        Update-Status "Uninstalling $DisplayName..."
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $UninstallCommand" -Wait -PassThru -WindowStyle Hidden

        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Log "Successfully uninstalled: $DisplayName" "SUCCESS"
            return $true
        }
        else {
            Write-Log "Uninstall may have failed for $DisplayName (exit code: $($process.ExitCode))" "WARN"
            return $false
        }
    }
    catch {
        Write-Log "Failed to uninstall $DisplayName : $_" "ERROR"
        return $false
    }
}

# ============================================================================
# APPX BLOATWARE FUNCTIONS
# ============================================================================

function Test-ProtectedApp {
    param([string]$PackageName)
    foreach ($protected in $ProtectedApps) {
        if ($PackageName -like $protected) { return $true }
    }
    return $false
}

function Get-InstalledBloatware {
    $installed = @()
    $allApps = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue

    foreach ($bloatEntry in $CommonBloatware) {
        $app = $allApps | Where-Object { $_.Name -eq $bloatEntry.PackageName }
        if ($app) {
            $installed += @{
                Name = $bloatEntry.Name
                PackageName = $bloatEntry.PackageName
                PreChecked = $bloatEntry.PreChecked
                Type = "Common"
            }
        }
    }

    foreach ($bloatEntry in $LenovoBloatware) {
        $app = $allApps | Where-Object { $_.Name -eq $bloatEntry.PackageName }
        if ($app) {
            $installed += @{
                Name = $bloatEntry.Name
                PackageName = $bloatEntry.PackageName
                PreChecked = $bloatEntry.PreChecked
                Type = "Lenovo"
            }
        }
    }

    return $installed
}

function Remove-BloatwareApp {
    param([string]$PackageName, [string]$DisplayName)

    if ($script:DryRun) {
        Write-Log "Would remove AppX: $DisplayName ($PackageName)" "INFO"
        return $true
    }

    try {
        Update-Status "Removing $DisplayName..."

        # Remove for current user
        Get-AppxPackage -Name $PackageName -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue

        # Remove for all users
        Get-AppxPackage -Name $PackageName -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue

        # Remove provisioned package (prevents reinstall for new users)
        Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -eq $PackageName } |
            Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue

        Write-Log "Successfully removed: $DisplayName" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to remove $DisplayName : $_" "ERROR"
        return $false
    }
}

# ============================================================================
# WINGET FUNCTIONS
# ============================================================================

function Test-WingetInstalled {
    try {
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        return ($null -ne $wingetPath)
    }
    catch { return $false }
}

function Update-WingetSources {
    if ($script:DryRun) {
        Write-Log "Would update winget sources" "INFO"
        return $true
    }

    try {
        Update-Status "Updating winget sources..."
        winget source update 2>&1 | Out-Null
        Write-Log "Winget sources updated" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to update winget sources: $_" "WARN"
        return $false
    }
}

function Install-WingetApp {
    param([string]$WingetId, [string]$DisplayName)

    if ($script:DryRun) {
        Write-Log "Would install via winget: $DisplayName ($WingetId)" "INFO"
        return $true
    }

    try {
        Update-Status "Installing $DisplayName..."
        $result = winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Successfully installed: $DisplayName" "SUCCESS"
            return $true
        }
        elseif ($LASTEXITCODE -eq -1978335189) {
            Write-Log "$DisplayName is already installed/up to date" "INFO"
            return $true
        }
        else {
            Write-Log "Failed to install $DisplayName (Exit code: $LASTEXITCODE)" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Error installing $DisplayName : $_" "ERROR"
        return $false
    }
}

# ============================================================================
# SETTINGS FUNCTIONS
# ============================================================================

function Set-TaskbarSearchIcon {
    if ($script:DryRun) {
        Write-Log "Would set taskbar search to icon only" "INFO"
        return $true
    }

    try {
        Update-Status "Setting Search to icon only..."

        $searchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        if (!(Test-Path $searchPath)) {
            New-Item -Path $searchPath -Force | Out-Null
        }

        # Set search to icon only
        Set-ItemProperty -Path $searchPath -Name "SearchboxTaskbarMode" -Type DWord -Value 1

        # IMPORTANT: Set cache key to prevent Windows from resetting (from research)
        Set-ItemProperty -Path $searchPath -Name "SearchboxTaskbarModeCache" -Type DWord -Value 1

        Write-Log "Taskbar search set to icon only" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to set taskbar search: $_" "ERROR"
        return $false
    }
}

function Set-TaskbarMultiMonitor {
    if ($script:DryRun) {
        Write-Log "Would configure multi-monitor taskbar" "INFO"
        return $true
    }

    try {
        Update-Status "Configuring multi-monitor taskbar..."

        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        # Show taskbar on all displays
        Set-ItemProperty -Path $advancedPath -Name "MMTaskbarEnabled" -Type DWord -Value 1

        # Show apps on taskbar where window is open (value 2)
        Set-ItemProperty -Path $advancedPath -Name "MMTaskbarMode" -Type DWord -Value 2

        Write-Log "Multi-monitor taskbar configured" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure multi-monitor taskbar: $_" "ERROR"
        return $false
    }
}

function Set-TaskbarCombineWhenFull {
    if ($script:DryRun) {
        Write-Log "Would set taskbar to combine when full" "INFO"
        return $true
    }

    try {
        Update-Status "Setting taskbar to combine when full..."

        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        # Combine buttons only when full (0=Always, 1=When full, 2=Never)
        Set-ItemProperty -Path $advancedPath -Name "TaskbarGlomLevel" -Type DWord -Value 1
        Set-ItemProperty -Path $advancedPath -Name "MMTaskbarGlomLevel" -Type DWord -Value 1

        Write-Log "Taskbar combine when full enabled" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to set taskbar combine: $_" "ERROR"
        return $false
    }
}

function Hide-StartMenuRecommended {
    if (-not $IsWindows11) {
        Write-Log "Start Menu Recommended section not applicable for Windows 10" "INFO"
        return $true
    }

    if ($script:DryRun) {
        Write-Log "Would hide Start Menu Recommended section" "INFO"
        return $true
    }

    try {
        Update-Status "Hiding Start Menu Recommended section..."

        # Create Explorer policy key if needed
        $explorerPolicyPath = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        if (!(Test-Path $explorerPolicyPath)) {
            New-Item -Path $explorerPolicyPath -Force | Out-Null
        }

        # Hide Recommended section (Note: Only officially works on Windows 11 SE, but worth trying)
        Set-ItemProperty -Path $explorerPolicyPath -Name "HideRecommendedSection" -Type DWord -Value 1

        # Disable document tracking (clears Recommended items)
        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        Set-ItemProperty -Path $advancedPath -Name "Start_TrackDocs" -Type DWord -Value 0

        # Disable Iris Recommendations (from research)
        Set-ItemProperty -Path $advancedPath -Name "Start_IrisRecommendations" -Type DWord -Value 0

        Write-Log "Start Menu Recommended section hidden (Note: Full effect on Win11 SE only)" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to hide Start Menu Recommended: $_" "ERROR"
        return $false
    }
}

function Disable-BingSearch {
    if ($script:DryRun) {
        Write-Log "Would disable Bing search in Start Menu" "INFO"
        return $true
    }

    try {
        Update-Status "Disabling Bing search in Start Menu..."

        $searchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
        Set-ItemProperty -Path $searchPath -Name "BingSearchEnabled" -Type DWord -Value 0

        $cortanaPath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
        if (!(Test-Path $cortanaPath)) {
            New-Item -Path $cortanaPath -Force | Out-Null
        }
        Set-ItemProperty -Path $cortanaPath -Name "DisableSearchBoxSuggestions" -Type DWord -Value 1

        Write-Log "Bing search disabled" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to disable Bing search: $_" "ERROR"
        return $false
    }
}

function Set-PowerSettings {
    if ($script:DryRun) {
        Write-Log "Would configure power settings (Display off: 10min, No hibernate on AC, Sleep 1hr on battery)" "INFO"
        return $true
    }

    try {
        Update-Status "Configuring power settings..."

        # Display turns off after 10 minutes (AC and DC)
        # AC = plugged in, DC = battery
        powercfg /change monitor-timeout-ac 10
        powercfg /change monitor-timeout-dc 10

        # Never hibernate when on AC power (0 = never)
        powercfg /change hibernate-timeout-ac 0

        # Sleep after 60 minutes on battery (DC)
        powercfg /change standby-timeout-dc 60

        # Never sleep on AC power (optional - keeps PC always on when plugged in)
        powercfg /change standby-timeout-ac 0

        Write-Log "Power settings configured: Display off 10min, Never hibernate/sleep on AC, Sleep 1hr on battery" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure power settings: $_" "ERROR"
        return $false
    }
}

function Enable-ClipboardHistory {
    if ($script:DryRun) {
        Write-Log "Would enable Clipboard History" "INFO"
        return $true
    }

    try {
        Update-Status "Enabling Clipboard History..."

        $clipboardPath = "HKCU:\Software\Microsoft\Clipboard"
        if (!(Test-Path $clipboardPath)) {
            New-Item -Path $clipboardPath -Force | Out-Null
        }

        Set-ItemProperty -Path $clipboardPath -Name "EnableClipboardHistory" -Type DWord -Value 1

        Write-Log "Clipboard History enabled (Note: May have issues on Win11 24H2)" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to enable Clipboard History: $_" "ERROR"
        return $false
    }
}

function Enable-StorageSense {
    if ($script:DryRun) {
        Write-Log "Would enable Storage Sense (Downloads: 14 days, Recycle Bin: 30 days)" "INFO"
        return $true
    }

    try {
        Update-Status "Configuring Storage Sense..."

        $storagePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy"
        if (!(Test-Path $storagePath)) {
            New-Item -Path $storagePath -Force | Out-Null
        }

        # Enable Storage Sense
        Set-ItemProperty -Path $storagePath -Name "01" -Type DWord -Value 1

        # Run Storage Sense automatically (every month = 30)
        Set-ItemProperty -Path $storagePath -Name "2048" -Type DWord -Value 30

        # Enable Recycle Bin cleanup
        Set-ItemProperty -Path $storagePath -Name "08" -Type DWord -Value 1

        # Recycle Bin: Delete files older than 30 days
        Set-ItemProperty -Path $storagePath -Name "256" -Type DWord -Value 30

        # Enable Downloads folder cleanup
        Set-ItemProperty -Path $storagePath -Name "32" -Type DWord -Value 1

        # Downloads: Delete files older than 14 days
        Set-ItemProperty -Path $storagePath -Name "512" -Type DWord -Value 14

        Write-Log "Storage Sense enabled: Downloads cleanup 14 days, Recycle Bin cleanup 30 days" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to configure Storage Sense: $_" "ERROR"
        return $false
    }
}

function Set-ExplorerToThisPC {
    if ($script:DryRun) {
        Write-Log "Would set Explorer to open to This PC" "INFO"
        return $true
    }

    try {
        Update-Status "Setting Explorer to open to This PC..."

        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        # LaunchTo: 1 = This PC, 2 = Quick Access/Home
        Set-ItemProperty -Path $advancedPath -Name "LaunchTo" -Type DWord -Value 1

        Write-Log "Explorer set to open to This PC" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to set Explorer default: $_" "ERROR"
        return $false
    }
}

function Enable-ShowFileExtensions {
    if ($script:DryRun) {
        Write-Log "Would enable Show File Extensions" "INFO"
        return $true
    }

    try {
        Update-Status "Enabling Show File Extensions..."

        $advancedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

        # HideFileExt: 0 = Show extensions, 1 = Hide extensions
        Set-ItemProperty -Path $advancedPath -Name "HideFileExt" -Type DWord -Value 0

        Write-Log "Show file extensions enabled" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to enable show file extensions: $_" "ERROR"
        return $false
    }
}

function Set-StartMenuPins {
    if (-not $IsWindows11) {
        Write-Log "Start Menu pin layout not applicable for Windows 10 (uses tiles)" "INFO"
        return $true
    }

    if ($script:DryRun) {
        Write-Log "Would set Start Menu pins to Explorer, Calculator, Snipping Tool only" "INFO"
        return $true
    }

    try {
        Update-Status "Configuring Start Menu pinned apps..."

        # Windows 11 Start Menu layout is stored in a binary blob
        # We can use a JSON layout file approach for Windows 11 22H2+

        $layoutPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.json"

        # Create minimal Start layout with only the 3 requested apps
        $startLayout = @{
            pinnedList = @(
                @{ desktopAppLink = "%APPDATA%\Microsoft\Windows\Start Menu\Programs\File Explorer.lnk" }
                @{ packagedAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App" }
                @{ packagedAppId = "Microsoft.ScreenSketch_8wekyb3d8bbwe!App" }
            )
        }

        $startLayoutJson = $startLayout | ConvertTo-Json -Depth 10

        # Backup existing layout if present
        if (Test-Path $layoutPath) {
            Copy-Item $layoutPath "$layoutPath.backup" -Force -ErrorAction SilentlyContinue
        }

        # Write new layout
        $startLayoutJson | Out-File -FilePath $layoutPath -Encoding UTF8 -Force

        Write-Log "Start Menu pins configured (Explorer, Calculator, Snipping Tool). May require sign-out to take full effect." "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to set Start Menu pins: $_" "ERROR"
        return $false
    }
}

function Restart-Explorer {
    if ($script:DryRun) {
        Write-Log "Would restart Explorer" "INFO"
        return $true
    }

    try {
        Update-Status "Restarting Explorer to apply changes..."
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process explorer
        Write-Log "Explorer restarted" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to restart Explorer: $_" "ERROR"
        return $false
    }
}

# ============================================================================
# CREATE MAIN FORM
# ============================================================================

$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = "Windows PC Setup Utility v2.0"
$MainForm.Size = New-Object System.Drawing.Size(700, 580)
$MainForm.StartPosition = "CenterScreen"
$MainForm.FormBorderStyle = "FixedSingle"
$MainForm.MaximizeBox = $false
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# ============================================================================
# CREATE TAB CONTROL
# ============================================================================

$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Point(10, 10)
$TabControl.Size = New-Object System.Drawing.Size(665, 430)

# ============================================================================
# TAB 1: BLOATWARE
# ============================================================================

$TabBloatware = New-Object System.Windows.Forms.TabPage
$TabBloatware.Text = "Remove Bloatware"
$TabBloatware.Padding = New-Object System.Windows.Forms.Padding(10)

$BloatwarePanel = New-Object System.Windows.Forms.Panel
$BloatwarePanel.Dock = "Fill"
$BloatwarePanel.AutoScroll = $true

$LblBloatwareInfo = New-Object System.Windows.Forms.Label
$LblBloatwareInfo.Text = "Scanning for installed bloatware..."
$LblBloatwareInfo.Location = New-Object System.Drawing.Point(20, 10)
$LblBloatwareInfo.Size = New-Object System.Drawing.Size(600, 20)
$BloatwarePanel.Controls.Add($LblBloatwareInfo)

# Create checkboxes for bloatware dynamically
$script:BloatwareCheckboxes = @()
$installedBloatware = Get-InstalledBloatware

$yPos = 40
foreach ($bloat in $installedBloatware) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $typeLabel = if ($bloat.Type -eq "Lenovo") { " [Lenovo]" } else { "" }
    $chk.Text = "$($bloat.Name)$typeLabel"
    $chk.Tag = $bloat
    $chk.Location = New-Object System.Drawing.Point(20, $yPos)
    $chk.Size = New-Object System.Drawing.Size(600, 22)
    $chk.Checked = $bloat.PreChecked
    $BloatwarePanel.Controls.Add($chk)
    $script:BloatwareCheckboxes += $chk
    $yPos += 25
}

$LblBloatwareInfo.Text = "Found $($installedBloatware.Count) AppX apps (pre-checked = recommended to remove)"

# Add Win32 bloatware section
$yPos += 15
$LblWin32 = New-Object System.Windows.Forms.Label
$LblWin32.Text = "--- Win32 Programs ---"
$LblWin32.Location = New-Object System.Drawing.Point(20, $yPos)
$LblWin32.Size = New-Object System.Drawing.Size(600, 20)
$LblWin32.ForeColor = [System.Drawing.Color]::DarkRed
$BloatwarePanel.Controls.Add($LblWin32)
$yPos += 22

$script:Win32Checkboxes = @()
$installedWin32 = Get-InstalledWin32Programs

if ($installedWin32.Count -eq 0) {
    $LblNoWin32 = New-Object System.Windows.Forms.Label
    $LblNoWin32.Text = "(No Win32 bloatware detected)"
    $LblNoWin32.Location = New-Object System.Drawing.Point(20, $yPos)
    $LblNoWin32.Size = New-Object System.Drawing.Size(600, 20)
    $LblNoWin32.ForeColor = [System.Drawing.Color]::Gray
    $BloatwarePanel.Controls.Add($LblNoWin32)
    $yPos += 25
}
else {
    foreach ($win32 in $installedWin32) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = "$($win32.DisplayName) [$($win32.Category)]"
        $chk.Tag = $win32
        $chk.Location = New-Object System.Drawing.Point(20, $yPos)
        $chk.Size = New-Object System.Drawing.Size(600, 22)
        $chk.Checked = $win32.PreChecked
        $chk.ForeColor = [System.Drawing.Color]::DarkRed
        $BloatwarePanel.Controls.Add($chk)
        $script:Win32Checkboxes += $chk
        $yPos += 25
    }
}

# Select All / Deselect All buttons
$BtnSelectAllBloat = New-Object System.Windows.Forms.Button
$BtnSelectAllBloat.Text = "Select All"
$BtnSelectAllBloat.Location = New-Object System.Drawing.Point(20, ($yPos + 10))
$BtnSelectAllBloat.Size = New-Object System.Drawing.Size(100, 28)
$BtnSelectAllBloat.Add_Click({
    foreach ($chk in $script:BloatwareCheckboxes) { $chk.Checked = $true }
    foreach ($chk in $script:Win32Checkboxes) { $chk.Checked = $true }
})
$BloatwarePanel.Controls.Add($BtnSelectAllBloat)

$BtnDeselectAllBloat = New-Object System.Windows.Forms.Button
$BtnDeselectAllBloat.Text = "Deselect All"
$BtnDeselectAllBloat.Location = New-Object System.Drawing.Point(130, ($yPos + 10))
$BtnDeselectAllBloat.Size = New-Object System.Drawing.Size(100, 28)
$BtnDeselectAllBloat.Add_Click({
    foreach ($chk in $script:BloatwareCheckboxes) { $chk.Checked = $false }
    foreach ($chk in $script:Win32Checkboxes) { $chk.Checked = $false }
})
$BloatwarePanel.Controls.Add($BtnDeselectAllBloat)

$TabBloatware.Controls.Add($BloatwarePanel)

# ============================================================================
# TAB 2: SETTINGS
# ============================================================================

$TabSettings = New-Object System.Windows.Forms.TabPage
$TabSettings.Text = "Settings"
$TabSettings.Padding = New-Object System.Windows.Forms.Padding(10)

$SettingsPanel = New-Object System.Windows.Forms.Panel
$SettingsPanel.Dock = "Fill"
$SettingsPanel.AutoScroll = $true

$script:SettingsCheckboxes = @()
$settingsYPos = 20

# Taskbar section header
$LblTaskbar = New-Object System.Windows.Forms.Label
$LblTaskbar.Text = "Taskbar Settings:"
$LblTaskbar.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$LblTaskbar.Size = New-Object System.Drawing.Size(600, 20)
$LblTaskbar.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SettingsPanel.Controls.Add($LblTaskbar)
$settingsYPos += 25

# Search Icon
$ChkSearchIcon = New-Object System.Windows.Forms.CheckBox
$ChkSearchIcon.Text = "Reduce taskbar search to icon only"
$ChkSearchIcon.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkSearchIcon.Size = New-Object System.Drawing.Size(600, 25)
$ChkSearchIcon.Checked = $true
$ChkSearchIcon.Tag = "SearchIcon"
$SettingsPanel.Controls.Add($ChkSearchIcon)
$script:SettingsCheckboxes += $ChkSearchIcon
$settingsYPos += 28

# Multi-monitor taskbar
$ChkMultiMonitor = New-Object System.Windows.Forms.CheckBox
$ChkMultiMonitor.Text = "Show taskbar on all displays + show apps where window is open"
$ChkMultiMonitor.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkMultiMonitor.Size = New-Object System.Drawing.Size(600, 25)
$ChkMultiMonitor.Checked = $true
$ChkMultiMonitor.Tag = "MultiMonitor"
$SettingsPanel.Controls.Add($ChkMultiMonitor)
$script:SettingsCheckboxes += $ChkMultiMonitor
$settingsYPos += 28

# Combine when full
$ChkCombineWhenFull = New-Object System.Windows.Forms.CheckBox
$ChkCombineWhenFull.Text = "Combine taskbar buttons only when full"
$ChkCombineWhenFull.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkCombineWhenFull.Size = New-Object System.Drawing.Size(600, 25)
$ChkCombineWhenFull.Checked = $true
$ChkCombineWhenFull.Tag = "CombineWhenFull"
$SettingsPanel.Controls.Add($ChkCombineWhenFull)
$script:SettingsCheckboxes += $ChkCombineWhenFull
$settingsYPos += 35

# Start Menu section header
$LblStartMenu = New-Object System.Windows.Forms.Label
$LblStartMenu.Text = "Start Menu Settings (Windows 11):"
$LblStartMenu.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$LblStartMenu.Size = New-Object System.Drawing.Size(600, 20)
$LblStartMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SettingsPanel.Controls.Add($LblStartMenu)
$settingsYPos += 25

# Start Menu Recommended
$ChkHideRecommended = New-Object System.Windows.Forms.CheckBox
$ChkHideRecommended.Text = "Hide Start Menu Recommended section"
$ChkHideRecommended.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkHideRecommended.Size = New-Object System.Drawing.Size(600, 25)
$ChkHideRecommended.Checked = $IsWindows11
$ChkHideRecommended.Tag = "HideRecommended"
$ChkHideRecommended.Enabled = $IsWindows11
$SettingsPanel.Controls.Add($ChkHideRecommended)
$script:SettingsCheckboxes += $ChkHideRecommended
$settingsYPos += 28

# Disable Bing Search
$ChkDisableBing = New-Object System.Windows.Forms.CheckBox
$ChkDisableBing.Text = "Disable Bing search in Start Menu"
$ChkDisableBing.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkDisableBing.Size = New-Object System.Drawing.Size(600, 25)
$ChkDisableBing.Checked = $true
$ChkDisableBing.Tag = "DisableBing"
$SettingsPanel.Controls.Add($ChkDisableBing)
$script:SettingsCheckboxes += $ChkDisableBing
$settingsYPos += 35

# Power section header
$LblPower = New-Object System.Windows.Forms.Label
$LblPower.Text = "Power Settings:"
$LblPower.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$LblPower.Size = New-Object System.Drawing.Size(600, 20)
$LblPower.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SettingsPanel.Controls.Add($LblPower)
$settingsYPos += 25

# Power Settings
$ChkPowerSettings = New-Object System.Windows.Forms.CheckBox
$ChkPowerSettings.Text = "Configure power: Display off 10min, Never sleep on AC, Sleep 1hr on battery"
$ChkPowerSettings.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkPowerSettings.Size = New-Object System.Drawing.Size(600, 25)
$ChkPowerSettings.Checked = $true
$ChkPowerSettings.Tag = "PowerSettings"
$SettingsPanel.Controls.Add($ChkPowerSettings)
$script:SettingsCheckboxes += $ChkPowerSettings
$settingsYPos += 35

# Other section header
$LblOther = New-Object System.Windows.Forms.Label
$LblOther.Text = "Other Settings:"
$LblOther.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$LblOther.Size = New-Object System.Drawing.Size(600, 20)
$LblOther.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SettingsPanel.Controls.Add($LblOther)
$settingsYPos += 25

# Clipboard History
$ChkClipboard = New-Object System.Windows.Forms.CheckBox
$ChkClipboard.Text = "Enable Windows Clipboard History (Win+V)"
$ChkClipboard.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkClipboard.Size = New-Object System.Drawing.Size(600, 25)
$ChkClipboard.Checked = $true
$ChkClipboard.Tag = "Clipboard"
$SettingsPanel.Controls.Add($ChkClipboard)
$script:SettingsCheckboxes += $ChkClipboard
$settingsYPos += 28

# Storage Sense
$ChkStorageSense = New-Object System.Windows.Forms.CheckBox
$ChkStorageSense.Text = "Enable Storage Sense (Downloads: 14 days, Recycle Bin: 30 days auto-cleanup)"
$ChkStorageSense.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkStorageSense.Size = New-Object System.Drawing.Size(600, 25)
$ChkStorageSense.Checked = $true
$ChkStorageSense.Tag = "StorageSense"
$SettingsPanel.Controls.Add($ChkStorageSense)
$script:SettingsCheckboxes += $ChkStorageSense
$settingsYPos += 35

# Explorer section header
$LblExplorer = New-Object System.Windows.Forms.Label
$LblExplorer.Text = "Explorer Settings:"
$LblExplorer.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$LblExplorer.Size = New-Object System.Drawing.Size(600, 20)
$LblExplorer.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$SettingsPanel.Controls.Add($LblExplorer)
$settingsYPos += 25

# Explorer opens to This PC
$ChkExplorerThisPC = New-Object System.Windows.Forms.CheckBox
$ChkExplorerThisPC.Text = "Set Explorer to open to 'This PC' instead of Quick Access/Home"
$ChkExplorerThisPC.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkExplorerThisPC.Size = New-Object System.Drawing.Size(600, 25)
$ChkExplorerThisPC.Checked = $true
$ChkExplorerThisPC.Tag = "ExplorerThisPC"
$SettingsPanel.Controls.Add($ChkExplorerThisPC)
$script:SettingsCheckboxes += $ChkExplorerThisPC
$settingsYPos += 28

# Show File Extensions
$ChkShowExtensions = New-Object System.Windows.Forms.CheckBox
$ChkShowExtensions.Text = "Show file extensions (e.g., .txt, .exe, .pdf)"
$ChkShowExtensions.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkShowExtensions.Size = New-Object System.Drawing.Size(600, 25)
$ChkShowExtensions.Checked = $true
$ChkShowExtensions.Tag = "ShowExtensions"
$SettingsPanel.Controls.Add($ChkShowExtensions)
$script:SettingsCheckboxes += $ChkShowExtensions
$settingsYPos += 28

# Start Menu Pins (Win11 only)
$ChkStartPins = New-Object System.Windows.Forms.CheckBox
$ChkStartPins.Text = "Set Start Menu pins: Explorer, Calculator, Snipping Tool only (Win11)"
$ChkStartPins.Location = New-Object System.Drawing.Point(30, $settingsYPos)
$ChkStartPins.Size = New-Object System.Drawing.Size(600, 25)
$ChkStartPins.Checked = $IsWindows11
$ChkStartPins.Tag = "StartPins"
$ChkStartPins.Enabled = $IsWindows11
$SettingsPanel.Controls.Add($ChkStartPins)
$script:SettingsCheckboxes += $ChkStartPins

$TabSettings.Controls.Add($SettingsPanel)

# ============================================================================
# TAB 3: INSTALL APPS
# ============================================================================

$TabInstallApps = New-Object System.Windows.Forms.TabPage
$TabInstallApps.Text = "Install Apps"
$TabInstallApps.Padding = New-Object System.Windows.Forms.Padding(10)

$InstallAppsPanel = New-Object System.Windows.Forms.Panel
$InstallAppsPanel.Dock = "Fill"
$InstallAppsPanel.AutoScroll = $true

$wingetAvailable = Test-WingetInstalled
$wingetStatus = if ($wingetAvailable) { "Winget is available - select apps to install" } else { "Winget NOT found - app installation will be skipped" }

$LblWingetStatus = New-Object System.Windows.Forms.Label
$LblWingetStatus.Text = $wingetStatus
$LblWingetStatus.Location = New-Object System.Drawing.Point(20, 10)
$LblWingetStatus.Size = New-Object System.Drawing.Size(600, 20)
$LblWingetStatus.ForeColor = if ($wingetAvailable) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
$InstallAppsPanel.Controls.Add($LblWingetStatus)

$script:AppCheckboxes = @()
$appYPos = 40
$currentCategory = ""

foreach ($app in $WingetApps) {
    if ($app.Category -ne $currentCategory) {
        $currentCategory = $app.Category
        $lblCategory = New-Object System.Windows.Forms.Label
        $lblCategory.Text = "--- $currentCategory ---"
        $lblCategory.Location = New-Object System.Drawing.Point(20, $appYPos)
        $lblCategory.Size = New-Object System.Drawing.Size(600, 20)
        $lblCategory.ForeColor = [System.Drawing.Color]::DarkBlue
        $InstallAppsPanel.Controls.Add($lblCategory)
        $appYPos += 22
    }

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "$($app.Name)"
    $chk.Tag = $app
    $chk.Location = New-Object System.Drawing.Point(35, $appYPos)
    $chk.Size = New-Object System.Drawing.Size(600, 22)
    $chk.Checked = $false
    $chk.Enabled = $wingetAvailable
    $InstallAppsPanel.Controls.Add($chk)
    $script:AppCheckboxes += $chk
    $appYPos += 24
}

# Select All / Deselect All buttons
$BtnSelectAllApps = New-Object System.Windows.Forms.Button
$BtnSelectAllApps.Text = "Select All"
$BtnSelectAllApps.Location = New-Object System.Drawing.Point(20, ($appYPos + 10))
$BtnSelectAllApps.Size = New-Object System.Drawing.Size(100, 28)
$BtnSelectAllApps.Enabled = $wingetAvailable
$BtnSelectAllApps.Add_Click({
    foreach ($chk in $script:AppCheckboxes) { $chk.Checked = $true }
})
$InstallAppsPanel.Controls.Add($BtnSelectAllApps)

$BtnDeselectAllApps = New-Object System.Windows.Forms.Button
$BtnDeselectAllApps.Text = "Deselect All"
$BtnDeselectAllApps.Location = New-Object System.Drawing.Point(130, ($appYPos + 10))
$BtnDeselectAllApps.Size = New-Object System.Drawing.Size(100, 28)
$BtnDeselectAllApps.Enabled = $wingetAvailable
$BtnDeselectAllApps.Add_Click({
    foreach ($chk in $script:AppCheckboxes) { $chk.Checked = $false }
})
$InstallAppsPanel.Controls.Add($BtnDeselectAllApps)

$TabInstallApps.Controls.Add($InstallAppsPanel)

# ============================================================================
# ADD TABS TO TAB CONTROL
# ============================================================================

$TabControl.Controls.AddRange(@($TabBloatware, $TabSettings, $TabInstallApps))
$MainForm.Controls.Add($TabControl)

# ============================================================================
# BOTTOM CONTROLS
# ============================================================================

# Dry Run checkbox
$ChkDryRun = New-Object System.Windows.Forms.CheckBox
$ChkDryRun.Text = "Dry Run (preview only, no changes)"
$ChkDryRun.Location = New-Object System.Drawing.Point(10, 450)
$ChkDryRun.Size = New-Object System.Drawing.Size(250, 22)
$ChkDryRun.Checked = $false
$ChkDryRun.ForeColor = [System.Drawing.Color]::DarkOrange
$MainForm.Controls.Add($ChkDryRun)

# Status label
$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text = "Ready | $OSName | Log: $script:LogPath"
$script:StatusLabel.Location = New-Object System.Drawing.Point(10, 510)
$script:StatusLabel.Size = New-Object System.Drawing.Size(660, 25)
$script:StatusLabel.BorderStyle = "FixedSingle"
$MainForm.Controls.Add($script:StatusLabel)

# Run button
$BtnRun = New-Object System.Windows.Forms.Button
$BtnRun.Text = "Run Selected Tasks"
$BtnRun.Location = New-Object System.Drawing.Point(520, 470)
$BtnRun.Size = New-Object System.Drawing.Size(150, 35)
$BtnRun.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$BtnRun.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$BtnRun.ForeColor = [System.Drawing.Color]::White
$BtnRun.FlatStyle = "Flat"

$BtnRun.Add_Click({
    $BtnRun.Enabled = $false
    $successCount = 0
    $failCount = 0

    $script:DryRun = $ChkDryRun.Checked
    if ($script:DryRun) {
        Write-Log "=== DRY RUN MODE - No changes will be made ===" "INFO"
        Update-Status "DRY RUN MODE - Previewing actions..."
    }

    try {
        # 1. Remove AppX Bloatware
        foreach ($chk in $script:BloatwareCheckboxes) {
            if ($chk.Checked) {
                $bloat = $chk.Tag
                if (Remove-BloatwareApp -PackageName $bloat.PackageName -DisplayName $bloat.Name) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        # 2. Remove Win32 Bloatware
        foreach ($chk in $script:Win32Checkboxes) {
            if ($chk.Checked) {
                $win32 = $chk.Tag
                if (Uninstall-Win32Program -DisplayName $win32.DisplayName -UninstallCommand $win32.UninstallCommand) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        # 3. Apply Settings
        foreach ($chk in $script:SettingsCheckboxes) {
            if ($chk.Checked) {
                $result = $false
                switch ($chk.Tag) {
                    "SearchIcon" { $result = Set-TaskbarSearchIcon }
                    "MultiMonitor" { $result = Set-TaskbarMultiMonitor }
                    "CombineWhenFull" { $result = Set-TaskbarCombineWhenFull }
                    "HideRecommended" { $result = Hide-StartMenuRecommended }
                    "DisableBing" { $result = Disable-BingSearch }
                    "PowerSettings" { $result = Set-PowerSettings }
                    "Clipboard" { $result = Enable-ClipboardHistory }
                    "StorageSense" { $result = Enable-StorageSense }
                    "ExplorerThisPC" { $result = Set-ExplorerToThisPC }
                    "ShowExtensions" { $result = Enable-ShowFileExtensions }
                    "StartPins" { $result = Set-StartMenuPins }
                }
                if ($result) { $successCount++ } else { $failCount++ }
            }
        }

        # 4. Install Apps
        $appsToInstall = $script:AppCheckboxes | Where-Object { $_.Checked }
        if ($appsToInstall.Count -gt 0 -and -not $script:DryRun) {
            Update-WingetSources
        }

        foreach ($chk in $script:AppCheckboxes) {
            if ($chk.Checked) {
                $app = $chk.Tag
                if (Install-WingetApp -WingetId $app.WingetId -DisplayName $app.Name) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        # 5. Restart Explorer if any settings were changed
        $settingsChanged = ($script:SettingsCheckboxes | Where-Object { $_.Checked }).Count -gt 0
        if ($settingsChanged -and -not $script:DryRun) {
            Restart-Explorer
        }

        $dryRunMsg = if ($script:DryRun) { "[DRY RUN] " } else { "" }
        Update-Status "${dryRunMsg}Complete! Success: $successCount, Failed: $failCount"

        [System.Windows.Forms.MessageBox]::Show(
            "${dryRunMsg}Setup complete!`n`nSuccessful operations: $successCount`nFailed operations: $failCount`n`nLog file: $script:LogPath",
            "Windows PC Setup Utility",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Critical error: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred: $_`n`nCheck log file: $script:LogPath",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $BtnRun.Enabled = $true
        $script:DryRun = $false
    }
})

$MainForm.Controls.Add($BtnRun)

# Open Log button
$BtnOpenLog = New-Object System.Windows.Forms.Button
$BtnOpenLog.Text = "Open Log"
$BtnOpenLog.Location = New-Object System.Drawing.Point(10, 475)
$BtnOpenLog.Size = New-Object System.Drawing.Size(100, 28)
$BtnOpenLog.Add_Click({
    if (Test-Path $script:LogPath) {
        Start-Process notepad.exe -ArgumentList $script:LogPath
    }
})
$MainForm.Controls.Add($BtnOpenLog)

# ============================================================================
# SHOW FORM
# ============================================================================

Write-Log "Application started"
[void]$MainForm.ShowDialog()

# Cleanup
Stop-Transcript
