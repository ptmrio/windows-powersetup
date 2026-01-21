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

# Global error handler - ensures errors are visible
trap {
    $errorMsg = "FATAL ERROR: $($_.Exception.Message)`nLine: $($_.InvocationInfo.ScriptLineNumber)`nStack: $($_.ScriptStackTrace)"
    Write-Host $errorMsg -ForegroundColor Red
    Add-Content -Path $script:LogPath -Value $errorMsg -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Windows PC Setup - Error", 0, 16) | Out-Null
    Stop-Transcript -ErrorAction SilentlyContinue
    exit 1
}

# ============================================================================
# UI CONSTANTS
# ============================================================================

$script:UI = @{
    # Colors
    AccentColor      = [System.Drawing.Color]::FromArgb(0, 120, 215)
    AccentHover      = [System.Drawing.Color]::FromArgb(0, 102, 183)
    AccentPressed    = [System.Drawing.Color]::FromArgb(0, 84, 153)
    DisabledBg       = [System.Drawing.Color]::FromArgb(204, 204, 204)
    DisabledFg       = [System.Drawing.Color]::FromArgb(136, 136, 136)
    WarningOrange    = [System.Drawing.Color]::FromArgb(255, 140, 0)
    WarningBg        = [System.Drawing.Color]::FromArgb(255, 250, 230)
    WarningBorder    = [System.Drawing.Color]::FromArgb(255, 193, 7)
    SuccessGreen     = [System.Drawing.Color]::FromArgb(40, 167, 69)
    ErrorRed         = [System.Drawing.Color]::FromArgb(220, 53, 69)
    SectionHeader    = [System.Drawing.Color]::FromArgb(51, 51, 51)
    SubtleGray       = [System.Drawing.Color]::FromArgb(128, 128, 128)

    # Spacing
    ItemSpacing      = 26
    SectionSpacing   = 35
    CheckboxIndent   = 30

    # Sizes
    CheckboxHeight   = 26
    ButtonHeight     = 32
    HeaderFontSize   = 10
    BodyFontSize     = 10
}

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
# UI HELPER FUNCTIONS
# ============================================================================

function Add-ButtonHoverEffect {
    param([System.Windows.Forms.Button]$Button)

    $Button.Add_MouseEnter({
        if ($this.Enabled) {
            $this.BackColor = $script:UI.AccentHover
        }
    })

    $Button.Add_MouseLeave({
        if ($this.Enabled) {
            $this.BackColor = $script:UI.AccentColor
        }
    })

    $Button.Add_MouseDown({
        if ($this.Enabled) {
            $this.BackColor = $script:UI.AccentPressed
        }
    })

    $Button.Add_MouseUp({
        if ($this.Enabled) {
            $this.BackColor = $script:UI.AccentHover
        }
    })
}

function Set-ButtonDisabled {
    param(
        [System.Windows.Forms.Button]$Button,
        [string]$WorkingText = "Working..."
    )
    $Button.Tag = @{ OriginalText = $Button.Text; OriginalBg = $Button.BackColor }
    $Button.Text = $WorkingText
    $Button.BackColor = $script:UI.DisabledBg
    $Button.ForeColor = $script:UI.DisabledFg
    $Button.Enabled = $false
}

function Set-ButtonEnabled {
    param([System.Windows.Forms.Button]$Button)
    if ($Button.Tag -and $Button.Tag.OriginalText) {
        $Button.Text = $Button.Tag.OriginalText
    }
    $Button.BackColor = $script:UI.AccentColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.Enabled = $true
}

function New-SectionHeader {
    param(
        [System.Windows.Forms.Panel]$Panel,
        [string]$Text,
        [int]$YPosition,
        [int]$XPosition = 20
    )

    $header = New-Object System.Windows.Forms.Label
    $header.Text = $Text
    $header.Location = New-Object System.Drawing.Point($XPosition, $YPosition)
    $header.Size = New-Object System.Drawing.Size(600, 24)
    $header.Font = New-Object System.Drawing.Font("Segoe UI Semibold", $script:UI.HeaderFontSize)
    $header.ForeColor = $script:UI.SectionHeader
    $Panel.Controls.Add($header)

    # Subtle underline separator
    $separator = New-Object System.Windows.Forms.Label
    $separator.Location = New-Object System.Drawing.Point($XPosition, ($YPosition + 22))
    $separator.Size = New-Object System.Drawing.Size(600, 1)
    $separator.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $Panel.Controls.Add($separator)

    return ($YPosition + 28)
}

function New-IntroText {
    param(
        [System.Windows.Forms.Panel]$Panel,
        [string]$Text,
        [int]$YPosition,
        [int]$XPosition = 20
    )

    $intro = New-Object System.Windows.Forms.Label
    $intro.Text = $Text
    $intro.Location = New-Object System.Drawing.Point($XPosition, $YPosition)
    $intro.Size = New-Object System.Drawing.Size(610, 36)
    $intro.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $intro.ForeColor = $script:UI.SubtleGray
    $Panel.Controls.Add($intro)

    return ($YPosition + 40)
}

function Update-SelectionCount {
    param(
        [System.Windows.Forms.Label]$Label,
        [array]$Checkboxes,
        [string]$ItemType = "item"
    )
    $count = @($Checkboxes | Where-Object { $_.Checked }).Count
    $plural = if ($count -eq 1) { "" } else { "s" }
    $Label.Text = "$count $ItemType$plural selected"
}

function Update-Progress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$CurrentItem = ""
    )
    if ($script:ProgressBar -and $script:ProgressLabel) {
        $script:ProgressBar.Maximum = $Total
        $script:ProgressBar.Value = [Math]::Min($Current, $Total)
        $script:ProgressBar.Visible = $true
        $script:ProgressLabel.Text = "Processing $Current of $Total$(if ($CurrentItem) { ": $CurrentItem" })"
        $script:ProgressLabel.Visible = $true
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Hide-Progress {
    if ($script:ProgressBar) {
        $script:ProgressBar.Visible = $false
        $script:ProgressBar.Value = 0
    }
    if ($script:ProgressLabel) {
        $script:ProgressLabel.Visible = $false
    }
}

function Show-ConfirmationDialog {
    param(
        [string]$Title,
        [int]$SelectedCount,
        [string]$ActionType,
        [bool]$IsDryRun
    )

    $dryRunNote = if ($IsDryRun) { "`n`n[DRY RUN MODE - No actual changes will be made]" } else { "" }
    $message = "You are about to $ActionType $SelectedCount item(s).$dryRunNote`n`nDo you want to continue?"

    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
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
                    $null -ne $_.DisplayName -and
                    $_.DisplayName.Trim() -ne "" -and
                    ($null -eq $_.SystemComponent -or $_.SystemComponent -ne 1) -and
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

    $maxRetries = 3
    $retryDelay = 5

    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            if ($attempt -gt 1) {
                Update-Status "Installing $DisplayName (attempt $attempt of $maxRetries)..."
                Start-Sleep -Seconds $retryDelay
            } else {
                Update-Status "Installing $DisplayName..."
            }

            # Use --disable-interactivity for cleaner automation
            $result = winget install --id $WingetId --silent --accept-package-agreements --accept-source-agreements --disable-interactivity 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Log "Successfully installed: $DisplayName" "SUCCESS"
                return $true
            }
            elseif ($LASTEXITCODE -eq -1978335189) {
                Write-Log "$DisplayName is already installed/up to date" "INFO"
                return $true
            }
            elseif ($LASTEXITCODE -eq -2147012889) {
                # Network timeout error - retry
                Write-Log "Network timeout for $DisplayName (attempt $attempt of $maxRetries)" "WARN"
                if ($attempt -eq $maxRetries) {
                    Write-Log "Failed to install $DisplayName after $maxRetries attempts - network timeout. Check internet connection or try again later." "ERROR"
                    return $false
                }
                # Continue to next retry
            }
            else {
                Write-Log "Failed to install $DisplayName (Exit code: $LASTEXITCODE)" "ERROR"
                return $false
            }
        }
        catch {
            Write-Log "Error installing $DisplayName : $_" "ERROR"
            if ($attempt -eq $maxRetries) {
                return $false
            }
        }
    }

    return $false
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
        Write-Log "Would enable Storage Sense (Daily cleanup: Downloads 14 days, Recycle Bin 30 days)" "INFO"
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

        # Run Storage Sense automatically (every day = 1)
        Set-ItemProperty -Path $storagePath -Name "2048" -Type DWord -Value 1

        # Enable Recycle Bin cleanup
        Set-ItemProperty -Path $storagePath -Name "08" -Type DWord -Value 1

        # Recycle Bin: Delete files older than 30 days
        Set-ItemProperty -Path $storagePath -Name "256" -Type DWord -Value 30

        # Enable Downloads folder cleanup
        Set-ItemProperty -Path $storagePath -Name "32" -Type DWord -Value 1

        # Downloads: Delete files older than 14 days
        Set-ItemProperty -Path $storagePath -Name "512" -Type DWord -Value 14

        Write-Log "Storage Sense enabled: Daily cleanup - Downloads 14 days, Recycle Bin 30 days" "SUCCESS"
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

function Clear-ExplorerQuickAccess {
    if ($script:DryRun) {
        Write-Log "Would clean Explorer Quick Access (keep only Desktop and Downloads)" "INFO"
        return $true
    }

    try {
        Update-Status "Cleaning Explorer Quick Access pinned folders..."

        # Load Shell COM object for Quick Access manipulation
        $shell = New-Object -ComObject Shell.Application

        # Get the Quick Access namespace (CLSID for Quick Access)
        $quickAccess = $shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")

        if ($null -eq $quickAccess) {
            Write-Log "Could not access Quick Access namespace" "WARN"
            return $false
        }

        # Folders to KEEP pinned (by name - localized names may vary)
        $keepFolders = @("Desktop", "Downloads")

        # Get all pinned items
        $pinnedItems = $quickAccess.Items()
        $removedCount = 0

        foreach ($item in $pinnedItems) {
            $itemName = $item.Name
            $itemPath = $item.Path

            # Check if this is a pinned folder (not a recent file)
            # Pinned folders have the "isfolder" property
            $isFolder = $item.IsFolder

            if ($isFolder) {
                # Check if folder name matches one we want to keep
                $shouldKeep = $false
                foreach ($keep in $keepFolders) {
                    if ($itemName -eq $keep -or $itemPath -like "*\$keep" -or $itemPath -like "*\$keep\") {
                        $shouldKeep = $true
                        break
                    }
                }

                if (-not $shouldKeep) {
                    # Get the verbs (context menu actions) for this item
                    $verbs = $item.Verbs()
                    foreach ($verb in $verbs) {
                        # Look for "Unpin from Quick access" or similar (varies by locale)
                        # The verb name is "unpinfromhome" internally
                        if ($verb.Name -match "Unpin|Entfernen|L.sen" -or $verb.Name -eq "&Unpin from Quick access") {
                            $verb.DoIt()
                            Write-Log "Unpinned from Quick Access: $itemName" "SUCCESS"
                            $removedCount++
                            Start-Sleep -Milliseconds 200
                            break
                        }
                    }
                }
            }
        }

        # Alternative approach: Clear via registry if COM doesn't work well
        # This clears the "frequent folders" that appear automatically
        $explorerPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer"

        # Disable "Show frequently used folders in Quick access"
        Set-ItemProperty -Path $explorerPath -Name "ShowFrequent" -Type DWord -Value 0 -ErrorAction SilentlyContinue

        # Disable "Show recently used files in Quick access"
        Set-ItemProperty -Path $explorerPath -Name "ShowRecent" -Type DWord -Value 0 -ErrorAction SilentlyContinue

        Write-Log "Explorer Quick Access cleaned. Removed $removedCount items. Kept: Desktop, Downloads" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to clean Explorer Quick Access: $_" "ERROR"
        return $false
    }
    finally {
        # Release COM object
        if ($null -ne $shell) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
        }
    }
}

function Set-StartMenuPins {
    if (-not $IsWindows11) {
        Write-Log "Start Menu pin layout not applicable for Windows 10 (uses tiles)" "INFO"
        return $true
    }

    if ($script:DryRun) {
        Write-Log "Would reset Start Menu pins (delete start2.bin to reset to defaults)" "INFO"
        return $true
    }

    try {
        Update-Status "Resetting Start Menu pinned apps..."

        # Windows 11 stores Start Menu pins in a binary file that cannot be easily edited
        # The most reliable approach is to delete the binary layout file to reset to defaults
        # User will need to manually pin the 3 desired apps after reset

        $startMenuHost = Get-AppxPackage -Name "Microsoft.Windows.StartMenuExperienceHost" -ErrorAction SilentlyContinue
        if ($startMenuHost) {
            $startBinPath = Join-Path $env:LOCALAPPDATA "Packages\$($startMenuHost.PackageFamilyName)\LocalState\start2.bin"

            if (Test-Path $startBinPath) {
                # Backup existing layout
                Copy-Item $startBinPath "$startBinPath.backup" -Force -ErrorAction SilentlyContinue
                Write-Log "Backed up existing Start layout to $startBinPath.backup" "INFO"

                # Stop StartMenuExperienceHost to release file lock
                Get-Process -Name "StartMenuExperienceHost" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500

                # Delete the start layout file
                Remove-Item $startBinPath -Force -ErrorAction SilentlyContinue
                Write-Log "Deleted Start Menu layout file - pins will reset to defaults on next login" "SUCCESS"
            }
            else {
                Write-Log "Start layout file not found at expected location" "WARN"
            }
        }
        else {
            Write-Log "StartMenuExperienceHost package not found" "WARN"
        }

        # Also clear the LayoutModification.json if it exists (for consistency)
        $layoutJsonPath = "$env:LOCALAPPDATA\Microsoft\Windows\Shell\LayoutModification.json"
        if (Test-Path $layoutJsonPath) {
            Remove-Item $layoutJsonPath -Force -ErrorAction SilentlyContinue
        }

        Write-Log "Start Menu reset complete. After restart/sign-out, please manually pin: Explorer, Calculator, Snipping Tool" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to reset Start Menu pins: $_" "ERROR"
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
$MainForm.Size = New-Object System.Drawing.Size(700, 620)
$MainForm.StartPosition = "CenterScreen"
$MainForm.FormBorderStyle = "FixedSingle"
$MainForm.MaximizeBox = $false
$MainForm.Font = New-Object System.Drawing.Font("Segoe UI", $script:UI.BodyFontSize)

# Create tooltip component for the entire form
$script:MainTooltip = New-Object System.Windows.Forms.ToolTip
$script:MainTooltip.AutoPopDelay = 10000
$script:MainTooltip.InitialDelay = 500
$script:MainTooltip.ReshowDelay = 200
$script:MainTooltip.ShowAlways = $true

# ============================================================================
# CREATE TAB CONTROL
# ============================================================================

$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Point(10, 10)
$TabControl.Size = New-Object System.Drawing.Size(665, 450)

# ============================================================================
# TAB 1: BLOATWARE
# ============================================================================

$TabBloatware = New-Object System.Windows.Forms.TabPage
$TabBloatware.Text = "&Remove Bloatware"
$TabBloatware.Padding = New-Object System.Windows.Forms.Padding(10)

$BloatwarePanel = New-Object System.Windows.Forms.Panel
$BloatwarePanel.Dock = "Fill"
$BloatwarePanel.AutoScroll = $true

# Introduction text
$yPos = New-IntroText -Panel $BloatwarePanel -Text "Select pre-installed apps to remove. Pre-checked items are safe for most users. Hover over items for more details." -YPosition 8

# AppX Apps section header
$yPos = New-SectionHeader -Panel $BloatwarePanel -Text "Store Apps (AppX)" -YPosition $yPos

# Create checkboxes for bloatware dynamically
$script:BloatwareCheckboxes = @()
$installedBloatware = Get-InstalledBloatware

# Tooltip descriptions for common apps
$bloatwareTooltips = @{
    "Candy Crush Saga" = "Promotional puzzle game - safe to remove"
    "Candy Crush Soda Saga" = "Promotional puzzle game - safe to remove"
    "Candy Crush Friends" = "Promotional puzzle game - safe to remove"
    "Bubble Witch 3 Saga" = "Promotional puzzle game - safe to remove"
    "TikTok" = "Short-form video app - safe to remove"
    "Spotify" = "Music streaming app - safe to remove unless you use it"
    "Netflix" = "Video streaming app - safe to remove unless you use it"
    "Disney+" = "Video streaming app - safe to remove unless you use it"
    "Amazon" = "Shopping app - safe to remove"
    "Prime Video" = "Video streaming app - safe to remove unless you use it"
    "Facebook" = "Social media app - safe to remove"
    "Instagram" = "Social media app - safe to remove"
    "Twitter/X" = "Social media app - safe to remove"
    "LinkedIn" = "Professional networking app - safe to remove unless you use it"
    "Duolingo" = "Language learning app - safe to remove unless you use it"
    "Clipchamp" = "Video editor by Microsoft - safe to remove"
    "Microsoft Solitaire" = "Card games collection - safe to remove"
    "Bing News" = "News aggregator - safe to remove"
    "Bing Weather" = "Weather app - safe to remove"
    "Bing Finance" = "Finance news app - safe to remove"
    "Bing Sports" = "Sports news app - safe to remove"
    "Get Help" = "Microsoft support app - safe to remove"
    "Get Started (Tips)" = "Windows tips app - safe to remove"
    "Mixed Reality Portal" = "VR headset app - safe to remove unless you have a VR headset"
    "3D Viewer" = "3D model viewer - safe to remove"
    "Office Hub" = "Office promotion app - safe to remove"
    "People" = "Contacts app - safe to remove"
    "Skype" = "Video calling app - safe to remove unless you use it"
    "Groove Music" = "Legacy music player - safe to remove"
    "Movies & TV" = "Video player - safe to remove"
    "Feedback Hub" = "Windows feedback app - safe to remove"
    "Maps" = "Offline maps - safe to remove unless you need offline navigation"
    "Power Automate" = "Automation tool - safe to remove unless you use it"
    "Microsoft To Do" = "Task manager - safe to remove unless you use it"
    "Xbox App" = "Xbox gaming platform - may be needed for some PC games"
    "Xbox Game Bar" = "Gaming overlay (Win+G) - may be needed for game recording"
    "Phone Link (Your Phone)" = "Connect Android/iPhone to PC - keep if you use this feature"
    "Cortana" = "Voice assistant - safe to remove"
    "Copilot" = "AI assistant - keep if you use it, otherwise safe to remove"
    "OneDrive" = "Cloud storage - keep if you use Microsoft cloud storage"
    "Mail and Calendar" = "Email and calendar app - keep if you use it"
    "Outlook (New)" = "New Outlook app - keep if you use it for email"
    "OneNote" = "Note-taking app - keep if you use it"
}

foreach ($bloat in $installedBloatware) {
    $chk = New-Object System.Windows.Forms.CheckBox
    $typeLabel = if ($bloat.Type -eq "Lenovo") { " [Lenovo]" } else { "" }
    $chk.Text = "$($bloat.Name)$typeLabel"
    $chk.Tag = $bloat
    $chk.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $yPos)
    $chk.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
    $chk.Checked = $bloat.PreChecked

    # Add tooltip
    $tooltipText = if ($bloatwareTooltips.ContainsKey($bloat.Name)) { $bloatwareTooltips[$bloat.Name] } else { "Package: $($bloat.PackageName)" }
    $script:MainTooltip.SetToolTip($chk, $tooltipText)

    # Update selection count on check change
    $chk.Add_CheckedChanged({
        Update-SelectionCount -Label $script:BloatwareSelectionLabel -Checkboxes ($script:BloatwareCheckboxes + $script:Win32Checkboxes) -ItemType "app"
    })

    $BloatwarePanel.Controls.Add($chk)
    $script:BloatwareCheckboxes += $chk
    $yPos += $script:UI.ItemSpacing
}

# Win32 Programs section
$yPos += 10
$yPos = New-SectionHeader -Panel $BloatwarePanel -Text "Win32 Programs" -YPosition $yPos

$script:Win32Checkboxes = @()
$installedWin32 = Get-InstalledWin32Programs

if ($null -eq $installedWin32 -or @($installedWin32).Count -eq 0) {
    $LblNoWin32 = New-Object System.Windows.Forms.Label
    $LblNoWin32.Text = "(No Win32 bloatware detected)"
    $LblNoWin32.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $yPos)
    $LblNoWin32.Size = New-Object System.Drawing.Size(600, 20)
    $LblNoWin32.ForeColor = $script:UI.SubtleGray
    $BloatwarePanel.Controls.Add($LblNoWin32)
    $yPos += $script:UI.ItemSpacing
}
else {
    foreach ($win32 in $installedWin32) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = "$($win32.DisplayName)"
        $chk.Tag = $win32
        $chk.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $yPos)
        $chk.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
        $chk.Checked = $win32.PreChecked
        $chk.ForeColor = $script:UI.ErrorRed

        # Add tooltip with category info
        $script:MainTooltip.SetToolTip($chk, "Category: $($win32.Category)")

        # Update selection count on check change
        $chk.Add_CheckedChanged({
            Update-SelectionCount -Label $script:BloatwareSelectionLabel -Checkboxes ($script:BloatwareCheckboxes + $script:Win32Checkboxes) -ItemType "app"
        })

        $BloatwarePanel.Controls.Add($chk)
        $script:Win32Checkboxes += $chk
        $yPos += $script:UI.ItemSpacing
    }
}

$yPos += 10

# Selection count label
$script:BloatwareSelectionLabel = New-Object System.Windows.Forms.Label
$script:BloatwareSelectionLabel.Location = New-Object System.Drawing.Point(250, ($yPos + 14))
$script:BloatwareSelectionLabel.Size = New-Object System.Drawing.Size(150, 20)
$script:BloatwareSelectionLabel.ForeColor = $script:UI.SubtleGray
$BloatwarePanel.Controls.Add($script:BloatwareSelectionLabel)
Update-SelectionCount -Label $script:BloatwareSelectionLabel -Checkboxes ($script:BloatwareCheckboxes + $script:Win32Checkboxes) -ItemType "app"

# Select All / Deselect All buttons
$BtnSelectAllBloat = New-Object System.Windows.Forms.Button
$BtnSelectAllBloat.Text = "Select All"
$BtnSelectAllBloat.Location = New-Object System.Drawing.Point(20, ($yPos + 10))
$BtnSelectAllBloat.Size = New-Object System.Drawing.Size(100, $script:UI.ButtonHeight)
$BtnSelectAllBloat.Add_Click({
    foreach ($chk in $script:BloatwareCheckboxes) { $chk.Checked = $true }
    foreach ($chk in $script:Win32Checkboxes) { $chk.Checked = $true }
})
$BloatwarePanel.Controls.Add($BtnSelectAllBloat)

$BtnDeselectAllBloat = New-Object System.Windows.Forms.Button
$BtnDeselectAllBloat.Text = "Deselect All"
$BtnDeselectAllBloat.Location = New-Object System.Drawing.Point(130, ($yPos + 10))
$BtnDeselectAllBloat.Size = New-Object System.Drawing.Size(100, $script:UI.ButtonHeight)
$BtnDeselectAllBloat.Add_Click({
    foreach ($chk in $script:BloatwareCheckboxes) { $chk.Checked = $false }
    foreach ($chk in $script:Win32Checkboxes) { $chk.Checked = $false }
})
$BloatwarePanel.Controls.Add($BtnDeselectAllBloat)

# Remove Selected Bloatware button with keyboard shortcut
$BtnRemoveBloatware = New-Object System.Windows.Forms.Button
$BtnRemoveBloatware.Text = "&Remove Selected"
$BtnRemoveBloatware.Location = New-Object System.Drawing.Point(480, ($yPos + 10))
$BtnRemoveBloatware.Size = New-Object System.Drawing.Size(150, $script:UI.ButtonHeight)
$BtnRemoveBloatware.BackColor = $script:UI.AccentColor
$BtnRemoveBloatware.ForeColor = [System.Drawing.Color]::White
$BtnRemoveBloatware.FlatStyle = "Flat"
Add-ButtonHoverEffect -Button $BtnRemoveBloatware
$script:MainTooltip.SetToolTip($BtnRemoveBloatware, "Remove all selected bloatware apps (Alt+R)")

$BtnRemoveBloatware.Add_Click({
    $selectedCount = @($script:BloatwareCheckboxes | Where-Object { $_.Checked }).Count + @($script:Win32Checkboxes | Where-Object { $_.Checked }).Count

    if ($selectedCount -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No items selected for removal.", "Remove Bloatware", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $script:DryRun = $script:ChkDryRun.Checked
    if (-not (Show-ConfirmationDialog -Title "Confirm Removal" -SelectedCount $selectedCount -ActionType "remove" -IsDryRun $script:DryRun)) {
        return
    }

    Set-ButtonDisabled -Button $BtnRemoveBloatware -WorkingText "Removing..."
    $successCount = 0
    $failCount = 0
    $currentItem = 0

    if ($script:DryRun) {
        Write-Log "=== DRY RUN MODE - No changes will be made ===" "INFO"
    }

    try {
        # Remove AppX Bloatware
        foreach ($chk in $script:BloatwareCheckboxes) {
            if ($chk.Checked) {
                $currentItem++
                $bloat = $chk.Tag
                Update-Progress -Current $currentItem -Total $selectedCount -CurrentItem $bloat.Name
                if (Remove-BloatwareApp -PackageName $bloat.PackageName -DisplayName $bloat.Name) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        # Remove Win32 Bloatware
        foreach ($chk in $script:Win32Checkboxes) {
            if ($chk.Checked) {
                $currentItem++
                $win32 = $chk.Tag
                Update-Progress -Current $currentItem -Total $selectedCount -CurrentItem $win32.DisplayName
                if (Uninstall-Win32Program -DisplayName $win32.DisplayName -UninstallCommand $win32.UninstallCommand) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        Hide-Progress

        $dryRunMsg = if ($script:DryRun) { "[DRY RUN] " } else { "" }
        Update-Status "${dryRunMsg}Bloatware removal complete! Success: $successCount, Failed: $failCount"

        [System.Windows.Forms.MessageBox]::Show(
            "${dryRunMsg}Bloatware removal complete!`n`nSuccessful: $successCount`nFailed: $failCount`n`nLog: $script:LogPath",
            "Remove Bloatware",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error removing bloatware: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        Hide-Progress
        Set-ButtonEnabled -Button $BtnRemoveBloatware
        $script:DryRun = $false
    }
})
$BloatwarePanel.Controls.Add($BtnRemoveBloatware)

$TabBloatware.Controls.Add($BloatwarePanel)

# ============================================================================
# TAB 2: SETTINGS
# ============================================================================

$TabSettings = New-Object System.Windows.Forms.TabPage
$TabSettings.Text = "&Settings"
$TabSettings.Padding = New-Object System.Windows.Forms.Padding(10)

$SettingsPanel = New-Object System.Windows.Forms.Panel
$SettingsPanel.Dock = "Fill"
$SettingsPanel.AutoScroll = $true

# Introduction text
$settingsYPos = New-IntroText -Panel $SettingsPanel -Text "Configure Windows settings. Changes take effect immediately after applying. Explorer will restart automatically." -YPosition 8

$script:SettingsCheckboxes = @()

# Tooltip descriptions for settings
$settingsTooltips = @{
    "SearchIcon" = "Reduces the large search box to a compact icon, saving taskbar space"
    "MultiMonitor" = "Shows the taskbar on all monitors; apps appear on the display where their window is open"
    "CombineWhenFull" = "Shows separate buttons for each window until taskbar runs out of space"
    "HideRecommended" = "Hides the 'Recommended' section in the Start Menu that shows recent files (Win11 only)"
    "DisableBing" = "Prevents web searches when you type in the Start Menu - only shows local results"
    "PowerSettings" = "Optimizes power settings: display off after 10 min, never sleep when plugged in, sleep after 1 hour on battery"
    "Clipboard" = "Enables clipboard history - press Win+V to access previously copied items"
    "StorageSense" = "Automatically cleans temporary files daily: Downloads older than 14 days, Recycle Bin items older than 30 days"
    "ExplorerThisPC" = "Opens File Explorer to 'This PC' view showing drives instead of Quick Access/Home"
    "ShowExtensions" = "Shows file extensions like .txt, .exe, .pdf - helps identify file types and spot malware"
    "CleanQuickAccess" = "Removes all pinned folders from Quick Access except Desktop and Downloads"
    "StartPins" = "Resets pinned apps in Start Menu to Windows defaults (requires sign-out to take effect)"
}

# Taskbar section
$settingsYPos = New-SectionHeader -Panel $SettingsPanel -Text "Taskbar" -YPosition $settingsYPos

# Search Icon
$ChkSearchIcon = New-Object System.Windows.Forms.CheckBox
$ChkSearchIcon.Text = "Reduce taskbar search to icon only"
$ChkSearchIcon.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkSearchIcon.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkSearchIcon.Checked = $true
$ChkSearchIcon.Tag = "SearchIcon"
$script:MainTooltip.SetToolTip($ChkSearchIcon, $settingsTooltips["SearchIcon"])
$SettingsPanel.Controls.Add($ChkSearchIcon)
$script:SettingsCheckboxes += $ChkSearchIcon
$settingsYPos += $script:UI.ItemSpacing

# Multi-monitor taskbar
$ChkMultiMonitor = New-Object System.Windows.Forms.CheckBox
$ChkMultiMonitor.Text = "Show taskbar on all displays + show apps where window is open"
$ChkMultiMonitor.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkMultiMonitor.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkMultiMonitor.Checked = $true
$ChkMultiMonitor.Tag = "MultiMonitor"
$script:MainTooltip.SetToolTip($ChkMultiMonitor, $settingsTooltips["MultiMonitor"])
$SettingsPanel.Controls.Add($ChkMultiMonitor)
$script:SettingsCheckboxes += $ChkMultiMonitor
$settingsYPos += $script:UI.ItemSpacing

# Combine when full
$ChkCombineWhenFull = New-Object System.Windows.Forms.CheckBox
$ChkCombineWhenFull.Text = "Combine taskbar buttons only when full"
$ChkCombineWhenFull.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkCombineWhenFull.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkCombineWhenFull.Checked = $true
$ChkCombineWhenFull.Tag = "CombineWhenFull"
$script:MainTooltip.SetToolTip($ChkCombineWhenFull, $settingsTooltips["CombineWhenFull"])
$SettingsPanel.Controls.Add($ChkCombineWhenFull)
$script:SettingsCheckboxes += $ChkCombineWhenFull
$settingsYPos += $script:UI.SectionSpacing

# Start Menu section
$settingsYPos = New-SectionHeader -Panel $SettingsPanel -Text "Start Menu" -YPosition $settingsYPos

# Start Menu Recommended
$ChkHideRecommended = New-Object System.Windows.Forms.CheckBox
$ChkHideRecommended.Text = "Hide Start Menu Recommended section"
$ChkHideRecommended.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkHideRecommended.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkHideRecommended.Checked = $IsWindows11
$ChkHideRecommended.Tag = "HideRecommended"
$ChkHideRecommended.Enabled = $IsWindows11
$script:MainTooltip.SetToolTip($ChkHideRecommended, $settingsTooltips["HideRecommended"])
$SettingsPanel.Controls.Add($ChkHideRecommended)
$script:SettingsCheckboxes += $ChkHideRecommended
$settingsYPos += $script:UI.ItemSpacing

# Disable Bing Search
$ChkDisableBing = New-Object System.Windows.Forms.CheckBox
$ChkDisableBing.Text = "Disable Bing search in Start Menu"
$ChkDisableBing.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkDisableBing.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkDisableBing.Checked = $true
$ChkDisableBing.Tag = "DisableBing"
$script:MainTooltip.SetToolTip($ChkDisableBing, $settingsTooltips["DisableBing"])
$SettingsPanel.Controls.Add($ChkDisableBing)
$script:SettingsCheckboxes += $ChkDisableBing
$settingsYPos += $script:UI.SectionSpacing

# Power section
$settingsYPos = New-SectionHeader -Panel $SettingsPanel -Text "Power" -YPosition $settingsYPos

# Power Settings
$ChkPowerSettings = New-Object System.Windows.Forms.CheckBox
$ChkPowerSettings.Text = "Configure power: Display off 10min, Never sleep on AC, Sleep 1hr on battery"
$ChkPowerSettings.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkPowerSettings.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkPowerSettings.Checked = $true
$ChkPowerSettings.Tag = "PowerSettings"
$script:MainTooltip.SetToolTip($ChkPowerSettings, $settingsTooltips["PowerSettings"])
$SettingsPanel.Controls.Add($ChkPowerSettings)
$script:SettingsCheckboxes += $ChkPowerSettings
$settingsYPos += $script:UI.SectionSpacing

# System section
$settingsYPos = New-SectionHeader -Panel $SettingsPanel -Text "System" -YPosition $settingsYPos

# Clipboard History
$ChkClipboard = New-Object System.Windows.Forms.CheckBox
$ChkClipboard.Text = "Enable Windows Clipboard History (Win+V)"
$ChkClipboard.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkClipboard.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkClipboard.Checked = $true
$ChkClipboard.Tag = "Clipboard"
$script:MainTooltip.SetToolTip($ChkClipboard, $settingsTooltips["Clipboard"])
$SettingsPanel.Controls.Add($ChkClipboard)
$script:SettingsCheckboxes += $ChkClipboard
$settingsYPos += $script:UI.ItemSpacing

# Storage Sense
$ChkStorageSense = New-Object System.Windows.Forms.CheckBox
$ChkStorageSense.Text = "Enable Storage Sense (auto-cleanup temporary files)"
$ChkStorageSense.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkStorageSense.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkStorageSense.Checked = $true
$ChkStorageSense.Tag = "StorageSense"
$script:MainTooltip.SetToolTip($ChkStorageSense, $settingsTooltips["StorageSense"])
$SettingsPanel.Controls.Add($ChkStorageSense)
$script:SettingsCheckboxes += $ChkStorageSense
$settingsYPos += $script:UI.SectionSpacing

# Explorer section
$settingsYPos = New-SectionHeader -Panel $SettingsPanel -Text "File Explorer" -YPosition $settingsYPos

# Explorer opens to This PC
$ChkExplorerThisPC = New-Object System.Windows.Forms.CheckBox
$ChkExplorerThisPC.Text = "Set Explorer to open to 'This PC' instead of Quick Access/Home"
$ChkExplorerThisPC.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkExplorerThisPC.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkExplorerThisPC.Checked = $true
$ChkExplorerThisPC.Tag = "ExplorerThisPC"
$script:MainTooltip.SetToolTip($ChkExplorerThisPC, $settingsTooltips["ExplorerThisPC"])
$SettingsPanel.Controls.Add($ChkExplorerThisPC)
$script:SettingsCheckboxes += $ChkExplorerThisPC
$settingsYPos += $script:UI.ItemSpacing

# Show File Extensions
$ChkShowExtensions = New-Object System.Windows.Forms.CheckBox
$ChkShowExtensions.Text = "Show file extensions (e.g., .txt, .exe, .pdf)"
$ChkShowExtensions.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkShowExtensions.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkShowExtensions.Checked = $true
$ChkShowExtensions.Tag = "ShowExtensions"
$script:MainTooltip.SetToolTip($ChkShowExtensions, $settingsTooltips["ShowExtensions"])
$SettingsPanel.Controls.Add($ChkShowExtensions)
$script:SettingsCheckboxes += $ChkShowExtensions
$settingsYPos += $script:UI.ItemSpacing

# Clean Quick Access
$ChkCleanQuickAccess = New-Object System.Windows.Forms.CheckBox
$ChkCleanQuickAccess.Text = "Clean Quick Access pinned folders (keep only Desktop and Downloads)"
$ChkCleanQuickAccess.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkCleanQuickAccess.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkCleanQuickAccess.Checked = $true
$ChkCleanQuickAccess.Tag = "CleanQuickAccess"
$script:MainTooltip.SetToolTip($ChkCleanQuickAccess, $settingsTooltips["CleanQuickAccess"])
$SettingsPanel.Controls.Add($ChkCleanQuickAccess)
$script:SettingsCheckboxes += $ChkCleanQuickAccess
$settingsYPos += $script:UI.ItemSpacing

# Start Menu Pins (Win11 only)
$ChkStartPins = New-Object System.Windows.Forms.CheckBox
$ChkStartPins.Text = "Reset Start Menu pins to defaults (requires sign-out)"
$ChkStartPins.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $settingsYPos)
$ChkStartPins.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$ChkStartPins.Checked = $IsWindows11
$ChkStartPins.Tag = "StartPins"
$ChkStartPins.Enabled = $IsWindows11
$script:MainTooltip.SetToolTip($ChkStartPins, $settingsTooltips["StartPins"])
$SettingsPanel.Controls.Add($ChkStartPins)
$script:SettingsCheckboxes += $ChkStartPins
$settingsYPos += 40

# Apply Settings button with keyboard shortcut
$BtnApplySettings = New-Object System.Windows.Forms.Button
$BtnApplySettings.Text = "&Apply Settings"
$BtnApplySettings.Location = New-Object System.Drawing.Point(480, $settingsYPos)
$BtnApplySettings.Size = New-Object System.Drawing.Size(150, $script:UI.ButtonHeight)
$BtnApplySettings.BackColor = $script:UI.AccentColor
$BtnApplySettings.ForeColor = [System.Drawing.Color]::White
$BtnApplySettings.FlatStyle = "Flat"
Add-ButtonHoverEffect -Button $BtnApplySettings
$script:MainTooltip.SetToolTip($BtnApplySettings, "Apply all selected settings (Alt+A)")

$BtnApplySettings.Add_Click({
    $selectedCount = @($script:SettingsCheckboxes | Where-Object { $_.Checked -and $_.Enabled }).Count

    if ($selectedCount -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No settings selected to apply.", "Apply Settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    Set-ButtonDisabled -Button $BtnApplySettings -WorkingText "Applying..."
    $successCount = 0
    $failCount = 0
    $currentItem = 0

    $script:DryRun = $script:ChkDryRun.Checked
    if ($script:DryRun) {
        Write-Log "=== DRY RUN MODE - No changes will be made ===" "INFO"
    }

    try {
        foreach ($chk in $script:SettingsCheckboxes) {
            if ($chk.Checked -and $chk.Enabled) {
                $currentItem++
                Update-Progress -Current $currentItem -Total $selectedCount -CurrentItem $chk.Text
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
                    "CleanQuickAccess" { $result = Clear-ExplorerQuickAccess }
                    "StartPins" { $result = Set-StartMenuPins }
                }
                if ($result) { $successCount++ } else { $failCount++ }
            }
        }

        Hide-Progress

        # Restart Explorer to apply changes
        if (-not $script:DryRun -and $successCount -gt 0) {
            Restart-Explorer
        }

        $dryRunMsg = if ($script:DryRun) { "[DRY RUN] " } else { "" }
        Update-Status "${dryRunMsg}Settings applied! Success: $successCount, Failed: $failCount"

        [System.Windows.Forms.MessageBox]::Show(
            "${dryRunMsg}Settings applied!`n`nSuccessful: $successCount`nFailed: $failCount`n`nLog: $script:LogPath",
            "Apply Settings",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error applying settings: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        Hide-Progress
        Set-ButtonEnabled -Button $BtnApplySettings
        $script:DryRun = $false
    }
})
$SettingsPanel.Controls.Add($BtnApplySettings)

$TabSettings.Controls.Add($SettingsPanel)

# ============================================================================
# TAB 3: INSTALL APPS
# ============================================================================

$TabInstallApps = New-Object System.Windows.Forms.TabPage
$TabInstallApps.Text = "&Install Apps"
$TabInstallApps.Padding = New-Object System.Windows.Forms.Padding(10)

$InstallAppsPanel = New-Object System.Windows.Forms.Panel
$InstallAppsPanel.Dock = "Fill"
$InstallAppsPanel.AutoScroll = $true

# Introduction text
$appYPos = New-IntroText -Panel $InstallAppsPanel -Text "Select applications to install via Windows Package Manager (winget). Installations run silently in the background." -YPosition 8

$wingetAvailable = Test-WingetInstalled

# Winget status indicator
$LblWingetStatus = New-Object System.Windows.Forms.Label
$LblWingetStatus.Location = New-Object System.Drawing.Point(20, $appYPos)
$LblWingetStatus.Size = New-Object System.Drawing.Size(600, 20)
if ($wingetAvailable) {
    $LblWingetStatus.Text = "Winget is available"
    $LblWingetStatus.ForeColor = $script:UI.SuccessGreen
} else {
    $LblWingetStatus.Text = "Winget NOT found - please install App Installer from the Microsoft Store"
    $LblWingetStatus.ForeColor = $script:UI.ErrorRed
}
$InstallAppsPanel.Controls.Add($LblWingetStatus)
$appYPos += 25

# App tooltips
$appTooltips = @{
    "Google Chrome" = "Popular web browser by Google"
    "Brave Browser" = "Privacy-focused web browser with built-in ad blocking"
    "Adobe Acrobat Reader" = "Standard PDF viewer and annotator"
    "Google Drive" = "Cloud storage and file sync service"
    "Proton Pass" = "Secure password manager by Proton"
    "Todoist" = "Popular task management and to-do list app"
    "PhraseVault" = "Password and phrase management utility"
    "Microsoft PowerToys" = "Power user utilities including FancyZones, PowerRename, and more"
    "LocalSend" = "Cross-platform file sharing app (like AirDrop)"
    "Visual Studio Code" = "Lightweight code editor with extensions support"
    "7-Zip" = "Free file archiver with high compression ratio"
    "VLC Media Player" = "Universal media player supporting most formats"
}

$script:AppCheckboxes = @()
$currentCategory = ""

foreach ($app in $WingetApps) {
    if ($app.Category -ne $currentCategory) {
        $currentCategory = $app.Category
        $appYPos = New-SectionHeader -Panel $InstallAppsPanel -Text $currentCategory -YPosition $appYPos
    }

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "$($app.Name)"
    $chk.Tag = $app
    $chk.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $appYPos)
    $chk.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
    $chk.Checked = $false
    $chk.Enabled = $wingetAvailable

    # Add tooltip
    $tooltipText = if ($appTooltips.ContainsKey($app.Name)) { $appTooltips[$app.Name] } else { "Winget ID: $($app.WingetId)" }
    $script:MainTooltip.SetToolTip($chk, $tooltipText)

    # Update selection count on check change
    $chk.Add_CheckedChanged({
        Update-SelectionCount -Label $script:AppsSelectionLabel -Checkboxes $script:AppCheckboxes -ItemType "app"
    })

    $InstallAppsPanel.Controls.Add($chk)
    $script:AppCheckboxes += $chk
    $appYPos += $script:UI.ItemSpacing
}

$appYPos += 10

# Selection count label
$script:AppsSelectionLabel = New-Object System.Windows.Forms.Label
$script:AppsSelectionLabel.Location = New-Object System.Drawing.Point(250, ($appYPos + 14))
$script:AppsSelectionLabel.Size = New-Object System.Drawing.Size(150, 20)
$script:AppsSelectionLabel.ForeColor = $script:UI.SubtleGray
$InstallAppsPanel.Controls.Add($script:AppsSelectionLabel)
Update-SelectionCount -Label $script:AppsSelectionLabel -Checkboxes $script:AppCheckboxes -ItemType "app"

# Select All / Deselect All buttons
$BtnSelectAllApps = New-Object System.Windows.Forms.Button
$BtnSelectAllApps.Text = "Select All"
$BtnSelectAllApps.Location = New-Object System.Drawing.Point(20, ($appYPos + 10))
$BtnSelectAllApps.Size = New-Object System.Drawing.Size(100, $script:UI.ButtonHeight)
$BtnSelectAllApps.Enabled = $wingetAvailable
$BtnSelectAllApps.Add_Click({
    foreach ($chk in $script:AppCheckboxes) { $chk.Checked = $true }
})
$InstallAppsPanel.Controls.Add($BtnSelectAllApps)

$BtnDeselectAllApps = New-Object System.Windows.Forms.Button
$BtnDeselectAllApps.Text = "Deselect All"
$BtnDeselectAllApps.Location = New-Object System.Drawing.Point(130, ($appYPos + 10))
$BtnDeselectAllApps.Size = New-Object System.Drawing.Size(100, $script:UI.ButtonHeight)
$BtnDeselectAllApps.Enabled = $wingetAvailable
$BtnDeselectAllApps.Add_Click({
    foreach ($chk in $script:AppCheckboxes) { $chk.Checked = $false }
})
$InstallAppsPanel.Controls.Add($BtnDeselectAllApps)

# Install Selected Apps button with keyboard shortcut
$BtnInstallApps = New-Object System.Windows.Forms.Button
$BtnInstallApps.Text = "&Install Selected"
$BtnInstallApps.Location = New-Object System.Drawing.Point(480, ($appYPos + 10))
$BtnInstallApps.Size = New-Object System.Drawing.Size(150, $script:UI.ButtonHeight)
$BtnInstallApps.BackColor = $script:UI.AccentColor
$BtnInstallApps.ForeColor = [System.Drawing.Color]::White
$BtnInstallApps.FlatStyle = "Flat"
$BtnInstallApps.Enabled = $wingetAvailable
Add-ButtonHoverEffect -Button $BtnInstallApps
$script:MainTooltip.SetToolTip($BtnInstallApps, "Install all selected applications via winget (Alt+I)")

$BtnInstallApps.Add_Click({
    $selectedCount = @($script:AppCheckboxes | Where-Object { $_.Checked }).Count

    if ($selectedCount -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No apps selected for installation.", "Install Apps", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $script:DryRun = $script:ChkDryRun.Checked
    if (-not (Show-ConfirmationDialog -Title "Confirm Installation" -SelectedCount $selectedCount -ActionType "install" -IsDryRun $script:DryRun)) {
        return
    }

    Set-ButtonDisabled -Button $BtnInstallApps -WorkingText "Installing..."
    $successCount = 0
    $failCount = 0
    $currentItem = 0

    if ($script:DryRun) {
        Write-Log "=== DRY RUN MODE - No changes will be made ===" "INFO"
    }

    try {
        $appsToInstall = @($script:AppCheckboxes | Where-Object { $_.Checked })
        if ($appsToInstall.Count -gt 0 -and -not $script:DryRun) {
            Update-WingetSources
        }

        foreach ($chk in $script:AppCheckboxes) {
            if ($chk.Checked) {
                $currentItem++
                $app = $chk.Tag
                Update-Progress -Current $currentItem -Total $selectedCount -CurrentItem $app.Name
                if (Install-WingetApp -WingetId $app.WingetId -DisplayName $app.Name) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
        }

        Hide-Progress

        $dryRunMsg = if ($script:DryRun) { "[DRY RUN] " } else { "" }
        Update-Status "${dryRunMsg}App installation complete! Success: $successCount, Failed: $failCount"

        [System.Windows.Forms.MessageBox]::Show(
            "${dryRunMsg}App installation complete!`n`nSuccessful: $successCount`nFailed: $failCount`n`nLog: $script:LogPath",
            "Install Apps",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error installing apps: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        Hide-Progress
        Set-ButtonEnabled -Button $BtnInstallApps
        $script:DryRun = $false
    }
})
$InstallAppsPanel.Controls.Add($BtnInstallApps)

$TabInstallApps.Controls.Add($InstallAppsPanel)

# ============================================================================
# ADD TABS TO TAB CONTROL
# ============================================================================

$TabControl.Controls.AddRange(@($TabBloatware, $TabSettings, $TabInstallApps))
$MainForm.Controls.Add($TabControl)

# ============================================================================
# BOTTOM CONTROLS
# ============================================================================

# Dry Run panel with enhanced visibility
$DryRunPanel = New-Object System.Windows.Forms.Panel
$DryRunPanel.Location = New-Object System.Drawing.Point(10, 468)
$DryRunPanel.Size = New-Object System.Drawing.Size(280, 36)
$DryRunPanel.BackColor = $script:UI.WarningBg
$DryRunPanel.BorderStyle = "FixedSingle"

# Dry Run checkbox with keyboard shortcut (script-scope so tab buttons can access it)
$script:ChkDryRun = New-Object System.Windows.Forms.CheckBox
$script:ChkDryRun.Text = "Dr&y Run (preview only, no changes)"
$script:ChkDryRun.Location = New-Object System.Drawing.Point(8, 8)
$script:ChkDryRun.Size = New-Object System.Drawing.Size(260, 20)
$script:ChkDryRun.Checked = $false
$script:ChkDryRun.Font = New-Object System.Drawing.Font("Segoe UI Semibold", $script:UI.BodyFontSize)
$script:ChkDryRun.ForeColor = $script:UI.WarningOrange
$script:ChkDryRun.BackColor = $script:UI.WarningBg
$script:MainTooltip.SetToolTip($script:ChkDryRun, "When enabled, actions will be simulated without making actual changes (Alt+Y)")
$DryRunPanel.Controls.Add($script:ChkDryRun)
$MainForm.Controls.Add($DryRunPanel)

# Open Log button
$BtnOpenLog = New-Object System.Windows.Forms.Button
$BtnOpenLog.Text = "Open Log"
$BtnOpenLog.Location = New-Object System.Drawing.Point(300, 472)
$BtnOpenLog.Size = New-Object System.Drawing.Size(90, 28)
$script:MainTooltip.SetToolTip($BtnOpenLog, "Open the log file in Notepad")
$BtnOpenLog.Add_Click({
    if (Test-Path $script:LogPath) {
        Start-Process notepad.exe -ArgumentList $script:LogPath
    }
})
$MainForm.Controls.Add($BtnOpenLog)

# Progress label (hidden by default, shown above progress bar)
$script:ProgressLabel = New-Object System.Windows.Forms.Label
$script:ProgressLabel.Location = New-Object System.Drawing.Point(400, 470)
$script:ProgressLabel.Size = New-Object System.Drawing.Size(275, 16)
$script:ProgressLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:ProgressLabel.ForeColor = $script:UI.SubtleGray
$script:ProgressLabel.Visible = $false
$MainForm.Controls.Add($script:ProgressLabel)

# Progress bar (hidden by default)
$script:ProgressBar = New-Object System.Windows.Forms.ProgressBar
$script:ProgressBar.Location = New-Object System.Drawing.Point(400, 486)
$script:ProgressBar.Size = New-Object System.Drawing.Size(275, 18)
$script:ProgressBar.Style = "Continuous"
$script:ProgressBar.Visible = $false
$MainForm.Controls.Add($script:ProgressBar)

# Status label
$script:StatusLabel = New-Object System.Windows.Forms.Label
$script:StatusLabel.Text = "Ready | $OSName | Log: $script:LogPath"
$script:StatusLabel.Location = New-Object System.Drawing.Point(10, 550)
$script:StatusLabel.Size = New-Object System.Drawing.Size(665, 25)
$script:StatusLabel.BorderStyle = "FixedSingle"
$script:StatusLabel.Padding = New-Object System.Windows.Forms.Padding(5, 4, 0, 0)
$MainForm.Controls.Add($script:StatusLabel)

# Version and help info
$LblVersion = New-Object System.Windows.Forms.Label
$LblVersion.Text = "v2.0 | Keyboard shortcuts: Alt+R (Remove), Alt+A (Apply), Alt+I (Install), Alt+Y (Dry Run)"
$LblVersion.Location = New-Object System.Drawing.Point(10, 530)
$LblVersion.Size = New-Object System.Drawing.Size(665, 18)
$LblVersion.ForeColor = $script:UI.SubtleGray
$LblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$MainForm.Controls.Add($LblVersion)

# ============================================================================
# SHOW FORM
# ============================================================================

try {
    Write-Log "Application started"
    [void]$MainForm.ShowDialog()
}
catch {
    $errorMsg = "FATAL ERROR: $($_.Exception.Message)`n`nStack trace:`n$($_.ScriptStackTrace)"
    Write-Log $errorMsg "ERROR"
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Windows PC Setup - Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
finally {
    Stop-Transcript
}
