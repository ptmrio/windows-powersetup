# Project: Windows PowerSetup

## Overview
PowerShell GUI utility for IT admins to set up Windows PCs. Single-file script (~1970 lines).

## Key Facts
- **Company**: SPQRK Web Solutions
- **License**: MIT
- **Supports**: Windows 10 and 11
- **Requires**: PowerShell 5.1+, Admin privileges

## Architecture
- Single file: `Windows-PC-Setup.ps1`
- Uses Windows Forms GUI (`System.Windows.Forms`)
- Auto-elevates to admin if needed
- Logs to `$env:TEMP\PCSetup_*.log`

## UI Constants
Located at top of script in `$script:UI` hashtable - colors, spacing, sizes.

## Syntax Validation
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

## Git Push (WSL)
```bash
git config --global credential.helper manager-core
git push
```

## Tabs
1. **Remove Bloatware** - AppX and Win32 apps
2. **Settings** - Taskbar, Start Menu, Power, Explorer
3. **Install Apps** - Via winget

## Keyboard Shortcuts
- Alt+R: Remove Selected
- Alt+A: Apply Settings
- Alt+I: Install Selected
- Alt+Y: Dry Run toggle
