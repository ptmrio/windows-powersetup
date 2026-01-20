# Windows PowerSetup

A PowerShell GUI utility for IT admins to quickly set up and configure new Windows PCs.

## Features

- **Remove Bloatware** - Uninstall pre-installed apps (Candy Crush, TikTok, McAfee, etc.)
- **Configure Settings** - Taskbar, Start Menu, Power, Explorer preferences
- **Install Apps** - Via winget (Chrome, Brave, VS Code, 7-Zip, VLC, etc.)
- **Dry Run Mode** - Preview all changes before applying

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1+
- Administrator privileges

## How to Run

### Option 1: Right-click
1. Right-click `Windows-PC-Setup.ps1`
2. Select "Run with PowerShell"
3. Accept the admin elevation prompt

### Option 2: PowerShell
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Windows-PC-Setup.ps1
```

### Option 3: One-liner (download and run)
```powershell
irm https://raw.githubusercontent.com/ptmrio/windows-powersetup/main/Windows-PC-Setup.ps1 | iex
```

## Settings Applied

| Setting | Value |
|---------|-------|
| Taskbar search | Icon only |
| Taskbar buttons | Combine when full |
| Multi-monitor taskbar | Show on all displays |
| Explorer opens to | This PC |
| File extensions | Visible |
| Storage Sense | Downloads 14 days, Recycle Bin 30 days |
| Power (AC) | Display off 10min, never sleep |
| Power (Battery) | Sleep after 1 hour |

## License

MIT
