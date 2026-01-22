# Windows PowerSetup

A PowerShell GUI utility for IT admins to quickly set up and configure new Windows PCs.

**By [SPQRK Web Solutions](https://spqrk.net)**

## Why This Exists

I set up new PCs regularly—whether it's a fresh install, a client machine, or fixing a broken system. Every time, it's the same routine: remove bloatware, tweak taskbar settings, run `sfc /scannow`, install the usual apps...

This is my **opinionated automation** of that entire process. Instead of clicking through 47 settings dialogs and typing the same PowerShell commands, I click a few checkboxes and walk away.

If you also find yourself:
- Setting up Windows PCs more than once a year
- Running SFC/DISM to fix corrupted system files
- Uninstalling Candy Crush for the hundredth time
- Wishing Windows had sane defaults

...this tool is for you.

## Features

- **Remove Bloatware** - Uninstall pre-installed apps (Candy Crush, TikTok, McAfee, etc.)
- **Configure Settings** - Taskbar, Start Menu, Power, Explorer preferences
- **Install Apps** - Via winget (Chrome, Brave, VS Code, 7-Zip, VLC, etc.)
- **Dry Run Mode** - Preview all changes before applying
- **Tooltips** - Hover over any item for detailed explanations
- **Keyboard Shortcuts** - Alt+R (Remove), Alt+A (Apply), Alt+I (Install), Alt+Y (Dry Run)

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1+
- Administrator privileges

## How to Run

### Option 1: Download and Run
```powershell
# Download
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ptmrio/windows-powersetup/master/Windows-PC-Setup.ps1" -OutFile "$env:TEMP\Windows-PC-Setup.ps1"

# Run
powershell -ExecutionPolicy Bypass -File "$env:TEMP\Windows-PC-Setup.ps1"
```

### Option 2: Right-Click
1. Download `Windows-PC-Setup.ps1`
2. Right-click → "Run with PowerShell"
3. Accept the admin elevation prompt

### Option 3: PowerShell (local file)
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Windows-PC-Setup.ps1
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

## Support This Project ☕

If this saved you time, consider buying me a beer!

[![GitHub Sponsors](https://img.shields.io/badge/Sponsor-GitHub-ea4aaa?logo=github)](https://github.com/sponsors/ptmrio)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-FF5E5B?logo=ko-fi)](https://ko-fi.com/spqrk)
[![PayPal](https://img.shields.io/badge/PayPal-Donate-00457C?logo=paypal)](https://paypal.me/realSPQRK)

## Disclaimer

**USE AT YOUR OWN RISK.** This software is provided "AS IS", without warranty of any kind, express or implied. While this utility has been tested, it modifies Windows system settings and removes applications, which may have unintended effects on your system.

Before running:
- **Review the source code** to understand what changes will be made
- **Create a system restore point** or backup
- **Use Dry Run mode** first to preview all actions

In no event shall the authors or copyright holders be liable for any claim, damages, or other liability arising from the use of this software. You are solely responsible for determining the appropriateness of using this utility and assume all risks associated with its use.

## License

MIT

---

*Windows is a registered trademark of Microsoft Corporation. This utility is not affiliated with or endorsed by Microsoft.*
