# Windows PowerSetup

Public MIT WinForms PowerShell GUI for IT admins setting up Windows 10/11 PCs. Company: SPQRK Web Solutions. Version 2.6.

Cursor and other harnesses: read `AGENTS.md` first. Traps: `docs/known.md`.

## Architecture

- Single file: `Windows-PC-Setup.ps1` (~2800 lines). Do not split unless asked.
- `System.Windows.Forms`. Auto-elevates if not admin. Logs to `$env:TEMP\PCSetup_*.log`.
- UI constants: `$script:UI` hashtable at top of script.
- Tabs: Bloatware · Settings · Install Apps (winget) · Brave (HKCU policies/shortcuts) · Harden · Repair (DISM / SFC / CHKDSK).

## Verify

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

Do not run the GUI script on this host unless the user explicitly asks.

## Shortcuts

Alt+R remove · Alt+A apply settings · Alt+H harden · Alt+U unlock · Alt+I install · Alt+B Brave settings · Alt+P profile shortcuts · Alt+Y dry-run toggle.
