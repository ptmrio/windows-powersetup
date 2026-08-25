# AGENTS.md

Public MIT PowerShell GUI (`Windows-PC-Setup.ps1`, v2.3) for IT admins setting up Windows 10/11 PCs. Single-file WinForms. Do not split the script unless asked.

This machine is a daily driver. The script auto-elevates and mutates the OS. Edit and parse it; do not apply it.

## Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

That parse is the merge bar (`exit 1` on parse errors). Harden helpers: `test-harden.ps1`. Do not call work done on red syntax. Do not run `Windows-PC-Setup.ps1` against this host without an explicit OK.

## Hard rules

- Dry-run is the default mental model. Every mutating function already branches on `$script:DryRun`; keep that.
- Never auto-elevate, never `Start-Process -Verb RunAs`, never winget/uninstall/repair on this PC unless asked.
- Never `git push` or force-git unless asked. No Co-Authored-By trailers.
- Registry: HKCU preference + HKCU policy fallback. No HKLM Start policy. No `IsEducationEnvironment`. Harden may call Defender cmdlets and write only `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` `EnableSmartScreen` + `ShellSmartScreenLevel=Warn`. No Chrome policies.
- Local Users/Administrators: resolve by SID (`S-1-5-32-545` / `S-1-5-32-544`), never English names.
- DISM before SFC (see `CTO-SAFETY-REPORT.md`). Repair tab already exists; do not reorder from WinUtil.
- Traps and exact commands live in `docs/known.md`. Append what cost time; delete what stops being true.

## Tabs in the GUI

1. Remove Bloatware (AppX + Win32)
2. Settings (taskbar, Start, power, Explorer)
3. Install Apps (winget)
4. Harden (standard local user + Defender/SmartScreen/ASR)
5. Repair (DISM / SFC / CHKDSK)

UI constants: `$script:UI` at the top of the script. Logs: `$env:TEMP\PCSetup_*.log`.
