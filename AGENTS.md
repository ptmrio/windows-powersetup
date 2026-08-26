# AGENTS.md

Public MIT PowerShell GUI (`Windows-PC-Setup.ps1`, v2.6) for IT admins setting up Windows 10/11 PCs. Single-file WinForms. Do not split the script unless asked.

This machine is a daily driver. The script auto-elevates and mutates the OS. Edit and parse it; do not apply it.

## Gates

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

That parse is the merge bar (`exit 1` on parse errors). Harden helpers: `test-harden.ps1`. Brave helpers: `test-brave.ps1`. Do not call work done on red syntax. Do not run `Windows-PC-Setup.ps1` against this host without an explicit OK.

## Hard rules

- Dry-run is the default mental model. Every mutating function already branches on `$script:DryRun`; keep that.
- Never auto-elevate, never `Start-Process -Verb RunAs`, never winget/uninstall/repair on this PC unless asked.
- Never `git push` or force-git unless asked. No Co-Authored-By trailers.
- Registry: HKCU preference + HKCU policy fallback. No HKLM Start policy. No `IsEducationEnvironment`. Harden may call Defender cmdlets and write only: `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` `EnableSmartScreen` + `ShellSmartScreenLevel=Warn`; `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System` `InactivityTimeoutSecs`; `HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen` `ConfigureAppInstallControlEnabled` + `ConfigureAppInstallControl` (Unlock may delete those two Store-only values only); `HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy` `VerifiedAndReputablePolicyState` (SAC On/Off, Win11). Never write Explorer `AicEnabled`. No Chrome policies. Brave Apply may write only `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave` (`RestoreOnStartup`, `DefaultSearchProviderEnabled`, `DefaultSearchProviderName`, `DefaultSearchProviderKeyword`, `DefaultSearchProviderSearchURL`, `DefaultSearchProviderSuggestURL`, `PasswordManagerEnabled`, `AutofillAddressEnabled`, `AutofillCreditCardEnabled`, `PromptForDownloadLocation`, `DefaultDownloadDirectory`) and numeric REG_SZ values under `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave\ExtensionInstallForcelist`. HKCU Brave config does not follow a later standard user.
- Local Users/Administrators: resolve by SID (`S-1-5-32-545` / `S-1-5-32-544`), never English names.
- DISM before SFC (see `CTO-SAFETY-REPORT.md`). Repair tab already exists; do not reorder from WinUtil.
- Traps and exact commands live in `docs/known.md`. Append what cost time; delete what stops being true. No customer names, hostnames, or incident write-ups. Do not commit `.cursor/`, `.claude/settings.json`, `research-start-menu-2026.md`, or `docs/incidents/`.

## Tabs in the GUI

1. Remove Bloatware (AppX + Win32)
2. Settings (taskbar, Start, power, Explorer)
3. Install Apps (winget)
4. Brave (HKCU policies with per-setting checkboxes + profile shortcuts; Apply when brave.exe is on disk)
5. Harden (optional standard user, Defender/SmartScreen/ASR, 10-min lock, Store-only + SAC with Unlock)
6. Repair (DISM / SFC / CHKDSK)

UI constants: `$script:UI` at the top of the script. Logs: `$env:TEMP\PCSetup_*.log`.
