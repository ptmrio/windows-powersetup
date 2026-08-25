# Harden tab — design

Date: 2026-08-25  
Status: implemented (v2.3 Harden tab)  
Product: Windows PowerSetup (`Windows-PC-Setup.ps1`, single-file WinForms)

## Problem

Fake “PDF converter/reader” setups (`PDFConvertSetup.exe`, `PDFReaderSetup.exe`) are bundlers. They install junk browsers plus an updater that writes **Chrome machine policies**. Yahoo is locked, Google disappears from the list, new tabs hop through throwaway domains. It is per-PC. Shared drives are usually clean. Recovery is a **wipe**, not a Chrome reset. Workspace “Managed by …” can still be genuine; local Chrome policy still wins unless cloud is set to override.

Incident notes: `docs/incidents/2026-pdf-bundler-chrome-hijack.md`, trap in `docs/known.md`.

Audience: one-person IT, mostly Austrian shops with 1–5 PCs, local accounts (not Entra). Another technician must be able to understand and undo the box.

## Goal

On a **clean install**, after the technician has already created a local admin, patched, and installed drivers, PowerSetup offers a Harden tab that:

1. Creates a **standard local user** (name and password typed in; no canned names).
2. Applies a **conservative Microsoft Defender / SmartScreen / ASR baseline** (machine-wide).

Then the technician signs in as that user to set up the desktop, confirming UAC with the **technician-admin password the employee must not know**. Close the elevated PowerShell window this script leaves behind (`-NoExit` relaunch) before handing the PC over.

**Guarantee (honest):** a genuine standard user cannot write **HKLM** Chrome policy or install machine-wide software without that admin password. They **can** still run a downloaded EXE, write **HKCU** Chrome policy, and install per-user junk. The six ASR rules do not block “PDF converter from a search ad.” SmartScreen is **Warn** (click-through). Wipe remains the recovery.

## Non-goals

- Per-machine PUP/bloatware hunting (OneBrowser, Sheaf, etc.). Out of scope; hugely case-by-case. Existing Bloatware tab unchanged by this work.
- Installing a PDF reader (already Install Apps).
- Writing Chrome policies (HKLM or HKCU) from this script.
- Treating Google Admin “cloud over local” as automatic: platform machine policy still wins by default; `CloudPolicyOverridesPlatformPolicy` is machine-scope and must be verified in `chrome://policy`.
- Smart App Control, BitLocker, Controlled Folder Access, Office/Acrobat “block child processes” ASR, WDAC/AppLocker.
- Demoting the current admin, auto-logon, Microsoft/Entra accounts, default usernames (`admin`, `praxis`, …).
- Applying HKCU “user environment” tweaks while running as admin (they would stick to the wrong account).

## Operator workflow

1. Clean Windows install. Create **one local admin** (name chosen by the technician). Updates, drivers.
2. Run PowerSetup elevated (existing auto-elevate).
3. Use other tabs as today (Settings, Install Apps while still admin).
4. Harden: create daily user + apply baseline → Apply (confirm dialog). Dry Run available.
5. Sign out. **Close the elevated PowerShell window** this tool leaves (`-NoExit`). Sign in as the new standard user. Set up their environment. UAC asks for the technician-admin password (employee must not know it).

## UI

Fifth tab, after Install Apps, before Repair. Title: `Harden`.

Intro text (substance, not exact copy): install apps on this admin account first; then create the daily user; then switch. Harden is machine-wide. PDF reader is on Install Apps. Chrome Workspace cloud-over-local is done in Google Admin, not here.

**Block 1 — Create standard user**

- Username, password, confirm password. No placeholders that look like names.
- Checkbox: “User must change password at next logon” — default **off** (technician must sign in to set up the desktop).
- Checkbox: “Create this user” — default **on**. If off, Apply only does the baseline.

**Block 2 — Microsoft baseline** (each item a checkbox, all default **on**)

- Defender real-time + cloud protection (High)
- Potentially unwanted app blocking (PUA)
- SmartScreen (Warn)
- Network Protection
- Attack surface reduction (the six rules below)

One **Apply Harden** button (same confirm pattern as Remove/Install, not Settings). Dry Run uses the existing bottom checkbox.

## Create user — behavior

- Local account only. Members of **Users** (`S-1-5-32-545`), never Administrators (`S-1-5-32-544`). `Add-LocalGroupMember -SID` (do not resolve English/`Benutzer` names). Verify membership by **account SID**, not `*\name`.
- Never add the new account to Administrators. Never remove the current user from Administrators. Never delete accounts.
- Refuse: empty username, password mismatch, empty password, password shorter than 8 characters (workgroup policy often allows `1234`), invalid name (Windows illegal characters / too long), well-known RID accounts (`-500` Administrator, `-501` Guest/`Gast`, `-503`, `-504`) even if renamed.
- If the local user **already exists**: success only if enabled, not a well-known RID, in Users, **not** in Administrators. Do not change password or groups. If those postconditions fail (typed the admin, `Gast`, orphaned account), **fail closed and do not apply baseline**. Partial create (user exists but Users add / logonpasswordchg failed) is also fail-closed on retry.
- `New-LocalUser` on Windows PowerShell 5.1 has **no** `-UserMayChangePassword` and **no** `-ChangePasswordAtLogon`. Default already allows password change. Optional must-change-at-logon: `net user <name> /logonpasswordchg:yes`.
- Do not auto-logon. Do not log the password. Clear password boxes on **every** handler exit (validation fail, cancel, apply).
- Dry-run: log “Would create local user ‘X’ in Users (SID S-1-5-32-545)” without calling `New-LocalUser`.
- 64-bit PowerShell only (`LocalAccounts` is missing in 32-bit on 64-bit Windows).

## Microsoft baseline — behavior

All via Defender cmdlets where possible (`Set-MpPreference`, `Add-MpPreference`). Machine-wide. Dry-run logs the intended cmdlets/values only. Use enum **names** (`High`, `Enabled`, `Advanced`) not magic numbers. Mutators: `-ErrorAction Stop`. After each apply, **read back** `Get-MpPreference` / `Get-MpComputerStatus`. If Tamper Protection or passive mode blocked the change, log FAIL (not SUCCESS). Decide Cloud Block Level High vs “skipped, tamper-protected” in the log; do not claim High if it did not stick.

If Defender / `Set-MpPreference` is missing (stripped image): that block fails; user creation can still succeed. Do not throw the whole GUI.

ASR: merge the six GUIDs to Block **without** wiping other rules. Do **not** wrap `$null` properties in `@()` then test `$null -eq $ids` (that is a one-element array). Prefer `Add-MpPreference` one GUID at a time with action Enabled, **or** a dictionary merge into clean paired arrays for `Set-MpPreference`. Verify a pre-existing seventh rule still exists after apply (VM).

| Setting | Value | Notes |
|---|---|---|
| Real-time monitoring | on | `DisableRealtimeMonitoring $false` |
| Cloud-delivered protection | High | `MAPSReporting` Advanced; `CloudBlockLevel` High (2). Not HighPlus / ZeroTolerance |
| Sample submission | leave default unless already disabled; do not force off | |
| PUA | Enabled (1) | Not Audit |
| Network Protection | Enabled (1) | |
| SmartScreen | on, **Warn** | Explorer SmartScreen + `HKLM\SOFTWARE\Policies\Microsoft\Windows\System` `EnableSmartScreen=1`, `ShellSmartScreenLevel=Warn`. Not Block. Documented HKLM **exception**: Windows SmartScreen policy only, not Start/Chrome |
| ASR | Block (`Enabled`) on the six GUIDs below | Merge; do not replace the entire ASR list with only these six |

**ASR GUIDs (Block). Do not add others in v1.**

| Rule | GUID |
|---|---|
| Block credential stealing from LSASS | `9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2` |
| Block persistence through WMI event subscription | `e6db77e5-3df2-4cf1-b95a-636979351e5b` |
| Block abuse of exploited vulnerable signed drivers | `56a863a9-875e-4185-98a7-b882c64b5ce5` |
| Block executable content from email client and webmail | `be9ba2d9-53ea-4cdc-84e5-9b1eeee46550` |
| Block JavaScript or VBScript from launching downloaded executable content | `d3e037e1-3eb8-44c8-a917-57927947596d` |
| Block untrusted and unsigned processes that run from USB | `b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4` |

Undo for another technician: **not** Windows Security toggles (HKLM SmartScreen greys them out as “managed”). Remove `EnableSmartScreen` and `ShellSmartScreenLevel` under `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` (restore prior values if captured). ASR: `Set-MpPreference` / `Add-MpPreference` Audit or Disabled. Local user: Settings → Accounts or `Remove-LocalUser` (not offered in the GUI).

SmartScreen policy is documented for Pro/Enterprise/Education; **test Windows 11 Home** on a VM before treating Warn as guaranteed. It is Explorer/app-reputation SmartScreen, not Chrome/Edge.

Chrome Workspace: platform machine policy still beats cloud by default. `CloudPolicyOverridesPlatformPolicy` is a **machine** precedence setting and is not delivered as a user cloud policy. Ordinary Workspace user policy is not a reliable override of local hijack keys. Verify with `chrome://policy`. Out of script scope.

## Error handling

- Confirm dialog lists: username to create (or “skip user”), and which baseline boxes are on. No password in the dialog. Same Yes/No MessageBox as Remove/Install; Harden may pass a `-Detail` body.
- Per-item success/fail counts like other tabs. If user create is requested and fails (including existing-admin / Gast / membership postcondition), **stop before baseline**. If user create is unchecked, run baseline only.
- `try/catch` plus `$LASTEXITCODE` where native. Keep `$script:DryRun` branches on every mutating function.

## Product-rule delta

`AGENTS.md` today: no HKLM policy (meant Start vs Intune). **Amend `AGENTS.md` before writing SmartScreen.** Harden may:

- call Defender preference APIs (they persist in Defender’s store / HKLM under the hood);
- write **only** the SmartScreen values named above.

Still forbidden: HKLM Start policy, `IsEducationEnvironment`, Chrome `Policies\Google\Chrome`, HKLM as a dump for unrelated keys.

## Testing (this repo)

- `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1` is the merge bar. That file must `exit 1` on parse errors (today it always exits 0).
- `test-harden.ps1` must test production function bodies (extract from `Windows-PC-Setup.ps1` AST or a shared snippet), not a drifting copy. Also assert `New-LocalUser` 5.1 metadata has no `-UserMayChangePassword`.
- Do **not** run `Windows-PC-Setup.ps1` on the daily-driver checkout host unless the operator explicitly OK’s it.
- Behavioral proof is a disposable clean VM: German UI language (SID groups), create user, confirm not in Administrators, Defender ASR list, SmartScreen Warn, Dry Run creates nobody.

## Day-to-day (accepted)

Technician on the standard user: UAC + admin password for installs/drivers/winget. Employee cannot complete a **machine-wide** install or HKLM Chrome lock without that password. They can still click through SmartScreen Warn and write HKCU. Conservative ASR/PUA should be quiet; the USB ASR rule also applies after a file is copied off USB. No Controlled Folder Access / SAC / Office child-process rules.
