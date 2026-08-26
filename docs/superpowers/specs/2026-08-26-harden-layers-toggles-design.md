# Harden layers and toggles - design (v2.4 delta)

Date: 2026-08-26
Status: implemented (v2.4 Harden layers and toggles). Revised after Codex review (SAC Off is not one-way on 25H2; Store-only scope narrowed; AicEnabled hive corrected).
Product: Windows PowerSetup (`Windows-PC-Setup.ps1`, single-file WinForms)
Amends: `docs/superpowers/specs/2026-08-25-harden-tab-design.md` (v2.3, implemented). That file stays as written; where the two disagree, this one wins for v2.4.

ASCII only, on purpose. See the encoding trap in `docs/known.md`.

## Problem this delta addresses

v2.3 shipped two things: create a standard local user, apply a Defender/SmartScreen/ASR baseline. In the field the standard-user half is usually skipped. The Austrian 1-5 PC shops this tool targets want one account, and that account is the admin the technician already made. So the part of the box that actually stopped the fake-PDF bundler class is the part that gets turned off.

v2.4 does not fight that. It keeps standard-user creation, flips its checkbox default to off, and adds friction that works on a single-admin box:

- an idle lock the employee cannot shrug off,
- installs restricted to the Microsoft Store by machine policy,
- Smart App Control On (Windows 11) so unsigned, unreputable EXEs do not start at all.

Operator decisions behind this delta came from a Codex review plus a read-only probe of one 25H2 Pro daily driver (SAC Off, no `AicEnabled` value, no `InactivityTimeoutSecs`, SecHealthUI 1000.29628). No machine was mutated for this spec.

## Guarantee vs friction (honest)

Least privilege is the only real guarantee here, and it is the one thing this delta does not add. If the technician skips user creation, the daily account is a local administrator: it can write HKLM Chrome policy, install machine-wide software, remove every value this tab writes, and turn Smart App Control off in Windows Security - all behind a UAC prompt that account can approve itself. Everything in the new "blocking pair" is friction on that path, not a wall across it. Smart App Control is the strongest of the three because it blocks execution of unsigned/unreputable binaries before any consent prompt, but it is still admin-reversible in two clicks and a reboot. Store-only and the inactivity lock are HKLM policy values an admin can delete with `reg delete`. Against a genuine standard user, the same controls are close to absolute; against the admin who clicks the fake PDF converter, they buy a pause and an obvious "this is managed" signal. Wipe is still the recovery. Say this in the intro copy, not only in this file.

## Non-goals

Carried from v2.3 and still in force: no Chrome policies (HKLM or HKCU), no BitLocker, no Controlled Folder Access, no Office/Acrobat block-child-process ASR, no WDAC/AppLocker, no HKLM Start policy, no `IsEducationEnvironment`, no per-machine PUP hunting, no demoting the current admin, no auto-logon, no default usernames.

New for v2.4:

- **Unlock never disables Defender.** Unlock touches exactly two things: Smart App Control -> Off, and removal of the two Store-only policy values. It does not change real-time protection, cloud level, PUA, Network Protection, ASR, SmartScreen, or the inactivity limit. There is no "undo Harden" button and there will not be one.
- **Smart App Control is not Explorer `AicEnabled`.** Do not read or write `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer` `AicEnabled` (REG_SZ `Anywhere` / `PreferStore` / `StoreOnly` when present). That is the machine *preference* behind Settings > Choose where to get apps, not SAC and not the policy this tab uses. It was absent on the probed 25H2 box. Store-only here is set by machine **policy** only.
- **Auto-lock is not the Settings-tab power option.** The Settings tab's powercfg "display off after 10 minutes" is unrelated and stays where it is. This is the secpol machine inactivity limit, which locks the session.
- No signing of `Windows-PC-Setup.ps1`. PhraseVault's Azure Trusted Signing certificate does not cover this repo. Out of scope; consequences are documented under "Ordering trap" below.
- No new ASR rules. The six from v2.3 remain the whole list.

## What stays exactly as v2.3

- Standard-user creation logic, SID checks (`S-1-5-32-545` / `S-1-5-32-544`), fail-closed on an existing account that is disabled, well-known RID, or in Administrators, `net user /logonpasswordchg:yes` for must-change-at-logon, 64-bit-only `LocalAccounts`, password boxes cleared on every handler exit.
- The five baseline items and their read-back-or-FAIL behaviour.
- Apply-only semantics: **checked = set, unchecked = skip, never disable**. No checkbox in this tab turns protection off. The only subtraction in the tab is the Unlock button, and only for its two named items.
- Auto-elevate (`Windows-PC-Setup.ps1:31-37`), `-NoExit` relaunch, `$env:TEMP\PCSetup_*.log`, existing confirm-dialog pattern, existing Dry Run checkbox.

## UI (v2.4 Harden tab)

Four sections, in this order, in the existing `$HardenPanel`.

Intro text (substance, not exact copy): install apps on this admin account first, because the Store-only policy and Smart App Control both block ordinary installers afterwards. Creating a daily standard user is optional and off by default - **if you skip it, the account in daily use stays an administrator and can still write HKLM, including Chrome policy; the settings below are friction, not least privilege**. Harden is machine-wide. PDF reader is on Install Apps. Chrome Workspace cloud-over-local is set in Google Admin, not here.

**Section 1 - Standard user.** Unchanged controls (`Create this user`, username, password, confirm, must-change-at-logon). One change: `$script:ChkHardenCreateUser.Checked = $false`. Tooltip gains "Off is the common path; the daily account then stays an administrator." When unchecked, the three text boxes and the must-change checkbox are disabled (greyed) so the tab reads as a deliberate skip, not a half-filled form.

**Section 2 - Microsoft baseline.** The five existing checkboxes, all default on, plus a sixth in the same group:

- `Auto-lock the screen after 10 minutes idle` - default **on**.

It belongs here, not with the blocking pair, because it is apply-only like its five siblings: checked writes the value, unchecked skips, and **Unlock does not clear it**.

**Section 3 - Block unknown installers.** Two checkboxes, both default **on** where supported:

- `Allow apps from the Microsoft Store only` - machine policy.
- `Smart App Control (On)` - Windows 11 only. On Windows 10, `Enabled = $false` and `Checked = $false`, tooltip "Windows 11 only." Same pattern as the Start-menu checkboxes at `Windows-PC-Setup.ps1:2346-2374`, driven by the existing `$IsWindows11`.

Section header copy should carry one line: "Turn these off with Unlock for maintenance before installing anything with winget."

**Section 4 - Buttons.**

- `Apply &Harden` - unchanged position and accelerator (Alt+H).
- `&Unlock for maintenance` - new, to the left of Apply Harden. Alt+U is free (`&R`, `&A`, `&I`, `&H`, `&R`un, `Dr&y` are the current accelerators). Not accent-coloured; use the plain button style so it does not read as the primary action.

Dry Run is the existing bottom checkbox and applies to both buttons.

## Per-control reference

| Control | Mechanism | Apply (checked) | Unchecked | Unlock button | Win10 | Win11 | Who can flip it afterwards |
|---|---|---|---|---|---|---|---|
| Create standard user | `New-LocalUser` + `Add-LocalGroupMember -SID S-1-5-32-545` | create/verify | skip entirely | untouched | yes | yes | admin (Settings > Accounts, `Remove-LocalUser`) |
| Defender real-time + cloud High | `Set-MpPreference -DisableRealtimeMonitoring $false -MAPSReporting Advanced -CloudBlockLevel High` | set + read back | skip | untouched | yes | yes | admin; Tamper Protection may block even admin |
| PUA | `Set-MpPreference -PUAProtection Enabled` | set + read back | skip | untouched | yes | yes | admin |
| SmartScreen Warn | HKLM `SOFTWARE\Policies\Microsoft\Windows\System`: `EnableSmartScreen`=1 (DWORD), `ShellSmartScreenLevel`=`Warn` (String) | set + read back | skip | untouched | yes | yes | admin via regedit; Windows Security toggle stays greyed while the policy exists |
| Network Protection | `Set-MpPreference -EnableNetworkProtection Enabled` | set + read back | skip | untouched | yes | yes | admin |
| ASR six rules | merge to Block via `Set-MpPreference -AttackSurfaceReductionRules_*` | merge + read back | skip | untouched | yes | yes | admin |
| **Auto-lock 10 min** | HKLM `SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System`: `InactivityTimeoutSecs` DWORD `600` | write 600 + read back | skip | **untouched (by design)** | yes | yes | admin only (secpol or regedit). Standard user cannot raise or remove it |
| **Store only** | HKLM `SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen`: `ConfigureAppInstallControlEnabled` DWORD `1`, `ConfigureAppInstallControl` String `StoreOnly` | write both + read back; then treat as FAIL if edition is Home (policy CSP is Pro/Enterprise/Education, not Home) | skip | **delete both values** | Pro/Edu/Ent: yes. Home: write may stick, enforcement unverified - log FAIL | same | admin only. Greys Settings > Apps > Erweiterte App-Einstellungen > Quellen zum Abrufen von Apps until the values are gone. Does **not** block USB / network-share / offline copies |
| **Smart App Control On** | HKLM `SYSTEM\CurrentControlSet\Control\CI\Policy`: `VerifiedAndReputablePolicyState` DWORD `1`; state verified against `(Get-MpComputerStatus).SmartAppControlState` | write 1 from Eval **or Off** when Windows Security would allow On (25H2 / SecHealthUI >= 1000.29554); else FAIL and continue | skip | **write 0 (Off)** | **no - disabled** | yes | admin, via Windows Security > App & browser control > Smart App Control. On current 25H2 this is reversible without a Windows reset. After enforcement the unsigned GUI may not start; Windows Security is then the SAC Off path |

### Auto-lock detail

`InactivityTimeoutSecs` is the backing value of secpol "Interactive logon: Machine inactivity limit" (`secpol.msc` > Local Policies > Security Options). 600 = lock the session after ten minutes with no input. It is not the Settings-tab powercfg display timeout, and it is not the personalisation lock-screen timeout a user can change under Settings > Personalization. Applying it does not need a reboot; it takes effect on the next policy evaluation / sign-in. Read back the DWORD after writing; if it is not 600, log FAIL for that item and continue.

Do not write `InactivityTimeoutSecs` to `HKLM\SOFTWARE\Policies\...`; the value the LSA reads lives under `CurrentVersion\Policies\System`. Do not touch `ScreenSaverGracePeriod` or the HKCU screensaver keys.

### Store-only detail

The Explorer-facing setting ("Choose where to get apps" / "Quellen zum Abrufen von Apps") also has an HKLM preference (`Explorer` `AicEnabled`). This tab sets the **policy** only. Rationale: the preference stays changeable in Settings by a daily admin; the policy greys the dropdown.

Both values are required. `ConfigureAppInstallControlEnabled=1` turns the policy on; `ConfigureAppInstallControl` carries the mode string. Write mode `StoreOnly` exactly - not `PreferStore`, not `Recommendations`, not `Anywhere`. Read both back after writing. If `EditionID` is Home, log FAIL even if the values read back (CSP editions omit Home); continue.

Honest scope: Microsoft documents that App Install Control does not cover USB, network shares, or other non-internet sources, and that the policy blocks installation only while online unless shell SmartScreen + PreventOverride are also on. This baseline keeps SmartScreen at Warn, so those holes stay. Store-only is friction on internet-downloaded installers and on the Settings path, not a universal installer gate.

Consequence the technician must know, and the reason Unlock exists: with this policy on, typical web-downloaded / winget Win32 installers are expected to be blocked. That includes `winget install --source winget` on the Install Apps tab (`Windows-PC-Setup.ps1:1141`) **if** AIC intercepts that path (prove on a VM; do not claim it in the UI until proven). Install Apps first, or Unlock first.

### Smart App Control detail

SAC has three states: 0 Off, 1 On/Enforced, 2 Evaluation. Windows often ships new installs in Evaluation and may move to On or Off on its own.

Current Microsoft FAQ (2026): recent Windows 11 updates allow turning SAC on, off, and **back on** from Windows Security without a clean install. Independent write-up: reversible as of 25H2 / Windows Security App v1000.29554+. The probed daily driver is 25H2 Pro with SecHealthUI 1000.29628, SAC currently Off. Older Learn pages that still say "reset to re-enable" are stale for those builds.

Supported automation is not fully documented. The CI DWORD is what the field uses; Windows Security is the supported UI. This tool may:

- Current On (1) -> nothing to do, log SUCCESS ("already On").
- Current Eval (2) or Off (0) -> write 1 **only if** the OS looks like a reversible build (Win11, SecHealthUI >= 1000.29554, or `SmartAppControlState` is readable and the Windows Security toggle is not greyed as unavailable). Then read back. If the value does not stick or Get-MpComputerStatus still says Off, log FAIL and continue. Do not claim SUCCESS from a write that did not change reported state.
- Windows 10, or Win11 where On is unavailable (enterprise-managed, developer mode, diagnostics off, greyed Off) -> **do not write.** Log FAIL with the reason, continue.
- Reboot may still be required before enforcement is live. Read-back after apply proves the *intent* was recorded; the log line must say so.

Order SAC **last** in the apply sequence, after Store-only and after everything else, so an unexpected SAC failure cannot strand the rest.

VM must prove Off->On->Off->On on a 25H2 box before treating the registry write as the product path. If that VM fails, fall back to: checkbox stays, Apply logs FAIL on Off machines, and the intro tells the technician to flip SAC in Windows Security. Do not ship "never write from Off" as the default - that silently drops the locked SAC On control on the usual in-use PC.

### Ordering trap: SAC vs this script

SAC evaluates binaries and scripts. This GUI is an unsigned `.ps1` that loads WinForms (`Add-Type`, `New-Object`). After SAC is enforcing, **the Unlock button may be unreachable because the tool that hosts it will not start** (block, or Constrained Language Mode). PhraseVault's Azure signature does not help this file. Signing the `.ps1` stays out of scope.

That does not make Unlock fake; it splits the two-control toggle:

1. Do all first-time installs while SAC is still Off/Eval (Install Apps tab, drivers, printer software).
2. Apply Harden with the blocking pair on.
3. Reboot. Verify SAC shows On in Windows Security.
4. Later maintenance:
   - If PowerSetup still starts: **Unlock for maintenance** (SAC Off + delete Store-only), install, **Apply Harden** again.
   - If PowerSetup does not start: Windows Security > App & browser control > Smart App Control > **Off**, reboot, then run PowerSetup. Unlock then only needs to clear Store-only if it is still present. Apply Harden turns both back on.

On 25H2, SAC Off is not a one-way trap. The cost is extra clicks and a reboot, not a Windows reset. Intro copy and the completion MessageBox must say: reboot may be required; this unsigned tool may not start while SAC is on; Windows Security is the SAC Off path in that case; then Unlock / Apply in this tab.

Do not tell the technician that re-enabling SAC requires resetting Windows. That was true on older builds and is the wrong instruction on the 25H2 audience.

## Unlock for maintenance - behaviour

Button text `&Unlock for maintenance`. Tooltip: "Smart App Control Off + allow non-Store installers, so you can use Install Apps. Run Apply Harden again afterwards. Does not touch Defender."

Confirm dialog (same `Show-ConfirmationDialog` pattern, with `-Detail`) lists exactly what will happen, both lines always shown with their current state resolved:

- `Smart App Control: On -> Off (reboot may be required; turn back on with Apply Harden or Windows Security)` - or `already Off, nothing to do`, or on Win10 `not applicable`.
- `Remove Store-only policy (ConfigureAppInstallControlEnabled, ConfigureAppInstallControl)` - or `policy not present, nothing to do`.
- A closing line: `Defender, SmartScreen, ASR and the 10-minute lock are NOT changed.`

Apply order: Store-only removal first (immediate effect, lets winget work now), then SAC Off. Each is independently pass/fail and counted like other tabs. Missing values are SUCCESS-as-noop, not FAIL.

Dry run: log `Would delete HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen ConfigureAppInstallControlEnabled, ConfigureAppInstallControl` and `Would set VerifiedAndReputablePolicyState=0 (Smart App Control Off)`. Write nothing.

Removal uses `Remove-ItemProperty -ErrorAction SilentlyContinue` per value; do not delete the `SmartScreen` key itself (other policies live there), and do not delete the parent `Windows Defender` key under any circumstances.

## Apply Harden - sequencing and error handling

1. Validate username/password only if `Create this user` is checked.
2. Confirm dialog lists the selected items, including the three new ones by name. No password in the dialog.
3. If user create is checked and fails: stop before everything else, exactly as v2.3 does - baseline **and** blocking pair are skipped. Message unchanged in shape.
4. Otherwise run, in order: five baseline items -> auto-lock -> Store-only -> SAC. Each item is independent: FAIL increments the fail count and the run continues.
5. Completion MessageBox: existing success/fail counts, plus the reboot/SAC note when SAC was applied, plus the existing "close the elevated PowerShell window before handover" line.

Every new mutating function branches on `$script:DryRun` before touching anything, returns `$true`/`$false`, and writes one SUCCESS or one ERROR line. Keep all log strings ASCII (the em-dash-eats-the-AST trap in `docs/known.md`).

Suggested names, to keep the existing `Set-Harden*` shape and the `test-harden.ps1` AST list:

- `Set-HardenInactivityLock`
- `Set-HardenStoreOnly` / `Remove-HardenStoreOnly`
- `Get-HardenSacState` (returns `On` / `Off` / `Eval` / `Unsupported` plus a reason string; no writes)
- `Set-HardenSmartAppControlOn` / `Set-HardenSmartAppControlOff`

New controls: `$script:ChkHardenAutoLock`, `$script:ChkHardenStoreOnly`, `$script:ChkHardenSac`, `$script:BtnHardenUnlock`.

## Product-rule delta (AGENTS.md)

`AGENTS.md` today allows Harden exactly one HKLM location: `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` `EnableSmartScreen` + `ShellSmartScreenLevel=Warn`. **Amend that bullet before writing any of the new values.** The allowlist after this delta is:

| Path | Values | Access |
|---|---|---|
| `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` | `EnableSmartScreen` (DWORD 1), `ShellSmartScreenLevel` (String `Warn`) | write (existing) |
| `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System` | `InactivityTimeoutSecs` (DWORD 600) | write (new) |
| `HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen` | `ConfigureAppInstallControlEnabled` (DWORD 1), `ConfigureAppInstallControl` (String `StoreOnly`) | write + delete-these-two (new) |
| `HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy` | `VerifiedAndReputablePolicyState` (DWORD 1 or 0) | read always; write only per the SAC state rules above (new) |

Plus Defender preference APIs, as today. Still forbidden, unchanged: HKLM Start policy, `IsEducationEnvironment`, `Policies\Google\Chrome`, and HKLM as a dump for anything not in the table.

Also update at implementation time (not part of this spec's edits): version strings to 2.4 in `Windows-PC-Setup.ps1`, `AGENTS.md`, `.claude/CLAUDE.md` (which still says 2.2), the tab description in `AGENTS.md`, README, and new trap lines in `docs/known.md` for: SAC Off->On is build-dependent (FAQ vs stale Learn); unsigned `.ps1` may not start under enforcing SAC so Windows Security is the Off path; Store-only is internet/AIC not USB; Store-only vs winget must be VM-proven; `InactivityTimeoutSecs` is not powercfg; `AicEnabled` is HKLM Explorer preference, not SAC.

## Undo, for the next technician

Assume the next person has this box, no notes, and no PowerSetup. Everything below is reachable from the Windows UI.

| What | Undo |
|---|---|
| Store-only + SAC together | Run PowerSetup as admin, Harden tab, **Unlock for maintenance**. Only works if the script still starts - see the next row |
| Smart App Control | Windows Security > App & browser control > Smart App Control > **Off**, then reboot. Primary route once SAC is enforcing, because the unsigned script may not start. On 25H2, On again from the same screen or via Apply Harden - not a Windows reset |
| Store-only, without the script | `HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\SmartScreen`: delete `ConfigureAppInstallControlEnabled` and `ConfigureAppInstallControl`. The Settings dropdown stays greyed until both are gone |
| 10-minute auto-lock | `secpol.msc` > Local Policies > Security Options > "Interactive logon: Machine inactivity limit" > `0` (or delete `InactivityTimeoutSecs` under `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System`). Not offered by any button in this tool |
| SmartScreen Warn | Delete `EnableSmartScreen` / `ShellSmartScreenLevel` under `HKLM\SOFTWARE\Policies\Microsoft\Windows\System`. Windows Security toggles stay greyed while they exist |
| ASR / Defender prefs | `Set-MpPreference` / `Add-MpPreference` to Audit or Disabled |
| Standard user | Settings > Accounts, or `Remove-LocalUser`. Not offered in the GUI |

The completion MessageBox and the log should both name the Windows Security route for SAC, so the log file left in `%TEMP%` is enough to undo the box.

## Testing

- `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1` is the merge bar.
- `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1` - extend the required-function list with the new `Set-Harden*` / `Get-HardenSacState` names so an encoding poisoning that eats them fails the build. Add a source assertion that the script never writes `AicEnabled` and never writes `VerifiedAndReputablePolicyState` outside the two guarded functions.
- Do **not** run `Windows-PC-Setup.ps1` on the checkout host. It auto-elevates and mutates. This delta specifically must not be "tested" by applying it to the daily driver: SAC Off there is permanent-by-design, and Store-only would break the operator's own winget.
- Behavioural proof is a disposable VM, four cases:
  1. **Win11 25H2 fresh, SAC in Evaluation** - Apply with everything on; reboot; SAC reads On; confirm whether PowerSetup still starts; confirm Windows Security > SAC > Off recovers it; confirm Off->On again without reset; confirm the Settings apps-source dropdown is greyed; confirm the session locks after 10 minutes idle.
  2. **Win11 25H2 with SAC already Off** (the field case) - Apply must **attempt** SAC On, not skip. Pass only if state becomes On (possibly after reboot). If On is greyed/unavailable, FAIL that item and still apply Store-only and the lock.
  3. **Win10 22H2** - SAC checkbox disabled, Store-only and lock both apply, ASR merge still keeps a pre-existing seventh rule.
  4. **Win11 Home** - Store-only must not log SUCCESS unless enforcement is real (greyed dropdown and a web-downloaded EXE blocked). USB-copied EXE must be checked so we do not claim a hole is closed. Also isolate `winget install --source winget` under Store-only with SAC Off.
- Dry Run on both buttons: nothing written, log lines present, `reg query` shows no new values.
- German-language VM for the Settings-page copy and the SID group checks, as in v2.3.

## Day-to-day (accepted)

Single-admin box, blocking pair on: the employee's Store apps and already-installed software keep working. New downloads from the web do not install, and on Windows 11 unsigned unreputable EXEs do not run at all. The screen locks after ten minutes. The admin can undo all of it, and if they do it themselves at a fake-installer prompt, the box is wiped, not repaired. With the standard user created instead, the same controls become close to absolute for that account, which remains the recommended shape and the reason the checkbox is still in the tab.
