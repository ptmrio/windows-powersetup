# Brave tab with per-setting checkboxes - design (v2.6)

Date: 2026-08-26
Status: implemented (v2.6)
Product: Windows PowerSetup (`Windows-PC-Setup.ps1`, single-file WinForms)
Amends: `docs/superpowers/specs/2026-08-26-brave-config-design.md` (v2.5). Policy payload, hive, helpers, shortcut rules, and out-of-scope list stay unless this file says otherwise.

ASCII only. See the encoding trap in `docs/known.md`.

## Problem

v2.5 put Brave apply and profile shortcuts on the Install Apps tab as two always-on payload buttons. That has three product bugs:

1. Brave config is mixed with winget. The technician cannot treat "configure an already-installed browser" as its own job.
2. There are no per-setting checkboxes. Settings and Harden let the operator uncheck one row. Brave applies the whole payload or nothing.
3. Copy and mental model still sound like this session's Install Selected. Detection already uses `brave.exe` on disk; the UI must match that. An existing PC with Brave already installed is a first-class target.

## Product locks

- Brave settings live on their **own tab**, not a section under Install Apps.
- The tab uses the **same verbose checkbox pattern** as Settings / Harden: one descriptive checkbox per in-scope policy group, tooltips, then Apply writes only checked rows.
- Mutating controls **do not depend on this session having run Install Selected.** They depend on whether `brave.exe` is on disk (`Get-BraveInstallPath` / `Test-BraveInstalled`).
- Shortcuts stay a **separate** button. Profiles are a different precondition (`Local State`), not a policy checkbox mixed into Apply.
- Policy payload, HKCU-only, Proton Pass CWS id, no Preferences JSON, no icon hiding, no Drive launcher, no HKLM Brave, no Chrome: unchanged from v2.5.

## Approaches considered

1. **Dedicated tab + Settings-style checkboxes (chosen).** Tab order: Install Apps, then Brave, then Harden. Checkboxes stay interactive so the technician can pre-select; Apply enables when `brave.exe` exists; Shortcuts enables when real profiles exist.
2. Keep Brave on Install Apps, add checkboxes there. Rejected: still mixed with winget; the user asked for its own tab.
3. Hide or disable the whole Brave tab until Brave is installed. Rejected: the operator must see the list and the "not found" reason; an existing-PC visit should not look empty.

## Tab order

1. Remove Bloatware
2. Settings
3. Install Apps (winget only)
4. **Brave** (new)
5. Harden
6. Repair

Brave after Install Apps, before Harden: install the browser, configure it, then apply Store-only / SAC.

## Enablement (disk, not session)

`Test-BraveInstalled` remains "resolved `brave.exe` path is non-empty." Search order unchanged: `ProgramFiles`, `ProgramFiles(x86)`, `$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\Application\brave.exe`.

| Control | Enabled when |
|---|---|
| Policy checkboxes | Always (pre-select is allowed while Brave is missing) |
| Apply Brave settings | `Test-BraveInstalled` |
| Create profile shortcuts | Local State has at least one real profile (same skip list as v2.5) |

Refresh `Update-BraveTabState` (rename from `Update-BraveButtonState`) when:

- The Brave tab is selected (`TabControl.SelectedIndexChanged` compares `$TabBrave`, not `$TabInstallApps`)
- Install Selected finishes (Brave may have just landed on disk)
- After the Brave tab controls are created

Do not treat "Brave was in this session's winget checkbox list" as a signal. Do not require Apply to have run this session before Shortcuts.

Status label at the top of the Brave tab (same idea as the winget status line):

- Found: `Brave found: <full path to brave.exe>` (success green)
- Missing: `Brave not found. Install it on the Install Apps tab. This tab enables when brave.exe is on disk.` (error red)

Apply / Shortcuts fail-closed copy if somehow clicked while disabled: same substance as the status line. Do **not** say "tick Brave Browser, then Install Selected."

## UI

Intro (substance, not exact copy): these HKCU policies apply to this Windows account only and do not follow a later standard user. Close Brave after Apply. Proton sign-in, PIN, and Sync stay manual. Do this before Harden Store-only / SAC so the Proton Pass CRX can download.

Section **Policies** — eight checkboxes, default **checked**, verbose labels:

| Tag | Checkbox text | Writes |
|---|---|---|
| `Startup` | Open New Tab page on startup | `RestoreOnStartup=5` |
| `Search` | Set search engine to Google | full search group, URLs first then Enabled |
| `PasswordManager` | Disable Brave password manager | `PasswordManagerEnabled=0` |
| `AutofillAddress` | Disable address autofill | `AutofillAddressEnabled=0` |
| `AutofillCreditCard` | Disable card autofill | `AutofillCreditCardEnabled=0` |
| `DownloadPrompt` | Ask where to save each download | `PromptForDownloadLocation=1` |
| `DownloadDirectory` | Set default download folder to Windows Downloads | `DefaultDownloadDirectory` Known Folder |
| `ProtonPass` | Force-install Proton Pass extension | `ExtensionInstallForcelist` upsert |

Tooltips: one sentence each, matching Settings. Password/autofill tooltips mention Proton Pass will own those. Proton Pass tooltip: extension installs on next Brave launch; sign-in stays manual.

`$script:BraveCheckboxes` array, same construction as `$script:SettingsCheckboxes` (Tag, tooltip, indent, height).

No Select All / Deselect All. Eight rows is small; Settings and Harden do not have those either.

Apply button: `Apply &Brave settings` (Alt+B), accent, right-aligned like Settings Apply. Disabled when Brave is missing. If zero checked boxes: info "No Brave settings selected to apply." Confirm dialog lists **only the checked tags**, plus the close-Brave and manual leftover lines.

Re-check `Test-BraveInstalled` **after** the confirm dialog returns Yes and **before** `Set-BraveHkcuPolicies`. If Brave disappeared while the dialog was open: info box, `Update-BraveTabState`, no registry writes. Same re-check of `Get-BraveInstallPath` after the Shortcuts confirm, before `New-BraveProfileShortcuts`.

After the Apply button is added, advance `$braveY` by `$script:UI.ButtonHeight` plus section spacing, then the Profiles header. Do not place Profiles at the same Y as Apply.

Apply result: if `Set-BraveHkcuPolicies` returns false, Warning/Error modal, not a success-only box. After work, `Set-ButtonEnabled` restores the original caption, then `Update-BraveTabState` sets Enabled from disk (so Apply stays grey if Brave is still missing). Do not leave the button labelled "Applying...".

Section **Profiles** — one button: `Create &profile shortcuts` (Alt+P). Unchanged shortcut behaviour from v2.5 (confirm lists actual `.lnk` names, DryRun, current-user Desktop).

Install Apps tab: remove the Brave section header and both Brave buttons. Restore a winget-only intro. One short line is allowed: configure Brave on the Brave tab after `brave.exe` exists. Install Selected still must not call `Set-BraveHkcuPolicies` or `New-BraveProfileShortcuts`. It may call `Update-BraveTabState` so a just-installed Brave enables Apply without a restart.

## Policy apply selection

`Set-BraveHkcuPolicies` gains `[string[]]$Groups`. Default when omitted: all eight tags (keeps DryRun tests and callers honest). Empty array: no writes, return `$true`.

Only listed groups are written. Unchecked groups are skipped, **not deleted**. Apply is set-only, like Settings. Undo remains: delete `HKCU\SOFTWARE\Policies\BraveSoftware\Brave` (known.md).

Search group write order and ExtensionInstallForcelist upsert rules stay as v2.5. Do not set `DefaultSearchProviderEnabled` if the four SZ writes failed.

DryRun logs `Would set ...` only for selected groups. `test-brave.ps1` must capture those log lines: default/omitted Groups logs startup and Proton Pass; `-Groups @('Startup')` logs RestoreOnStartup and does not log PasswordManager or ExtensionInstallForcelist; `-Groups @()` returns `$true` and logs no `Would set`.

Apply click AST must call `Set-BraveHkcuPolicies -Groups` with checkbox `.Tag` values, not a no-arg call that would default to all eight.

## Helpers

Keep the v2.5 helper list. Rename `Update-BraveButtonState` -> `Update-BraveTabState`. Add it to `test-brave.ps1` `$requiredFns`. Tests must import it and prove it does not throw when labels/buttons are still `$null`. Its function AST must call `Get-BraveInstallPath` or `Test-BraveInstalled` and must not mention `$script:AppCheckboxes` or a session-install flag.

It must:

- Set the status label text/color from `Get-BraveInstallPath`
- Set Apply.Enabled from `Test-BraveInstalled`
- Set Shortcuts.Enabled from profile presence (try/catch around Local State; StrictMode-safe `PSObject.Properties` as today)
- Never throw

`$script:BtnBraveApply`, `$script:BtnBraveShortcuts`, `$script:LblBraveStatus`, `$script:BraveCheckboxes` are script-scoped like Harden controls.

## Failure behaviour

Unchanged from v2.5 except copy and selection. Brave missing: Apply disabled; no registry writes. Local State unreadable: Shortcuts disabled / info box, no `.lnk`. Partial policy failure: continue other selected groups.

## Docs and version

v2.6 in script header, banner, form title, footer, `AGENTS.md`, `README.md`, `.claude/CLAUDE.md`.

AGENTS tabs: Install Apps is winget only; new line 4 Brave (HKCU policies + profile shortcuts); Harden becomes 5; Repair becomes 6.

Footer keeps Alt+B and Alt+P. Do not lengthen the footer string; the v2.5 review already clipped it.

`docs/known.md`: keep v2.5 Brave traps. Add: Brave tab enablement is `brave.exe` on disk, not this session's winget run. Add: unchecking a Brave row skips that write; it does not delete an already-written policy.

## Non-goals

No new policies. No toolbar icon work. No JSON prefs. No HKLM. No Brave Unlock button. No splitting the `.ps1`. No launching Brave.
