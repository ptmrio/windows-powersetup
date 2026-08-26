# Brave config after install - design (v2.5)

Date: 2026-08-26
Status: implemented (v2.5)
Product: Windows PowerSetup (`Windows-PC-Setup.ps1`, single-file WinForms)
Amends: Install Apps tab only. Harden v2.4 is unchanged.

ASCII only. See the encoding trap in `docs/known.md`.

## Problem

The technician installs Brave with winget, then repeats the same Brave settings by hand on every profile: Google search, new-tab startup, Proton Pass owns credentials, ask-where-to-save downloads, and a desktop shortcut per named profile. That hand work does not survive a second profile unless it is policy, and it is easy to do in the wrong order (configure before Brave exists, or create shortcuts before any profile exists).

## Product locks (from the 2026-08-26 spike)

Keep only mechanisms that are policy or a documented shortcut. Leave out everything the spike marked fragile or impossible.

**In scope**

- Startup opens the New Tab page (`RestoreOnStartup=5`).
- Default search is Google (full default-search policy group, not a display name).
- Disable Brave's password manager and address/card autofill (`PasswordManagerEnabled`, `AutofillAddressEnabled`, `AutofillCreditCardEnabled`).
- Ask where to save each download (`PromptForDownloadLocation=1`).
- Default download folder is the Windows Downloads Known Folder (`DefaultDownloadDirectory`, not the locked `DownloadDirectory`).
- Force-install the Proton Pass browser extension.
- Desktop shortcuts named `Brave <ProfileName>` for existing profiles, on a **separate** button.

**Out of scope (do not implement)**

- Icon-only hiding of Rewards, Sidebar, Wallet, Leo, VPN. Feature-disable policies used as a proxy for hiding icons are also out.
- Per-profile `Preferences` / `Secure Preferences` JSON edits.
- `master_preferences` / `initial_preferences`.
- Setting Proton Pass as the selected default password provider (no desktop policy).
- Proton Pass PIN or other in-extension settings.
- Brave Sync chain join or `BraveSyncUrl`.
- Google "Application Launcher for Drive" CWS force-install (native-messaging is unproven; Drive for desktop stays on the winget list).
- "Show downloads when done" (no policy; Chromium default is already on).
- HKLM Brave policies.
- Chrome policies (`Google\Chrome`). Still forbidden.
- Launching, killing, or restarting Brave.
- Signing the `.ps1`.

Honest leftover for the technician (say it in the confirm dialog, not only here): sign into Proton Pass once; PIN and Sync stay manual; close Brave after Apply so search/extension take effect.

**Daily-admin path only.** HKCU policies and current-user Desktop shortcuts belong to the Windows account that launched the elevated script. That is the common Harden path (create-user off). If Harden later creates a standard user who will run Brave, that user does not get these policies or shortcuts. v2.5 does not write HKLM Brave keys and cannot fix that. Intro copy must say so. Do not tell the operator to "Harden last" as if the new standard user inherits Brave config.

## Guarantee vs friction

These are HKCU **policies** for the Windows account that is running the elevated script. They apply to every Brave profile that account creates, including ones that do not exist yet. They do not follow a later standard user. They grey out the matching rows in `brave://settings`. An admin can delete the key. Wipe is still the recovery for a hijacked profile.

## Hive and allowlist

Write only:

```
HKCU:\SOFTWARE\Policies\BraveSoftware\Brave
```

AGENTS.md today says "No Chrome policies." Keep that. Add an explicit Brave HKCU allowlist (values listed under "Policy payload"). Do not write HKLM Brave keys. Do not write Explorer `AicEnabled`.

If the elevated process is a different Windows user than the daily account, these policies land on the wrong hive. That matches the rest of this tool (Settings tab is also HKCU of the runner). Intro copy: apply while signed in as the daily admin.

## Workflow (this is the product)

Three separate actions. Never one "apply all".

```
1. Install Selected          winget only
        |
        |  Brave.Brave must succeed (or already be present)
        v
2. Apply Brave settings      HKCU policies + Proton Pass force-install
        |
        |  technician opens Brave, names profiles, signs into Proton Pass
        v
3. Create profile shortcuts  .lnk files from Local State; rerun after new profiles
```

Order rules:

1. **Install Selected does not configure Brave.** Winget stays winget. Mixing install-fail with policy-fail in one dialog is a lie.
2. **Apply Brave settings is disabled until `brave.exe` exists.** Tooltip: "Install Brave first." Detection is a resolved file path, not "was it in this session's checkbox list." Already-installed Brave on an existing PC is a valid target.
3. **Apply Brave settings does not create shortcuts.** Profiles usually do not exist yet. First launch is when Brave creates `User Data`.
4. **Create profile shortcuts is disabled until `Local State` has at least one real profile.** Skip `Guest Profile` and `System Profile`. Tooltip if Brave is installed but never launched: "Open Brave, create and name profiles, then retry."
5. **Apply Brave settings before Harden Store-only / SAC.** Extension force-install is Brave fetching a CWS CRX on next launch. Intro on Install Apps: do Brave settings before the Harden blocking pair.
6. **Do not start Brave** from this tool. First-run UI belongs to the technician.
7. Re-apply is idempotent for our payload: upsert the Proton Pass force-install entry (see policy table); overwrite `.lnk` files that this run computed for that profile id.
8. Button enablement is evaluated when the Install Apps tab is selected (`TabControl.SelectedIndexChanged`), after Install Selected finishes, and once after the buttons are created. Profiles created while the GUI stays open must enable Shortcuts on the next visit to the tab.

`Update-BraveButtonState` must not throw under `Set-StrictMode -Version Latest`. `Get-BraveProfilesForShortcuts` and the button-state helper must test JSON properties via `PSObject.Properties['name']` (missing `profile` / `info_cache` / empty `{}` are empty-list, not a crash). Wrap the Local State read in try/catch so a bad file cannot prevent the form from opening.

## Operator sequence on a new PC

1. Finish the admin account, updates, drivers.
2. Install Apps: tick Brave, Proton Pass (desktop), Google Drive, others. Install Selected.
3. Apply Brave settings. Close Brave if it is already open.
4. Open Brave. Confirm `brave://policy` shows the values. Sign into Proton Pass. Create extra profiles if needed.
5. Create profile shortcuts.
6. Other tabs. Harden last only if the daily account is this same admin. If you will create a standard user who uses Brave, do not expect these Brave settings to follow them.

On a PC where Brave is already installed: skip step 2 or leave Brave unticked; steps 3-5 still work.

## UI (Install Apps tab)

Keep the existing winget list and Install Selected (Alt+I) row.

Below that row, a section header `Brave` and two buttons on one line:

- `Apply Brave settings` (Alt+B). Accent button. Disabled when Brave is missing.
- `Create profile shortcuts` (Alt+P). Standard button. Disabled when there is no real profile in Local State.

Intro text (substance, not exact copy): install Brave first; then Apply Brave settings (Google, new tab, Proton Pass extension, download prompt); after you have named profiles, create desktop shortcuts. Do this before Harden. PIN, Sync, and Proton sign-in stay manual.

No per-setting checkboxes. Apply writes the whole in-scope payload. YAGNI.

Confirm dialogs use `Show-ConfirmationDialog -Detail` (same as Harden). Dry Run is the existing global checkbox.

**Apply confirm** lists: New Tab on startup; Google search; Brave password/address/card autofill off; ask where to save; Downloads folder; Proton Pass extension. One line: close Brave if it is open. One line: sign-in / PIN / Sync remain manual.

**Shortcuts confirm** lists the shortcuts that will be written (`Brave <name>` -> `--profile-directory=<id>`). If the list is empty after filtering, do not show Yes/No; show an info box instead.

Progress: reuse `Update-Progress` / `Update-Status`. Disable the clicked button with `Set-ButtonDisabled` / `Set-ButtonEnabled` like Install Selected.

## Policy payload

All values under `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave`. DWORD unless noted. Write the search group in one function so a partial write cannot leave Brave Search as default.

| Policy | Value |
|---|---|
| `RestoreOnStartup` | `5` (New Tab Page) |
| `DefaultSearchProviderEnabled` | `1` |
| `DefaultSearchProviderName` | `Google` (REG_SZ) |
| `DefaultSearchProviderKeyword` | `google.com` (REG_SZ) |
| `DefaultSearchProviderSearchURL` | `https://www.google.com/search?q={searchTerms}` (REG_SZ) |
| `DefaultSearchProviderSuggestURL` | `https://www.google.com/complete/search?output=chrome&q={searchTerms}` (REG_SZ) |
| `PasswordManagerEnabled` | `0` |
| `AutofillAddressEnabled` | `0` |
| `AutofillCreditCardEnabled` | `0` |
| `PromptForDownloadLocation` | `1` |
| `DefaultDownloadDirectory` | resolved Downloads Known Folder path (REG_SZ). Never `DownloadDirectory` (that locks the UI). |
| `ExtensionInstallForcelist` | child key. Each numeric REG_SZ is `id;https://clients2.google.com/service/update2/crx`. Upsert: if any numeric value already contains id `ghmbeldphafepmbegfdlkpapadhbakde`, leave it. Else write that string to the first unused numeric name (`1`, `2`, ...). Never overwrite a different extension's entry. Never delete sibling values. |

Search group write order: Name, Keyword, SearchURL, SuggestURL first, then `DefaultSearchProviderEnabled=1` last. If any of those four SZ writes fail, do not set Enabled. Read back after write. Any missing or mismatched value is FAIL for that item; continue other groups (startup, autofill, downloads, extension). DryRun logs `Would set ...` and returns success without writing.

Apply result UI: if any item failed, the modal is Warning/Error and the status line says finished with errors. Do not show a success-only "complete" box on partial failure.

Do not write `RestoreOnStartupURLs`. Do not write `session.restore_on_startup` in Preferences.

## Helpers (testable)

Place after Harden helpers, still in `Windows-PC-Setup.ps1`. Names:

- `Get-BraveInstallPath` -> `[string]` full path to `brave.exe`, or empty. Search in order: `ProgramFiles`, `ProgramFiles(x86)`, `$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\Application\brave.exe`.
- `Test-BraveInstalled` -> `[bool]`
- `Get-BraveUserDataDir` -> `[string]` `$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data`
- `Get-BraveDownloadsFolder` -> `[string]` Windows Downloads Known Folder for the current user. Not `Join-Path $HOME Downloads` (redirected folders).
- `Get-BravePolicyPath` -> `[string]` `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave`
- `Get-BraveSearchPolicyMap` -> `[hashtable]` the search-group names and values (pure).
- `ConvertTo-BraveShortcutFileName` `[string]$ProfileName` -> `[string]` `Brave <sanitized>`. Replace `<>:"/\|?*` and control chars with `_`; trim trailing dots and spaces; if empty after sanitize, use `Brave Profile`.
- `Get-BraveUniqueShortcutNames` `$Profiles` -> array of `Id`, `Name`, `FileName` (pure). Collision loop as specified under Shortcuts.
- `Get-BraveProfilesForShortcuts` `[string]$LocalStateJson` -> array of hashtables `Id`, `Name` (pure; parses JSON). Skip ids `Guest Profile` and `System Profile`. Use `info_cache` keys as `Id` and `.name` as `Name`. If `.name` is missing, use the id.
- `Set-BraveHkcuPolicies` -> `[bool]` overall (false if any item failed). Branches on `$script:DryRun` first.
- `New-BraveProfileShortcuts` `[string]$DesktopDir`, `[string]$BraveExe`, `$Profiles` -> `[bool]`

Tests live in `test-brave.ps1`, same AST-extract pattern as `test-harden.ps1`. Tests must not write HKCU policy, must not create `.lnk` on the real Desktop (use a temp dir), must not run the GUI.

## Shortcuts

Target: resolved `brave.exe`.
Arguments: `--profile-directory="<Id>"` with quotes. `Id` is `Default` or `Profile 1`, never the friendly name.
WorkingDirectory: directory of `brave.exe`.
Icon: `brave.exe,0`.
Path: `[Environment]::GetFolderPath('Desktop')` + `ConvertTo-BraveShortcutFileName`. Current-user Desktop, not Public Desktop.
Implementation: `WScript.Shell` `CreateShortcut` (PowerShell 5.1 has no better built-in).
If `brave.exe` is missing at click time, FAIL closed and write nothing.
Precompute a unique filename per profile before writing or showing the confirm dialog. If `ConvertTo-BraveShortcutFileName` collides, suffix `_` + Id with spaces as `_`, then sanitize again; if that still collides, suffix `_` + an incrementing integer. Confirm dialog lists the actual `.lnk` names that will be written, not the raw profile names.

## Failure behaviour

- Apply with Brave missing: button disabled; if somehow clicked, info box, no registry writes.
- Policy key create failed: FAIL that item, continue.
- Proton Pass CWS id stays the constant above. Do not scrape the store.
- Shortcuts with Brave running: still write `.lnk` files (they do not need a restart).
- Local State unreadable / invalid JSON: info box, no shortcuts.

## Docs and version

v2.5 in the script header, startup banner, form title, footer (add Alt+B and Alt+P), `AGENTS.md`, `README.md`, `.claude/CLAUDE.md`.

`docs/known.md` traps:

- Brave config is HKCU policy, not Preferences JSON (Secure Preferences would drop startup/search).
- Shortcuts need the directory id, not the profile display name.
- Apply Brave settings before Harden Store-only/SAC.
- HKCU Brave config does not follow a later standard user.
- Proton sign-in / PIN / Sync are manual.
- Downloads path is the Known Folder, not `$HOME\Downloads`.

## Non-goals already stated, repeated for implementers

No toolbar icon work. No JSON prefs. No Drive launcher. No HKLM Brave. No Unlock button (these policies do not block winget; technician undo is delete the Brave policy key, documented in known.md).
