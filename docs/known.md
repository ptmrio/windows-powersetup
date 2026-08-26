# Known

Commands and traps only. Delete what stops being true.

## Commands

| Command | Proves |
|---|---|
| `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1` | Parser accepts `Windows-PC-Setup.ps1` (`exit 1` on errors) |
| `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1` | Production Harden helpers extracted from the script AST |

Do not run `Windows-PC-Setup.ps1` on this host without an explicit OK. It auto-elevates.

## Traps

- **Win11 `ProductName` lies.** Registry often still says "Windows 10". Use `CurrentBuild -ge 22000` (`$IsWindows11`). Also read `DisplayVersion` / `EditionID` / UBR for lifecycle text. Build family `28000` is 26H1 (Snapdragon X2 OEM-only, not an in-place feature update). 26H2 is not GA as of 2026-08-25.
- **No `IsEducationEnvironment`.** EDU CSP only. Sophia removes it. Win11Debloat dropped the HKLM education hack (2026-06-11); do not reintroduce it.
- **No HKLM Start policy.** Initial setup, not MDM. HKCU preference + HKCU policy fallback only. `ConfigureStartPins` exists (24H2+KB5062660) but locks pin/unpin — wrong for this tool.
- **Hide Recommended:** HKCU policy `HideRecommendedSection` is Pro+ CSP. `HKCU\...\Start\HideRecommendedSection` is not a Sophia key (invented; likely no-op). Sophia sets `ShowRecentList`/`ShowFrequentList`/`Start_IrisRecommendations`/`Start_TrackDocs`. `ShowRecentSection=0` is the section toggle on Insider builds 26200.9267+ / 26300.8553+; current GA is 26200.9168 (KB5121003) so treat it as forward-compat dual-write, not a GA proof. Win11Debloat #419 is closed (reporter used clear-pins, not hide-recommended) — do not cite it.
- **More pins:** `Start_Layout=1` (old Start) **and** `ShowAllPinsList=1` at `HKCU\...\CurrentVersion\Start` (new Start). Unknown keys ignored is harmless.
- **All Apps view:** `HideCategoryView` is Pro+ CSP/GPO since 24H2+KB5067036. All-editions pref is `HKCU\...\Start\AllAppsViewMode` (`0` Category / `1` Grid / `2` List). Policy removes the Category option; the pref only selects Grid.
- **`start2.bin` is unsupported on 24H2+.** Keep it for current-user empty pins; do not also treat `settings.dat` as required (Win11Debloat #282 closed, file often absent on 25H2). Default-profile `LayoutModification.json` should include `"applyOnce":true` (24H2+KB5062660) so new users can re-pin. Success must not be logged if the binary write, JSON write, or host stop failed independently.
- **New vs old Start cannot be detected by build alone.** Feature-flagged (ViveTool 47205210, KB5067036 / KB5074109). Dual-write keys safe on both.
- **Xbox names:** `Microsoft.XboxApp` is discontinued Console Companion. Current Store app is `Microsoft.GamingApp`. `Microsoft.XboxGamingOverlay` is Game Bar (Win+G). Do not precheck TCUI / IdentityProvider / Speech.
- **Apply Settings has no confirm dialog** (Remove/Install do). Dry Run defaults off (`$script:ChkDryRun.Checked = $false`).
- **`Test-ProtectedApp` has no callers.** `$ProtectedApps` is documentation unless wired into detect/remove. `Test-ProtectedWin32` is enforced.
- **`winget source reset --force`** on every install init removes non-default sources. First-run agreements are `--accept-source-agreements`, not reset.
- **Repair order is DISM then SFC**, then optional CHKDSK. DISM `/Source` must be a mounted `Windows` dir, `sources\SxS`, or `Wim:<path>:<index>` — not a random folder containing `install.wim`. Auto `/LimitAccess` with a bad source blocks Windows Update fallback.
- **Do not copy Sophia's full policy-clear + third-party Start11/StartAllBack check.** Out of scope for a one-shot setup tool.
- **Chrome policy hijack (fake PDF bundlers).** Search-ad “PDF converter/reader” setups (`PDFConvertSetup.exe`, `PDFReaderSetup.exe`) install OneBrowser/Sheaf plus an updater that writes Chrome **machine policies**. Yahoo is locked; new-tab fake-search hops (e.g. newspulsenow.net, healthygeorge.com). Workspace “Managed by …” can still be real. Lives on that PC; shared drives often clean. **Wipe, do not Chrome-reset; do not restore that profile or those downloads.**
- **Harden groups are SIDs.** Users `S-1-5-32-545`, Administrators `S-1-5-32-544`. German Windows names are `Benutzer` / `Administratoren`.
- **Harden ASR merge.** Set the six GUIDs to Block without wiping other ASR rules already on the device. Do not wrap `$null` Defender lists in `@()` then test `$null -eq $ids`.
- **Harden SmartScreen** is HKLM `Policies\Microsoft\Windows\System` `EnableSmartScreen=1` + `ShellSmartScreenLevel=Warn`, not Block, not Chrome. Undo = remove those properties (Windows Security toggles stay greyed while the policy exists).
- **Do not write Chrome policies from PowerSetup.** Workspace cloud-over-local belongs in Google Admin and is not automatic (`chrome://policy`). Fake PDF bundlers win by writing machine Chrome policy.
- **New-LocalUser has no `-UserMayChangePassword` and no `-ChangePasswordAtLogon` on Windows PowerShell 5.1.** Default already allows password change. Use `net user <name> /logonpasswordchg:yes` when that checkbox is on.
- **Harden existing-user is fail-closed.** Typing the current admin, German `Gast`, or a non-Users member must not skip to the Defender baseline. If “must change password at next logon” is on, retry an already-created standard user by running `net user /logonpasswordchg:yes` again; do not treat skip-exists as success until that postcondition holds.
- **PowerShell 5.1 + UTF-8 without BOM.** An em dash `—` inside a double-quoted string is decoded as Windows-1252 `â€”`; the `”` byte silently ends the string and later functions disappear from the AST with **zero parse errors**. `test-syntax.ps1` cannot catch that. Keep Harden log strings ASCII, or save the script UTF-8 with BOM. `test-harden.ps1` asserts Defender functions exist in the 5.1 AST and that `New-HardenStandardUser` does not span hundreds of lines.
- **LocalAccounts is 64-bit only** on 64-bit Windows. 32-bit PowerShell cannot create the standard user.
- **Harden auto-lock is `InactivityTimeoutSecs=600`**, not the Settings-tab powercfg display-off. Unlock does not clear it. Undo: secpol Interactive logon: Machine inactivity limit = 0, or delete the DWORD.
- **Harden Store-only is machine policy** `ConfigureAppInstallControl=StoreOnly`, not Explorer `AicEnabled`. Does not cover USB/network-share/offline copies. Home SKU may not enforce. Unlock deletes those two values only, never the SmartScreen key.
- **Harden SAC** writes `VerifiedAndReputablePolicyState`. Win11 only. After enforcement this unsigned `.ps1` may not start; Windows Security is the SAC Off path, then Unlock for Store-only. Do not tell technicians that 25H2 needs a Windows reset to turn SAC back on.

SKU snapshot (Microsoft release-health, 2026-08 B): 25H2 `26200.9168`; 24H2 `26100.9168` Home/Pro EOS 2026-10-13; 26H1 `28000.2704` OEM-only; Win10 consumer EOS 2025-10-14 (ESU/LTSC still serviced).

Research dump (April 2026, several Start rows now stale): `research-start-menu-2026.md`. Plan (gitignored): `.todo/start-menu-implementation-plan.md`.
