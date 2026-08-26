# Harden Layers and Toggles Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship v2.4 Harden: user-create default off, 10-minute inactivity lock in the baseline, Store-only policy + Smart App Control as an Unlock pair.

**Architecture:** Same single-file WinForms script. Add pure helpers (testable without HKLM writes), then mutating `Set-Harden*` / `Remove-Harden*` that branch on `$script:DryRun`, then UI. Never run the GUI on the checkout host.

**Tech Stack:** Windows PowerShell 5.1, WinForms, `test-syntax.ps1`, `test-harden.ps1`

**Spec:** `docs/superpowers/specs/2026-08-26-harden-layers-toggles-design.md`

## Global Constraints

- ASCII-only Harden log/UI strings in `Windows-PC-Setup.ps1` (em-dash AST trap).
- Dry-run on every mutator. Do not write HKLM from tests.
- Do not run `Windows-PC-Setup.ps1` on this host.
- Do not write Explorer `AicEnabled`.
- Unlock never disables Defender, ASR, SmartScreen, or the inactivity lock.
- SAC Win11 only. Store-only FAIL on Home SKU after write-back.
- Amend `AGENTS.md` HKLM allowlist before new registry helpers.
- Gates: `test-syntax.ps1` then `test-harden.ps1`. No commit unless asked.

---

### Task 1: Failing tests + AGENTS allowlist

**Files:**
- Modify: `AGENTS.md`
- Modify: `test-harden.ps1`

**Interfaces:**
- Produces required AST names listed in Task 2.

- [ ] **Step 1: Amend AGENTS.md Harden HKLM allowlist** to the four-row table in the spec (SmartScreen pair, `InactivityTimeoutSecs`, Store-only policy two values, CI `VerifiedAndReputablePolicyState`). Mention Unlock deletes only the two Store-only values. Tabs line: Harden still exists; note layers/unlock.

- [ ] **Step 2: Write failing tests in `test-harden.ps1`**

Add to `$requiredFns`:
`Test-HardenIsHomeEdition`, `Test-HardenSacUiVersionAllowsOn`, `Convert-HardenSacState`, `Set-HardenInactivityLock`, `Set-HardenStoreOnly`, `Remove-HardenStoreOnly`, `Get-HardenSacState`, `Set-HardenSmartAppControlOn`, `Set-HardenSmartAppControlOff`.

After existing tests, import the three pure functions and assert:

```powershell
Assert-Harden (Test-HardenIsHomeEdition 'Core') 'Core is Home'
Assert-Harden (Test-HardenIsHomeEdition 'CoreSingleLanguage') 'CoreSingleLanguage is Home'
Assert-Harden (-not (Test-HardenIsHomeEdition 'Professional')) 'Pro is not Home'
Assert-Harden (Test-HardenSacUiVersionAllowsOn '1000.29628.1000.0') '29628 allows On'
Assert-Harden (-not (Test-HardenSacUiVersionAllowsOn '1000.29553.0.0')) '29553 does not allow On'
Assert-Harden (-not (Test-HardenSacUiVersionAllowsOn '')) 'empty version does not allow On'
Assert-Harden ((Convert-HardenSacState 0) -eq 'Off') '0 is Off'
Assert-Harden ((Convert-HardenSacState 1) -eq 'On') '1 is On'
Assert-Harden ((Convert-HardenSacState 2) -eq 'Eval') '2 is Eval'

Assert-Harden ($scriptText -notmatch '(?i)Explorer[''\\]AicEnabled' -and $scriptText -notmatch "Name 'AicEnabled'") 'must not write Explorer AicEnabled'
```

- [ ] **Step 3: Run tests, expect FAIL** because the new names are missing.

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: FAIL `Test-HardenIsHomeEdition must exist in the PowerShell 5.1 AST` (or first missing name).

---

### Task 2: Pure helpers + DryRun mutators

**Files:**
- Modify: `Windows-PC-Setup.ps1` after `Set-HardenSmartScreenWarn`

**Interfaces:**
- `Test-HardenIsHomeEdition [string]$EditionId` -> `[bool]`
- `Test-HardenSacUiVersionAllowsOn [string]$Version` -> `[bool]` (true if parseable and `>= 1000.29554.0.0`)
- `Convert-HardenSacState $RegistryDword` -> `[string]` `Off`/`On`/`Eval`
- `Get-HardenSacState` -> hashtable `State`, `Reason`, `Reversible` (reads only)
- `Set-HardenInactivityLock` / `Set-HardenStoreOnly` / `Remove-HardenStoreOnly` / `Set-HardenSmartAppControlOn` / `Set-HardenSmartAppControlOff` -> `[bool]`, DryRun first

- [ ] **Step 1: Implement the three pure functions** (ASCII strings). Home SKUs: `Core`, `CoreN`, `CoreSingleLanguage`, `CoreCountrySpecific` (case-insensitive). Version compare with `[version]`; pad to four parts if needed.

- [ ] **Step 2: Implement mutators.** Every mutator: if `$script:DryRun` then `Write-Log "Would ..."` INFO and `return $true`. Live paths use `$ErrorActionPreference` Stop, read back, FAIL if not sticky. Store-only: do not delete the SmartScreen key. After Store-only write, if `Test-HardenIsHomeEdition` on current `EditionID`, log FAIL. SAC On: Win11 + reversible UI version (or Eval/On already). SAC Off: Win10 no-op SUCCESS. Do not write `AicEnabled`.

- [ ] **Step 3: In `test-harden.ps1`, stub `Write-Log`, set `$script:DryRun = $true`, import mutators, call each, assert `$true`.** Then `$script:DryRun = $false` is not required if tests exit.

- [ ] **Step 4: Run `test-harden.ps1` and `test-syntax.ps1`. Expect PASS.**

---

### Task 3: Harden tab UI + handlers

**Files:**
- Modify: `Windows-PC-Setup.ps1` Harden tab (~2774-3023), version strings 2.3 -> 2.4, `Set-ButtonEnabled` restore `OriginalBg`.

**Interfaces:** `$script:ChkHardenAutoLock`, `$script:ChkHardenStoreOnly`, `$script:ChkHardenSac`, `$script:BtnHardenUnlock`. Create-user default `$false`. Grey username/password/must-change when create is off.

- [ ] **Step 1: Source assertions in `test-harden.ps1`:** `ChkHardenCreateUser.Checked = $false`; `Unlock for maintenance`; `ChkHardenAutoLock`; `ChkHardenStoreOnly`; `ChkHardenSac`. Run, expect FAIL.

- [ ] **Step 2: UI + Apply sequence:** five baseline, auto-lock, Store-only, SAC last. Unlock: Store-only remove then SAC Off. Confirm copy per spec. SAC completion note. Alt+U. Intro Height ~80.

- [ ] **Step 3: Re-run both test scripts. Expect PASS.**

---

### Task 4: Docs

**Files:** `README.md`, `docs/known.md`, `.claude/CLAUDE.md`, spec status line `implemented`.

Traps: inactivity is not powercfg; Store-only is policy not USB; SAC may block this unsigned `.ps1`; no Explorer `AicEnabled`; Unlock does not clear the lock.

- [ ] **Step 1: Write the docs.**
- [ ] **Step 2: Re-run both test scripts.**
