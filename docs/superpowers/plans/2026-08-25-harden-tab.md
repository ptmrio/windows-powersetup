# Harden Tab Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a fifth Harden tab to `Windows-PC-Setup.ps1` that creates a named standard local user and applies a conservative machine-wide Defender/SmartScreen/ASR baseline on clean installs.

**Architecture:** Keep the single-file WinForms script. Add pure validation helpers (unit-tested without launching the GUI), then local-user and Defender functions with `$script:DryRun` branches, then the tab UI and Apply handler. Groups are resolved by SID so German Windows works. No Chrome policies. Do not run `Windows-PC-Setup.ps1` on the daily-driver checkout host.

**Tech Stack:** PowerShell 5.1, WinForms, `Microsoft.PowerShell.LocalAccounts` (`New-LocalUser`), Defender `Set-MpPreference` / `Add-MpPreference` / `Get-MpPreference`.

**Spec:** `docs/superpowers/specs/2026-08-25-harden-tab-design.md`

## Global Constraints

- Single file: `Windows-PC-Setup.ps1`. Do not split the GUI script.
- `#Requires -Version 5.1`. `Set-StrictMode -Version Latest`. `$script:DryRun` on every mutating function.
- Merge bar: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1`. Extra gate: `test-harden.ps1` (no elevation, no GUI).
- Do not execute `Windows-PC-Setup.ps1` on this host unless the operator explicitly OK’s it.
- No canned usernames. No HKLM Start policy. No `IsEducationEnvironment`. No `Policies\Google\Chrome`. No PUP/bloatware list on this tab. No Smart App Control, BitLocker, Controlled Folder Access, Office/Acrobat child-process ASR.
- Local Users group SID `S-1-5-32-545`. Administrators SID `S-1-5-32-544`. Never add the new user to Administrators. Never demote the current user.
- If “Create this user” is on and create fails, **stop before baseline**. If the user already exists, skip (no password/group change) and continue.
- HKLM allowed only for SmartScreen: `HKLM:\SOFTWARE\Policies\Microsoft\Windows\System` `EnableSmartScreen=1` (DWord), `ShellSmartScreenLevel=Warn` (String), plus Defender cmdlets.
- ASR: Block (Enabled) only these six GUIDs; **merge** into existing ASR lists, do not replace the whole set.
- Confirm dialog like Remove/Install (not Settings). Dry Run is the existing bottom checkbox.
- Commit only if the operator asked for commits this session. Never push. No Co-Authored-By.

## Post-review amendments (supersede code below)

Independent Claude + Codex review (2026-08-25). Host-verified: `New-LocalUser` has **no** `-UserMayChangePassword`; `CloudBlockLevel` High=2.

When implementing, **do not paste** the pre-review snippets that contradict this list.

1. `New-LocalUser`: drop `-UserMayChangePassword`. Optional: `-UserMayNotChangePassword:$false`.
2. `Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $name -ErrorAction Stop`. Verify membership by SID. Do not `SilentlyContinue` the Users add. Do not treat “exists” as success unless Users-yes and Administrators-no.
3. Reserved accounts: well-known RID (`-500/-501/-503/-504`), not English `Guest` (German `Gast`).
4. ASR: never `@($null)` then `$null -eq $ids`. Merge with a dictionary / `Add-MpPreference` per GUID. `-ErrorAction Stop`. Read back. Preserve unrelated rules.
5. Defender: enum names (`-CloudBlockLevel High`, `-PUAProtection Enabled`, `-EnableNetworkProtection Enabled`, `-MAPSReporting Advanced`). Preflight `Get-MpComputerStatus`. Read back; tamper/passive = FAIL not SUCCESS.
6. One Apply handler only: explicit `if` chain, `Apply &Harden`, real progress totals. Clear passwords on **all** exits. Fail closed on `net user /logonpasswordchg` after create (retry must not skip to baseline).
7. `Test-HardenPasswordPair`: also reject length < 8.
8. `test-syntax.ps1`: `exit 1` on parse errors (Task 0 / fold into Task 1). `test-harden.ps1` must exercise production bodies + `New-LocalUser` parameter metadata.
9. Amend `AGENTS.md` HKLM exception in Task 1, before Task 3 writes SmartScreen.
10. Spec threat model: HKLM blocked without admin password; HKCU and per-user EXE remain possible.

---

## File map

| File | Role |
|---|---|
| `test-harden.ps1` | Create. Pure validation tests (no admin, no GUI). |
| `Windows-PC-Setup.ps1` | Modify. Helpers, mutating functions, Harden tab, wire into `$TabControl`. |
| `test-syntax.ps1` | Unchanged command; run after every script edit. |
| `AGENTS.md` | Fifth tab + Harden HKLM exception. |
| `README.md` | Harden / standard-user workflow. |
| `docs/known.md` | SID groups, ASR merge, SmartScreen HKLM exception, do not Chrome-policy. |

---

### Task 1: Username/password validation helpers + tests

The main script auto-elevates and shows a GUI, so tests cannot dot-source it. Put the **same function bodies** in `test-harden.ps1` first, prove them, then paste identical bodies into `Windows-PC-Setup.ps1`. If they drift, `test-harden.ps1` is wrong.

**Files:**
- Create: `test-harden.ps1`
- Modify: `Windows-PC-Setup.ps1` (insert after `Show-ConfirmationDialog`, before Win32 bloatware)

**Interfaces:**
- Consumes: nothing
- Produces:
  - `Test-HardenUsername` `[string]$Name` → `[hashtable]@{ Ok = [bool]; Reason = [string] }`
  - `Test-HardenPasswordPair` `[string]$Password`, `[string]$Confirm` → `[hashtable]@{ Ok = [bool]; Reason = [string] }`
  - `$script:HardenReservedUsernames` = `@('Administrator','Guest','DefaultAccount','WDAGUtilityAccount','krbtgt')`

- [ ] **Step 1: Write `test-harden.ps1` with functions + assertions**

```powershell
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:HardenReservedUsernames = @(
    'Administrator', 'Guest', 'DefaultAccount', 'WDAGUtilityAccount', 'krbtgt'
)

function Test-HardenUsername {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return @{ Ok = $false; Reason = 'Username is empty' }
    }
    $trimmed = $Name.Trim()
    if ($trimmed.Length -gt 20) {
        return @{ Ok = $false; Reason = 'Username longer than 20 characters' }
    }
    if ($trimmed.EndsWith('.')) {
        return @{ Ok = $false; Reason = 'Username cannot end with a period' }
    }
    if ($trimmed -match '[/\\\[\]:;|=,+*?<>@"\s]') {
        return @{ Ok = $false; Reason = 'Username contains illegal characters or spaces' }
    }
    foreach ($reserved in $script:HardenReservedUsernames) {
        if ($trimmed -ieq $reserved) {
            return @{ Ok = $false; Reason = "Username '$trimmed' is a reserved built-in name" }
        }
    }
    return @{ Ok = $true; Reason = '' }
}

function Test-HardenPasswordPair {
    param([string]$Password, [string]$Confirm)
    if ([string]::IsNullOrEmpty($Password)) {
        return @{ Ok = $false; Reason = 'Password is empty' }
    }
    if ($Password.Length -lt 8) {
        return @{ Ok = $false; Reason = 'Password shorter than 8 characters' }
    }
    if ($Password -cne $Confirm) {
        return @{ Ok = $false; Reason = 'Password and confirmation do not match' }
    }
    return @{ Ok = $true; Reason = '' }
}

$failures = 0
function Assert-Harden($cond, $msg) {
    if (-not $cond) {
        Write-Host "FAIL: $msg" -ForegroundColor Red
        $script:failures++
    }
}

$ok = Test-HardenUsername 'jsmith'
Assert-Harden $ok.Ok 'jsmith should pass'

$bad = Test-HardenUsername ''
Assert-Harden (-not $bad.Ok) 'empty should fail'

$admin = Test-HardenUsername 'Administrator'
Assert-Harden (-not $admin.Ok) 'Administrator should fail'

$guest = Test-HardenUsername 'guest'
Assert-Harden (-not $guest.Ok) 'guest (case) should fail'

$space = Test-HardenUsername 'jan schmidt'
Assert-Harden (-not $space.Ok) 'spaces should fail'

$long = Test-HardenUsername ('a' * 21)
Assert-Harden (-not $long.Ok) '21 chars should fail'

$pw = Test-HardenPasswordPair 'x' 'y'
Assert-Harden (-not $pw.Ok) 'mismatch should fail'

$pw2 = Test-HardenPasswordPair 'x' 'x'
Assert-Harden $pw2.Ok 'match should pass'

$pw3 = Test-HardenPasswordPair '' ''
Assert-Harden (-not $pw3.Ok) 'empty password should fail'

$pw4 = Test-HardenPasswordPair '1234567' '1234567'
Assert-Harden (-not $pw4.Ok) '7-char password should fail'

$pw5 = Test-HardenPasswordPair '12345678' '12345678'
Assert-Harden $pw5.Ok '8-char password should pass'

if ($failures -eq 0) {
    Write-Host 'HARDEN HELPER TESTS PASSED' -ForegroundColor Green
    exit 0
}
Write-Host "HARDEN HELPER TESTS FAILED: $failures" -ForegroundColor Red
exit 1
```

- [ ] **Step 2: Run tests — expect PASS (functions are in the test file)**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: `HARDEN HELPER TESTS PASSED` and exit 0.

- [ ] **Step 3: Paste the same `$script:HardenReservedUsernames`, `Test-HardenUsername`, and `Test-HardenPasswordPair` into `Windows-PC-Setup.ps1` immediately after `Show-ConfirmationDialog` (after its closing `}`). Do not change bodies.**

- [ ] **Step 4: Syntax gate**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: `SYNTAX CHECK PASSED` and `HARDEN HELPER TESTS PASSED`.

- [ ] **Step 5: Commit (only if the operator asked)**

```powershell
git add test-harden.ps1 Windows-PC-Setup.ps1
git commit -m "Add Harden username and password validation helpers."
```

---

### Task 2: SID group lookup + create standard user

**Files:**
- Modify: `Windows-PC-Setup.ps1` (after Task 1 helpers)

**Interfaces:**
- Consumes: `Test-HardenUsername`, `Write-Log`, `$script:DryRun`
- Produces:
  - `Get-HardenLocalGroupName` `[string]$Sid` → `[string]` (group name) or throws if missing
  - `New-HardenStandardUser` `[string]$UserName`, `[string]$Password`, `[bool]$ChangePasswordAtLogon` → `[bool]` (`$true` success or skip-exists; `$false` failure)

- [ ] **Step 1: Add these functions after Task 1 helpers**

```powershell
function Get-HardenLocalGroupName {
    param([Parameter(Mandatory)][string]$Sid)
    $group = Get-LocalGroup | Where-Object { $_.SID.Value -eq $Sid }
    if ($null -eq $group) {
        throw "Local group with SID $Sid not found"
    }
    return $group.Name
}

function Test-HardenAccountIsStandardUser {
    param($User)
    if ($null -eq $User -or -not $User.Enabled) { return $false }
    if ($User.SID.Value -match '-50[0134]$') { return $false }
    $userSid = $User.SID.Value
    $inUsers = @(Get-LocalGroupMember -SID 'S-1-5-32-545' -ErrorAction Stop | Where-Object { $_.SID.Value -eq $userSid })
    $inAdmins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | Where-Object { $_.SID.Value -eq $userSid })
    return ($inUsers.Count -gt 0 -and $inAdmins.Count -eq 0)
}

function New-HardenStandardUser {
    param(
        [Parameter(Mandatory)][string]$UserName,
        [Parameter(Mandatory)][string]$Password,
        [bool]$ChangePasswordAtLogon = $false
    )

    $check = Test-HardenUsername -Name $UserName
    if (-not $check.Ok) {
        Write-Log "Refusing user create: $($check.Reason)" "ERROR"
        return $false
    }
    $name = $UserName.Trim()

    $existing = Get-LocalUser -Name $name -ErrorAction SilentlyContinue
    if ($existing) {
        if (-not (Test-HardenAccountIsStandardUser -User $existing)) {
            Write-Log "Existing account '$name' is not an enabled standard Users member (or is an Administrator / built-in). Refusing." "ERROR"
            return $false
        }
        Write-Log "Local user '$name' already exists and is a standard Users member — skipping create" "WARN"
        return $true
    }

    if ($script:DryRun) {
        Write-Log "Would create local user '$name' in Users (SID S-1-5-32-545)" "INFO"
        return $true
    }

    try {
        $secure = ConvertTo-SecureString -String $Password -AsPlainText -Force
        New-LocalUser -Name $name -Password $secure -AccountNeverExpires -PasswordNeverExpires:$false | Out-Null

        Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $name -ErrorAction Stop

        $created = Get-LocalUser -Name $name -ErrorAction Stop
        if (-not (Test-HardenAccountIsStandardUser -User $created)) {
            throw "Postcondition failed: '$name' is not an enabled standard Users member"
        }

        if ($ChangePasswordAtLogon) {
            $net = Start-Process -FilePath "$env:SystemRoot\System32\net.exe" -ArgumentList @('user', $name, '/logonpasswordchg:yes') -Wait -PassThru -WindowStyle Hidden
            if ($net.ExitCode -ne 0) {
                Write-Log "Created '$name' but failed to set change-password-at-logon (net.exe exit $($net.ExitCode))" "ERROR"
                return $false
            }
        }

        Write-Log "Created local standard user '$name' (Users SID S-1-5-32-545)" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to create local user '$name': $_" "ERROR"
        return $false
    }
}
```

Do **not** call `Remove-LocalGroupMember` on the currently logged-on admin. The new account is a different name.

`New-LocalUser` on Windows 10 5.1 has no `-ChangePasswordAtLogon`. Use `net user … /logonpasswordchg:yes` only when that checkbox is on.

- [ ] **Step 2: Syntax gate**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

Expected: `SYNTAX CHECK PASSED`.

- [ ] **Step 3: Commit (only if the operator asked)**

```powershell
git add Windows-PC-Setup.ps1
git commit -m "Add SID-safe standard local user creation for Harden."
```

---

### Task 3: Defender baseline + SmartScreen (no Chrome)

**Files:**
- Modify: `Windows-PC-Setup.ps1` (after Task 2)

**Interfaces:**
- Consumes: `Write-Log`, `$script:DryRun`
- Produces:
  - `$script:HardenAsrGuids` = the six spec GUIDs (string array, fixed order)
  - `Test-HardenDefenderAvailable` → `[bool]`
  - `Set-HardenDefenderCloud` → `[bool]`
  - `Set-HardenPua` → `[bool]`
  - `Set-HardenNetworkProtection` → `[bool]`
  - `Set-HardenAsrBaseline` → `[bool]`
  - `Set-HardenSmartScreenWarn` → `[bool]`

- [ ] **Step 1: Add constants and functions**

```powershell
$script:HardenAsrGuids = @(
    '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2',
    'e6db77e5-3df2-4cf1-b95a-636979351e5b',
    '56a863a9-875e-4185-98a7-b882c64b5ce5',
    'be9ba2d9-53ea-4cdc-84e5-9b1eeee46550',
    'd3e037e1-3eb8-44c8-a917-57927947596d',
    'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4'
)

function Test-HardenDefenderAvailable {
    return [bool](Get-Command Set-MpPreference -ErrorAction SilentlyContinue)
}

function Set-HardenDefenderCloud {
    if ($script:DryRun) {
        Write-Log "Would enable Defender real-time monitoring, MAPS Advanced, CloudBlockLevel High (2)" "INFO"
        return $true
    }
    if (-not (Test-HardenDefenderAvailable)) {
        Write-Log "Set-MpPreference not found — cannot apply Defender cloud/realtime" "ERROR"
        return $false
    }
    try {
        Set-MpPreference -DisableRealtimeMonitoring $false -MAPSReporting Advanced -CloudBlockLevel High
        Write-Log "Defender real-time on, cloud protection High" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed Defender cloud/realtime: $_" "ERROR"
        return $false
    }
}

function Set-HardenPua {
    if ($script:DryRun) {
        Write-Log "Would set PUAProtection Enabled (1)" "INFO"
        return $true
    }
    if (-not (Test-HardenDefenderAvailable)) {
        Write-Log "Set-MpPreference not found — cannot apply PUA" "ERROR"
        return $false
    }
    try {
        Set-MpPreference -PUAProtection Enabled
        Write-Log "PUA protection enabled" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed PUA protection: $_" "ERROR"
        return $false
    }
}

function Set-HardenNetworkProtection {
    if ($script:DryRun) {
        Write-Log "Would set EnableNetworkProtection Enabled (1)" "INFO"
        return $true
    }
    if (-not (Test-HardenDefenderAvailable)) {
        Write-Log "Set-MpPreference not found — cannot apply Network Protection" "ERROR"
        return $false
    }
    try {
        Set-MpPreference -EnableNetworkProtection Enabled
        Write-Log "Network Protection enabled" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed Network Protection: $_" "ERROR"
        return $false
    }
}

function Set-HardenAsrBaseline {
    if ($script:DryRun) {
        Write-Log "Would merge six ASR rules to Block (Enabled): $($script:HardenAsrGuids -join ', ')" "INFO"
        return $true
    }
    if (-not (Test-HardenDefenderAvailable)) {
        Write-Log "Set-MpPreference not found — cannot apply ASR" "ERROR"
        return $false
    }
    try {
        $pref = Get-MpPreference
        $rawIds = $pref.AttackSurfaceReductionRules_Ids
        $rawActions = $pref.AttackSurfaceReductionRules_Actions
        $ids = @()
        $actions = @()
        if ($null -ne $rawIds) { $ids = @($rawIds) }
        if ($null -ne $rawActions) { $actions = @($rawActions) }

        foreach ($guid in $script:HardenAsrGuids) {
            $idx = [array]::IndexOf(@($ids | ForEach-Object { $_.ToString().ToLower() }), $guid.ToLower())
            if ($idx -ge 0) {
                $actions[$idx] = 1
            }
            else {
                $ids += $guid
                $actions += 1
            }
        }

        Set-MpPreference -AttackSurfaceReductionRules_Ids $ids -AttackSurfaceReductionRules_Actions $actions
        Write-Log "ASR baseline merged (six rules Block); existing other rules kept" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed ASR baseline: $_" "ERROR"
        return $false
    }
}

function Set-HardenSmartScreenWarn {
    if ($script:DryRun) {
        Write-Log "Would set HKLM Policies\Microsoft\Windows\System EnableSmartScreen=1 ShellSmartScreenLevel=Warn" "INFO"
        return $true
    }
    try {
        $key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'
        if (-not (Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        Set-ItemProperty -Path $key -Name 'EnableSmartScreen' -Type DWord -Value 1
        Set-ItemProperty -Path $key -Name 'ShellSmartScreenLevel' -Type String -Value 'Warn'
        Write-Log "SmartScreen Warn (machine policy) applied" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed SmartScreen policy: $_" "ERROR"
        return $false
    }
}
```

Do not write any `Google\Chrome` keys. Do not set `CloudBlockLevel` 4 or 6. Do not enable Controlled Folder Access.

- [ ] **Step 2: Syntax gate**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

Expected: `SYNTAX CHECK PASSED`.

- [ ] **Step 3: Commit (only if the operator asked)**

```powershell
git add Windows-PC-Setup.ps1
git commit -m "Add Harden Defender, PUA, ASR merge, and SmartScreen Warn."
```

---

### Task 4: Harden tab UI (no Apply logic yet)

**Files:**
- Modify: `Windows-PC-Setup.ps1`
  - Optional: add `[int]$Height = 36` to `New-IntroText` and use it for the taller Harden intro (`$Height = 52`).
  - Insert a new section **after** `$TabInstallApps.Controls.Add($InstallAppsPanel)` and **before** `# TAB 4: SYSTEM REPAIR`. Renumber comments: Harden is tab 4, Repair becomes tab 5.
  - Store controls in `$script:` so the later click handler can read them:
    - `$script:ChkHardenCreateUser`
    - `$script:TxtHardenUser`
    - `$script:TxtHardenPassword`
    - `$script:TxtHardenPasswordConfirm`
    - `$script:ChkHardenMustChangePassword`
    - `$script:ChkHardenDefenderCloud`
    - `$script:ChkHardenPua`
    - `$script:ChkHardenSmartScreen`
    - `$script:ChkHardenNetwork`
    - `$script:ChkHardenAsr`
    - `$script:BtnApplyHarden`

**Interfaces:**
- Consumes: `New-IntroText`, `New-SectionHeader`, `$script:UI`, `$script:MainTooltip`
- Produces: `$TabHarden` (TabPage) ready to add to `$TabControl`

- [ ] **Step 1: Extend `New-IntroText` with optional height**

In `New-IntroText`, change `$intro.Size` to:

```powershell
    [int]$Height = 36
    # ... in param() block
    $intro.Size = New-Object System.Drawing.Size(610, $Height)
    return ($YPosition + $Height + 4)
```

Existing callers omit `$Height` and keep 36.

- [ ] **Step 2: Build `$TabHarden` UI**

```powershell
# ============================================================================
# TAB 4: HARDEN
# ============================================================================

$TabHarden = New-Object System.Windows.Forms.TabPage
$TabHarden.Text = "Harden"
$TabHarden.Padding = New-Object System.Windows.Forms.Padding(10)

$HardenPanel = New-Object System.Windows.Forms.Panel
$HardenPanel.Dock = "Fill"
$HardenPanel.AutoScroll = $true

$hardenY = New-IntroText -Panel $HardenPanel -Height 52 -YPosition 8 -Text "Install apps on this admin account first. Then create the daily standard user and apply machine-wide Defender settings. Sign in as that user to set up their desktop. PDF reader is on Install Apps. Chrome Workspace cloud-over-local is set in Google Admin, not here."

$hardenY = New-SectionHeader -Panel $HardenPanel -Text "Standard user" -YPosition $hardenY

$script:ChkHardenCreateUser = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenCreateUser.Text = "Create this user"
$script:ChkHardenCreateUser.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenCreateUser.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenCreateUser.Checked = $true
$script:MainTooltip.SetToolTip($script:ChkHardenCreateUser, "Creates a local Users-group account. Never an administrator. Off = apply Defender baseline only.")
$HardenPanel.Controls.Add($script:ChkHardenCreateUser)
$hardenY += $script:UI.ItemSpacing

$LblHardenUser = New-Object System.Windows.Forms.Label
$LblHardenUser.Text = "Username"
$LblHardenUser.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$LblHardenUser.Size = New-Object System.Drawing.Size(120, 20)
$HardenPanel.Controls.Add($LblHardenUser)
$script:TxtHardenUser = New-Object System.Windows.Forms.TextBox
$script:TxtHardenUser.Location = New-Object System.Drawing.Point(160, ($hardenY - 2))
$script:TxtHardenUser.Size = New-Object System.Drawing.Size(280, 22)
$script:TxtHardenUser.Text = ""
$script:MainTooltip.SetToolTip($script:TxtHardenUser, "Local account name you choose. No default. Max 20 characters, no spaces.")
$HardenPanel.Controls.Add($script:TxtHardenUser)
$hardenY += $script:UI.ItemSpacing

$LblHardenPw = New-Object System.Windows.Forms.Label
$LblHardenPw.Text = "Password"
$LblHardenPw.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$LblHardenPw.Size = New-Object System.Drawing.Size(120, 20)
$HardenPanel.Controls.Add($LblHardenPw)
$script:TxtHardenPassword = New-Object System.Windows.Forms.TextBox
$script:TxtHardenPassword.Location = New-Object System.Drawing.Point(160, ($hardenY - 2))
$script:TxtHardenPassword.Size = New-Object System.Drawing.Size(280, 22)
$script:TxtHardenPassword.UseSystemPasswordChar = $true
$HardenPanel.Controls.Add($script:TxtHardenPassword)
$hardenY += $script:UI.ItemSpacing

$LblHardenPw2 = New-Object System.Windows.Forms.Label
$LblHardenPw2.Text = "Confirm password"
$LblHardenPw2.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$LblHardenPw2.Size = New-Object System.Drawing.Size(120, 20)
$HardenPanel.Controls.Add($LblHardenPw2)
$script:TxtHardenPasswordConfirm = New-Object System.Windows.Forms.TextBox
$script:TxtHardenPasswordConfirm.Location = New-Object System.Drawing.Point(160, ($hardenY - 2))
$script:TxtHardenPasswordConfirm.Size = New-Object System.Drawing.Size(280, 22)
$script:TxtHardenPasswordConfirm.UseSystemPasswordChar = $true
$HardenPanel.Controls.Add($script:TxtHardenPasswordConfirm)
$hardenY += $script:UI.ItemSpacing

$script:ChkHardenMustChangePassword = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenMustChangePassword.Text = "User must change password at next logon"
$script:ChkHardenMustChangePassword.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenMustChangePassword.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenMustChangePassword.Checked = $false
$script:MainTooltip.SetToolTip($script:ChkHardenMustChangePassword, "Leave off so you can sign in as this user and set up their desktop.")
$HardenPanel.Controls.Add($script:ChkHardenMustChangePassword)
$hardenY += $script:UI.SectionSpacing

$hardenY = New-SectionHeader -Panel $HardenPanel -Text "Microsoft baseline" -YPosition $hardenY

$script:ChkHardenDefenderCloud = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenDefenderCloud.Text = "Defender real-time + cloud protection (High)"
$script:ChkHardenDefenderCloud.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenDefenderCloud.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenDefenderCloud.Checked = $true
$HardenPanel.Controls.Add($script:ChkHardenDefenderCloud)
$hardenY += $script:UI.ItemSpacing

$script:ChkHardenPua = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenPua.Text = "Potentially unwanted app blocking (PUA)"
$script:ChkHardenPua.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenPua.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenPua.Checked = $true
$HardenPanel.Controls.Add($script:ChkHardenPua)
$hardenY += $script:UI.ItemSpacing

$script:ChkHardenSmartScreen = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenSmartScreen.Text = "SmartScreen (Warn)"
$script:ChkHardenSmartScreen.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenSmartScreen.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenSmartScreen.Checked = $true
$HardenPanel.Controls.Add($script:ChkHardenSmartScreen)
$hardenY += $script:UI.ItemSpacing

$script:ChkHardenNetwork = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenNetwork.Text = "Network Protection"
$script:ChkHardenNetwork.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenNetwork.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenNetwork.Checked = $true
$HardenPanel.Controls.Add($script:ChkHardenNetwork)
$hardenY += $script:UI.ItemSpacing

$script:ChkHardenAsr = New-Object System.Windows.Forms.CheckBox
$script:ChkHardenAsr.Text = "Attack surface reduction (six conservative rules)"
$script:ChkHardenAsr.Location = New-Object System.Drawing.Point($script:UI.CheckboxIndent, $hardenY)
$script:ChkHardenAsr.Size = New-Object System.Drawing.Size(600, $script:UI.CheckboxHeight)
$script:ChkHardenAsr.Checked = $true
$script:MainTooltip.SetToolTip($script:ChkHardenAsr, "LSASS credential steal, WMI persistence, vulnerable drivers, email/webmail executables, script-launched downloads, unsigned USB executables. Does not block Office or Acrobat child processes.")
$HardenPanel.Controls.Add($script:ChkHardenAsr)
$hardenY += 40

$script:BtnApplyHarden = New-Object System.Windows.Forms.Button
$script:BtnApplyHarden.Text = "Apply &Harden"
$script:BtnApplyHarden.Location = New-Object System.Drawing.Point(480, $hardenY)
$script:BtnApplyHarden.Size = New-Object System.Drawing.Size(150, $script:UI.ButtonHeight)
$script:BtnApplyHarden.BackColor = $script:UI.AccentColor
$script:BtnApplyHarden.ForeColor = [System.Drawing.Color]::White
$script:BtnApplyHarden.FlatStyle = "Flat"
Add-ButtonHoverEffect -Button $script:BtnApplyHarden
$script:MainTooltip.SetToolTip($script:BtnApplyHarden, "Create the user if selected, then apply checked Defender settings (Alt+A is Settings; this is Alt+P via &Apply)")
$HardenPanel.Controls.Add($script:BtnApplyHarden)

$TabHarden.Controls.Add($HardenPanel)
```

Use `&Apply Harden` so Alt+A does not steal Settings’ shortcut. Settings already owns Alt+A. The button text `&Apply Harden` uses Alt+A in WinForms if A is the first `&` — **conflict**. Set button text to `Apply &Harden` so the shortcut is Alt+H.

Do **not** add the click handler yet (Task 5). Leave `$script:BtnApplyHarden` with hover only.

- [ ] **Step 3: Wire the tab (partial)**

Change

```powershell
$TabControl.Controls.AddRange(@($TabBloatware, $TabSettings, $TabInstallApps, $TabRepair))
```

to

```powershell
$TabControl.Controls.AddRange(@($TabBloatware, $TabSettings, $TabInstallApps, $TabHarden, $TabRepair))
```

- [ ] **Step 4: Syntax gate**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
```

Expected: `SYNTAX CHECK PASSED`.

- [ ] **Step 5: Commit (only if the operator asked)**

```powershell
git add Windows-PC-Setup.ps1
git commit -m "Add Harden tab layout after Install Apps."
```

---

### Task 5: Apply handler, confirm dialog, stop-before-baseline

**Files:**
- Modify: `Windows-PC-Setup.ps1` — `Show-ConfirmationDialog` and `$script:BtnApplyHarden.Add_Click`

**Interfaces:**
- Consumes: all `$script:ChkHarden*` / `$script:TxtHarden*` / functions from Tasks 1–3
- Produces: Apply Harden click behavior

- [ ] **Step 1: Add optional `-Detail` to `Show-ConfirmationDialog`**

Replace the function body so existing callers stay valid:

```powershell
function Show-ConfirmationDialog {
    param(
        [string]$Title,
        [int]$SelectedCount,
        [string]$ActionType,
        [bool]$IsDryRun,
        [string]$Detail = ""
    )

    $dryRunNote = if ($IsDryRun) { "`n`n[DRY RUN MODE - No actual changes will be made]" } else { "" }
    if ($Detail) {
        $message = "$Detail$dryRunNote`n`nDo you want to continue?"
    }
    else {
        $message = "You are about to $ActionType $SelectedCount item(s).$dryRunNote`n`nDo you want to continue?"
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        $message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}
```

- [ ] **Step 2: Add the click handler on `$script:BtnApplyHarden` (after the button is created)**

```powershell
$script:BtnApplyHarden.Add_Click({
    $script:DryRun = $script:ChkDryRun.Checked

    $createUser = $script:ChkHardenCreateUser.Checked
    $userName = $script:TxtHardenUser.Text
    $password = $script:TxtHardenPassword.Text
    $confirm = $script:TxtHardenPasswordConfirm.Text

    if ($createUser) {
        $u = Test-HardenUsername -Name $userName
        if (-not $u.Ok) {
            [System.Windows.Forms.MessageBox]::Show($u.Reason, "Harden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $p = Test-HardenPasswordPair -Password $password -Confirm $confirm
        if (-not $p.Ok) {
            [System.Windows.Forms.MessageBox]::Show($p.Reason, "Harden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
    }

    $lines = New-Object System.Collections.Generic.List[string]
    if ($createUser) {
        $lines.Add("Create standard user: $($userName.Trim())")
        if ($script:ChkHardenMustChangePassword.Checked) {
            $lines.Add("Must change password at next logon: yes")
        }
    }
    else {
        $lines.Add("Skip user create")
    }
    if ($script:ChkHardenDefenderCloud.Checked) { $lines.Add("Defender real-time + cloud High") }
    if ($script:ChkHardenPua.Checked) { $lines.Add("PUA") }
    if ($script:ChkHardenSmartScreen.Checked) { $lines.Add("SmartScreen Warn") }
    if ($script:ChkHardenNetwork.Checked) { $lines.Add("Network Protection") }
    if ($script:ChkHardenAsr.Checked) { $lines.Add("ASR six-rule baseline") }

    if ($lines.Count -eq 0 -or ($lines.Count -eq 1 -and -not $createUser -and -not $script:ChkHardenDefenderCloud.Checked -and -not $script:ChkHardenPua.Checked -and -not $script:ChkHardenSmartScreen.Checked -and -not $script:ChkHardenNetwork.Checked -and -not $script:ChkHardenAsr.Checked)) {
        [System.Windows.Forms.MessageBox]::Show("Nothing selected to apply.", "Harden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $detail = "You are about to apply:`n`n" + ($lines -join "`n")
    $selectedCount = $lines.Count
    if (-not (Show-ConfirmationDialog -Title "Confirm Harden" -SelectedCount $selectedCount -ActionType "apply" -IsDryRun $script:DryRun -Detail $detail)) {
        return
    }

    Set-ButtonDisabled -Button $script:BtnApplyHarden -WorkingText "Applying..."
    $successCount = 0
    $failCount = 0

    if ($script:DryRun) {
        Write-Log "=== DRY RUN MODE - No changes will be made ===" "INFO"
    }

    try {
        if ($createUser) {
            Update-Progress -Current 1 -Total 6 -CurrentItem "Local user"
            if (-not (New-HardenStandardUser -UserName $userName.Trim() -Password $password -ChangePasswordAtLogon:$script:ChkHardenMustChangePassword.Checked)) {
                Hide-Progress
                Update-Status "Harden stopped: user create failed. Baseline not applied."
                [System.Windows.Forms.MessageBox]::Show(
                    "User create failed. Defender baseline was not applied. Fix the name/password and retry.`n`nLog: $script:LogPath",
                    "Harden",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            $successCount++
        }

        $steps = @()
        if ($script:ChkHardenDefenderCloud.Checked) { $steps += @{ Name = 'Defender cloud'; Fn = { Set-HardenDefenderCloud } } }
        if ($script:ChkHardenPua.Checked) { $steps += @{ Name = 'PUA'; Fn = { Set-HardenPua } } }
        if ($script:ChkHardenSmartScreen.Checked) { $steps += @{ Name = 'SmartScreen'; Fn = { Set-HardenSmartScreenWarn } } }
        if ($script:ChkHardenNetwork.Checked) { $steps += @{ Name = 'Network Protection'; Fn = { Set-HardenNetworkProtection } } }
        if ($script:ChkHardenAsr.Checked) { $steps += @{ Name = 'ASR'; Fn = { Set-HardenAsrBaseline } } }

        $i = 0
        foreach ($step in $steps) {
            $i++
            Update-Progress -Current $i -Total $steps.Count -CurrentItem $step.Name
            if ((& $step.Fn)) { $successCount++ } else { $failCount++ }
        }

        Hide-Progress
        $dryRunMsg = if ($script:DryRun) { "[DRY RUN] " } else { "" }
        Update-Status "${dryRunMsg}Harden complete. Success: $successCount, Failed: $failCount"
        [System.Windows.Forms.MessageBox]::Show(
            "${dryRunMsg}Harden complete.`n`nSuccessful: $successCount`nFailed: $failCount`n`nLog: $script:LogPath",
            "Harden",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        Write-Log "Error during Harden: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Harden", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        Hide-Progress
        Set-ButtonEnabled -Button $script:BtnApplyHarden
        $script:DryRun = $false
        $script:TxtHardenPassword.Text = ""
        $script:TxtHardenPasswordConfirm.Text = ""
    }
})
```

Clear password boxes in `finally` so the GUI does not keep the password after Apply. Do not `Write-Log` the password.

WinForms scriptblocks: `$step.Fn` scriptblocks inside `Add_Click` can lose `$script:` functions. **Do not use the hashtable-of-scriptblocks pattern if it fails StrictMode.** Prefer an explicit `if` chain:

```powershell
        if ($script:ChkHardenDefenderCloud.Checked) {
            Update-Progress -Current 1 -Total 5 -CurrentItem "Defender cloud"
            if (Set-HardenDefenderCloud) { $successCount++ } else { $failCount++ }
        }
        if ($script:ChkHardenPua.Checked) {
            if (Set-HardenPua) { $successCount++ } else { $failCount++ }
        }
        if ($script:ChkHardenSmartScreen.Checked) {
            if (Set-HardenSmartScreenWarn) { $successCount++ } else { $failCount++ }
        }
        if ($script:ChkHardenNetwork.Checked) {
            if (Set-HardenNetworkProtection) { $successCount++ } else { $failCount++ }
        }
        if ($script:ChkHardenAsr.Checked) {
            if (Set-HardenAsrBaseline) { $successCount++ } else { $failCount++ }
        }
```

Use this explicit chain in the actual click handler, not `$step.Fn`.

Empty selection: if create-user is off and all baseline boxes are off, show “Nothing selected” and return.

- [ ] **Step 3: Syntax + helper tests**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: both PASS.

- [ ] **Step 4: Commit (only if the operator asked)**

```powershell
git add Windows-PC-Setup.ps1
git commit -m "Wire Harden Apply with confirm, user-first stop, and password clear."
```

---

### Task 6: Docs (AGENTS, README, known)

**Files:**
- Modify: `AGENTS.md`, `README.md`, `docs/known.md`
- Modify: `Windows-PC-Setup.ps1` comment header `.DESCRIPTION` to mention Harden

**Interfaces:** none

- [ ] **Step 1: `AGENTS.md` — replace the tabs list and add Harden to hard rules**

Tabs:

```
1. Remove Bloatware (AppX + Win32)
2. Settings (taskbar, Start, power, Explorer)
3. Install Apps (winget)
4. Harden (standard local user + Defender/SmartScreen/ASR)
5. Repair (DISM / SFC / CHKDSK)
```

Add under Hard rules:

```
- Harden may use Defender cmdlets and HKLM SmartScreen (`Policies\Microsoft\Windows\System` EnableSmartScreen + ShellSmartScreenLevel=Warn) only. Still no HKLM Start policy, no IsEducationEnvironment, no Chrome policies.
- Local Users/Administrators groups: resolve by SID (S-1-5-32-545 / S-1-5-32-544), never English names.
```

- [ ] **Step 2: `README.md` — add a short Harden bullet under Features and a workflow sentence under How to Run**

Features bullet: `**Harden** - Create a standard local user and apply conservative Defender / SmartScreen / ASR settings`

How to Run note: `On a clean PC: finish admin account, Windows Update, and drivers, run this tool (install apps while still admin), Harden, then sign in as the new standard user.`

- [ ] **Step 3: `docs/known.md` — append traps**

```
- **Harden groups are SIDs.** Users `S-1-5-32-545`, Administrators `S-1-5-32-544`. German Windows names are `Benutzer` / `Administratoren`.
- **Harden ASR merge.** Set the six GUIDs to Block without wiping other ASR rules already on the device.
- **Harden SmartScreen** is HKLM `Policies\Microsoft\Windows\System` Warn, not Block, not Chrome.
- **Do not write Chrome policies from PowerSetup.** Workspace cloud-over-local belongs in Google Admin. Fake PDF bundlers win by writing machine Chrome policy.
- **New-LocalUser has no -ChangePasswordAtLogon on Windows PowerShell 5.1.** Use `net user <name> /logonpasswordchg:yes` when that checkbox is on.
```

- [ ] **Step 4: Synopsis DESCRIPTION in `Windows-PC-Setup.ps1`** — add `- Create a standard local user and apply a Defender baseline (Harden tab)`

- [ ] **Step 5: Syntax gate**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: both PASS.

- [ ] **Step 6: Commit (only if the operator asked)**

```powershell
git add AGENTS.md README.md docs/known.md Windows-PC-Setup.ps1
git commit -m "Document Harden tab, SID groups, and SmartScreen HKLM exception."
```

---

### Task 7: Final gate (no host apply)

**Files:** none new

- [ ] **Step 1: Run both test files**

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-syntax.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-harden.ps1
```

Expected: `SYNTAX CHECK PASSED`, `HARDEN HELPER TESTS PASSED`.

- [ ] **Step 2: Grep the script for forbidden Chrome policy writes**

```powershell
powershell.exe -NoProfile -Command "Select-String -Path Windows-PC-Setup.ps1 -Pattern 'Policies\\Google\\Chrome','IsEducationEnvironment' "
```

Expected: no Harden-era matches (existing script should still have no `IsEducationEnvironment`; Chrome policy path must not appear).

- [ ] **Step 3: Do not run `Windows-PC-Setup.ps1` on this checkout host.** Behavioral proof is a disposable VM (German UI): Dry Run creates nobody; live create makes a Users-only account; ASR six GUIDs present; SmartScreen Warn.

---

## Spec coverage (self-review)

| Spec item | Task |
|---|---|
| Fifth tab after Install Apps | 4, AddRange in 4 |
| Create user fields, no canned names, create checkbox default on, must-change default off | 4 |
| Baseline five checkboxes default on | 4 |
| SID Users/Admins, never demote current admin, never add new user to Admins | 2 |
| Refuse empty/illegal/reserved; exists → skip | 1, 2 |
| Dry-run user log line | 2 |
| Password policy error from `New-LocalUser` | 2 catch |
| Defender missing → that block fails, user can still succeed if create already done | 3, 5 (create first; defender functions return false) |
| Cloud High, PUA 1, Network 1, ASR six GUIDs merge | 3 |
| SmartScreen Warn HKLM exception | 3 |
| Confirm lists username, no password | 5 |
| Stop baseline if create requested and fails | 5 |
| No Chrome / no PUP list / no SAC / CFA / Office ASR | 3, 6 |
| test-syntax; do not run GUI on daily driver | 7 |
| AGENTS/README/known | 6 |
| `net user /logonpasswordchg` for 5.1 | 2, known.md |

**Gap closed in plan (not explicit in spec UI but required):** password boxes cleared after Apply; Alt+H not Alt+A; ASR merge vs replace; German group SIDs.
