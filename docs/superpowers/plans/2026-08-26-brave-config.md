# Brave Config Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship v2.5 Install Apps Brave section: HKCU policies (Google, new tab, no Brave autofill, download prompt, Downloads folder, Proton Pass force-install) on one button, profile desktop shortcuts on a second button.

**Architecture:** Same single-file WinForms script. Pure helpers first (path, search map, JSON profiles, shortcut names), then DryRun mutators, then two buttons on the Install Apps tab. Install Selected stays winget-only.

**Tech Stack:** Windows PowerShell 5.1, WinForms, WScript.Shell shortcuts, `test-syntax.ps1`, new `test-brave.ps1`

**Spec:** `docs/superpowers/specs/2026-08-26-brave-config-design.md`

## Global Constraints

- ASCII-only new log/UI strings in `Windows-PC-Setup.ps1` (em-dash AST trap).
- Dry-run on every mutator. Tests must not write `HKCU:\SOFTWARE\Policies\BraveSoftware`.
- Do not run `Windows-PC-Setup.ps1` on this host.
- No Preferences / Secure Preferences JSON. No HKLM Brave. No Chrome policies. No icon-hiding policies.
- Install Selected does not call Brave helpers.
- Shortcuts button does not call `Set-BraveHkcuPolicies`.
- Proton Pass CWS id is `ghmbeldphafepmbegfdlkpapadhbakde`.
- Gates: `test-syntax.ps1` then `test-brave.ps1`. No commit unless asked.

## Files

- Create: `test-brave.ps1`
- Modify: `AGENTS.md` (v2.5, HKCU Brave allowlist, gates line)
- Modify: `Windows-PC-Setup.ps1` (helpers after Harden block, Install Apps UI after Install Selected, version strings)
- Modify: `docs/known.md`, `README.md`, `.claude/CLAUDE.md`

---

### Task 1: Failing tests + AGENTS allowlist

**Files:**
- Modify: `AGENTS.md`
- Create: `test-brave.ps1`

**Interfaces:**
- Produces required AST names listed in Task 2.

- [ ] **Step 1: Amend AGENTS.md**

Version `v2.5`. Gates paragraph: also `test-brave.ps1`. Hard rules registry sentence: keep "No Chrome policies." Add: Brave Apply may write only `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave` (DWORD/SZ values named in the spec payload table) and numeric REG_SZ values under `HKCU:\SOFTWARE\Policies\BraveSoftware\Brave\ExtensionInstallForcelist`. Tabs line 3: Install Apps (winget + optional Brave policies and profile shortcuts). Note: HKCU Brave config does not follow a later standard user.

- [ ] **Step 2: Create `test-brave.ps1`** modeled on `test-harden.ps1` (same `Assert-Harden` renamed `Assert-Brave`, same `Import-HardenFunctionFromScript` copied as `Import-BraveFunctionFromScript`).

```powershell
#Requires -Version 5.1
# Tests production Brave helpers by extracting function ASTs from Windows-PC-Setup.ps1.
# Does not run the GUI script. Does not write HKCU Brave policy.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot "Windows-PC-Setup.ps1"
$failures = 0

function Assert-Brave {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        Write-Host "FAIL: $Message" -ForegroundColor Red
        $script:failures++
    }
}

function Import-BraveFunctionFromScript {
    param([string]$Path, [string]$FunctionName)
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw "Parse errors in ${Path}: $($parseErrors[0].Message)"
    }
    $fn = $ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $FunctionName
        }, $true) | Select-Object -First 1
    if ($null -eq $fn) {
        throw "Function $FunctionName not found in $Path"
    }
    $define = [scriptblock]::Create("function script:$FunctionName $($fn.Body.Extent.Text)")
    . $define
}

$scriptText = Get-Content -Path $scriptPath -Raw
Assert-Brave ($scriptText -notmatch '(?i)HKLM:\\SOFTWARE\\Policies\\BraveSoftware') 'must not write HKLM Brave policies'
Assert-Brave ($scriptText -match '(?i)HKCU:\\SOFTWARE\\Policies\\BraveSoftware\\Brave') 'must write HKCU Brave policies'
Assert-Brave ($scriptText -match 'ExtensionInstallForcelist') 'must use ExtensionInstallForcelist subkey'
Assert-Brave ($scriptText -notmatch "Google\\Chrome") 'must not write Google Chrome policies'
Assert-Brave ($scriptText -notmatch 'Secure Preferences') 'must not touch Secure Preferences'
Assert-Brave ($scriptText -notmatch 'show_brave_rewards_button') 'must not hide Rewards via Preferences'
Assert-Brave ($scriptText -notmatch 'BraveWalletDisabled') 'must not use icon-proxy feature policies'
Assert-Brave ($scriptText -notmatch 'lmjegmlicamnimmfhcmpkclmigmmcbeh') 'must not force-install Drive launcher'
Assert-Brave ($scriptText -match 'ghmbeldphafepmbegfdlkpapadhbakde') 'Proton Pass CWS id must be present'

$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$parseErrors)
$fnNames = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | ForEach-Object { $_.Name })
$requiredFns = @(
    'Get-BraveInstallPath',
    'Test-BraveInstalled',
    'Get-BraveUserDataDir',
    'Get-BraveDownloadsFolder',
    'Get-BravePolicyPath',
    'Get-BraveSearchPolicyMap',
    'ConvertTo-BraveShortcutFileName',
    'Get-BraveUniqueShortcutNames',
    'Get-BraveProfilesForShortcuts',
    'Set-BraveHkcuPolicies',
    'New-BraveProfileShortcuts'
)
foreach ($need in $requiredFns) {
    Assert-Brave ($fnNames -contains $need) "$need must exist in the PowerShell 5.1 AST"
}

Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BravePolicyPath'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveSearchPolicyMap'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'ConvertTo-BraveShortcutFileName'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveUniqueShortcutNames'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveProfilesForShortcuts'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveUserDataDir'

Assert-Brave ((Get-BravePolicyPath) -eq 'HKCU:\SOFTWARE\Policies\BraveSoftware\Brave') 'policy path is HKCU Brave'
$map = Get-BraveSearchPolicyMap
Assert-Brave ($map.DefaultSearchProviderEnabled -eq 1) 'search enabled'
Assert-Brave ($map.DefaultSearchProviderName -eq 'Google') 'search name Google'
Assert-Brave ($map.DefaultSearchProviderKeyword -eq 'google.com') 'search keyword'
Assert-Brave ($map.DefaultSearchProviderSuggestURL -match 'google\.com/complete') 'suggest URL is Google'
Assert-Brave ($map.DefaultSearchProviderSearchURL -match 'google\.com/search') 'search URL is Google'
Assert-Brave ($map.DefaultSearchProviderSearchURL -match '\{searchTerms\}') 'search URL has searchTerms'

Assert-Brave ((ConvertTo-BraveShortcutFileName 'Work') -eq 'Brave Work') 'plain name'
Assert-Brave ((ConvertTo-BraveShortcutFileName 'A:B*C') -eq 'Brave A_B_C') 'illegal chars'
Assert-Brave ((ConvertTo-BraveShortcutFileName '  ') -eq 'Brave Profile') 'blank becomes Brave Profile'

$json = '{"profile":{"info_cache":{"Default":{"name":"Work"},"Profile 1":{"name":"YouTube"},"Guest Profile":{"name":"Guest"},"System Profile":{"name":"System"}}}}'
$profs = @(Get-BraveProfilesForShortcuts -LocalStateJson $json)
Assert-Brave ($profs.Count -eq 2) 'skip Guest and System'
Assert-Brave ($profs[0].Id -eq 'Default' -and $profs[0].Name -eq 'Work') 'Default id and name'
Assert-Brave ($profs[1].Id -eq 'Profile 1' -and $profs[1].Name -eq 'YouTube') 'Profile 1 id and name'
Assert-Brave (@(Get-BraveProfilesForShortcuts -LocalStateJson '{}').Count -eq 0) 'empty object is no profiles'
Assert-Brave (@(Get-BraveProfilesForShortcuts -LocalStateJson '{"profile":{}}').Count -eq 0) 'missing info_cache is no profiles'

$collide = @(
    [pscustomobject]@{ Id = 'Default'; Name = 'Work_Profile 2' },
    [pscustomobject]@{ Id = 'Profile 1'; Name = 'Work' },
    [pscustomobject]@{ Id = 'Profile 2'; Name = 'Work' }
)
$names = @(Get-BraveUniqueShortcutNames -Profiles $collide | ForEach-Object { $_.FileName })
Assert-Brave (($names | Select-Object -Unique).Count -eq 3) 'collision names are unique'

$ud = Get-BraveUserDataDir
Assert-Brave ($ud -match 'Brave-Browser\\User Data$') 'user data dir suffix'

if ($failures -gt 0) {
    Write-Host "Brave tests FAILED: $failures" -ForegroundColor Red
    exit 1
}
Write-Host "Brave tests OK" -ForegroundColor Green
exit 0
```

- [ ] **Step 3: Run tests, expect FAIL** because the functions are missing.

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File test-brave.ps1
```

Expected: FAIL `Get-BraveInstallPath must exist in the PowerShell 5.1 AST` (or first missing name).

---

### Task 2: Pure helpers + DryRun mutators

**Files:**
- Modify: `Windows-PC-Setup.ps1` immediately after the Harden helper block (after `Set-HardenSmartAppControlOff`)

**Interfaces:**
- Consumes: `$script:DryRun`, `Write-Log`
- Produces the ten functions in Task 1.

- [ ] **Step 1: Implement the pure functions** (ASCII strings).

```powershell
function Get-BravePolicyPath {
    return 'HKCU:\SOFTWARE\Policies\BraveSoftware\Brave'
}

function Get-BraveUserDataDir {
    return (Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\User Data')
}

function Get-BraveSearchPolicyMap {
    return @{
        DefaultSearchProviderEnabled   = 1
        DefaultSearchProviderName      = 'Google'
        DefaultSearchProviderKeyword   = 'google.com'
        DefaultSearchProviderSearchURL = 'https://www.google.com/search?q={searchTerms}'
        DefaultSearchProviderSuggestURL = 'https://www.google.com/complete/search?output=chrome&q={searchTerms}'
    }
}

function ConvertTo-BraveShortcutFileName {
    param([string]$ProfileName)
    $base = if ([string]::IsNullOrWhiteSpace($ProfileName)) { '' } else { $ProfileName.Trim() }
    foreach ($ch in @('<', '>', ':', '"', '/', '\', '|', '?', '*')) {
        $base = $base.Replace($ch, '_')
    }
    $base = $base -replace '[\x00-\x1F]', '_'
    $base = $base.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'Profile' }
    return "Brave $base"
}

function Get-BraveProfilesForShortcuts {
    param([string]$LocalStateJson)
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($LocalStateJson)) { return @() }
    try {
        $obj = $LocalStateJson | ConvertFrom-Json
    } catch {
        return @()
    }
    $profileProp = $obj.PSObject.Properties['profile']
    if (-not $profileProp -or $null -eq $profileProp.Value) { return @() }
    $cacheProp = $profileProp.Value.PSObject.Properties['info_cache']
    if (-not $cacheProp -or $null -eq $cacheProp.Value) { return @() }
    foreach ($prop in $cacheProp.Value.PSObject.Properties) {
        $id = $prop.Name
        if ($id -eq 'Guest Profile' -or $id -eq 'System Profile') { continue }
        $name = $id
        if ($prop.Value -and $prop.Value.PSObject.Properties['name'] -and -not [string]::IsNullOrWhiteSpace($prop.Value.name)) {
            $name = [string]$prop.Value.name
        }
        $out.Add([pscustomobject]@{ Id = $id; Name = $name })
    }
    return @($out.ToArray())
}

function Get-BraveUniqueShortcutNames {
    param($Profiles)
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($p in @($Profiles)) {
        $file = ConvertTo-BraveShortcutFileName $p.Name
        if ($used.Contains($file)) {
            $file = ConvertTo-BraveShortcutFileName ("{0}_{1}" -f $p.Name, ($p.Id -replace ' ', '_'))
        }
        $n = 2
        $candidate = $file
        while ($used.Contains($candidate)) {
            $candidate = ConvertTo-BraveShortcutFileName ("{0}_{1}_{2}" -f $p.Name, ($p.Id -replace ' ', '_'), $n)
            $n++
        }
        [void]$used.Add($candidate)
        $out.Add([pscustomobject]@{ Id = $p.Id; Name = $p.Name; FileName = $candidate })
    }
    return @($out.ToArray())
}
```

`Get-BraveInstallPath`: check, in order, `Join-Path ${env:ProgramFiles} 'BraveSoftware\Brave-Browser\Application\brave.exe'`, the same under `${env:ProgramFiles(x86)}`, then `Join-Path $env:LOCALAPPDATA 'BraveSoftware\Brave-Browser\Application\brave.exe'`. Return the first existing path, else `''`.

`Test-BraveInstalled`: `[bool](-not [string]::IsNullOrWhiteSpace((Get-BraveInstallPath)))`.

`Get-BraveDownloadsFolder`: read `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders` name `{374DE290-123F-4565-9164-39C4925E467B}` (expand REG_EXPAND_SZ). If missing, `Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'`.

- [ ] **Step 2: Implement mutators** that return `$true` only if every intended write succeeded or was DryRun.

`Set-BraveHkcuPolicies`: if `$script:DryRun`, log each would-set line and return `$true`. Else `New-Item -Force` the policy path.

Search group: write the four SZ values from `Get-BraveSearchPolicyMap` (Name, Keyword, SearchURL, SuggestURL) first; only then DWORD `DefaultSearchProviderEnabled=1`. If any of those SZ writes fail, skip Enabled.

Other DWORDs: `RestoreOnStartup=5`, `PasswordManagerEnabled=0`, `AutofillAddressEnabled=0`, `AutofillCreditCardEnabled=0`, `PromptForDownloadLocation=1`. `DefaultDownloadDirectory` = `Get-BraveDownloadsFolder` (REG_SZ). Do not create `DownloadDirectory`. Do not write `RestoreOnStartupURLs`.

`ExtensionInstallForcelist`: `New-Item -Force` the child key. Read existing numeric values. If any contains `ghmbeldphafepmbegfdlkpapadhbakde`, leave them. Else `Set-ItemProperty` on the first unused name `1`,`2`,... with `ghmbeldphafepmbegfdlkpapadhbakde;https://clients2.google.com/service/update2/crx`. Never overwrite a different id. Never delete siblings.

After each write, read back; log SUCCESS or ERROR; `$ok` starts `$true` and becomes `$false` on any mismatch. Return `$ok`.

`New-BraveProfileShortcuts`: params `[string]$DesktopDir`, `[string]$BraveExe`, `$Profiles`. DryRun: log each would-create path from `Get-BraveUniqueShortcutNames` and return `$true` without writing. If `$BraveExe` is missing or not a file, log ERROR and return `$false`. For each item from `Get-BraveUniqueShortcutNames`, path `Join-Path $DesktopDir ($item.FileName + '.lnk')`. `WScript.Shell` CreateShortcut: TargetPath `$BraveExe`, Arguments ``--profile-directory="$($item.Id)"``, WorkingDirectory `(Split-Path $BraveExe)`, IconLocation `"$BraveExe,0"`, Save(). If Save throws, FAIL that item and continue. Return `$false` if any failed or list empty (empty list: log WARN, return `$false`).

- [ ] **Step 3: Stub `Write-Log` in `test-brave.ps1` and DryRun-test the mutators** after the pure asserts (import `Set-BraveHkcuPolicies` and `New-BraveProfileShortcuts`). Do not call them with `$script:DryRun = $false`.

```powershell
function Write-Log { param($Message, $Level = 'INFO') }
$script:DryRun = $true
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveInstallPath'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveDownloadsFolder'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Set-BraveHkcuPolicies'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'New-BraveProfileShortcuts'
Assert-Brave (Set-BraveHkcuPolicies) 'DryRun Set-BraveHkcuPolicies returns true'
$dir = Join-Path $env:TEMP 'BraveShortcutTest'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$fakeExe = Join-Path $dir 'brave.exe'
Set-Content -Path $fakeExe -Value 'fake'
$plist = @([pscustomobject]@{ Id = 'Default'; Name = 'Work' })
Assert-Brave (New-BraveProfileShortcuts -DesktopDir $dir -BraveExe $fakeExe -Profiles $plist) 'DryRun shortcuts returns true'
Assert-Brave (-not (Test-Path (Join-Path $dir 'Brave Work.lnk'))) 'DryRun must not write lnk'
Remove-Item -Recurse -Force $dir
```

- [ ] **Step 4: Run `test-brave.ps1` and `test-syntax.ps1`.** Expect PASS.

---

### Task 3: Install Apps UI + two handlers

**Files:**
- Modify: `Windows-PC-Setup.ps1` Install Apps tab, after `$InstallAppsPanel.Controls.Add($BtnInstallApps)` and before `$TabInstallApps.Controls.Add($InstallAppsPanel)`
- Modify: version strings (header comment, startup banner, `$MainForm.Text`, `$LblVersion`)

**Interfaces:**
- Consumes: helpers from Task 2, `Show-ConfirmationDialog`, `Set-ButtonDisabled` / `Set-ButtonEnabled`, `Update-Progress`, `Hide-Progress`, `Update-Status`, `$script:ChkDryRun`
- Produces: `$script:BtnBraveApply`, `$script:BtnBraveShortcuts`

- [ ] **Step 1: Raise intro height** on Install Apps (`New-IntroText`) so the new sentence fits: install Brave first, then Apply Brave settings, then after named profiles create shortcuts; do this before Harden. PIN, Sync, Proton sign-in stay manual. These settings apply to this Windows account only; they do not follow a later standard user.

- [ ] **Step 2: Add section + buttons**

After the Install Selected row, `$appYPos += 50` (enough to clear the button row), then `New-SectionHeader -Text 'Brave'`. Two buttons:

`$script:BtnBraveApply` Text `Apply &Brave settings` Size 200x32 Accent, tooltip "Writes HKCU Brave policies and force-installs Proton Pass. Install Brave first. (Alt+B)"

`$script:BtnBraveShortcuts` Text `Create &profile shortcuts` Size 200x32, location to the right of Apply, tooltip "Writes Brave <name> shortcuts for existing profiles. Open Brave and name profiles first. (Alt+P)"

Helper in the click-scope (script function):

```powershell
function Update-BraveButtonState {
    $script:BtnBraveApply.Enabled = (Test-BraveInstalled)
    $ls = Join-Path (Get-BraveUserDataDir) 'Local State'
    $hasProfiles = $false
    try {
        if (Test-Path -LiteralPath $ls) {
            $raw = Get-Content -LiteralPath $ls -Raw -ErrorAction Stop
            $hasProfiles = (@(Get-BraveProfilesForShortcuts -LocalStateJson $raw).Count -gt 0)
        }
    } catch {
        $hasProfiles = $false
    }
    $script:BtnBraveShortcuts.Enabled = $hasProfiles
}
```

Call `Update-BraveButtonState` after creating the buttons, at the end of the existing Install Selected `finally`, and from `$TabControl.Add_SelectedIndexChanged` after the tabs are added (`Windows-PC-Setup.ps1` near the `Controls.AddRange` for tabs) when `$TabControl.SelectedTab -eq $TabInstallApps`.

- [ ] **Step 3: Apply click handler**

```powershell
$script:BtnBraveApply.Add_Click({
    if (-not (Test-BraveInstalled)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Install Brave first (tick Brave Browser, then Install Selected).',
            'Brave',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    $script:DryRun = $script:ChkDryRun.Checked
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('New Tab page on startup')
    [void]$lines.Add('Search engine: Google')
    [void]$lines.Add('Brave password, address, and card autofill: off')
    [void]$lines.Add('Ask where to save each download')
    [void]$lines.Add('Default folder: Windows Downloads')
    [void]$lines.Add('Force-install Proton Pass extension')
    [void]$lines.Add('')
    [void]$lines.Add('Close Brave if it is open so search and the extension apply.')
    [void]$lines.Add('Proton sign-in, PIN, and Sync stay manual.')
    $detail = ($lines -join "`n")
    if (-not (Show-ConfirmationDialog -Title 'Apply Brave settings' -SelectedCount 6 -ActionType 'apply' -IsDryRun $script:DryRun -Detail $detail)) {
        return
    }
    Set-ButtonDisabled -Button $script:BtnBraveApply -WorkingText 'Applying...'
    try {
        Update-Progress -Current 1 -Total 1 -CurrentItem 'Brave policies'
        $ok = Set-BraveHkcuPolicies
        Hide-Progress
        $dry = if ($script:DryRun) { '[DRY RUN] ' } else { '' }
        if ($ok) {
            Update-Status "${dry}Brave settings applied. Close Brave if it is running."
            $icon = [System.Windows.Forms.MessageBoxIcon]::Information
            $box = "${dry}Brave settings applied.`n`nLog: $script:LogPath"
        } else {
            Update-Status "${dry}Brave settings finished with errors. See log."
            $icon = [System.Windows.Forms.MessageBoxIcon]::Warning
            $box = "${dry}Brave settings finished with errors.`n`nLog: $script:LogPath"
        }
        [System.Windows.Forms.MessageBox]::Show(
            $box,
            'Brave',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            $icon
        )
    } catch {
        Write-Log "Brave apply error: $_" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show("Error: $_", 'Brave', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        Hide-Progress
        Set-ButtonEnabled -Button $script:BtnBraveApply
        $script:DryRun = $false
        Update-BraveButtonState
    }
})
```

- [ ] **Step 4: Shortcuts click handler**

Read Local State, `Get-BraveProfilesForShortcuts`. If count 0: info box "Open Brave, create and name profiles, then retry." return. Else `$plan = Get-BraveUniqueShortcutNames -Profiles $profs`. Confirm Detail is one line per item: `<FileName>.lnk -> --profile-directory="<Id>"`. Then `New-BraveProfileShortcuts -DesktopDir ([Environment]::GetFolderPath('Desktop')) -BraveExe (Get-BraveInstallPath) -Profiles $profs`. Same disable/enable/progress/DryRun/finally pattern as Apply. Empty-or-fail uses the function's `$false`.

- [ ] **Step 5: Version strings to v2.5.** Footer adds `Alt+B (Brave), Alt+P (profile shortcuts)`.

- [ ] **Step 6: Assert UI wiring in `test-brave.ps1`** (AST, not a 2500-char slice):

```powershell
Assert-Brave ($scriptText -match 'Apply &Brave settings') 'Apply Brave button text'
Assert-Brave ($scriptText -match 'Create &profile shortcuts') 'Shortcuts button text'
Assert-Brave ($scriptText -match 'Update-BraveButtonState') 'button state helper'
Assert-Brave ($scriptText -match 'Add_SelectedIndexChanged') 'tab select refreshes Brave buttons'
Assert-Brave ($scriptText -match 'LOCALAPPDATA.*Brave-Browser\\Application\\brave.exe' -or $scriptText -match "LOCALAPPDATA.+'BraveSoftware\\Brave-Browser\\Application\\brave.exe'") 'per-user Brave path is searched'

$installClick = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Member.ToString() -eq 'Add_Click' -and
        $node.Expression.ToString() -match 'BtnInstallApps'
    }, $true) | Select-Object -First 1
Assert-Brave ($null -ne $installClick) 'Install Add_Click AST found'
$installText = $installClick.Extent.Text
Assert-Brave ($installText -notmatch 'Set-BraveHkcuPolicies') 'Install Selected must not apply Brave policies'
Assert-Brave ($installText -notmatch 'New-BraveProfileShortcuts') 'Install Selected must not write shortcuts'
```

- [ ] **Step 7: Run `test-syntax.ps1` and `test-brave.ps1`.** Expect PASS.

---

### Task 4: Docs

**Files:**
- Modify: `docs/known.md`, `README.md`, `.claude/CLAUDE.md`

- [ ] **Step 1: known.md** add traps from the spec (HKCU policy not JSON; shortcut directory id; Apply before Harden blocking pair; Proton sign-in/PIN/Sync manual; Downloads Known Folder; HKCU does not follow a later standard user). Add `test-brave.ps1` to the Commands table.

- [ ] **Step 2: README** Install Apps / How to Run: install Brave, Apply Brave settings, name profiles, shortcuts. Keyboard shortcuts Alt+B / Alt+P. Version 2.5.

- [ ] **Step 3: `.claude/CLAUDE.md`** version 2.5, Install Apps mention, shortcuts line.

- [ ] **Step 4: Re-run both gates.** Expect PASS.

---

## Spec coverage (self-review)

| Spec item | Task |
|---|---|
| Install Selected winget-only | Task 3 step 6 assert |
| Apply disabled without brave.exe | Task 3 Update-BraveButtonState |
| Shortcuts disabled without profiles | Task 3 Update-BraveButtonState |
| Policy payload table | Task 2 Set-BraveHkcuPolicies |
| Proton CWS id | Task 1 string assert + Task 2 |
| Unique shortcut names / Guest skip | Task 1 tests + Task 2 |
| Tab select refresh / StrictMode Local State | Task 3 Update-BraveButtonState |
| Extension upsert / search write order | Task 2 Set-BraveHkcuPolicies |
| Per-user Brave.exe path | Task 2 Get-BraveInstallPath |
| Failure modal is Warning | Task 3 Apply handler |
| Daily-admin HKCU limit | Task 3 intro + Task 4 known.md |
| Known Folder downloads | Task 2 Get-BraveDownloadsFolder |
| No JSON prefs / no icon policies / no Drive launcher | Task 1 forbidden-string asserts |
| Docs + v2.5 | Task 4 |
| DryRun | Task 2 step 3 |
| Apply before Harden (copy only) | Task 3 intro + Task 4 known.md |
