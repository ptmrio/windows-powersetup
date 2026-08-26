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
    'New-BraveProfileShortcuts',
    'Update-BraveTabState'
)
foreach ($need in $requiredFns) {
    Assert-Brave ($fnNames -contains $need) "$need must exist in the PowerShell 5.1 AST"
}

if ($failures -gt 0) {
    Write-Host "Brave tests FAILED: $failures" -ForegroundColor Red
    exit 1
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
Assert-Brave (@(Get-BraveProfilesForShortcuts -LocalStateJson 'null').Count -eq 0) 'JSON null is no profiles'

$collide = @(
    [pscustomobject]@{ Id = 'Default'; Name = 'Work_Profile 2' },
    [pscustomobject]@{ Id = 'Profile 1'; Name = 'Work' },
    [pscustomobject]@{ Id = 'Profile 2'; Name = 'Work' }
)
$names = @(Get-BraveUniqueShortcutNames -Profiles $collide | ForEach-Object { $_.FileName })
Assert-Brave (($names | Select-Object -Unique).Count -eq 3) 'collision names are unique'

$ud = Get-BraveUserDataDir
Assert-Brave ($ud -match 'Brave-Browser\\User Data$') 'user data dir suffix'

$script:BraveDryLogs = New-Object System.Collections.Generic.List[string]
function Write-Log { param($Message, $Level = 'INFO'); [void]$script:BraveDryLogs.Add([string]$Message) }
$script:DryRun = $true
$script:BraveProtonPassExtensionId = 'ghmbeldphafepmbegfdlkpapadhbakde'
$script:BraveProtonPassForceInstall = 'ghmbeldphafepmbegfdlkpapadhbakde;https://clients2.google.com/service/update2/crx'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveInstallPath'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Get-BraveDownloadsFolder'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Set-BravePolicyDword'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Set-BravePolicyString'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Set-BraveHkcuPolicies'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'New-BraveProfileShortcuts'
Import-BraveFunctionFromScript -Path $scriptPath -FunctionName 'Update-BraveTabState'
$script:LblBraveStatus = $null
$script:BtnBraveApply = $null
$script:BtnBraveShortcuts = $null
$script:BraveDryLogs.Clear()
Assert-Brave (Set-BraveHkcuPolicies) 'DryRun Set-BraveHkcuPolicies returns true'
$allLogs = ($script:BraveDryLogs -join "`n")
Assert-Brave ($allLogs -match 'RestoreOnStartup') 'default Groups logs startup'
Assert-Brave ($allLogs -match 'ExtensionInstallForcelist') 'default Groups logs Proton Pass'
$script:BraveDryLogs.Clear()
Assert-Brave (Set-BraveHkcuPolicies -Groups @('Startup')) 'DryRun Startup-only returns true'
$subLogs = ($script:BraveDryLogs -join "`n")
Assert-Brave ($subLogs -match 'RestoreOnStartup') 'subset logs startup'
Assert-Brave ($subLogs -notmatch 'PasswordManagerEnabled') 'subset does not log password manager'
Assert-Brave ($subLogs -notmatch 'ExtensionInstallForcelist') 'subset does not log Proton Pass'
$script:BraveDryLogs.Clear()
Assert-Brave (Set-BraveHkcuPolicies -Groups @()) 'empty Groups returns true'
$emptyLogs = ($script:BraveDryLogs -join "`n")
Assert-Brave ($emptyLogs -notmatch 'Would set') 'empty Groups logs no Would set'
Assert-Brave ($(Update-BraveTabState; $true)) 'Update-BraveTabState does not throw when controls are null'
$dir = Join-Path $env:TEMP ("BraveShortcutTest-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $dir | Out-Null
try {
    $fakeExe = Join-Path $dir 'brave.exe'
    Set-Content -Path $fakeExe -Value 'fake'
    $plist = @([pscustomobject]@{ Id = 'Default'; Name = 'Work' })
    Assert-Brave (New-BraveProfileShortcuts -DesktopDir $dir -BraveExe $fakeExe -Profiles $plist) 'DryRun shortcuts returns true'
    Assert-Brave (-not (Test-Path (Join-Path $dir 'Brave Work.lnk'))) 'DryRun must not write lnk'
}
finally {
    Remove-Item -Recurse -Force $dir
}

Assert-Brave ($scriptText -match 'Apply &Brave settings') 'Apply Brave button text'
Assert-Brave ($scriptText -match 'Create &profile shortcuts') 'Shortcuts button text'
Assert-Brave ($scriptText -match 'Update-BraveTabState') 'button state helper renamed to tab state'
Assert-Brave ($scriptText -notmatch 'Update-BraveButtonState') 'old Update-BraveButtonState name is gone'
Assert-Brave ($scriptText -match '\$TabBrave') 'Brave tab page exists'
Assert-Brave ($scriptText -match '\$script:BraveCheckboxes') 'Brave checkbox array'
Assert-Brave ($scriptText -match 'Open New Tab page on startup') 'Startup checkbox text'
Assert-Brave ($scriptText -match 'Set search engine to Google') 'Search checkbox text'
Assert-Brave ($scriptText -match 'Disable Brave password manager') 'PasswordManager checkbox text'
Assert-Brave ($scriptText -match 'Disable address autofill') 'AutofillAddress checkbox text'
Assert-Brave ($scriptText -match 'Disable card autofill') 'AutofillCreditCard checkbox text'
Assert-Brave ($scriptText -match 'Ask where to save each download') 'DownloadPrompt checkbox text'
Assert-Brave ($scriptText -match 'Set default download folder to Windows Downloads') 'DownloadDirectory checkbox text'
Assert-Brave ($scriptText -match 'Force-install Proton Pass extension') 'ProtonPass checkbox text'
Assert-Brave ($scriptText -match 'Brave found:') 'status label found copy'
Assert-Brave ($scriptText -match 'brave\.exe is on disk') 'status label missing copy is disk-based'
Assert-Brave ($scriptText -notmatch 'tick Brave Browser, then Install Selected') 'must not imply this-session winget'
Assert-Brave ($scriptText -match '\[string\[\]\]\$Groups') 'Set-BraveHkcuPolicies takes Groups'
Assert-Brave ($scriptText -match "BraveSoftware\\Brave-Browser\\Application\\brave.exe") 'per-user and Program Files brave.exe paths are searched'
Assert-Brave ($scriptText -match '\$TabBloatware,\s*\$TabSettings,\s*\$TabInstallApps,\s*\$TabBrave,\s*\$TabHarden,\s*\$TabRepair') 'tab order Install Apps then Brave then Harden'
Assert-Brave ($scriptText -notmatch 'Then Apply Brave settings \(Google, new tab') 'Install Apps intro is no longer the Brave workflow dump'

$tabFn = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-BraveTabState'
    }, $true) | Select-Object -First 1
Assert-Brave ($null -ne $tabFn) 'Update-BraveTabState function AST exists'
if ($null -ne $tabFn) {
    $tabBody = $tabFn.Body.Extent.Text
    Assert-Brave ($tabBody -match 'Get-BraveInstallPath' -or $tabBody -match 'Test-BraveInstalled') 'tab state uses disk detection'
    Assert-Brave ($tabBody -notmatch 'AppCheckboxes') 'tab state does not use winget checkboxes'
}

$sel = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Member.ToString() -eq 'Add_SelectedIndexChanged'
    }, $true) | Select-Object -First 1
Assert-Brave ($null -ne $sel) 'SelectedIndexChanged AST found'
if ($null -ne $sel) {
    $selText = $sel.Extent.Text
    Assert-Brave ($selText -match 'TabBrave') 'tab select refreshes on Brave tab'
    Assert-Brave ($selText -notmatch 'TabInstallApps') 'tab select must not key Brave refresh off Install Apps'
}

$installClick = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Member.ToString() -eq 'Add_Click' -and
        $node.Expression.ToString() -match 'BtnInstallApps'
    }, $true) | Select-Object -First 1
Assert-Brave ($null -ne $installClick) 'Install Add_Click AST found'
if ($null -ne $installClick) {
    $installText = $installClick.Extent.Text
    Assert-Brave ($installText -notmatch 'Set-BraveHkcuPolicies') 'Install Selected must not apply Brave policies'
    Assert-Brave ($installText -notmatch 'New-BraveProfileShortcuts') 'Install Selected must not write shortcuts'
}

$braveApplyClick = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
        $node.Member.ToString() -eq 'Add_Click' -and
        $node.Expression.ToString() -match 'BtnBraveApply'
    }, $true) | Select-Object -First 1
Assert-Brave ($null -ne $braveApplyClick) 'Brave Apply Add_Click AST found'
if ($null -ne $braveApplyClick) {
    $applyText = $braveApplyClick.Extent.Text
    Assert-Brave ($applyText -match 'Set-BraveHkcuPolicies\s+-Groups') 'Apply passes -Groups'
    Assert-Brave ($applyText -match 'BraveCheckboxes') 'Apply reads checkbox selection'
    Assert-Brave ($applyText -match '\.Tag') 'Apply uses checkbox Tag values'
    Assert-Brave (([regex]::Matches($applyText, 'Test-BraveInstalled')).Count -ge 2) 'Apply rechecks install after confirm'
    Assert-Brave ($applyText -notmatch 'tick Brave Browser') 'Apply missing-Brave copy is not session-install'
}

if ($failures -gt 0) {
    Write-Host "Brave tests FAILED: $failures" -ForegroundColor Red
    exit 1
}
Write-Host "Brave tests OK" -ForegroundColor Green
exit 0
