#Requires -Version 5.1
# Tests production Harden helpers by extracting function ASTs from Windows-PC-Setup.ps1.
# Does not run the GUI script.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot "Windows-PC-Setup.ps1"
$failures = 0

function Assert-Harden {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) {
        Write-Host "FAIL: $Message" -ForegroundColor Red
        $script:failures++
    }
}

function Import-HardenFunctionFromScript {
    param(
        [string]$Path,
        [string]$FunctionName
    )
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

# --- New-LocalUser 5.1 must not have the fake parameter ---
$nl = Get-Command New-LocalUser -ErrorAction Stop
Assert-Harden (-not $nl.Parameters.ContainsKey('UserMayChangePassword')) 'New-LocalUser must not have -UserMayChangePassword on this host'
Assert-Harden $nl.Parameters.ContainsKey('UserMayNotChangePassword') 'New-LocalUser must have -UserMayNotChangePassword'

$scriptText = Get-Content -Path $scriptPath -Raw
Assert-Harden ($scriptText -notmatch 'UserMayChangePassword') 'Windows-PC-Setup.ps1 must not call -UserMayChangePassword'

# --- Extract production functions (must exist in the 5.1 AST, not a copy) ---
$parseErrors = $null
$tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors -and $parseErrors.Count -gt 0) {
    throw "Parse errors in ${scriptPath}: $($parseErrors[0].Message)"
}
$fnNames = @($ast.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
        }, $true) | ForEach-Object { $_.Name })
$requiredFns = @(
    'Test-HardenUsername',
    'Test-HardenPasswordPair',
    'Merge-HardenAsrState',
    'New-HardenStandardUser',
    'Test-HardenDefenderAvailable',
    'Get-HardenDefenderReady',
    'Set-HardenDefenderCloud',
    'Set-HardenPua',
    'Set-HardenNetworkProtection',
    'Set-HardenAsrBaseline',
    'Set-HardenSmartScreenWarn'
)
foreach ($need in $requiredFns) {
    Assert-Harden ($fnNames -contains $need) "$need must exist in the PowerShell 5.1 AST"
}
$nuAst = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-HardenStandardUser'
    }, $true) | Select-Object -First 1
Assert-Harden ($null -ne $nuAst) 'New-HardenStandardUser AST node'
if ($nuAst) {
    $span = $nuAst.Extent.EndLineNumber - $nuAst.Extent.StartLineNumber
    Assert-Harden ($span -lt 120) "New-HardenStandardUser AST span is $span lines (encoding poison?)"
}

Import-HardenFunctionFromScript -Path $scriptPath -FunctionName 'Test-HardenUsername'
Import-HardenFunctionFromScript -Path $scriptPath -FunctionName 'Test-HardenPasswordPair'
Import-HardenFunctionFromScript -Path $scriptPath -FunctionName 'Merge-HardenAsrState'

$ok = Test-HardenUsername 'jsmith'
Assert-Harden $ok.Ok 'jsmith should pass'

$bad = Test-HardenUsername ''
Assert-Harden (-not $bad.Ok) 'empty username should fail'

$admin = Test-HardenUsername 'Administrator'
Assert-Harden (-not $admin.Ok) 'Administrator should fail'

$guest = Test-HardenUsername 'guest'
Assert-Harden (-not $guest.Ok) 'guest (case) should fail'

$gast = Test-HardenUsername 'Gast'
Assert-Harden (-not $gast.Ok) 'Gast (German Guest) should fail'

$dot = Test-HardenUsername 'jsmith.'
Assert-Harden (-not $dot.Ok) 'trailing period should fail'

$space = Test-HardenUsername 'jan schmidt'
Assert-Harden (-not $space.Ok) 'spaces should fail'

$long = Test-HardenUsername ('a' * 21)
Assert-Harden (-not $long.Ok) '21 chars should fail'

$pw = Test-HardenPasswordPair 'x' 'y'
Assert-Harden (-not $pw.Ok) 'mismatch should fail'

$pwEmpty = Test-HardenPasswordPair '' ''
Assert-Harden (-not $pwEmpty.Ok) 'empty password should fail'

$pwShort = Test-HardenPasswordPair '1234567' '1234567'
Assert-Harden (-not $pwShort.Ok) '7-char password should fail'

$pwOk = Test-HardenPasswordPair '12345678' '12345678'
Assert-Harden $pwOk.Ok '8-char password should pass'

# ASR merge: null existing lists
$empty = Merge-HardenAsrState -ExistingIds $null -ExistingActions $null
Assert-Harden ($empty.Ids.Count -eq 6) 'null ASR lists should yield six GUIDs'
Assert-Harden (-not $empty.Mismatched) 'null ASR lists are not a length mismatch'
Assert-Harden (@($empty.Actions | Where-Object { $_ -ne 1 }).Count -eq 0) 'all merged ASR actions should be Block (1)'

$mismatch = Merge-HardenAsrState -ExistingIds @('01443614-cd74-433a-b99e-2ecdc07bfc25') -ExistingActions $null
Assert-Harden $mismatch.Mismatched 'IDs without a matching actions list must be flagged'

# Preserve a seventh rule (Audit = 2)
$seventh = '01443614-cd74-433a-b99e-2ecdc07bfc25'
$kept = Merge-HardenAsrState -ExistingIds @($seventh) -ExistingActions @(2)
Assert-Harden ($kept.Ids.Count -eq 7) 'seventh ASR rule must be preserved'
$idx = [array]::IndexOf(@($kept.Ids | ForEach-Object { $_.ToString().ToLower() }), $seventh.ToLower())
Assert-Harden ($idx -ge 0) 'seventh GUID present'
Assert-Harden ([int]$kept.Actions[$idx] -eq 2) 'seventh rule Audit action preserved'

# Promote an existing wanted GUID from Audit to Block
$firstWanted = '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2'
$promoted = Merge-HardenAsrState -ExistingIds @($firstWanted) -ExistingActions @(2)
$pidx = [array]::IndexOf(@($promoted.Ids | ForEach-Object { $_.ToString().ToLower() }), $firstWanted.ToLower())
Assert-Harden ([int]$promoted.Actions[$pidx] -eq 1) 'existing wanted ASR rule must be set to Block'

if ($failures -eq 0) {
    Write-Host 'HARDEN HELPER TESTS PASSED' -ForegroundColor Green
    exit 0
}
Write-Host "HARDEN HELPER TESTS FAILED: $failures" -ForegroundColor Red
exit 1
