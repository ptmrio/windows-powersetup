# Syntax check only - does NOT run the main script
$scriptPath = Join-Path $PSScriptRoot "Windows-PC-Setup.ps1"
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$null, [ref]$errors)

if ($errors.Count -eq 0) {
    Write-Host "SYNTAX CHECK PASSED" -ForegroundColor Green
    Write-Host "Script lines: approximately $($ast.Extent.EndLineNumber)"
} else {
    Write-Host "SYNTAX ERRORS FOUND:" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "  Line $($e.Extent.StartLineNumber): $($e.Message)" -ForegroundColor Red
    }
}
