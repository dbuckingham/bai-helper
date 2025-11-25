# Test script for Get-NASPScoreSheet.ps1 parameter handling
# This script tests the parameter validation without actually connecting to the website

Write-Host "Testing Get-NASPScoreSheet.ps1 Parameter Handling..." -ForegroundColor Cyan
Write-Host ""

# Test 1: Check if script file exists
Write-Host "Test 1: Script file exists" -ForegroundColor Yellow
if (Test-Path "./Get-NASPScoreSheet.ps1") {
    Write-Host "  PASSED - Script file found" -ForegroundColor Green
} else {
    Write-Host "  FAILED - Script file not found" -ForegroundColor Red
    exit 1
}

# Test 2: Parse script and verify parameters
Write-Host "`nTest 2: Script parameters are correctly defined" -ForegroundColor Yellow
$scriptContent = Get-Content "./Get-NASPScoreSheet.ps1" -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path './Get-NASPScoreSheet.ps1').Path, [ref]$null, [ref]$null)
$params = $ast.ParamBlock.Parameters

if ($params.Count -eq 3) {
    Write-Host "  PASSED - Found 3 parameters" -ForegroundColor Green
    foreach ($param in $params) {
        Write-Host "    - $($param.Name.VariablePath.UserPath)" -ForegroundColor Gray
    }
} else {
    Write-Host "  FAILED - Expected 3 parameters, found $($params.Count)" -ForegroundColor Red
}

# Test 3: Verify parameter names
Write-Host "`nTest 3: Required parameters exist" -ForegroundColor Yellow
$expectedParams = @('Credential', 'OutputPath', 'SeasonId')
$paramNames = $params | ForEach-Object { $_.Name.VariablePath.UserPath }
$allFound = $true

foreach ($expected in $expectedParams) {
    if ($paramNames -contains $expected) {
        Write-Host "  PASSED - Parameter '$expected' found" -ForegroundColor Green
    } else {
        Write-Host "  FAILED - Parameter '$expected' not found" -ForegroundColor Red
        $allFound = $false
    }
}

# Test 4: Verify default values
Write-Host "`nTest 4: Default values are set correctly" -ForegroundColor Yellow
if ($scriptContent -match 'OutputPath.*=.*"NASPScoreSheet.csv"') {
    Write-Host "  PASSED - OutputPath default is 'NASPScoreSheet.csv'" -ForegroundColor Green
} else {
    Write-Host "  FAILED - OutputPath default not found or incorrect" -ForegroundColor Red
}

if ($scriptContent -match 'SeasonId.*=.*5232') {
    Write-Host "  PASSED - SeasonId default is 5232" -ForegroundColor Green
} else {
    Write-Host "  FAILED - SeasonId default not found or incorrect" -ForegroundColor Red
}

# Test 5: Verify help documentation
Write-Host "`nTest 5: Help documentation is available" -ForegroundColor Yellow
$help = Get-Help "./Get-NASPScoreSheet.ps1" -ErrorAction SilentlyContinue
if ($help.Synopsis) {
    Write-Host "  PASSED - Synopsis: $($help.Synopsis.Trim())" -ForegroundColor Green
} else {
    Write-Host "  FAILED - No synopsis found" -ForegroundColor Red
}

if ($help.Description) {
    Write-Host "  PASSED - Description available" -ForegroundColor Green
} else {
    Write-Host "  FAILED - No description found" -ForegroundColor Red
}

if ($help.Examples.Count -gt 0) {
    Write-Host "  PASSED - Found $($help.Examples.Count) examples" -ForegroundColor Green
} else {
    Write-Host "  FAILED - No examples found" -ForegroundColor Red
}

# Test 6: Verify key functionality keywords
Write-Host "`nTest 6: Script contains required functionality" -ForegroundColor Yellow
$requiredKeywords = @(
    'Invoke-WebRequest',
    'login.aspx',
    'SeasonScoreSheet.aspx',
    'Export-Csv',
    'Get-Credential',
    'UseForTeam'
)

foreach ($keyword in $requiredKeywords) {
    if ($scriptContent -match [regex]::Escape($keyword)) {
        Write-Host "  PASSED - Found '$keyword'" -ForegroundColor Green
    } else {
        Write-Host "  FAILED - Missing '$keyword'" -ForegroundColor Red
    }
}

Write-Host "`n=======================" -ForegroundColor Cyan
Write-Host "All Tests Completed!" -ForegroundColor Cyan
Write-Host "=======================" -ForegroundColor Cyan
