# Example: Basic usage of the BaiHelper module
# This script demonstrates how to use the Export-SeasonScoreSheet function

#requires -Module BaiHelper

<#
.SYNOPSIS
    Example script showing basic usage of the BaiHelper module.

.DESCRIPTION
    This script demonstrates various ways to use the Export-SeasonScoreSheet function
    from the BaiHelper PowerShell module.

.NOTES
    Make sure to import the BaiHelper module before running this script:
    Import-Module BaiHelper
#>

# Import the module (if not already imported)
if (-not (Get-Module BaiHelper)) {
    Import-Module BaiHelper -Force
}

Write-Host "BaiHelper Module Example Usage" -ForegroundColor Green
Write-Host "=============================" -ForegroundColor Green
Write-Host

# Example 1: Basic usage with credential prompt
Write-Host "Example 1: Basic usage (will prompt for credentials)" -ForegroundColor Yellow
Write-Host "Export-SeasonScoreSheet" -ForegroundColor Cyan

try {
    # Uncomment the line below to run the basic example
    # $result1 = Export-SeasonScoreSheet
    # Write-Host "Export completed: $result1" -ForegroundColor Green
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 1: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 2: Get available seasons" -ForegroundColor Yellow
Write-Host "Get-AvailableSeasons" -ForegroundColor Cyan

try {
    # Uncomment the lines below to run the seasons example
    # $seasons = Get-AvailableSeasons
    # Write-Host "Available seasons: $($seasons -join ', ')" -ForegroundColor Green
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 2: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 3: Using pre-created credentials" -ForegroundColor Yellow
Write-Host "Export-SeasonScoreSheet -Credential `$cred" -ForegroundColor Cyan

try {
    # Create credentials (you would replace this with actual credentials)
    Write-Host "Getting credentials..." -ForegroundColor Gray
    # $cred = Get-Credential -Message "Enter your NASP Tournaments credentials"
    
    # Uncomment the lines below to run with credentials
    # $result3 = Export-SeasonScoreSheet -Credential $cred
    # Write-Host "Export completed: $result3" -ForegroundColor Green
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 3: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 4: Discover and export specific season" -ForegroundColor Yellow
Write-Host "Get-AvailableSeasons + Export-SeasonScoreSheet" -ForegroundColor Cyan

try {
    # Uncomment the lines below to run the discover and export example
    # $cred = Get-Credential -Message "Enter your NASP Tournaments credentials"
    # $seasons = Get-AvailableSeasons -Credential $cred
    # Write-Host "Available seasons: $($seasons -join ', ')" -ForegroundColor Green
    # $selectedSeason = $seasons | Where-Object { $_ -like "*2023*" } | Select-Object -First 1
    # if ($selectedSeason) {
    #     $result4 = Export-SeasonScoreSheet -Credential $cred -Season $selectedSeason
    #     Write-Host "Export completed: $result4" -ForegroundColor Green
    # }
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 4: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 5: Create Enhanced Score Sheet" -ForegroundColor Yellow
Write-Host "New-EnhancedScoreSheet -SchoolName 'Sample_High_School' -Season '2023-2024'" -ForegroundColor Cyan

try {
    # Uncomment the lines below to run the enhancement example
    # $enhancedPath = New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024"
    # Write-Host "Enhanced score sheet created: $enhancedPath" -ForegroundColor Green
    
    # Example with IgnoreEmpty parameter
    # $enhancedPath = New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024" -IgnoreEmpty
    # if ($enhancedPath) {
    #     Write-Host "Enhanced score sheet created: $enhancedPath" -ForegroundColor Green
    # } else {
    #     Write-Host "Score sheet was empty, skipped enhancement" -ForegroundColor Yellow
    # }
    
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
    Write-Host "Note: This requires an existing score sheet file to enhance" -ForegroundColor Gray
    Write-Host "Adds AS_* (arrow counts), E_* (end scores), H1/H2 (half scores), and ES_* (end analysis)" -ForegroundColor Gray
    Write-Host "Use -IgnoreEmpty to skip processing empty files instead of failing" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 5: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 6: Specifying season and output path" -ForegroundColor Yellow
Write-Host "Export-SeasonScoreSheet -Season '2023-2024' -OutputPath 'C:\\Exports'" -ForegroundColor Cyan

try {
    # Create output directory if it doesn't exist
    $outputPath = Join-Path $env:TEMP "BaiHelper_Examples"
    if (-not (Test-Path $outputPath)) {
        New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
    }
    
    Write-Host "Using output path: $outputPath" -ForegroundColor Gray
    
    # Uncomment the lines below to run with specific season and output path
    # $result6 = Export-SeasonScoreSheet -Season "2023-2024" -OutputPath $outputPath
    # Write-Host "Export completed: $result6" -ForegroundColor Green
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 6: $($_.Exception.Message)"
}

Write-Host
Write-Host "Example 7: Full parameter example" -ForegroundColor Yellow
Write-Host "Export-SeasonScoreSheet -Credential `$cred -Season '2023-2024' -OrganizationId 5232 -OutputPath 'C:\Exports'" -ForegroundColor Cyan

try {
    # Uncomment the lines below to run the full example
    # $cred = Get-Credential -Message "Enter your NASP Tournaments credentials"
    # $outputPath = Join-Path $env:TEMP "BaiHelper_Full_Example"
    # $result7 = Export-SeasonScoreSheet -Credential $cred -Season "2023-2024" -OrganizationId 5232 -OutputPath $outputPath
    # Write-Host "Export completed: $result7" -ForegroundColor Green
    Write-Host "Example commented out - remove comments to test" -ForegroundColor Gray
}
catch {
    Write-Warning "Error in Example 7: $($_.Exception.Message)"
}

Write-Host
Write-Host "To run these examples:" -ForegroundColor Green
Write-Host "1. Uncomment the example you want to test" -ForegroundColor White
Write-Host "2. Make sure you have valid NASP Tournaments credentials" -ForegroundColor White
Write-Host "3. Run the script" -ForegroundColor White
Write-Host
Write-Host "For help with the functions, use:" -ForegroundColor Green
Write-Host "Get-Help Export-SeasonScoreSheet -Full" -ForegroundColor Cyan
Write-Host "Get-Help Get-AvailableSeasons -Full" -ForegroundColor Cyan
Write-Host "Get-Help New-EnhancedScoreSheet -Full" -ForegroundColor Cyan