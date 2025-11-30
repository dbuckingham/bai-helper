# BaiHelper PowerShell Module

A PowerShell module for exporting season score sheets from NASP Tournaments website.

## Overview

The BaiHelper module provides functions for working with NASP Tournaments website:

- **Export-SeasonScoreSheet**: Logs into the NASP Tournaments website, navigates to the Season Score Sheet page, and exports the data to a CSV file. The exported file is saved in an organized folder structure: "Season Score Sheets/{SchoolName}/{Season}".

- **Get-AvailableSeasons**: Retrieves the list of available seasons for a given school organization, which is useful for discovering which seasons can be exported.

- **New-EnhancedScoreSheet**: Creates an enhanced copy of an existing Season Score Sheet with comprehensive analysis columns: arrow counts (AS_*), end scores (E_*), half scores (H1, H2), and end score distribution analysis (ES_*).

The functions work together to provide a complete workflow for downloading, analyzing, and enhancing archery score data.

## Installation

### Option 1: Manual Installation
1. Download or clone this repository
2. Copy the entire `BaiHelper` folder to one of your PowerShell module paths:
   - User modules: `$HOME\Documents\PowerShell\Modules\` (PowerShell Core)
   - User modules: `$HOME\Documents\WindowsPowerShell\Modules\` (Windows PowerShell)
   - System modules: `$env:ProgramFiles\PowerShell\Modules\`

### Option 2: Import from Local Path
```powershell
Import-Module "C:\Path\To\BaiHelper" -Force
```

### Verify Installation
```powershell
Get-Module BaiHelper -ListAvailable
Get-Command -Module BaiHelper
```

## Usage

### Import the Module
```powershell
Import-Module BaiHelper
```

#### Basic Usage (Prompts for Credentials)

```powershell
Export-SeasonScoreSheet
```

#### Get Available Seasons

```powershell
# Get list of available seasons for default organization
$seasons = Get-AvailableSeasons
$seasons | ForEach-Object { Write-Host "Available season: $_" }

# Get seasons for specific organization
Get-AvailableSeasons -OrganizationId 1234
```

#### Create Enhanced Score Sheet

```powershell
# Create enhanced score sheet with arrow count analysis
New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024"

# Specify custom base path
New-EnhancedScoreSheet -SchoolName "My_School" -Season "2024-2025" -BasePath "C:\Exports"

# Skip empty files instead of failing
New-EnhancedScoreSheet -SchoolName "School_Name" -Season "2023-2024" -IgnoreEmpty
```

#### Provide Credentials

```powershell
$cred = Get-Credential
Export-SeasonScoreSheet -Credential $cred
```

#### Select a Specific Season

```powershell
Export-SeasonScoreSheet -Season "2023-2024"
```

#### Specify Organization ID

```powershell
Export-SeasonScoreSheet -OrganizationId 5232
```

#### Specify Output Location

```powershell
Export-SeasonScoreSheet -OutputPath "C:\MyExports"
```

#### Full Example with All Parameters

```powershell
$cred = Get-Credential
Export-SeasonScoreSheet -Credential $cred -Season "2023-2024" -OrganizationId 5232 -OutputPath "C:\MyExports"
```

#### Discover and Export Specific Season

```powershell
# First, get available seasons
$cred = Get-Credential
$seasons = Get-AvailableSeasons -Credential $cred -OrganizationId 5232
Write-Host "Available seasons: $($seasons -join ', ')"

# Then export a specific season
$selectedSeason = $seasons | Where-Object { $_ -like "*2023*" } | Select-Object -First 1
Export-SeasonScoreSheet -Credential $cred -Season $selectedSeason -OrganizationId 5232
```

#### Complete Workflow - Export and Enhance

```powershell
# Export score sheet and create enhanced version with arrow counts
$cred = Get-Credential
$exportPath = Export-SeasonScoreSheet -Credential $cred -Season "2023-2024" -OrganizationId 5232
Write-Host "Exported to: $exportPath"

# Create enhanced version with arrow count analysis
$enhancedPath = New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024"
Write-Host "Enhanced version created: $enhancedPath"
```

### Parameters

#### Export-SeasonScoreSheet Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Credential` | No | A PSCredential object containing login credentials. If not provided, you will be prompted to enter credentials. |
| `-Season` | No | The season to select from the dropdown menu. If not provided, the default (current) season is used. |
| `-OrganizationId` | No | The organization ID for the score sheet URL. Defaults to 5232. |
| `-OutputPath` | No | Base path where "Season Score Sheets" folder will be created. Defaults to current directory. |

#### Get-AvailableSeasons Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Credential` | No | A PSCredential object containing login credentials. If not provided, you will be prompted to enter credentials. |
| `-OrganizationId` | No | The organization ID to retrieve seasons for. Defaults to 5232. |

#### New-EnhancedScoreSheet Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-SchoolName` | Yes | The name of the school (must match the folder name in the Season Score Sheets directory). |
| `-Season` | Yes | The season name (must match the folder name under the school directory). |
| `-BasePath` | No | The base path where the "Season Score Sheets" folder is located. Defaults to current directory. |
| `-IgnoreEmpty` | No | When specified, skips processing and returns null instead of throwing an error if the score sheet file is empty. |

## Requirements

- PowerShell 5.1 or later
- Valid NASP Tournaments account credentials

## Module Structure

```
BaiHelper/
├── BaiHelper.psd1          # Module manifest
├── BaiHelper.psm1          # Main module file
├── Public/                 # Public functions
│   ├── Export-SeasonScoreSheet.ps1
│   ├── Get-AvailableSeasons.ps1
│   └── New-EnhancedScoreSheet.ps1
├── Private/                # Private helper functions
│   ├── Config.ps1
│   ├── CommonHelpers.ps1
│   ├── AuthenticationHelpers.ps1
│   ├── DataExtractionHelpers.ps1
│   ├── ExportHelpers.ps1
│   ├── FileOperationHelpers.ps1
│   └── ScoreSheetHelpers.ps1
└── Examples/               # Example scripts
    └── BasicUsage.ps1
```

### Output

The module creates the following organized folder structure:

```
{OutputPath}/
└── Season Score Sheets/
    └── {SchoolName}/
        └── {Season}/
            ├── SeasonScoreSheet_{SchoolName}_{Season}.csv
            └── Enhanced_SeasonScoreSheet_{SchoolName}_{Season}.csv
```

**Example:**
```
C:\MyExports/
└── Season Score Sheets/
    └── Sample_High_School/
        └── 2023-2024/
            ├── SeasonScoreSheet_Sample_High_School_2023-2024.csv
            └── Enhanced_SeasonScoreSheet_Sample_High_School_2023-2024.csv
```

### Features

- **Automatic School Detection**: Extracts school name from the website and creates clean folder names
- **Season Management**: Automatically uses the default season or allows selection of a specific season
- **Robust Export Methods**: 
  - Primary: Direct CSV export from the website
  - Fallback: HTML table parsing when direct export is unavailable
- **Arrow Count Analysis**: Enhanced score sheets include AS_10 through AS_0 columns that count arrows scoring each value
- **End Score Analysis**: Enhanced score sheets include E_1 through E_6 columns that sum scores for each 5-arrow end
- **Half Score Analysis**: Enhanced score sheets include H1 and H2 columns for first half vs second half performance comparison
- **End Score Distribution**: Enhanced score sheets include ES_* columns that analyze end score patterns and consistency
- **Data Enhancement**: Automatically processes existing score sheets to add statistical analysis columns
- **Error Handling**: Comprehensive error handling with detailed feedback
- **Clean File Naming**: Removes invalid characters from school names and creates descriptive file names

### Enhanced Score Sheet Columns

When using `New-EnhancedScoreSheet`, the following additional columns are added to provide detailed analysis:

#### Arrow Score Count Columns (AS_*)
- **AS_10**: Count of arrows scoring 10 points (perfect shots)
- **AS_9**: Count of arrows scoring 9 points
- **AS_8**: Count of arrows scoring 8 points
- ... down to ...
- **AS_0**: Count of arrows scoring 0 points (misses)

#### End Score Columns (E_*)
- **E_1**: Sum of arrows 1-5 (first end)
- **E_2**: Sum of arrows 6-10 (second end)
- **E_3**: Sum of arrows 11-15 (third end)
- **E_4**: Sum of arrows 16-20 (fourth end)
- **E_5**: Sum of arrows 21-25 (fifth end)
- **E_6**: Sum of arrows 26-30 (sixth end)

#### Half Score Columns (H_*)
- **H1**: Sum of first half (ends 1-3)
- **H2**: Sum of second half (ends 4-6)

#### End Score Analysis Columns (ES_*)
- **ES_50**: Count of ends scoring exactly 50 points (perfect ends)
- **ES_49**: Count of ends scoring exactly 49 points
- **ES_48**: Count of ends scoring exactly 48 points
- **ES_47**: Count of ends scoring exactly 47 points
- **ES_46**: Count of ends scoring exactly 46 points
- **ES_45**: Count of ends scoring exactly 45 points
- **ES_40_44**: Count of ends scoring 40-44 points
- **ES_35_39**: Count of ends scoring 35-39 points
- **ES_30_34**: Count of ends scoring 30-34 points
- **ES_25_29**: Count of ends scoring 25-29 points
- **ES_20_24**: Count of ends scoring 20-24 points
- **ES_15_19**: Count of ends scoring 15-19 points
- **ES_10_14**: Count of ends scoring 10-14 points
- **ES_5_9**: Count of ends scoring 5-9 points
- **ES_0_4**: Count of ends scoring 0-4 points

These columns enable detailed analysis of shooting patterns, consistency across ends, scoring distribution, end score performance trends, and first half vs second half performance comparison.

### Notes

- **Authentication**: The script maintains a web session to handle authentication cookies securely
- **ASP.NET Compatibility**: ViewState and EventValidation tokens are handled automatically for proper form submissions
- **Multiple Export Methods**: 
  - Attempts direct CSV export first for optimal performance
  - Falls back to HTML table parsing if direct export is unavailable
  - Handles various table formats and nested HTML content
- **Smart Season Selection**: Only performs postback operations when a different season is explicitly requested
- **File Safety**: School names are automatically cleaned to remove invalid file system characters
- **Organized Storage**: Creates nested folder structure by school and season for easy organization