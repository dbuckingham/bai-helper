# BaiHelper PowerShell Module

A PowerShell module for exporting season score sheets from NASP Tournaments website.

## Overview

The BaiHelper module provides functions for working with NASP Tournaments website:

- **Export-SeasonScoreSheet**: Logs into the NASP Tournaments website, navigates to the Season Score Sheet page, and exports the data to a CSV file. The exported file is saved in an organized folder structure: "Season Score Sheets/{SchoolName}/{Season}".

- **Get-AvailableSeasons**: Retrieves the list of available seasons for a given school organization, which is useful for discovering which seasons can be exported.

Both functions automatically handle authentication and provide multiple export methods for maximum compatibility.

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
│   └── Get-AvailableSeasons.ps1
├── Private/                # Private helper functions
│   ├── Config.ps1
│   ├── CommonHelpers.ps1
│   ├── AuthenticationHelpers.ps1
│   ├── DataExtractionHelpers.ps1
│   ├── ExportHelpers.ps1
│   └── FileOperationHelpers.ps1
└── Examples/               # Example scripts
    └── BasicUsage.ps1
```

### Output

The script creates the following organized folder structure:

```
{OutputPath}/
└── Season Score Sheets/
    └── {SchoolName}/
        └── {Season}/
            └── SeasonScoreSheet_{SchoolName}_{Season}.csv
```

**Example:**
```
C:\MyExports/
└── Season Score Sheets/
    └── Sample_High_School/
        └── 2023-2024/
            └── SeasonScoreSheet_Sample_High_School_2023-2024.csv
```

### Features

- **Automatic School Detection**: Extracts school name from the website and creates clean folder names
- **Season Management**: Automatically uses the default season or allows selection of a specific season
- **Robust Export Methods**: 
  - Primary: Direct CSV export from the website
  - Fallback: HTML table parsing when direct export is unavailable
- **Error Handling**: Comprehensive error handling with detailed feedback
- **Clean File Naming**: Removes invalid characters from school names and creates descriptive file names

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