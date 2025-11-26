# bai-helper

A PowerShell utility for exporting season score sheets from NASP Tournaments.

## Export-SeasonScoreSheet.ps1

This script logs into the NASP Tournaments website, navigates to the Season Score Sheet page, and exports the data to a CSV file. The exported file is saved in an organized folder structure: "Season Score Sheets/{SchoolName}/{Season}". The script automatically extracts the school name from the website and handles multiple export methods for maximum compatibility.

### Requirements

- PowerShell 5.1 or later
- Valid NASP Tournaments account credentials

### Usage

#### Basic Usage (Prompts for Credentials)

```powershell
.\Export-SeasonScoreSheet.ps1
```

#### Provide Credentials

```powershell
$cred = Get-Credential
.\Export-SeasonScoreSheet.ps1 -Credential $cred
```

#### Select a Specific Season

```powershell
.\Export-SeasonScoreSheet.ps1 -Season "2023-2024"
```

#### Specify Organization ID

```powershell
.\Export-SeasonScoreSheet.ps1 -OrganizationId 5232
```

#### Specify Output Location

```powershell
.\Export-SeasonScoreSheet.ps1 -OutputPath "C:\MyExports"
```

#### Full Example with All Parameters

```powershell
$cred = Get-Credential
.\Export-SeasonScoreSheet.ps1 -Credential $cred -Season "2023-2024" -OrganizationId 5232 -OutputPath "C:\MyExports"
```

### Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-Credential` | No | A PSCredential object containing login credentials. If not provided, you will be prompted to enter credentials. |
| `-Season` | No | The season to select from the dropdown menu. If not provided, the default (current) season is used. |
| `-OrganizationId` | No | The organization ID for the score sheet URL. Defaults to 5232. |
| `-OutputPath` | No | Base path where "Season Score Sheets" folder will be created. Defaults to current directory. |

### Output

The script creates the following organized folder structure:

```
{OutputPath}/
└── Season Score Sheets/
    └── {SchoolName}/
        └── {Season}/
            └── SeasonScoreSheet_{Season}_{Timestamp}.csv
```

**Example:**
```
C:\MyExports/
└── Season Score Sheets/
    └── Sample_High_School/
        └── 2023-2024/
            └── SeasonScoreSheet_2023-2024_20231126_143022.csv
```

### Features

- **Automatic School Detection**: Extracts school name from the website and creates clean folder names
- **Season Management**: Automatically uses the default season or allows selection of a specific season
- **Robust Export Methods**: 
  - Primary: Direct CSV export from the website
  - Fallback: HTML table parsing when direct export is unavailable
- **Error Handling**: Comprehensive error handling with detailed feedback
- **Clean File Naming**: Removes invalid characters from school names and creates timestamped files

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