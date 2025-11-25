# bai-helper

A PowerShell utility for exporting season score sheets from NASP Tournaments.

## Export-SeasonScoreSheet.ps1

This script logs into the NASP Tournaments website, navigates to the Season Score Sheet page, and exports the data to a CSV file. The exported file is saved in a "Season Score Sheets/{SchoolName}" folder structure.

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

The script creates the following folder structure:

```
{OutputPath}/
└── Season Score Sheets/
    └── {SchoolName}/
        └── SeasonScoreSheet_{Season}_{Timestamp}.csv
```

### Notes

- The script maintains a web session to handle authentication cookies
- ASP.NET ViewState and EventValidation tokens are handled automatically
- If direct CSV export is not available, the script will attempt to parse table data from the HTML page