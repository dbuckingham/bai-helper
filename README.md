# bai-helper

PowerShell helper scripts for NASP (National Archery in the Schools Program) tournament management.

## Scripts

### Get-NASPScoreSheet.ps1

Downloads and processes NASP tournament score sheet data from nasptournaments.org.

#### Features
- Accepts credentials as a parameter or prompts the user to enter them
- Logs into nasptournaments.org
- Downloads the 2025-2026 season score sheet (or any specified season)
- Parses archer data from the score sheet
- Calculates team score by summing scores where "Use For Team" is Y
- Exports data to CSV format

#### Usage

**Basic usage (will prompt for credentials):**
```powershell
.\Get-NASPScoreSheet.ps1
```

**With credentials:**
```powershell
$cred = Get-Credential
.\Get-NASPScoreSheet.ps1 -Credential $cred
```

**Custom output path:**
```powershell
.\Get-NASPScoreSheet.ps1 -OutputPath "C:\Data\scores.csv"
```

**Different season:**
```powershell
.\Get-NASPScoreSheet.ps1 -SeasonId 5232
```

#### Parameters

- **Credential** (optional): PSCredential object containing username and password
- **OutputPath** (optional): Path to save the CSV file (default: "NASPScoreSheet.csv")
- **SeasonId** (optional): The season ID to download (default: 5232 for 2025-2026)

#### Output

The script creates a CSV file with the following columns:
- Name: Archer name
- Score: Archer's score
- UseForTeam: Y/N indicator for team inclusion

The last row contains the team total score.

#### Requirements

- PowerShell 5.1 or higher
- Internet connection
- Valid nasptournaments.org account credentials