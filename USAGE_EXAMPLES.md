# Usage Examples for Get-NASPScoreSheet.ps1

This document provides detailed examples of how to use the Get-NASPScoreSheet.ps1 script.

## Basic Usage

### Example 1: Interactive Mode (Recommended for first-time users)

Simply run the script without any parameters. You'll be prompted to enter your credentials:

```powershell
.\Get-NASPScoreSheet.ps1
```

**What happens:**
1. You'll see a Windows credential prompt
2. Enter your nasptournaments.org username and password
3. The script will log in, download the score sheet, and save it to `NASPScoreSheet.csv`

### Example 2: Using Pre-stored Credentials

If you want to store your credentials in a variable first:

```powershell
# Get credentials once
$cred = Get-Credential -Message "Enter your NASP credentials"

# Use the credentials with the script
.\Get-NASPScoreSheet.ps1 -Credential $cred
```

### Example 3: Custom Output Location

Save the CSV file to a specific location:

```powershell
.\Get-NASPScoreSheet.ps1 -OutputPath "C:\Users\YourName\Documents\NASP\ScoreSheet_2025.csv"
```

### Example 4: Different Season

Download a different season's score sheet by specifying the season ID:

```powershell
# Download season 5232 (2025-2026)
.\Get-NASPScoreSheet.ps1 -SeasonId 5232

# Download a different season
.\Get-NASPScoreSheet.ps1 -SeasonId 5100
```

### Example 5: All Parameters Combined

```powershell
$cred = Get-Credential
.\Get-NASPScoreSheet.ps1 -Credential $cred -OutputPath ".\Scores\2025-2026.csv" -SeasonId 5232
```

## Advanced Usage

### Automating with Scheduled Tasks

You can create a scheduled task to run this script automatically. Here's how to set it up:

1. **Store credentials securely (Windows only):**

```powershell
# Save encrypted credentials to file
$cred = Get-Credential
$cred | Export-Clixml -Path "$env:USERPROFILE\nasp_cred.xml"

# Later, import and use them
$cred = Import-Clixml -Path "$env:USERPROFILE\nasp_cred.xml"
.\Get-NASPScoreSheet.ps1 -Credential $cred
```

2. **Create a wrapper script:**

```powershell
# SavedCredentialRunner.ps1
$credPath = "$env:USERPROFILE\nasp_cred.xml"
$cred = Import-Clixml -Path $credPath
.\Get-NASPScoreSheet.ps1 -Credential $cred -OutputPath "C:\Reports\DailyScoreSheet.csv"
```

3. **Schedule it in Task Scheduler:**
   - Open Task Scheduler
   - Create a new task
   - Set the action to: `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\SavedCredentialRunner.ps1"`
   - Set the schedule (e.g., daily at 8 AM)

### Batch Processing Multiple Seasons

```powershell
# Download multiple seasons
$seasons = @(5232, 5100, 4950)
$cred = Get-Credential

foreach ($season in $seasons) {
    $outputFile = "ScoreSheet_Season_$season.csv"
    Write-Host "Downloading season $season..." -ForegroundColor Cyan
    .\Get-NASPScoreSheet.ps1 -Credential $cred -OutputPath $outputFile -SeasonId $season
}
```

## Output Format

The script creates a CSV file with the following structure:

```csv
Name,Score,UseForTeam
John Doe,285,Y
Jane Smith,290,Y
Bob Johnson,275,N
Alice Williams,280,Y
TEAM TOTAL,855,-
```

### Understanding the Output

- **Name**: Archer's name
- **Score**: Archer's tournament score
- **UseForTeam**: Whether the archer's score counts toward the team total (Y/N)
- **TEAM TOTAL**: Sum of all scores where UseForTeam = Y

## Troubleshooting

### Issue: "Login failed. Please check your credentials."

**Solution:** Verify your username and password are correct. Try logging in manually at https://nasptournaments.org/userutilities/login.aspx

### Issue: "Could not automatically parse score sheet data."

**Solution:** The script saves the raw HTML to `scoresheet_raw.html` for manual review. The website structure may have changed. Check this file and report the issue to the script maintainer.

### Issue: Script takes a long time to run

**Solution:** This is normal. The script needs to:
1. Connect to the website
2. Authenticate
3. Navigate to the score sheet page
4. Download and parse data

Typical execution time is 10-30 seconds depending on network speed.

### Issue: CSV file is empty or missing data

**Solution:** 
1. Check if `scoresheet_raw.html` was created
2. Verify you have access to the season score sheet on the website
3. Ensure the season ID is correct

## Security Best Practices

1. **Never store credentials in plain text**
2. **Use `Export-Clixml`** for encrypted credential storage (Windows only)
3. **Limit access** to scripts and credential files
4. **Use least privilege** - only run with necessary permissions
5. **Regularly update** your password on nasptournaments.org

## Getting Help

View the built-in help:

```powershell
Get-Help .\Get-NASPScoreSheet.ps1 -Detailed
Get-Help .\Get-NASPScoreSheet.ps1 -Examples
Get-Help .\Get-NASPScoreSheet.ps1 -Full
```

## System Requirements

- **PowerShell**: 5.1 or higher (Windows PowerShell or PowerShell Core)
- **Operating System**: Windows, macOS, or Linux
- **Network**: Internet connection required
- **Credentials**: Valid nasptournaments.org account
