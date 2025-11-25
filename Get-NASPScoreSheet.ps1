<#
.SYNOPSIS
    Downloads and processes NASP tournament score sheet data.

.DESCRIPTION
    This script logs into nasptournaments.org, downloads the 2025-2026 season score sheet,
    calculates team scores, and exports the data to a CSV file.

.PARAMETER Credential
    PSCredential object containing username and password for nasptournaments.org.
    If not provided, the user will be prompted to enter credentials.

.PARAMETER OutputPath
    Path to save the CSV file. Defaults to "NASPScoreSheet.csv" in the current directory.

.PARAMETER SeasonId
    The season ID to download. Defaults to 5232 (2025-2026 season).

.EXAMPLE
    .\Get-NASPScoreSheet.ps1
    Prompts for credentials and downloads the score sheet to the default location.

.EXAMPLE
    $cred = Get-Credential
    .\Get-NASPScoreSheet.ps1 -Credential $cred -OutputPath "C:\Data\scores.csv"
    Uses provided credentials and saves to a custom location.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "NASPScoreSheet.csv",
    
    [Parameter(Mandatory=$false)]
    [int]$SeasonId = 5232
)

# If no credential provided, prompt user
if (-not $Credential) {
    Write-Host "Please enter your nasptournaments.org credentials:" -ForegroundColor Cyan
    $Credential = Get-Credential -Message "Enter your nasptournaments.org credentials"
    
    if (-not $Credential) {
        Write-Error "Credentials are required to proceed."
        exit 1
    }
}

# Extract username and password from credential
$username = $Credential.UserName
$password = $Credential.GetNetworkCredential().Password

Write-Host "Connecting to nasptournaments.org..." -ForegroundColor Green

# Create a web session to maintain cookies
$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

try {
    # Step 1: Navigate to the login page to get any initial cookies/viewstate
    Write-Host "Accessing login page..." -ForegroundColor Yellow
    $loginPageUrl = "https://nasptournaments.org/userutilities/login.aspx"
    $loginPage = Invoke-WebRequest -Uri $loginPageUrl -SessionVariable session -UseBasicParsing
    
    # Parse the login page to get ViewState and other hidden fields
    $viewState = $null
    $viewStateGenerator = $null
    $eventValidation = $null
    
    if ($loginPage.Content -match '__VIEWSTATE"\s+value="([^"]+)"') {
        $viewState = $matches[1]
    }
    if ($loginPage.Content -match '__VIEWSTATEGENERATOR"\s+value="([^"]+)"') {
        $viewStateGenerator = $matches[1]
    }
    if ($loginPage.Content -match '__EVENTVALIDATION"\s+value="([^"]+)"') {
        $eventValidation = $matches[1]
    }
    
    # Step 2: Post login credentials
    Write-Host "Logging in as $username..." -ForegroundColor Yellow
    
    # Build the login form data
    $loginBody = @{
        '__VIEWSTATE' = $viewState
        '__VIEWSTATEGENERATOR' = $viewStateGenerator
        '__EVENTVALIDATION' = $eventValidation
        'ctl00$ContentPlaceHolder1$txtUserName' = $username
        'ctl00$ContentPlaceHolder1$txtPassword' = $password
        'ctl00$ContentPlaceHolder1$btnLogin' = 'Login'
    }
    
    # Perform login
    $loginResponse = Invoke-WebRequest -Uri $loginPageUrl -Method Post -Body $loginBody -WebSession $session -UseBasicParsing
    
    # Check if login was successful (look for redirect or specific success indicators)
    if ($loginResponse.Content -match "Invalid|Error|failed" -or $loginResponse.StatusCode -ne 200) {
        Write-Error "Login failed. Please check your credentials."
        exit 1
    }
    
    Write-Host "Login successful!" -ForegroundColor Green
    
    # Step 3: Navigate to the score sheet page
    Write-Host "Accessing score sheet page..." -ForegroundColor Yellow
    $scoreSheetUrl = "https://nasptournaments.org/Schoolmgr/SeasonScoreSheet.aspx?oid=$SeasonId"
    $scoreSheetPage = Invoke-WebRequest -Uri $scoreSheetUrl -WebSession $session -UseBasicParsing
    
    # Step 4: Parse the score sheet data
    Write-Host "Parsing score sheet data..." -ForegroundColor Yellow
    
    # The score sheet is likely in an HTML table. We need to parse it.
    # We'll look for table rows and extract archer data
    $content = $scoreSheetPage.Content
    
    # Initialize array to store archer data
    $archers = @()
    $teamScore = 0
    
    # Parse HTML table - looking for data patterns
    # This regex will find table rows with archer data
    $tableRowPattern = '<tr[^>]*>.*?</tr>'
    $tableRows = [regex]::Matches($content, $tableRowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    
    foreach ($row in $tableRows) {
        $rowContent = $row.Value
        
        # Extract cell data from the row
        $cellPattern = '<td[^>]*>(.*?)</td>'
        $cells = [regex]::Matches($rowContent, $cellPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        
        if ($cells.Count -gt 0) {
            # Try to parse archer data
            # Expected columns might include: Name, Score, Use For Team, etc.
            # We'll need to identify the structure
            
            $cellValues = @()
            foreach ($cell in $cells) {
                $cellText = $cell.Groups[1].Value -replace '<[^>]+>', '' -replace '&nbsp;', '' -replace '\s+', ' '
                $cellText = $cellText.Trim()
                $cellValues += $cellText
            }
            
            # Skip header rows and empty rows
            if ($cellValues.Count -ge 3 -and $cellValues[0] -notmatch '^(Name|Archer|#)$' -and $cellValues[0] -ne '') {
                # Try to find score and "Use For Team" indicator
                # Common patterns: Name, Score, UseForTeam or similar
                
                # Look for numeric score
                $score = $null
                $useForTeam = $null
                $name = $cellValues[0]
                
                foreach ($value in $cellValues) {
                    # Check if this is a score (numeric)
                    if ($value -match '^\d+$') {
                        $score = [int]$value
                    }
                    # Check if this is Use For Team indicator
                    if ($value -match '^[YN]$') {
                        $useForTeam = $value
                    }
                }
                
                if ($score -ne $null -and $useForTeam -ne $null) {
                    $archerData = [PSCustomObject]@{
                        Name = $name
                        Score = $score
                        UseForTeam = $useForTeam
                    }
                    $archers += $archerData
                    
                    # Add to team score if Use For Team is Y
                    if ($useForTeam -eq 'Y') {
                        $teamScore += $score
                    }
                }
            }
        }
    }
    
    # If we didn't find data in tables, try alternative parsing
    if ($archers.Count -eq 0) {
        Write-Host "Trying alternative data parsing method..." -ForegroundColor Yellow
        
        # Look for JSON data or JavaScript variables that might contain the data
        if ($content -match 'var\s+scoreData\s*=\s*(\[.*?\]);') {
            $jsonData = $matches[1]
            $archers = ConvertFrom-Json $jsonData
            
            # Calculate team score
            foreach ($archer in $archers) {
                if ($archer.UseForTeam -eq 'Y') {
                    $teamScore += $archer.Score
                }
            }
        }
        else {
            # Try to parse grid view or other controls
            # Look for specific ASP.NET control patterns
            Write-Warning "Could not automatically parse score sheet data. Saving raw HTML for manual review."
            
            # Save raw HTML for debugging
            $content | Out-File -FilePath "scoresheet_raw.html" -Encoding UTF8
            Write-Host "Raw HTML saved to scoresheet_raw.html" -ForegroundColor Cyan
        }
    }
    
    # Step 5: Display results and export to CSV
    if ($archers.Count -gt 0) {
        Write-Host "`nFound $($archers.Count) archers in the score sheet" -ForegroundColor Green
        Write-Host "Team Score (sum of archers with Use For Team = Y): $teamScore" -ForegroundColor Green
        
        # Add team score as a summary row
        $summaryData = $archers + [PSCustomObject]@{
            Name = "TEAM TOTAL"
            Score = $teamScore
            UseForTeam = "-"
        }
        
        # Export to CSV
        $summaryData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "`nData exported to: $OutputPath" -ForegroundColor Green
        
        # Display preview
        Write-Host "`nPreview of data:" -ForegroundColor Cyan
        $archers | Format-Table -AutoSize
        Write-Host "`nTeam Score: $teamScore" -ForegroundColor Magenta
    }
    else {
        Write-Warning "No archer data could be parsed from the score sheet."
        Write-Host "Please check scoresheet_raw.html for the raw data." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "An error occurred: $_"
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
finally {
    Write-Host "`nScript completed." -ForegroundColor Green
}
