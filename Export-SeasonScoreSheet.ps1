<#
.SYNOPSIS
    Exports a season score sheet from NASP Tournaments website to a CSV file.

.DESCRIPTION
    This script logs into the NASP Tournaments website using provided or prompted credentials,
    navigates to the Season Score Sheet page, and exports the data to a CSV file.
    The exported file is saved in a "Season Score Sheets/{SchoolName}" folder structure.

.PARAMETER Credential
    A PSCredential object containing the username and password for login.
    If not provided, the user will be prompted to enter credentials.

.PARAMETER Season
    Optional. The season to select from the dropdown menu on the page.
    If not provided, the default (current) season will be used.

.PARAMETER OrganizationId
    Optional. The organization ID to use in the URL. Defaults to 5232.

.PARAMETER OutputPath
    Optional. The base path where the "Season Score Sheets" folder will be created.
    Defaults to the current directory.

.EXAMPLE
    .\Export-SeasonScoreSheet.ps1
    Prompts for credentials and exports the default season score sheet.

.EXAMPLE
    .\Export-SeasonScoreSheet.ps1 -Credential (Get-Credential)
    Uses provided credentials to export the default season score sheet.

.EXAMPLE
    .\Export-SeasonScoreSheet.ps1 -Season "2023-2024"
    Exports the score sheet for the specified season.

.NOTES
    Author: BAI Helper
    Requires: PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [string]
    $Season,

    [Parameter(Mandatory = $false)]
    [int]
    $OrganizationId = 5232,

    [Parameter(Mandatory = $false)]
    [string]
    $OutputPath = (Get-Location).Path
)

# Base URLs
$LoginUrl = "https://nasptournaments.org/userutilities/login.aspx"
$ScoreSheetUrl = "https://nasptournaments.org/Schoolmgr/SeasonScoreSheet.aspx?oid=$OrganizationId"

# If credentials not provided, prompt the user
if (-not $Credential) {
    Write-Host "Please enter your NASP Tournaments credentials:" -ForegroundColor Cyan
    $Credential = Get-Credential -Message "Enter your NASP Tournaments username and password"
    
    if (-not $Credential) {
        Write-Error "Credentials are required to proceed."
        exit 1
    }
}

# Extract username and password from credential
$Username = $Credential.UserName
$Password = $Credential.GetNetworkCredential().Password

# Create a web session to maintain cookies
$WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

try {
    Write-Host "Connecting to NASP Tournaments..." -ForegroundColor Yellow
    
    # First, get the login page to retrieve any necessary tokens (like __VIEWSTATE)
    $LoginPageResponse = Invoke-WebRequest -Uri $LoginUrl -SessionVariable WebSession -UseBasicParsing
    
    # Parse the form to get __EVENTTARGET, __EVENTARGUMENT, __VIEWSTATE, __VIEWSTATEGENERATOR, __EVENTVALIDATION
    $EventTarget = ""
    $EventArgument = ""
    $ViewState = ""
    $ViewStateGenerator = ""
    $EventValidation = ""
    
    # Try to extract hidden fields from the login page
    if ($LoginPageResponse.Content -match 'id="__EVENTTARGET"\s+value="([^"]*)"') {
        $EventTarget = $Matches[1]
    }
    if ($LoginPageResponse.Content -match 'id="__EVENTARGUMENT"\s+value="([^"]*)"') {
        $EventArgument = $Matches[1]
    }
    if ($LoginPageResponse.Content -match 'id="__VIEWSTATE"\s+value="([^"]*)"') {
        $ViewState = $Matches[1]
    }
    if ($LoginPageResponse.Content -match 'id="__VIEWSTATEGENERATOR"\s+value="([^"]*)"') {
        $ViewStateGenerator = $Matches[1]
    }
    if ($LoginPageResponse.Content -match 'id="__EVENTVALIDATION"\s+value="([^"]*)"') {
        $EventValidation = $Matches[1]
    }

    # Prepare login form data
    # Note: The actual field names may vary - these are common ASP.NET patterns
    $LoginBody = @{
        "__EVENTTARGET" = $EventTarget
        "__EVENTARGUMENT" = $EventArgument
        "__VIEWSTATE" = $ViewState
        "__VIEWSTATEGENERATOR" = $ViewStateGenerator
        "__EVENTVALIDATION" = $EventValidation
        "ctl00`$ContentPlaceHolder1`$TextBox_username" = $Username
        "ctl00`$ContentPlaceHolder1`$TextBox_password" = $Password
        "ctl00`$ContentPlaceHolder1`$Button_login" = "Sign In"
    }
    
    Write-Host "Logging in as $Username..." -ForegroundColor Yellow
    
    # Submit the login form
    $LoginResponse = Invoke-WebRequest -Uri $LoginUrl -Method POST -Body $LoginBody -WebSession $WebSession -UseBasicParsing
    
    # Check if login was successful using the HTTP status code
    if ($LoginResponse.StatusCode -ne 200) {
        Write-Error "Login failed with status code: $($LoginResponse.StatusCode). Please check your credentials."
        exit 1
    }
    
    Write-Host "Login successful!" -ForegroundColor Green

    # Navigate to the Season Score Sheet page
    Write-Host "Navigating to Season Score Sheet page..." -ForegroundColor Yellow
    $ScoreSheetResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -WebSession $WebSession -UseBasicParsing

    # Extract school name from the page by searching for the Label_school_name element
    $SchoolName = "Unknown School"
    if ($ScoreSheetResponse.Content -match '<span[^>]*id="ctl00_ContentPlaceHolder1_Label_school_name"[^>]*>([^<]+)</span>') {
        $SchoolName = $Matches[1].Trim()
    } elseif ($ScoreSheetResponse.Content -match 'id="[^"]*Label_school_name[^"]*"[^>]*>([^<]+)<') {
        $SchoolName = $Matches[1].Trim()
    } elseif ($ScoreSheetResponse.Content -match '<span[^>]*id="[^"]*lblSchoolName[^"]*"[^>]*>([^<]+)</span>') {
        $SchoolName = $Matches[1].Trim()
    }
    
    # Clean the school name for use in file paths
    $SchoolNameClean = $SchoolName -replace '[<>:"/\\|?*\[\]]', '_'
    $SchoolNameClean = $SchoolNameClean.Trim()
    if ([string]::IsNullOrWhiteSpace($SchoolNameClean)) {
        $SchoolNameClean = "Organization_$OrganizationId"
    }
    
    Write-Host "School: $SchoolName" -ForegroundColor Cyan
    
    # Extract form fields for postback
    $ScoreSheetViewState = ""
    $ScoreSheetViewStateGenerator = ""
    $ScoreSheetEventValidation = ""
    
    if ($ScoreSheetResponse.Content -match 'id="__VIEWSTATE"\s+value="([^"]*)"') {
        $ScoreSheetViewState = $Matches[1]
    }
    if ($ScoreSheetResponse.Content -match 'id="__VIEWSTATEGENERATOR"\s+value="([^"]*)"') {
        $ScoreSheetViewStateGenerator = $Matches[1]
    }
    if ($ScoreSheetResponse.Content -match 'id="__EVENTVALIDATION"\s+value="([^"]*)"') {
        $ScoreSheetEventValidation = $Matches[1]
    }
    
    # Find the season dropdown and extract default value if not provided
    $SeasonDropdownId = ""
    if ($ScoreSheetResponse.Content -match '<select[^>]*id="([^"]*Season[^"]*)"') {
        $SeasonDropdownId = $Matches[1]
    } elseif ($ScoreSheetResponse.Content -match '<select[^>]*id="([^"]*ddl[^"]*)"') {
        $SeasonDropdownId = $Matches[1]
    }
    
    # If Season parameter not provided, extract the default selected value from the dropdown
    $DefaultSeason = ""
    if ($ScoreSheetResponse.Content -match '<option[^>]*selected[^>]*>([^<]+)</option>') {
        $DefaultSeason = $Matches[1].Trim()
    } elseif ($ScoreSheetResponse.Content -match '<select[^>]*id="[^"]*Season[^"]*"[^>]*>[\s\S]*?<option[^>]*value="[^"]*"[^>]*>([^<]+)</option>') {
        # Fall back to first option if no selected attribute
        $DefaultSeason = $Matches[1].Trim()
    }
    
    if (-not $Season) {
        $Season = $DefaultSeason
        Write-Host "Using default season: $Season" -ForegroundColor Cyan
    }
    
    # Only perform postback if user specified a different season than the default
    $NeedsPostback = $Season -and ($Season -ne $DefaultSeason)
    
    if ($NeedsPostback) {
        Write-Host "Selecting season: $Season..." -ForegroundColor Yellow
        
        # Find the season value
        $SeasonValue = ""
        $SeasonPattern1 = '<option[^>]*value="([^"]*)"[^>]*>' + [regex]::Escape($Season) + '</option>'
        $SeasonPattern2 = '<option[^>]*value="([^"]*)"[^>]*>[^<]*' + [regex]::Escape($Season) + '[^<]*</option>'
        if ($ScoreSheetResponse.Content -match $SeasonPattern1) {
            $SeasonValue = $Matches[1]
        } elseif ($ScoreSheetResponse.Content -match $SeasonPattern2) {
            $SeasonValue = $Matches[1]
        }
        
        if ($SeasonValue -and $SeasonDropdownId) {
            # Convert ASP.NET client ID to server control ID format for form field name
            $SeasonDropdownName = $SeasonDropdownId -replace '_', '`$'
            
            # Perform postback to update the data table with the selected season
            $SeasonSelectBody = @{
                "__VIEWSTATE" = $ScoreSheetViewState
                "__VIEWSTATEGENERATOR" = $ScoreSheetViewStateGenerator
                "__EVENTVALIDATION" = $ScoreSheetEventValidation
                "__EVENTTARGET" = $SeasonDropdownId  # Use client ID format for __EVENTTARGET
                "__EVENTARGUMENT" = ""
                $SeasonDropdownName = $SeasonValue
            }
            
            Write-Host "Performing postback to update data table..." -ForegroundColor Yellow
            $ScoreSheetResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -Method POST -Body $SeasonSelectBody -WebSession $WebSession -UseBasicParsing
            
            # Re-extract form fields after season selection postback
            if ($ScoreSheetResponse.Content -match 'id="__VIEWSTATE"\s+value="([^"]*)"') {
                $ScoreSheetViewState = $Matches[1]
            }
            if ($ScoreSheetResponse.Content -match 'id="__VIEWSTATEGENERATOR"\s+value="([^"]*)"') {
                $ScoreSheetViewStateGenerator = $Matches[1]
            }
            if ($ScoreSheetResponse.Content -match 'id="__EVENTVALIDATION"\s+value="([^"]*)"') {
                $ScoreSheetEventValidation = $Matches[1]
            }
            
            Write-Host "Season selected: $Season" -ForegroundColor Green
        } else {
            Write-Warning "Could not find season '$Season' in dropdown. Using default season."
            $Season = $DefaultSeason
        }
    } else {
        Write-Host "Using season: $Season" -ForegroundColor Cyan
    }
    
    # Find and click the export button
    Write-Host "Exporting score sheet to CSV..." -ForegroundColor Yellow
    
    $ExportButtonName = "ctl00`$ContentPlaceHolder1`$Button_export"

    # Build export request body
    $ExportBody = @{
        "__VIEWSTATE" = $ScoreSheetViewState
        "__VIEWSTATEGENERATOR" = $ScoreSheetViewStateGenerator
        "__EVENTVALIDATION" = $ScoreSheetEventValidation
    }
    
    if ($ExportButtonName) {
        $ExportBody[$ExportButtonName] = "Export"
    } else {
        # Try common ASP.NET button names
        $ExportBody["ctl00`$ContentPlaceHolder1`$btnExport"] = "Export"
    }
    
    # Make the export request
    $ExportResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -Method POST -Body $ExportBody -WebSession $WebSession -UseBasicParsing
    
    # Check if we got CSV content
    $ContentType = $ExportResponse.Headers["Content-Type"]
    $ContentDisposition = $ExportResponse.Headers["Content-Disposition"]
    
    $CsvContent = $null
    $FileName = "SeasonScoreSheet.csv"
    
    if (($ContentDisposition -join " ") -match 'filename="?([^";]+)"?') {
        $FileName = $Matches[1]
    }
    
    # Check if the response is CSV data
    if ($ContentType -match "text/csv|application/csv|application/octet-stream" -or $ContentDisposition) {
        $CsvContent = $ExportResponse.Content
    } elseif ($ExportResponse.Content -match '^[\w\s,"]+\r?\n') {
        # Response might be CSV without proper headers
        $CsvContent = $ExportResponse.Content
    }
    
    if ($CsvContent) {
        # Create output directory structure
        $OutputFolder = Join-Path -Path $OutputPath -ChildPath "Season Score Sheets"
        $SchoolFolder = Join-Path -Path $OutputFolder -ChildPath $SchoolNameClean
        
        if (-not (Test-Path -Path $SchoolFolder)) {
            New-Item -Path $SchoolFolder -ItemType Directory -Force | Out-Null
            Write-Host "Created folder: $SchoolFolder" -ForegroundColor Cyan
        }
        
        # Generate filename with timestamp
        $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $SeasonSuffix = if ($Season) { "_$($Season -replace '\s+', '_')" } else { "" }
        $OutputFileName = "SeasonScoreSheet${SeasonSuffix}_$Timestamp.csv"
        $OutputFilePath = Join-Path -Path $SchoolFolder -ChildPath $OutputFileName
        
        # Save the CSV file
        if ($CsvContent -is [byte[]]) {
            [System.IO.File]::WriteAllBytes($OutputFilePath, $CsvContent)
        } else {
            $CsvContent | Out-File -FilePath $OutputFilePath -Encoding UTF8 -Force
        }
        
        Write-Host "Score sheet exported successfully!" -ForegroundColor Green
        Write-Host "File saved to: $OutputFilePath" -ForegroundColor Cyan
        
        return $OutputFilePath
    } else {
        # If no CSV was returned, the page might use a different export mechanism
        # Try to parse the HTML table directly
        Write-Host "Direct export not available. Attempting to parse table data..." -ForegroundColor Yellow
        
        # Helper function to strip HTML tags and decode entities
        function Get-CleanText {
            param([string]$HtmlText)
            # Remove HTML tags
            $text = $HtmlText -replace '<[^>]+>', ''
            # Decode common HTML entities
            $text = $text -replace '&nbsp;', ' '
            $text = $text -replace '&amp;', '&'
            $text = $text -replace '&lt;', '<'
            $text = $text -replace '&gt;', '>'
            $text = $text -replace '&quot;', '"'
            $text = $text -replace '&#(\d+);', { [char][int]$_.Groups[1].Value }
            return $text.Trim()
        }
        
        # Extract table data from the page
        $TableData = @()
        
        # Find the main data table - try multiple patterns
        $TablePattern = '<table[^>]*(?:class="[^"]*(?:grid|data|score)[^"]*"|id="[^"]*(?:gv|grid|tbl)[^"]*")[^>]*>([\s\S]*?)</table>'
        if ($ScoreSheetResponse.Content -match $TablePattern) {
            $TableHtml = $Matches[1]
            
            # Extract headers - handle content that may contain nested tags
            $Headers = @()
            [regex]::Matches($TableHtml, '<th[^>]*>([\s\S]*?)</th>') | ForEach-Object {
                $Headers += Get-CleanText $_.Groups[1].Value
            }
            
            # Extract rows - handle content that may contain nested tags
            [regex]::Matches($TableHtml, '<tr[^>]*>([\s\S]*?)</tr>') | ForEach-Object {
                $RowHtml = $_.Groups[1].Value
                $RowData = @()
                [regex]::Matches($RowHtml, '<td[^>]*>([\s\S]*?)</td>') | ForEach-Object {
                    $RowData += Get-CleanText $_.Groups[1].Value
                }
                
                if ($RowData.Count -gt 0) {
                    $RowObject = [PSCustomObject]@{}
                    for ($i = 0; $i -lt [Math]::Min($Headers.Count, $RowData.Count); $i++) {
                        $HeaderName = if ([string]::IsNullOrWhiteSpace($Headers[$i])) { "Column$i" } else { $Headers[$i] }
                        $RowObject | Add-Member -NotePropertyName $HeaderName -NotePropertyValue $RowData[$i]
                    }
                    $TableData += $RowObject
                }
            }
        }
        
        if ($TableData.Count -gt 0) {
            # Create output directory structure
            $OutputFolder = Join-Path -Path $OutputPath -ChildPath "Season Score Sheets"
            $SchoolFolder = Join-Path -Path $OutputFolder -ChildPath $SchoolNameClean
            
            if (-not (Test-Path -Path $SchoolFolder)) {
                New-Item -Path $SchoolFolder -ItemType Directory -Force | Out-Null
                Write-Host "Created folder: $SchoolFolder" -ForegroundColor Cyan
            }
            
            # Generate filename with timestamp
            $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $SeasonSuffix = if ($Season) { "_$($Season -replace '\s+', '_')" } else { "" }
            $OutputFileName = "SeasonScoreSheet${SeasonSuffix}_$Timestamp.csv"
            $OutputFilePath = Join-Path -Path $SchoolFolder -ChildPath $OutputFileName
            
            # Export to CSV
            $TableData | Export-Csv -Path $OutputFilePath -NoTypeInformation -Encoding UTF8
            
            Write-Host "Score sheet exported successfully!" -ForegroundColor Green
            Write-Host "File saved to: $OutputFilePath" -ForegroundColor Cyan
            
            return $OutputFilePath
        } else {
            Write-Warning "Could not extract score sheet data. The page structure may have changed."
            Write-Host "Please verify the page manually and update the script if needed." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "An error occurred: $_"
    Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Response) {
        Write-Host "Response status: $($_.Exception.Response.StatusCode)" -ForegroundColor Red
    }
    exit 1
}
