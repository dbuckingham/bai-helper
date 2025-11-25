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
$WebSession = $null

try {
    Write-Host "Connecting to NASP Tournaments..." -ForegroundColor Yellow
    
    # First, get the login page to retrieve any necessary tokens (like __VIEWSTATE)
    $LoginPageResponse = Invoke-WebRequest -Uri $LoginUrl -SessionVariable WebSession -UseBasicParsing
    
    # Parse the form to get __VIEWSTATE, __VIEWSTATEGENERATOR, __EVENTVALIDATION
    $ViewState = ""
    $ViewStateGenerator = ""
    $EventValidation = ""
    
    # Try to extract hidden fields from the login page
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
        "__VIEWSTATE" = $ViewState
        "__VIEWSTATEGENERATOR" = $ViewStateGenerator
        "__EVENTVALIDATION" = $EventValidation
        "ctl00`$ContentPlaceHolder1`$txtUserName" = $Username
        "ctl00`$ContentPlaceHolder1`$txtPassword" = $Password
        "ctl00`$ContentPlaceHolder1`$btnLogin" = "Login"
    }
    
    Write-Host "Logging in as $Username..." -ForegroundColor Yellow
    
    # Submit the login form
    $LoginResponse = Invoke-WebRequest -Uri $LoginUrl -Method POST -Body $LoginBody -WebSession $WebSession -UseBasicParsing
    
    # Check if login was successful by looking for common indicators
    if ($LoginResponse.Content -match "Invalid|incorrect|failed|error" -and $LoginResponse.Content -notmatch "logout|welcome|dashboard") {
        Write-Error "Login may have failed. Please check your credentials."
        exit 1
    }
    
    Write-Host "Login successful!" -ForegroundColor Green
    
    # Navigate to the Season Score Sheet page
    Write-Host "Navigating to Season Score Sheet page..." -ForegroundColor Yellow
    $ScoreSheetResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -WebSession $WebSession -UseBasicParsing
    
    # Extract school name from the page
    $SchoolName = "Unknown School"
    if ($ScoreSheetResponse.Content -match '<span[^>]*id="[^"]*lblSchoolName[^"]*"[^>]*>([^<]+)</span>') {
        $SchoolName = $Matches[1].Trim()
    } elseif ($ScoreSheetResponse.Content -match '<h\d[^>]*>([^<]*School[^<]*)</h\d>') {
        $SchoolName = $Matches[1].Trim()
    } elseif ($ScoreSheetResponse.Content -match 'class="school[^"]*"[^>]*>([^<]+)<') {
        $SchoolName = $Matches[1].Trim()
    }
    
    # Clean the school name for use in file paths
    $SchoolNameClean = $SchoolName -replace '[<>:"/\\|?*]', '_'
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
    
    # If a specific season was requested, select it from the dropdown
    if ($Season) {
        Write-Host "Selecting season: $Season..." -ForegroundColor Yellow
        
        # Find the dropdown and its options
        # Look for select element with season options
        $SeasonDropdownId = ""
        if ($ScoreSheetResponse.Content -match '<select[^>]*id="([^"]*Season[^"]*)"') {
            $SeasonDropdownId = $Matches[1]
        } elseif ($ScoreSheetResponse.Content -match '<select[^>]*id="([^"]*ddl[^"]*)"') {
            $SeasonDropdownId = $Matches[1]
        }
        
        # Find the season value
        $SeasonValue = ""
        if ($ScoreSheetResponse.Content -match "<option[^>]*value=`"([^`"]*)`"[^>]*>$Season</option>") {
            $SeasonValue = $Matches[1]
        } elseif ($ScoreSheetResponse.Content -match "<option[^>]*value=`"([^`"]*)`"[^>]*>[^<]*$Season[^<]*</option>") {
            $SeasonValue = $Matches[1]
        }
        
        if ($SeasonValue -and $SeasonDropdownId) {
            # Convert ASP.NET client ID to server control ID format for postback
            $SeasonDropdownName = $SeasonDropdownId -replace '_', '`$'
            
            $SeasonSelectBody = @{
                "__VIEWSTATE" = $ScoreSheetViewState
                "__VIEWSTATEGENERATOR" = $ScoreSheetViewStateGenerator
                "__EVENTVALIDATION" = $ScoreSheetEventValidation
                "__EVENTTARGET" = $SeasonDropdownName
                "__EVENTARGUMENT" = ""
                $SeasonDropdownName = $SeasonValue
            }
            
            $ScoreSheetResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -Method POST -Body $SeasonSelectBody -WebSession $WebSession -UseBasicParsing
            
            # Re-extract form fields after season selection
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
        }
    }
    
    # Find and click the export button
    Write-Host "Exporting score sheet to CSV..." -ForegroundColor Yellow
    
    # Look for export/download button
    $ExportButtonName = ""
    if ($ScoreSheetResponse.Content -match '<input[^>]*id="([^"]*(?:btnExport|btnDownload|btnCSV|Export)[^"]*)"[^>]*type="submit"') {
        $ExportButtonName = $Matches[1] -replace '_', '`$'
    } elseif ($ScoreSheetResponse.Content -match '<input[^>]*type="submit"[^>]*id="([^"]*(?:btnExport|btnDownload|btnCSV|Export)[^"]*)"') {
        $ExportButtonName = $Matches[1] -replace '_', '`$'
    } elseif ($ScoreSheetResponse.Content -match '<a[^>]*id="([^"]*(?:lnkExport|lnkDownload|Export)[^"]*)"') {
        $ExportButtonName = $Matches[1] -replace '_', '`$'
    }
    
    # If no explicit export button, look for any CSV-related link or button
    if (-not $ExportButtonName) {
        # Try alternative patterns
        if ($ScoreSheetResponse.Content -match 'name="([^"]*Export[^"]*)"') {
            $ExportButtonName = $Matches[1]
        } elseif ($ScoreSheetResponse.Content -match 'id="([^"]*csv[^"]*)"') {
            $ExportButtonName = $Matches[1] -replace '_', '`$'
        }
    }
    
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
    
    if ($ContentDisposition -match 'filename="?([^";]+)"?') {
        $FileName = $Matches[1]
    }
    
    # Check if the response is CSV data
    if ($ContentType -match "text/csv|application/csv|application/octet-stream" -or $ContentDisposition) {
        $CsvContent = $ExportResponse.Content
    } elseif ($ExportResponse.Content -match "^[\w\s,`"]+\r?\n") {
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
        
        # Extract table data from the page
        $TableData = @()
        
        # Find the main data table
        if ($ScoreSheetResponse.Content -match '<table[^>]*class="[^"]*(?:grid|data|score)[^"]*"[^>]*>(.*?)</table>') {
            $TableHtml = $Matches[1]
            
            # Extract headers
            $Headers = @()
            [regex]::Matches($TableHtml, '<th[^>]*>([^<]*)</th>') | ForEach-Object {
                $Headers += $_.Groups[1].Value.Trim()
            }
            
            # Extract rows
            [regex]::Matches($TableHtml, '<tr[^>]*>(.*?)</tr>') | ForEach-Object {
                $RowHtml = $_.Groups[1].Value
                $RowData = @()
                [regex]::Matches($RowHtml, '<td[^>]*>([^<]*)</td>') | ForEach-Object {
                    $RowData += $_.Groups[1].Value.Trim()
                }
                
                if ($RowData.Count -gt 0) {
                    $RowObject = [PSCustomObject]@{}
                    for ($i = 0; $i -lt [Math]::Min($Headers.Count, $RowData.Count); $i++) {
                        $RowObject | Add-Member -NotePropertyName $Headers[$i] -NotePropertyValue $RowData[$i]
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
