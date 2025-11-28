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
    .\Export-SeasonScoreSheet-Refactored.ps1
    Prompts for credentials and exports the default season score sheet.

.EXAMPLE
    .\Export-SeasonScoreSheet-Refactored.ps1 -Credential (Get-Credential)
    Uses provided credentials to export the default season score sheet.

.EXAMPLE
    .\Export-SeasonScoreSheet-Refactored.ps1 -Season "2023-2024"
    Exports the score sheet for the specified season.

.NOTES
    Author: BAI Helper
    Version: 2.0 (Refactored)
    Requires: PowerShell 5.1 or later
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [System.Management.Automation.PSCredential]
    $Credential,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]
    $Season,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 99999)]
    [int]
    $OrganizationId = 5232,

    [Parameter(Mandatory = $false)]
    [ValidateScript({Test-Path $_ -IsValid})]
    [string]
    $OutputPath = (Get-Location).Path
)

#region Configuration
# Application Configuration
$Script:Config = @{
    # Base URLs
    BaseUrl = "https://nasptournaments.org"
    LoginPath = "/userutilities/login.aspx"
    ScoreSheetPath = "/Schoolmgr/SeasonScoreSheet.aspx"
    
    # ASP.NET Control IDs
    Controls = @{
        Username = "ctl00`$ContentPlaceHolder1`$TextBox_username"
        Password = "ctl00`$ContentPlaceHolder1`$TextBox_password"
        ExportButton = "ctl00`$ContentPlaceHolder1`$Button_export"
        ReturnToSchoolButton = "ctl00`$ContentPlaceHolder1`$Button_return_school"
        SchoolLabel = "ctl00_ContentPlaceHolder1_Label_school_name"
    }
    
    # Regex Patterns for HTML parsing
    Patterns = @{
        SchoolName = '<span[^>]*id="ctl00_ContentPlaceHolder1_Label_school_name"[^>]*>([^<]+)</span>'
        SeasonDropdown1 = '<select[^>]*id="([^"]*Season[^"]*)"'
        SeasonDropdown2 = '<select[^>]*id="([^"]*ddl[^"]*)"'
        DefaultSeason1 = '<option[^>]*selected[^>]*>([^<]+)</option>'
        DefaultSeason2 = '<select[^>]*id="[^"]*Season[^"]*"[^>]*>[\s\S]*?<option[^>]*value="[^"]*"[^>]*>([^<]+)</option>'
        DataTable = '<table[^>]*(?:class="[^"]*(?:grid|data|score)[^"]*"|id="[^"]*(?:gv|grid|tbl)[^"]*")[^>]*>([\s\S]*?)</table>'
        TableHeader = '<th[^>]*>([\s\S]*?)</th>'
        TableRow = '<tr[^>]*>([\s\S]*?)</tr>'
        TableCell = '<td[^>]*>([\s\S]*?)</td>'
        CsvContent = '^[\w\s,"]+\r?\n'
    }
    
    # File and path settings
    FileSettings = @{
        InvalidPathChars = '[<>:"/\\|?*\[\]]'
        TimestampFormat = "yyyyMMdd_HHmmss"
        Encoding = "UTF8"
    }
}

# Script-scoped variables
$Script:WebSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$Script:LoginUrl = "$($Script:Config.BaseUrl)$($Script:Config.LoginPath)"
$Script:ScoreSheetUrl = "$($Script:Config.BaseUrl)$($Script:Config.ScoreSheetPath)?oid=$OrganizationId"
#endregion

#region Helper Functions
function Write-StatusMessage {
    <#
    .SYNOPSIS
        Writes a standardized status message.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Information', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Information'
    )
    
    $color = switch ($Level) {
        'Information' { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
        'Success' { 'Green' }
        default { 'White' }
    }
    
    Write-Host $Message -ForegroundColor $color
    Write-Verbose $Message
}

function Get-CleanText {
    <#
    .SYNOPSIS
        Strips HTML tags and decodes HTML entities from text.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$HtmlText
    )
    
    if ([string]::IsNullOrWhiteSpace($HtmlText)) {
        return ""
    }
    
    # Remove HTML tags
    $text = $HtmlText -replace '<[^>]+>', ''
    
    # Decode common HTML entities
    $entityMap = @{
        '&nbsp;' = ' '
        '&amp;' = '&'
        '&lt;' = '<'
        '&gt;' = '>'
        '&quot;' = '"'
        '&#39;' = "'"
        '&apos;' = "'"
    }
    
    foreach ($entity in $entityMap.GetEnumerator()) {
        $text = $text -replace [regex]::Escape($entity.Key), $entity.Value
    }
    
    # Handle numeric entities
    $text = $text -replace '&#(\d+);', { [char][int]$_.Groups[1].Value }
    
    return $text.Trim()
}

function Get-SafeFileName {
    <#
    .SYNOPSIS
        Creates a safe filename by removing invalid characters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,
        
        [Parameter(Mandatory = $false)]
        [string]$Replacement = '_'
    )
    
    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return "Unknown"
    }
    
    $safeName = $FileName -replace $Script:Config.FileSettings.InvalidPathChars, $Replacement
    return $safeName.Trim()
}
#endregion

#region Authentication Functions
function Get-UserCredentials {
    <#
    .SYNOPSIS
        Gets or prompts for user credentials.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    if ($Credential) {
        return $Credential
    }
    
    Write-StatusMessage "Please enter your NASP Tournaments credentials:" -Level Information
    $promptedCredential = Get-Credential -Message "Enter your NASP Tournaments username and password"
    
    if (-not $promptedCredential) {
        throw "Credentials are required to proceed."
    }
    
    return $promptedCredential
}

function Initialize-WebSession {
    <#
    .SYNOPSIS
        Initializes the web session by getting the login page.
    #>
    [CmdletBinding()]
    param()
    
    Write-StatusMessage "Connecting to NASP Tournaments..." -Level Information
    
    try {
        $loginPageResponse = Invoke-WebRequest -Uri $Script:LoginUrl -SessionVariable $Script:WebSession -UseBasicParsing -ErrorAction Stop
        return $loginPageResponse
    }
    catch {
        throw "Failed to connect to NASP Tournaments: $($_.Exception.Message)"
    }
}

function Get-FormFields {
    <#
    .SYNOPSIS
        Extracts form fields from a web response.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    $formFields = @{}
    
    foreach ($field in $Response.InputFields) {
        if ($field.name -and $field.name -ne "") {
            $formFields[$field.name] = $field.value
        }
    }
    
    return $formFields
}

function Invoke-Login {
    <#
    .SYNOPSIS
        Performs login to the NASP Tournaments website.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$LoginPageResponse
    )
    
    # Extract form fields from login page
    $formFields = Get-FormFields -Response $LoginPageResponse
    
    # Set credentials
    $formFields[$Script:Config.Controls.Username] = $Credential.UserName
    $formFields[$Script:Config.Controls.Password] = $Credential.GetNetworkCredential().Password

    Write-StatusMessage "Logging in as $($Credential.UserName)..." -Level Information
    
    try {
        $loginResponse = Invoke-WebRequest -Uri $Script:LoginUrl -Method POST -Body $formFields -WebSession $Script:WebSession -UseBasicParsing -ErrorAction Stop
        
        if ($loginResponse.StatusCode -eq 200) {
            Write-StatusMessage "Login successful!" -Level Success
            return $loginResponse
        }
        else {
            throw "Login failed with status code: $($loginResponse.StatusCode)"
        }
    }
    catch {
        throw "Login failed: $($_.Exception.Message)"
    }
}
#endregion

#region Data Extraction Functions
function Get-ScoreSheetPage {
    <#
    .SYNOPSIS
        Navigates to the Season Score Sheet page.
    #>
    [CmdletBinding()]
    param()
    
    Write-StatusMessage "Navigating to Season Score Sheet page..." -Level Information
    
    try {
        $scoreSheetResponse = Invoke-WebRequest -Uri $Script:ScoreSheetUrl -WebSession $Script:WebSession -UseBasicParsing -ErrorAction Stop
        return $scoreSheetResponse
    }
    catch {
        throw "Failed to navigate to score sheet page: $($_.Exception.Message)"
    }
}

function Get-SchoolInformation {
    <#
    .SYNOPSIS
        Extracts school information from the score sheet page.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    $schoolName = "Unknown School"
    
    if ($Response.Content -match $Script:Config.Patterns.SchoolName) {
        $schoolName = $Matches[1].Trim()
    }
    
    $schoolNameClean = Get-SafeFileName -FileName $schoolName
    if ([string]::IsNullOrWhiteSpace($schoolNameClean)) {
        $schoolNameClean = "Organization_$OrganizationId"
    }
    
    Write-StatusMessage "School: $schoolName" -Level Information
    
    return @{
        Name = $schoolName
        SafeName = $schoolNameClean
    }
}

function Get-SeasonInformation {
    <#
    .SYNOPSIS
        Extracts season information from the score sheet page.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response,
        
        [Parameter(Mandatory = $false)]
        [string]$RequestedSeason
    )
    
    # Find season dropdown ID
    $dropdownId = ""
    if ($Response.Content -match $Script:Config.Patterns.SeasonDropdown1) {
        $dropdownId = $Matches[1]
    }
    elseif ($Response.Content -match $Script:Config.Patterns.SeasonDropdown2) {
        $dropdownId = $Matches[1]
    }
    
    # Find default season
    $defaultSeason = ""
    if ($Response.Content -match $Script:Config.Patterns.DefaultSeason1) {
        $defaultSeason = $Matches[1].Trim()
    }
    elseif ($Response.Content -match $Script:Config.Patterns.DefaultSeason2) {
        $defaultSeason = $Matches[1].Trim()
    }

    # Determine selected season
    $selectedSeason = if ($RequestedSeason) { $RequestedSeason } else { $defaultSeason }
    $needsPostback = $RequestedSeason -and ($RequestedSeason -ne $defaultSeason)
    
    return @{
        DropdownId = $dropdownId
        DefaultSeason = $defaultSeason
        SelectedSeason = $selectedSeason
        NeedsPostback = $needsPostback
    }
}

function Set-SeasonSelection {
    <#
    .SYNOPSIS
        Updates the season selection if needed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SeasonInfo,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    if (-not $SeasonInfo.NeedsPostback) {
        return $Response
    }
    
    Write-StatusMessage "Selecting $($SeasonInfo.SelectedSeason) season..." -Level Information
    
    # Find season value
    $seasonValue = ""
    $pattern1 = '<option[^>]*value="([^"]*)"[^>]*>' + [regex]::Escape($SeasonInfo.SelectedSeason) + '</option>'
    $pattern2 = '<option[^>]*value="([^"]*)"[^>]*>[^<]*' + [regex]::Escape($SeasonInfo.SelectedSeason) + '[^<]*</option>'
    
    if ($Response.Content -match $pattern1) {
        $seasonValue = $Matches[1]
    }
    elseif ($Response.Content -match $pattern2) {
        $seasonValue = $Matches[1]
    }
    
    if (-not $seasonValue -or -not $SeasonInfo.DropdownId) {
        Write-StatusMessage "Could not find season '$($SeasonInfo.SelectedSeason)' in dropdown. Using default season." -Level Warning
        $SeasonInfo.SelectedSeason = $SeasonInfo.DefaultSeason
        return $Response
    }
    
    try {
        # Prepare form for postback
        $formFields = Get-FormFields -Response $Response
        $dropdownName = $SeasonInfo.DropdownId -replace '_', '$'
        $dropdownName = $dropdownName -replace '\$season', '_season'
        
        $formFields["__EVENTTARGET"] = $dropdownName
        $formFields["__EVENTARGUMENT"] = ""
        $formFields[$dropdownName] = $seasonValue
        
        # Remove export button to prevent accidental export
        $formFields.Remove($Script:Config.Controls.ExportButton)
        $formFields.Remove($Script:Config.Controls.ReturnToSchoolButton)

        Write-StatusMessage "Performing postback to select season..." -Level Information
        $updatedResponse = Invoke-WebRequest -Uri $Script:ScoreSheetUrl -Method POST -Body $formFields -WebSession $Script:WebSession -UseBasicParsing -ErrorAction Stop
        
        Write-StatusMessage "Season selection successful!" -Level Success

        return $updatedResponse
    }
    catch {
        Write-StatusMessage "Failed to update season selection: $($_.Exception.Message)" -Level Warning
        return $Response
    }
}
#endregion

#region Export Functions
function Export-ScoreSheetData {
    <#
    .SYNOPSIS
        Exports the score sheet data by triggering the export button.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$SeasonInfo
    )
    
    Write-StatusMessage "Exporting score sheet data..." -Level Information
    
    try {
        $formFields = Get-FormFields -Response $Response
        $formFields.Remove($Script:Config.Controls.ReturnToSchoolButton)
        
        # Add season selection if needed
        if ($SeasonInfo.NeedsPostback -and $SeasonInfo.DropdownId) {
            $dropdownName = $SeasonInfo.DropdownId -replace '_', '$'
            $dropdownName = $dropdownName -replace '\$season', '_season'
            
            # Find season value
            $seasonPattern = '<option[^>]*value="([^"]*)"[^>]*>' + [regex]::Escape($SeasonInfo.SelectedSeason) + '</option>'
            if ($Response.Content -match $seasonPattern) {
                $formFields[$dropdownName] = $Matches[1]
            }
        }
        
        # Trigger export
        $formFields[$Script:Config.Controls.ExportButton] = "Export"
        
        $exportResponse = Invoke-WebRequest -Uri $Script:ScoreSheetUrl -Method POST -Body $formFields -WebSession $Script:WebSession -UseBasicParsing -ErrorAction Stop
        return $exportResponse
    }
    catch {
        throw "Failed to export score sheet data: $($_.Exception.Message)"
    }
}

function Test-CsvResponse {
    <#
    .SYNOPSIS
        Tests if the response contains CSV data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    $contentType = $Response.Headers["Content-Type"]
    $contentDisposition = $Response.Headers["Content-Disposition"]
    
    # Check for CSV content type or disposition
    if ($contentType -match "text/csv|application/csv|application/octet-stream" -or $contentDisposition) {
        return $Response.Content
    }
    
    # Check if content looks like CSV
    if ($Response.Content -match $Script:Config.Patterns.CsvContent) {
        return $Response.Content
    }
    
    return $null
}

function Get-TableDataFromHtml {
    <#
    .SYNOPSIS
        Extracts table data from HTML content as fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    Write-StatusMessage "Direct export not available. Parsing HTML table..." -Level Warning
    
    $tableData = @()
    
    if ($Response.Content -match $Script:Config.Patterns.DataTable) {
        $tableHtml = $Matches[1]
        
        # Extract headers
        $headers = @()
        [regex]::Matches($tableHtml, $Script:Config.Patterns.TableHeader) | ForEach-Object {
            $headers += Get-CleanText -HtmlText $_.Groups[1].Value
        }
        
        # Extract rows
        [regex]::Matches($tableHtml, $Script:Config.Patterns.TableRow) | ForEach-Object {
            $rowHtml = $_.Groups[1].Value
            $rowData = @()
            
            [regex]::Matches($rowHtml, $Script:Config.Patterns.TableCell) | ForEach-Object {
                $rowData += Get-CleanText -HtmlText $_.Groups[1].Value
            }
            
            if ($rowData.Count -gt 0) {
                $rowObject = [PSCustomObject]@{}
                for ($i = 0; $i -lt [Math]::Min($headers.Count, $rowData.Count); $i++) {
                    $headerName = if ([string]::IsNullOrWhiteSpace($headers[$i])) { "Column$i" } else { $headers[$i] }
                    $rowObject | Add-Member -NotePropertyName $headerName -NotePropertyValue $rowData[$i]
                }
                $tableData += $rowObject
            }
        }
    }
    
    return $tableData
}
#endregion

#region File Operations
function New-OutputDirectory {
    <#
    .SYNOPSIS
        Creates the output directory structure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BasePath,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$SchoolInfo,
        
        [Parameter(Mandatory = $true)]
        [string]$Season
    )
    
    $baseFolder = Join-Path -Path $BasePath -ChildPath "Season Score Sheets"
    $schoolFolder = Join-Path -Path $baseFolder -ChildPath $SchoolInfo.SafeName
    $seasonFolder = Join-Path -Path $schoolFolder -ChildPath (Get-SafeFileName -FileName $Season)
    
    if (-not (Test-Path -Path $seasonFolder)) {
        New-Item -Path $seasonFolder -ItemType Directory -Force | Out-Null
        Write-StatusMessage "Created folder: $seasonFolder" -Level Information
    }
    
    return $seasonFolder
}

function Save-CsvData {
    <#
    .SYNOPSIS
        Saves CSV content to a file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CsvContent,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,
        
        [Parameter(Mandatory = $true)]
        [string]$School,

        [Parameter(Mandatory = $true)]
        [string]$Season
    )
    
    # $timestamp = Get-Date -Format $Script:Config.FileSettings.TimestampFormat
    $schoolSafe = Get-SafeFileName -FileName $School
    $seasonSafe = Get-SafeFileName -FileName $Season
    $fileName = "SeasonScoreSheet_${schoolSafe}_${seasonSafe}.csv"
    $filePath = Join-Path -Path $OutputDirectory -ChildPath $fileName
    
    try {
        if ($CsvContent -is [byte[]]) {
            [System.IO.File]::WriteAllBytes($filePath, $CsvContent)
        }
        else {
            $CsvContent | Out-File -FilePath $filePath -Encoding $Script:Config.FileSettings.Encoding -Force
        }
        
        Write-StatusMessage "Score sheet exported successfully!" -Level Success
        Write-StatusMessage "File saved to: $filePath" -Level Information
        
        return $filePath
    }
    catch {
        throw "Failed to save CSV file: $($_.Exception.Message)"
    }
}

function Save-TableDataAsCsv {
    <#
    .SYNOPSIS
        Saves table data as CSV file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$TableData,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,
        
        [Parameter(Mandatory = $true)]
        [string]$Season
    )
    
    if ($TableData.Count -eq 0) {
        throw "No table data found to export. The page structure may have changed."
    }
    
    $timestamp = Get-Date -Format $Script:Config.FileSettings.TimestampFormat
    $seasonSafe = Get-SafeFileName -FileName $Season
    $fileName = "SeasonScoreSheet_${seasonSafe}_$timestamp.csv"
    $filePath = Join-Path -Path $OutputDirectory -ChildPath $fileName
    
    try {
        $TableData | Export-Csv -Path $filePath -NoTypeInformation -Encoding $Script:Config.FileSettings.Encoding
        
        Write-StatusMessage "Score sheet exported successfully!" -Level Success
        Write-StatusMessage "File saved to: $filePath" -Level Information
        
        return $filePath
    }
    catch {
        throw "Failed to save table data as CSV: $($_.Exception.Message)"
    }
}
#endregion

#region Main Execution
try {
    # Get credentials
    $credential = Get-UserCredentials -Credential $Credential
    
    # Initialize session and login
    $loginPageResponse = Initialize-WebSession

    Invoke-Login -Credential $credential -LoginPageResponse $loginPageResponse
    
    # Get score sheet page
    $scoreSheetResponse = Get-ScoreSheetPage
    
    # Extract information
    $schoolInfo = Get-SchoolInformation -Response $scoreSheetResponse
    $seasonInfo = Get-SeasonInformation -Response $scoreSheetResponse -RequestedSeason $Season
    
    # Update season selection if needed
    if ($seasonInfo.NeedsPostback) {
        $scoreSheetResponse = Set-SeasonSelection -SeasonInfo $seasonInfo -Response $scoreSheetResponse
    }
    
    Write-StatusMessage "Season: $($seasonInfo.SelectedSeason)" -Level Information
    
    # Export data
    $exportResponse = Export-ScoreSheetData -Response $scoreSheetResponse -SeasonInfo $seasonInfo
    
    # Create output directory
    $outputDirectory = New-OutputDirectory -BasePath $OutputPath -SchoolInfo $schoolInfo -Season $seasonInfo.SelectedSeason
    
    # Check if we got CSV content
    $csvContent = Test-CsvResponse -Response $exportResponse
    
    if ($csvContent) {
        # Save CSV data directly
        $outputFilePath = Save-CsvData -CsvContent $csvContent -OutputDirectory $outputDirectory -School $schoolInfo.Name -Season $seasonInfo.SelectedSeason
    }
    else {
        # Try to parse HTML table data
        $tableData = Get-TableDataFromHtml -Response $exportResponse
        $outputFilePath = Save-TableDataAsCsv -TableData $tableData -OutputDirectory $outputDirectory -Season $seasonInfo.SelectedSeason
    }
    
    return $outputFilePath
}
catch {
    Write-StatusMessage "An error occurred: $($_.Exception.Message)" -Level Error
    Write-Verbose "Error details: $($_.Exception)"
    throw
}
#endregion