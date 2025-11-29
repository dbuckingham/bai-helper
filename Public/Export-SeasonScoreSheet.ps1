function Export-SeasonScoreSheet {
    <#
    .SYNOPSIS
        Exports a season score sheet from NASP Tournaments website to a CSV file.

    .DESCRIPTION
        This function logs into the NASP Tournaments website using provided or prompted credentials,
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
        Export-SeasonScoreSheet
        Prompts for credentials and exports the default season score sheet.

    .EXAMPLE
        Export-SeasonScoreSheet -Credential (Get-Credential)
        Uses provided credentials to export the default season score sheet.

    .EXAMPLE
        Export-SeasonScoreSheet -Season "2023-2024"
        Exports the score sheet for the specified season.

    .EXAMPLE
        $cred = Get-Credential
        Export-SeasonScoreSheet -Credential $cred -Season "2023-2024" -OrganizationId 5232 -OutputPath "C:\Exports"
        Full example with all parameters specified.

    .OUTPUTS
        System.String
        Returns the full path to the exported CSV file.

    .NOTES
        Author: BAI Helper
        Version: 2.0 (Module version)
        Requires: PowerShell 5.1 or later
    #>
    [CmdletBinding()]
    [OutputType([string])]
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

    # Initialize URLs
    $loginUrl = "$($Script:Config.BaseUrl)$($Script:Config.LoginPath)"
    $scoreSheetUrl = "$($Script:Config.BaseUrl)$($Script:Config.ScoreSheetPath)?oid=$OrganizationId"
    
    # Initialize web session variable
    $webSession = $null

    try {
        # Get credentials
        $credential = Get-UserCredentials -Credential $Credential
        
        # Initialize session and login
        $loginPageResponse = Initialize-WebSession -LoginUrl $loginUrl -WebSession ([ref]$webSession)
        Invoke-Login -Credential $credential -LoginPageResponse $loginPageResponse -LoginUrl $loginUrl -WebSession $webSession | Out-Null
        
        # Get score sheet page
        $scoreSheetResponse = Get-ScoreSheetPage -ScoreSheetUrl $scoreSheetUrl -WebSession $webSession
        
        # Extract information
        $schoolInfo = Get-SchoolInformation -Response $scoreSheetResponse -OrganizationId $OrganizationId
        $seasonInfo = Get-SeasonInformation -Response $scoreSheetResponse -RequestedSeason $Season
        
        # Update season selection if needed
        if ($seasonInfo.NeedsPostback) {
            $scoreSheetResponse = Set-SeasonSelection -SeasonInfo $seasonInfo -Response $scoreSheetResponse -ScoreSheetUrl $scoreSheetUrl -WebSession $webSession
        }
        
        Write-StatusMessage "Season: $($seasonInfo.SelectedSeason)" -Level Information
        
        # Export data
        $exportResponse = Export-ScoreSheetData -Response $scoreSheetResponse -SeasonInfo $seasonInfo -ScoreSheetUrl $scoreSheetUrl -WebSession $webSession
        
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
}