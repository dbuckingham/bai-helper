function Get-AvailableSeasons {
    <#
    .SYNOPSIS
        Retrieves the list of available seasons for a given school organization.

    .DESCRIPTION
        This function logs into the NASP Tournaments website and retrieves the list of 
        available seasons from the Season Score Sheet page dropdown for the specified organization.
        This is useful for discovering which seasons are available before attempting to export data.

    .PARAMETER Credential
        A PSCredential object containing the username and password for login.
        If not provided, the user will be prompted to enter credentials.

    .PARAMETER OrganizationId
        The organization ID to retrieve seasons for. Defaults to 5232.

    .EXAMPLE
        Get-AvailableSeasons
        Prompts for credentials and retrieves seasons for the default organization (5232).

    .EXAMPLE
        Get-AvailableSeasons -Credential (Get-Credential) -OrganizationId 1234
        Uses provided credentials to retrieve seasons for organization 1234.

    .EXAMPLE
        $cred = Get-Credential
        $seasons = Get-AvailableSeasons -Credential $cred
        $seasons | ForEach-Object { Write-Host "Available season: $_" }
        Retrieves seasons and displays each one.

    .OUTPUTS
        System.String[]
        Returns an array of available season names.

    .NOTES
        Author: BAI Helper
        Version: 2.0 (Module version)
        Requires: PowerShell 5.1 or later
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 99999)]
        [int]
        $OrganizationId = 5232
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
        
        # Extract available seasons from the dropdown
        $availableSeasons = Get-SeasonsFromDropdown -Response $scoreSheetResponse
        
        if ($availableSeasons.Count -eq 0) {
            Write-StatusMessage "No seasons found for organization ID: $OrganizationId" -Level Warning
        }
        else {
            Write-StatusMessage "Found $($availableSeasons.Count) available seasons for organization ID: $OrganizationId" -Level Success
        }
        
        return $availableSeasons
    }
    catch {
        Write-StatusMessage "An error occurred while retrieving seasons: $($_.Exception.Message)" -Level Error
        Write-Verbose "Error details: $($_.Exception)"
        throw
    }
}