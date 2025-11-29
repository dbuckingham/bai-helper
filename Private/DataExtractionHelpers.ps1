function Get-ScoreSheetPage {
    <#
    .SYNOPSIS
        Navigates to the Season Score Sheet page.
    
    .PARAMETER ScoreSheetUrl
        The URL of the score sheet page.
    
    .PARAMETER WebSession
        The web session to use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScoreSheetUrl,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession
    )
    
    Write-StatusMessage "Navigating to Season Score Sheet page..." -Level Information
    
    try {
        $scoreSheetResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -WebSession $WebSession -UseBasicParsing -ErrorAction Stop
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
    
    .PARAMETER Response
        The web response containing the school information.
    
    .PARAMETER OrganizationId
        The organization ID to use as fallback.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response,
        
        [Parameter(Mandatory = $false)]
        [int]$OrganizationId = 5232
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
    
    .PARAMETER Response
        The web response containing the season information.
    
    .PARAMETER RequestedSeason
        The requested season to select.
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
    
    .PARAMETER SeasonInfo
        Hashtable containing season information.
    
    .PARAMETER Response
        The web response to update.
    
    .PARAMETER ScoreSheetUrl
        The score sheet URL for postback.
    
    .PARAMETER WebSession
        The web session to use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$SeasonInfo,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response,
        
        [Parameter(Mandatory = $true)]
        [string]$ScoreSheetUrl,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession
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
        $updatedResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -Method POST -Body $formFields -WebSession $WebSession -UseBasicParsing -ErrorAction Stop
        
        Write-StatusMessage "Season selection successful!" -Level Success

        return $updatedResponse
    }
    catch {
        Write-StatusMessage "Failed to update season selection: $($_.Exception.Message)" -Level Warning
        return $Response
    }
}

function Get-SeasonsFromDropdown {
    <#
    .SYNOPSIS
        Extracts all available seasons from the dropdown on the score sheet page.
    
    .PARAMETER Response
        The web response containing the season dropdown.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    $seasons = @()
    
    # Find season dropdown ID
    $dropdownId = ""
    if ($Response.Content -match $Script:Config.Patterns.SeasonDropdown1) {
        $dropdownId = $Matches[1]
    }
    elseif ($Response.Content -match $Script:Config.Patterns.SeasonDropdown2) {
        $dropdownId = $Matches[1]
    }
    
    if (-not $dropdownId) {
        Write-StatusMessage "Could not find season dropdown on the page" -Level Warning
        return $seasons
    }
    
    # Find the dropdown section in the HTML
    $dropdownPattern = '<select[^>]*id="' + [regex]::Escape($dropdownId) + '"[^>]*>([\s\S]*?)</select>'
    if ($Response.Content -match $dropdownPattern) {
        $dropdownContent = $Matches[1]
        
        # Extract all option values and text
        $optionPattern = '<option[^>]*value="([^"]*)"[^>]*>([^<]+)</option>'
        [regex]::Matches($dropdownContent, $optionPattern) | ForEach-Object {
            $seasonText = $_.Groups[2].Value.Trim()
            if (-not [string]::IsNullOrWhiteSpace($seasonText)) {
                $seasons += $seasonText
            }
        }
    }
    
    # Remove duplicates and sort
    $seasons = $seasons | Sort-Object -Unique
    
    return $seasons
}