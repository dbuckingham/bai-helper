function Write-StatusMessage {
    <#
    .SYNOPSIS
        Writes a standardized status message.
    
    .PARAMETER Message
        The message to display.
    
    .PARAMETER Level
        The message level (Information, Warning, Error, Success).
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
    
    .PARAMETER HtmlText
        The HTML text to clean.
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
    
    .PARAMETER FileName
        The filename to clean.
    
    .PARAMETER Replacement
        The replacement character for invalid characters.
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