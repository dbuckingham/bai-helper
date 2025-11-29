function New-OutputDirectory {
    <#
    .SYNOPSIS
        Creates the output directory structure.
    
    .PARAMETER BasePath
        The base path for the output directory.
    
    .PARAMETER SchoolInfo
        Hashtable containing school information.
    
    .PARAMETER Season
        The season name for the folder.
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
    
    .PARAMETER CsvContent
        The CSV content to save.
    
    .PARAMETER OutputDirectory
        The directory to save the file to.
    
    .PARAMETER School
        The school name for the filename.
    
    .PARAMETER Season
        The season name for the filename.
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
    
    .PARAMETER TableData
        The table data to save.
    
    .PARAMETER OutputDirectory
        The directory to save the file to.
    
    .PARAMETER Season
        The season name for the filename.
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