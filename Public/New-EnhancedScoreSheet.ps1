function New-EnhancedScoreSheet {
    <#
    .SYNOPSIS
        Creates an enhanced copy of a Season Score Sheet with arrow count analysis columns.

    .DESCRIPTION
        This function takes a school name and season, locates the corresponding Season Score Sheet CSV file,
        and creates an enhanced copy with additional columns for analysis:
        - Arrow count columns (AS_10 through AS_0) that count arrows by score value
        - End score columns (E_1 through E_6) that sum the scores for each 5-arrow end
        - Half score columns (H1, H2) that sum the first half (ends 1-3) and second half (ends 4-6)
        - End score analysis columns (ES_*) that count ends within specific score ranges
        The enhanced file is saved with "Enhanced_" prepended to the filename in the same directory.

    .PARAMETER SchoolName
        The name of the school (must match the folder name in the Season Score Sheets directory).

    .PARAMETER Season
        The season name (must match the folder name under the school directory).

    .PARAMETER BasePath
        Optional. The base path where the "Season Score Sheets" folder is located.
        Defaults to the current directory.

    .PARAMETER IgnoreEmpty
        Optional. When specified, the function will skip processing and return null 
        instead of throwing an error if the score sheet file is empty or contains no data.

    .EXAMPLE
        New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024"
        Creates an enhanced copy of the score sheet for Sample High School's 2023-2024 season.

    .EXAMPLE
        New-EnhancedScoreSheet -SchoolName "My_School" -Season "2024-2025" -BasePath "C:\Exports"
        Creates an enhanced copy using a specific base path.

    .EXAMPLE
        New-EnhancedScoreSheet -SchoolName "Sample_High_School" -Season "2023-2024" -IgnoreEmpty
        Creates an enhanced copy, skipping processing if the score sheet file is empty.

    .OUTPUTS
        System.String or $null
        Returns the full path to the enhanced CSV file, or $null if the file was empty and -IgnoreEmpty was specified.

    .NOTES
        Author: BAI Helper
        Version: 2.0 (Module version)
        Requires: PowerShell 5.1 or later
        
        The function adds columns:
        - AS_10 through AS_0: Count the number of arrows with each score value
        - E_1 through E_6: Sum the scores for each 5-arrow end (arrows 1-5, 6-10, 11-15, 16-20, 21-25, 26-30)
        - H1: Sum of first half score (ends 1-3)
        - H2: Sum of second half score (ends 4-6)
        - ES_50, ES_49, etc.: Count of ends with specific scores (50, 49, 48, 47, 46, 45)
        - ES_40_44, ES_35_39, etc.: Count of ends within score ranges (40-44, 35-39, 30-34, 25-29, 20-24, 15-19, 10-14, 5-9, 0-4)
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SchoolName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Season,

        [Parameter(Mandatory = $false)]
        [ValidateScript({Test-Path $_ -IsValid})]
        [string]
        $BasePath = (Get-Location).Path,

        [Parameter(Mandatory = $false)]
        [switch]
        $IgnoreEmpty
    )

    try {
        # Find the score sheet file
        $scoreSheetPath = Find-ScoreSheetFile -SchoolName $SchoolName -Season $Season -BasePath $BasePath
        
        if (-not $scoreSheetPath) {
            throw "Could not find score sheet file for school '$SchoolName' and season '$Season'"
        }

        Write-StatusMessage "Found score sheet: $scoreSheetPath" -Level Information

        # Read the CSV data
        $csvData = Import-Csv -Path $scoreSheetPath -Encoding UTF8

        if (-not $csvData -or $csvData.Count -eq 0) {
            if ($IgnoreEmpty) {
                Write-StatusMessage "Score sheet file is empty, skipping enhancement as requested: $scoreSheetPath" -Level Warning
                return $null
            }
            else {
                throw "Score sheet file is empty or could not be read: $scoreSheetPath"
            }
        }

        Write-StatusMessage "Processing $($csvData.Count) rows..." -Level Information

        # Enhance the data with arrow count columns
        $enhancedData = Add-ArrowCountColumns -CsvData $csvData

        # Create the enhanced filename
        $originalFileName = [System.IO.Path]::GetFileNameWithoutExtension($scoreSheetPath)
        $directory = [System.IO.Path]::GetDirectoryName($scoreSheetPath)
        $enhancedFileName = "Enhanced_$originalFileName.csv"
        $enhancedFilePath = Join-Path -Path $directory -ChildPath $enhancedFileName

        # Save the enhanced data
        $enhancedData | Export-Csv -Path $enhancedFilePath -NoTypeInformation -Encoding UTF8

        Write-StatusMessage "Enhanced score sheet created successfully!" -Level Success
        Write-StatusMessage "File saved to: $enhancedFilePath" -Level Information

        return $enhancedFilePath
    }
    catch {
        Write-StatusMessage "An error occurred while creating enhanced score sheet: $($_.Exception.Message)" -Level Error
        Write-Verbose "Error details: $($_.Exception)"
        throw
    }
}