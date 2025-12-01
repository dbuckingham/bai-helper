function New-SeasonSummaryExcel {
    <#
    .SYNOPSIS
        Creates an Excel workbook with a season summary of tournament results.

    .DESCRIPTION
        This function takes a school name and season, locates the corresponding Enhanced Season Score Sheet CSV file,
        and creates an Excel workbook with comprehensive tournament analysis. The workbook includes:
        - Tournament Results worksheet with detailed statistics for each tournament
        - Tournament name, date, range type, team scores (overall, H1, H2)
        - Individual archer statistics (min, max, average scores)
        - Arrow count distribution (10s through 0s) for team-scoring archers
        - End score distribution (50s through 45s) for team-scoring archers
        
        This function requires Enhanced score sheets created by New-EnhancedScoreSheet.
        The Excel workbook is saved with "SeasonSummary_" prepended to the filename in the same directory.

    .PARAMETER SchoolName
        The name of the school (must match the folder name in the Season Score Sheets directory).

    .PARAMETER Season
        The season name (must match the folder name under the school directory).

    .PARAMETER BasePath
        Optional. The base path where the "Season Score Sheets" folder is located.
        Defaults to the current directory.

    .EXAMPLE
        New-SeasonSummaryExcel -SchoolName "Eastern High School" -Season "2023-2024"
        Creates a season summary Excel workbook for Eastern High School's 2023-2024 season.

    .EXAMPLE
        New-SeasonSummaryExcel -SchoolName "My_School" -Season "2024-2025" -BasePath "C:\Exports"
        Creates a season summary Excel workbook using a specific base path.

    .OUTPUTS
        System.String
        Returns the full path to the Excel workbook file.

    .NOTES
        Author: BAI Helper
        Version: 2.0 (Module version)
        Requires: PowerShell 5.1 or later, ImportExcel module, Enhanced score sheets
        
        The function requires the ImportExcel module for Excel functionality.
        Install with: Install-Module ImportExcel -Force
        
        The function requires Enhanced score sheets created by New-EnhancedScoreSheet.
        Run New-EnhancedScoreSheet first to generate the required enhanced statistics.
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
        $BasePath = (Get-Location).Path
    )

    try {
        Write-StatusMessage "Starting season summary Excel creation for '$SchoolName' - '$Season'..." -Level Information
        
        # Check for ImportExcel module
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            throw "ImportExcel module is required but not installed. Please run: Install-Module ImportExcel -Force"
        }
        Import-Module ImportExcel -Force

        # Find the Enhanced score sheet file
        $seasonPath = Join-Path -Path $BasePath -ChildPath "Season Score Sheets"
        $schoolPath = Join-Path -Path $seasonPath -ChildPath $SchoolName
        $seasonFolder = Join-Path -Path $schoolPath -ChildPath $Season

        if (-not (Test-Path -Path $seasonFolder)) {
            throw "Season folder not found: $seasonFolder"
        }

        # Look for Enhanced CSV files in the season folder
        $enhancedFiles = Get-ChildItem -Path $seasonFolder -Filter "Enhanced_*.csv"

        if (-not $enhancedFiles) {
            throw "No Enhanced Season Score Sheet CSV files found in: $seasonFolder. Please run New-EnhancedScoreSheet first."
        }

        if ($enhancedFiles.Count -gt 1) {
            # If multiple files, get the most recent one
            $scoreSheetFile = $enhancedFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            Write-StatusMessage "Multiple Enhanced CSV files found, using most recent: $($scoreSheetFile.Name)" -Level Information
        }
        else {
            $scoreSheetFile = $enhancedFiles[0]
        }

        Write-StatusMessage "Found score sheet: $($scoreSheetFile.FullName)" -Level Information

        # Read the CSV data
        $csvData = Import-Csv -Path $scoreSheetFile.FullName
        if (-not $csvData -or $csvData.Count -eq 0) {
            throw "Score sheet file is empty or contains no data"
        }

        Write-StatusMessage "Processing $($csvData.Count) records..." -Level Information

        # Filter for team-scoring archers only
        $teamData = $csvData | Where-Object { $_.USE_FOR_TEAM -eq 'Y' }
        
        if (-not $teamData -or $teamData.Count -eq 0) {
            throw "No team-scoring archers found in the data"
        }

        Write-StatusMessage "Found $($teamData.Count) team-scoring archer records..." -Level Information

        # Group by tournament to get tournament-level statistics
        $tournamentGroups = $teamData | Group-Object TOURNAMENT_ID, TOURNAMENT_NAME, RANGE_TYPE, END_DATE

        $tournamentResults = @()

        foreach ($group in $tournamentGroups) {
            Write-StatusMessage "Processing tournament: $($group.Name)" -Level Information
            
            $tournament = $group.Group[0]
            $archers = $group.Group
            
            # Calculate team score (sum of all team archers' scores)
            $teamScore = ($archers | Measure-Object -Property SCORE -Sum).Sum
            
            # Calculate H1 and H2 scores using pre-calculated values from Enhanced sheet
            $h1Score = ($archers | Measure-Object -Property H1 -Sum).Sum
            $h2Score = ($archers | Measure-Object -Property H2 -Sum).Sum
            
            # Calculate min, max, average scores
            $scoreStats = $archers | Measure-Object -Property SCORE -Minimum -Maximum -Average
            
            # Sum arrow counts from pre-calculated Enhanced values
            $arrowCounts = @{
                AS_10 = ($archers | Measure-Object -Property AS_10 -Sum).Sum
                AS_9 = ($archers | Measure-Object -Property AS_9 -Sum).Sum
                AS_8 = ($archers | Measure-Object -Property AS_8 -Sum).Sum
                AS_7 = ($archers | Measure-Object -Property AS_7 -Sum).Sum
                AS_6 = ($archers | Measure-Object -Property AS_6 -Sum).Sum
                AS_5 = ($archers | Measure-Object -Property AS_5 -Sum).Sum
                AS_4 = ($archers | Measure-Object -Property AS_4 -Sum).Sum
                AS_3 = ($archers | Measure-Object -Property AS_3 -Sum).Sum
                AS_2 = ($archers | Measure-Object -Property AS_2 -Sum).Sum
                AS_1 = ($archers | Measure-Object -Property AS_1 -Sum).Sum
                AS_0 = ($archers | Measure-Object -Property AS_0 -Sum).Sum
            }
            
            # Sum end score counts from pre-calculated Enhanced values
            $endCounts = @{
                ES_50 = ($archers | Measure-Object -Property ES_50 -Sum).Sum
                ES_49 = ($archers | Measure-Object -Property ES_49 -Sum).Sum
                ES_48 = ($archers | Measure-Object -Property ES_48 -Sum).Sum
                ES_47 = ($archers | Measure-Object -Property ES_47 -Sum).Sum
                ES_46 = ($archers | Measure-Object -Property ES_46 -Sum).Sum
                ES_45 = ($archers | Measure-Object -Property ES_45 -Sum).Sum
            }
            
            # Create tournament result object
            $result = [PSCustomObject]@{
                TournamentName = $tournament.TOURNAMENT_NAME
                Date = [DateTime]::Parse($tournament.END_DATE).ToString("MM/dd/yyyy")
                RangeType = $tournament.RANGE_TYPE
                TeamScore = $teamScore
                H1TeamScore = $h1Score
                H2TeamScore = $h2Score
                MinScore = $scoreStats.Minimum
                MaxScore = $scoreStats.Maximum
                AvgScore = [math]::Round($scoreStats.Average, 2)
                Total_10s = $arrowCounts.AS_10
                Total_9s = $arrowCounts.AS_9
                Total_8s = $arrowCounts.AS_8
                Total_7s = $arrowCounts.AS_7
                Total_6s = $arrowCounts.AS_6
                Total_5s = $arrowCounts.AS_5
                Total_4s = $arrowCounts.AS_4
                Total_3s = $arrowCounts.AS_3
                Total_2s = $arrowCounts.AS_2
                Total_1s = $arrowCounts.AS_1
                Total_0s = $arrowCounts.AS_0
                Ends_50 = $endCounts.ES_50
                Ends_49 = $endCounts.ES_49
                Ends_48 = $endCounts.ES_48
                Ends_47 = $endCounts.ES_47
                Ends_46 = $endCounts.ES_46
                Ends_45 = $endCounts.ES_45
            }
            
            $tournamentResults += $result
        }

        # Sort by date
        $tournamentResults = $tournamentResults | Sort-Object { [DateTime]::ParseExact($_.Date, "MM/dd/yyyy", $null) }

        # Create output file path
        $outputFileName = "SeasonSummary_$(Get-SafeFileName -FileName $SchoolName)_$(Get-SafeFileName -FileName $Season).xlsx"
        $outputPath = Join-Path -Path $scoreSheetFile.Directory.FullName -ChildPath $outputFileName

        Write-StatusMessage "Creating Excel workbook: $outputPath" -Level Information

        # Export to Excel
        $tournamentResults | Export-Excel -Path $outputPath -WorksheetName "Tournament Results" -AutoSize -TableStyle Medium2 -FreezeTopRow

        Write-StatusMessage "Season summary Excel workbook created successfully!" -Level Success
        Write-StatusMessage "File saved to: $outputPath" -Level Success

        return $outputPath
    }
    catch {
        Write-StatusMessage "An error occurred while creating season summary Excel: $($_.Exception.Message)" -Level Error
        throw
    }
}