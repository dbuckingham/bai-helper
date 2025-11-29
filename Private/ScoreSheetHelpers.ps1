function Find-ScoreSheetFile {
    <#
    .SYNOPSIS
        Locates the Season Score Sheet CSV file for a given school and season.
    
    .PARAMETER SchoolName
        The name of the school (folder name).
    
    .PARAMETER Season
        The season name (folder name).
    
    .PARAMETER BasePath
        The base path where the Season Score Sheets folder is located.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SchoolName,
        
        [Parameter(Mandatory = $true)]
        [string]$Season,
        
        [Parameter(Mandatory = $true)]
        [string]$BasePath
    )

    # Construct the expected path
    $seasonPath = Join-Path -Path $BasePath -ChildPath "Season Score Sheets"
    $schoolPath = Join-Path -Path $seasonPath -ChildPath $SchoolName
    $seasonFolder = Join-Path -Path $schoolPath -ChildPath $Season

    if (-not (Test-Path -Path $seasonFolder)) {
        Write-StatusMessage "Season folder not found: $seasonFolder" -Level Warning
        return $null
    }

    # Look for CSV files in the season folder
    $csvFiles = Get-ChildItem -Path $seasonFolder -Filter "*.csv" | Where-Object { $_.Name -like "SeasonScoreSheet*" -and $_.Name -notlike "Enhanced_*" }

    if (-not $csvFiles) {
        Write-StatusMessage "No Season Score Sheet CSV files found in: $seasonFolder" -Level Warning
        return $null
    }

    if ($csvFiles.Count -gt 1) {
        # If multiple files, get the most recent one
        $csvFile = $csvFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        Write-StatusMessage "Multiple CSV files found, using most recent: $($csvFile.Name)" -Level Information
    }
    else {
        $csvFile = $csvFiles[0]
    }

    return $csvFile.FullName
}

function Add-ArrowCountColumns {
    <#
    .SYNOPSIS
        Adds analysis columns to CSV data:
        - Arrow count columns (AS_10 through AS_0) that count arrows by score value
        - End score columns (E_1 through E_6) that sum scores for each 5-arrow end
        - Half score columns (H1, H2) that sum the first half (ends 1-3) and second half (ends 4-6)
        - End score analysis columns (ES_*) that count ends within specific score ranges
    
    .PARAMETER CsvData
        The CSV data object from Import-Csv.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$CsvData
    )

    # Get the column names to identify A_ columns (arrow score columns)
    $firstRow = $CsvData[0]
    $columnNames = $firstRow.PSObject.Properties.Name
    $arrowColumns = $columnNames | Where-Object { $_ -match '^A_\d+$' } | Sort-Object {
        # Sort numerically by the number after A_
        [int]($_ -replace 'A_', '')
    }

    if (-not $arrowColumns) {
        Write-StatusMessage "No arrow score columns (A_*) found in the data" -Level Warning
        return $CsvData
    }

    Write-StatusMessage "Found arrow columns: $($arrowColumns -join ', ')" -Level Information

    # Process each row
    foreach ($row in $CsvData) {
        # Initialize arrow count columns (AS_10 through AS_0)
        for ($score = 10; $score -ge 0; $score--) {
            $columnName = "AS_$score"
            $count = 0

            # Count arrows with this score value
            foreach ($arrowColumn in $arrowColumns) {
                $arrowValue = $row.$arrowColumn
                
                # Handle different possible formats (numeric, string, empty)
                if ([string]::IsNullOrWhiteSpace($arrowValue)) {
                    continue
                }
                
                # Try to convert to integer
                $numericValue = $null
                if ([int]::TryParse($arrowValue.ToString().Trim(), [ref]$numericValue)) {
                    if ($numericValue -eq $score) {
                        $count++
                    }
                }
            }

            # Add the count column to the row
            $row | Add-Member -NotePropertyName $columnName -NotePropertyValue $count -Force
        }

        # Calculate End scores (E_1 through E_6) and store them for ES analysis
        $endScores = @()
        for ($endNumber = 1; $endNumber -le 6; $endNumber++) {
            $columnName = "E_$endNumber"
            $endScore = 0
            
            # Calculate the arrow range for this end
            $startArrow = ($endNumber - 1) * 5 + 1
            $endArrow = $endNumber * 5
            
            # Sum the arrows in this end
            for ($arrowNum = $startArrow; $arrowNum -le $endArrow; $arrowNum++) {
                $arrowColumn = "A_$arrowNum"
                
                # Check if this arrow column exists in the data
                if ($arrowColumns -contains $arrowColumn) {
                    $arrowValue = $row.$arrowColumn
                    
                    # Handle different possible formats (numeric, string, empty)
                    if (-not [string]::IsNullOrWhiteSpace($arrowValue)) {
                        # Try to convert to integer
                        $numericValue = $null
                        if ([int]::TryParse($arrowValue.ToString().Trim(), [ref]$numericValue)) {
                            $endScore += $numericValue
                        }
                    }
                }
            }
            
            # Add the end score column to the row
            $row | Add-Member -NotePropertyName $columnName -NotePropertyValue $endScore -Force
            
            # Store end score for ES analysis
            $endScores += $endScore
        }

        # Add Half scores (H1 and H2)
        # H1 = sum of ends 1-3 (first half)
        # H2 = sum of ends 4-6 (second half)
        $h1Score = 0
        $h2Score = 0
        
        for ($i = 0; $i -lt $endScores.Count; $i++) {
            if ($i -lt 3) {
                # First half (ends 1-3, which are indices 0-2)
                $h1Score += $endScores[$i]
            } else {
                # Second half (ends 4-6, which are indices 3-5)
                $h2Score += $endScores[$i]
            }
        }
        
        # Add the half score columns to the row
        $row | Add-Member -NotePropertyName "H1" -NotePropertyValue $h1Score -Force
        $row | Add-Member -NotePropertyName "H2" -NotePropertyValue $h2Score -Force

        # Add End Score analysis columns (ES_*)
        # Define the score ranges for end score analysis
        $scoreRanges = @(
            @{ Name = "ES_50"; Min = 50; Max = 50 }
            @{ Name = "ES_49"; Min = 49; Max = 49 }
            @{ Name = "ES_48"; Min = 48; Max = 48 }
            @{ Name = "ES_47"; Min = 47; Max = 47 }
            @{ Name = "ES_46"; Min = 46; Max = 46 }
            @{ Name = "ES_45"; Min = 45; Max = 45 }
            @{ Name = "ES_40_44"; Min = 40; Max = 44 }
            @{ Name = "ES_35_39"; Min = 35; Max = 39 }
            @{ Name = "ES_30_34"; Min = 30; Max = 34 }
            @{ Name = "ES_25_29"; Min = 25; Max = 29 }
            @{ Name = "ES_20_24"; Min = 20; Max = 24 }
            @{ Name = "ES_15_19"; Min = 15; Max = 19 }
            @{ Name = "ES_10_14"; Min = 10; Max = 14 }
            @{ Name = "ES_5_9"; Min = 5; Max = 9 }
            @{ Name = "ES_0_4"; Min = 0; Max = 4 }
        )

        # Count ends in each score range
        foreach ($range in $scoreRanges) {
            $count = 0
            foreach ($endScore in $endScores) {
                if ($endScore -ge $range.Min -and $endScore -le $range.Max) {
                    $count++
                }
            }
            
            # Add the end score analysis column to the row
            $row | Add-Member -NotePropertyName $range.Name -NotePropertyValue $count -Force
        }
    }

    return $CsvData
}

function Get-SafeSchoolName {
    <#
    .SYNOPSIS
        Converts a school display name to a safe folder name format.
    
    .PARAMETER SchoolName
        The school name to convert.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SchoolName
    )

    return Get-SafeFileName -FileName $SchoolName
}