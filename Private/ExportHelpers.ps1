function Export-ScoreSheetData {
    <#
    .SYNOPSIS
        Exports the score sheet data by triggering the export button.
    
    .PARAMETER Response
        The web response containing the form.
    
    .PARAMETER SeasonInfo
        Hashtable containing season information.
    
    .PARAMETER ScoreSheetUrl
        The score sheet URL.
    
    .PARAMETER WebSession
        The web session to use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$SeasonInfo,
        
        [Parameter(Mandatory = $true)]
        [string]$ScoreSheetUrl,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession
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
        
        $exportResponse = Invoke-WebRequest -Uri $ScoreSheetUrl -Method POST -Body $formFields -WebSession $WebSession -UseBasicParsing -ErrorAction Stop
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
    
    .PARAMETER Response
        The web response to test.
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
    
    .PARAMETER Response
        The web response containing HTML table data.
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