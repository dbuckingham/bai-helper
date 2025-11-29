# BaiHelper PowerShell Module
# Provides functionality for exporting season score sheets from NASP Tournaments website

#Requires -Version 5.1

# Get public and private function definition files
$publicFunctions = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$privateFunctions = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the files
foreach ($import in @($publicFunctions + $privateFunctions)) {
    try {
        Write-Verbose "Importing $($import.FullName)"
        . $import.FullName
    }
    catch {
        Write-Error -Message "Failed to import function $($import.FullName): $($_.Exception.Message)"
    }
}

# Export public functions
Export-ModuleMember -Function $publicFunctions.BaseName

# Module variables
Write-Verbose "BaiHelper module loaded successfully"
Write-Verbose "Available functions: $($publicFunctions.BaseName -join ', ')"