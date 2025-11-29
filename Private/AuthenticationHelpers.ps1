function Get-UserCredentials {
    <#
    .SYNOPSIS
        Gets or prompts for user credentials.
    
    .PARAMETER Credential
        Optional credentials object. If not provided, user will be prompted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    if ($Credential) {
        return $Credential
    }
    
    Write-StatusMessage "Please enter your NASP Tournaments credentials:" -Level Information
    $promptedCredential = Get-Credential -Message "Enter your NASP Tournaments username and password"
    
    if (-not $promptedCredential) {
        throw "Credentials are required to proceed."
    }
    
    return $promptedCredential
}

function Initialize-WebSession {
    <#
    .SYNOPSIS
        Initializes the web session by getting the login page.
    
    .PARAMETER LoginUrl
        The URL to initialize the session with.
    
    .PARAMETER WebSession
        The web session variable to use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LoginUrl,
        
        [Parameter(Mandatory = $true)]
        [ref]$WebSession
    )
    
    Write-StatusMessage "Connecting to NASP Tournaments..." -Level Information
    
    try {
        $loginPageResponse = Invoke-WebRequest -Uri $LoginUrl -SessionVariable session -UseBasicParsing -ErrorAction Stop
        $WebSession.Value = $session
        return $loginPageResponse
    }
    catch {
        throw "Failed to connect to NASP Tournaments: $($_.Exception.Message)"
    }
}

function Get-FormFields {
    <#
    .SYNOPSIS
        Extracts form fields from a web response.
    
    .PARAMETER Response
        The web response to extract form fields from.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$Response
    )
    
    $formFields = @{}
    
    foreach ($field in $Response.InputFields) {
        if ($field.name -and $field.name -ne "") {
            $formFields[$field.name] = $field.value
        }
    }
    
    return $formFields
}

function Invoke-Login {
    <#
    .SYNOPSIS
        Performs login to the NASP Tournaments website.
    
    .PARAMETER Credential
        The credentials to use for login.
    
    .PARAMETER LoginPageResponse
        The response from the login page request.
    
    .PARAMETER LoginUrl
        The login URL.
    
    .PARAMETER WebSession
        The web session to use.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]$LoginPageResponse,
        
        [Parameter(Mandatory = $true)]
        [string]$LoginUrl,
        
        [Parameter(Mandatory = $true)]
        [Microsoft.PowerShell.Commands.WebRequestSession]$WebSession
    )
    
    # Extract form fields from login page
    $formFields = Get-FormFields -Response $LoginPageResponse
    
    # Set credentials
    $formFields[$Script:Config.Controls.Username] = $Credential.UserName
    $formFields[$Script:Config.Controls.Password] = $Credential.GetNetworkCredential().Password

    Write-StatusMessage "Logging in as $($Credential.UserName)..." -Level Information
    
    try {
        $loginResponse = Invoke-WebRequest -Uri $LoginUrl -Method POST -Body $formFields -WebSession $WebSession -UseBasicParsing -ErrorAction Stop
        
        if ($loginResponse.StatusCode -eq 200) {
            Write-StatusMessage "Login successful!" -Level Success
            return $loginResponse
        }
        else {
            throw "Login failed with status code: $($loginResponse.StatusCode)"
        }
    }
    catch {
        throw "Login failed: $($_.Exception.Message)"
    }
}