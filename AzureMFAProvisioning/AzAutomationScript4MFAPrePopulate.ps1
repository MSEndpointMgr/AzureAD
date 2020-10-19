<#
    .SYNOPSIS
        This script is for getting all non-MFA registered users - and setting their authentication phone for SMS
    .DESCRIPTION
        You need to configure the required strings as Azure Automation variables.
        You also need to create an account that has authentication admin rights and no MFA requirement.
        Option 1: Use and configure the Azure Run AS Account as Service Principal, and give it the following permissions: user.read.all reports.read.all and delegate: UserAuthenticationMethod.ReadWrite.All
        Option 2: Create an app registration that has application: user.read.all reports.read.all and delegate: UserAuthenticationMethod.ReadWrite.All
        Either app needs to be configured with "Treat application as a public client = YES"

    .NOTES
        Authors:     Jan Ketil Skanke, Sandy Zeng, Michael Mardahl
        Contact:     @JankeSkanke + @Sandy_Tsang + @Michael_Mardahl
        Created:     2020-10-07
        Updated:     2020-10-07
        Version history:
        1.0.0 - (2020-05-20) Initial Version
    #>    

#requires -module MSAL.PS

#region declarations
$connection = Get-AutomationConnection -Name AzureRunAsConnection
$script:Tenant = $connection.TenantID 
$script:AppId = $connection.ApplicationID
$script:graphVersion = 'beta'
$script:appCert = Get-AutomationCertificate -Name "AzureRunAsCertificate"
#$script:appSecret = ConvertTo-SecureString (Get-AutomationVariable -Name "AppSecret") -AsPlainText -Force #Change this to your own Azure Automation App Secret
$script:authenticationCredentials = Get-AutomationPSCredential -Name 'DelegateServiceAccount' #Change this to your own Azure Automation credential
$UseStagingGroup = Get-AutomationVariable -Name "EnableStagingGroup" #Set this Automation variable to $true or $false to enable use of the Staging/Pilot group.
$StagingGroupName = Get-AutomationVariable -Name "StagingGroupName" #Set this Automation variable to the name of your staging group or comment out if you absolutely never are going to use staging.
#endregion declarations

#region functions
function Invoke-MSGraphOperation {
    <#
    .SYNOPSIS
        Perform a specific call to Intune Graph API, either as GET, POST, PATCH or DELETE methods.

    .DESCRIPTION
        Perform a specific call to Intune Graph API, either as GET, POST, PATCH or DELETE methods.
        This function handles nextLink objects including throttling based on retry-after value from Graph response.

    .PARAMETER Get
        Switch parameter used to specify the method operation as 'GET'.

    .PARAMETER Post
        Switch parameter used to specify the method operation as 'POST'.

    .PARAMETER Patch
        Switch parameter used to specify the method operation as 'PATCH'.

    .PARAMETER Put
        Switch parameter used to specify the method operation as 'PUT'.

    .PARAMETER Delete
        Switch parameter used to specify the method operation as 'DELETE'.

    .PARAMETER Resource
        Specify the full resource path, e.g. deviceManagement/auditEvents.

    .PARAMETER Headers
        Specify a hash-table as the header containing minimum the authentication token.

    .PARAMETER Body
        Specify the body construct.

    .PARAMETER APIVersion
        Specify to use either 'Beta' or 'v1.0' API version.

    .PARAMETER ContentType
        Specify the content type for the graph request.

    .NOTES
        Author:      Nickolaj Andersen & Jan Ketil Skanke
        Contact:     @JankeSkanke @NickolajA
        Created:     2020-10-11
        Updated:     2020-10-11

        Version history:
        1.0.0 - (2020-10-11) Function created
    #>    
    param(
        [parameter(Mandatory = $true, ParameterSetName = "GET", HelpMessage = "Switch parameter used to specify the method operation as 'GET'.")]
        [switch]$Get,

        [parameter(Mandatory = $true, ParameterSetName = "POST", HelpMessage = "Switch parameter used to specify the method operation as 'POST'.")]
        [switch]$Post,

        [parameter(Mandatory = $true, ParameterSetName = "PATCH", HelpMessage = "Switch parameter used to specify the method operation as 'PATCH'.")]
        [switch]$Patch,

        [parameter(Mandatory = $true, ParameterSetName = "PUT", HelpMessage = "Switch parameter used to specify the method operation as 'PUT'.")]
        [switch]$Put,

        [parameter(Mandatory = $true, ParameterSetName = "DELETE", HelpMessage = "Switch parameter used to specify the method operation as 'DELETE'.")]
        [switch]$Delete,

        [parameter(Mandatory = $true, ParameterSetName = "GET", HelpMessage = "Specify the full resource path, e.g. deviceManagement/auditEvents.")]
        [parameter(Mandatory = $true, ParameterSetName = "POST")]
        [parameter(Mandatory = $true, ParameterSetName = "PATCH")]
        [parameter(Mandatory = $true, ParameterSetName = "PUT")]
        [parameter(Mandatory = $true, ParameterSetName = "DELETE")]
        [ValidateNotNullOrEmpty()]
        [string]$Resource,

        [parameter(Mandatory = $true, ParameterSetName = "GET", HelpMessage = "Specify a hash-table as the header containing minimum the authentication token.")]
        [parameter(Mandatory = $true, ParameterSetName = "POST")]
        [parameter(Mandatory = $true, ParameterSetName = "PATCH")]
        [parameter(Mandatory = $true, ParameterSetName = "PUT")]
        [parameter(Mandatory = $true, ParameterSetName = "DELETE")]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable]$Headers,

        [parameter(Mandatory = $true, ParameterSetName = "POST", HelpMessage = "Specify the body construct.")]
        [parameter(Mandatory = $true, ParameterSetName = "PATCH")]
        [parameter(Mandatory = $true, ParameterSetName = "PUT")]
        [ValidateNotNullOrEmpty()]
        [System.Object]$Body,

        [parameter(Mandatory = $false, ParameterSetName = "GET", HelpMessage = "Specify to use either 'Beta' or 'v1.0' API version.")]
        [parameter(Mandatory = $false, ParameterSetName = "POST")]
        [parameter(Mandatory = $false, ParameterSetName = "PATCH")]
        [parameter(Mandatory = $false, ParameterSetName = "PUT")]
        [parameter(Mandatory = $false, ParameterSetName = "DELETE")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Beta", "v1.0")]
        [string]$APIVersion = "v1.0",

        [parameter(Mandatory = $false, ParameterSetName = "GET", HelpMessage = "Specify the content type for the graph request.")]
        [parameter(Mandatory = $false, ParameterSetName = "POST")]
        [parameter(Mandatory = $false, ParameterSetName = "PATCH")]
        [parameter(Mandatory = $false, ParameterSetName = "PUT")]
        [parameter(Mandatory = $false, ParameterSetName = "DELETE")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("application/json", "image/png")]
        [string]$ContentType = "application/json"
    )
    Begin {
        # Construct list as return value for handling both single and multiple instances in response from call
        $GraphResponseList = New-Object -TypeName "System.Collections.ArrayList"

        # Construct full URI
        $GraphURI = "https://graph.microsoft.com/$($APIVersion)/$($Resource)"
        Write-Verbose -Message "$($PSCmdlet.ParameterSetName) $($GraphURI)"
    }
    Process {
        # Call Graph API and get JSON response
        do {
            try {
                # Construct table of default request parameters
                $RequestParams = @{
                    "Uri" = $GraphURI
                    "Headers" = $Headers
                    "Method" = $PSCmdlet.ParameterSetName
                    "ErrorAction" = "Stop"
                    "Verbose" = $false
                }

                switch ($PSCmdlet.ParameterSetName) {
                    "POST" {
                        $RequestParams.Add("Body", $Body)
                        $RequestParams.Add("ContentType", $ContentType)
                    }
                    "PATCH" {
                        $RequestParams.Add("Body", $Body)
                        $RequestParams.Add("ContentType", $ContentType)
                    }
                    "PUT" {
                        $RequestParams.Add("Body", $Body)
                        $RequestParams.Add("ContentType", $ContentType)
                    }
                }

                # Invoke Graph request
                $GraphResponse = Invoke-RestMethod @RequestParams

                # Handle paging in response
                if ($GraphResponse.'@odata.nextLink' -ne $null) {
                    $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                    $GraphURI = $GraphResponse.'@odata.nextLink'
                    Write-Verbose -Message "NextLink: $($GraphURI)"
                }
                else {
                    # NextLink from response was null, assuming last page but also handle if a single instance is returned
                    if (-not([string]::IsNullOrEmpty($GraphResponse.value))) {
                        $GraphResponseList.AddRange($GraphResponse.value) | Out-Null
                    }
                    else {
                        $GraphResponseList.Add($GraphResponse) | Out-Null
                    }
                    
                    # Set graph response as handled and stop processing loop
                    $GraphResponseProcess = $false
                }
            }
            catch [System.Exception] {
                if ($PSItem.Exception.Response.StatusCode -like "429") {
                    # Detected throttling based from response status code
                    $RetryInSeconds = $PSItem.Exception.Response.Headers["Retry-After"]

                    # Wait for given period of time specified in response headers
                    Write-Verbose -Message "Graph is throttling the request, will retry in '$($RetryInSeconds)' seconds"
                    Start-Sleep -Seconds $RetryInSeconds
                }
                else {
                    # Read the response stream
                    $StreamReader = New-Object -TypeName "System.IO.StreamReader" -ArgumentList @($PSItem.Exception.Response.GetResponseStream())
                    $StreamReader.BaseStream.Position = 0
                    $StreamReader.DiscardBufferedData()
                    $ResponseBody = ($StreamReader.ReadToEnd() | ConvertFrom-Json)
                    
                    switch ($PSCmdlet.ParameterSetName) {
                        "GET" {
                            # Output warning message that the request failed with error message description from response stream
                            Write-Warning -Message "Graph request failed with status code '$($PSItem.Exception.Response.StatusCode)'. Error message: $($ResponseBody.error.message)"

                            # Set graph response as handled and stop processing loop
                            $GraphResponseProcess = $false
                        }
                        default {
                            # Construct new custom error record
                            $SystemException = New-Object -TypeName "System.Management.Automation.RuntimeException" -ArgumentList ("{0}: {1}" -f $ResponseBody.error.code, $ResponseBody.error.message)
                            $ErrorRecord = New-Object -TypeName "System.Management.Automation.ErrorRecord" -ArgumentList @($SystemException, $ErrorID, [System.Management.Automation.ErrorCategory]::NotImplemented, [string]::Empty)

                            # Throw a terminating custom error record
                            $PSCmdlet.ThrowTerminatingError($ErrorRecord)
                        }
                    }

                    # Set graph response as handled and stop processing loop
                    $GraphResponseProcess = $false
                }
            }
        }
        until ($GraphResponseProcess -eq $false)

        # Handle return value
        #return $GraphResponseResult
        return $GraphResponseList
    }
}
#Get Access Token with either App or Delegate permission
function get-authToken {
    <#
        .SYNOPSIS
            Requests either a Delegated or Application based token grant form Microsoft Graph API
    #>
    [CmdletBinding(DefaultParameterSetName='Delegated')]
    Param(
        [Parameter(Mandatory=$true, ParameterSetName="Delegated")]
        [Switch]
        $DelegatedGrant,

        [Parameter(Mandatory=$true, ParameterSetName="Delegated", ValueFromPipeline=$true)]
        [System.Management.Automation.PSCredential]
        $UserCredential,

        [Parameter(Mandatory=$true, ParameterSetName="ApplicationCert")]
        [Parameter(Mandatory=$true, ParameterSetName="ApplicationSecret")]
        [Switch]
        $ApplicationGrant,

        [Parameter(Mandatory=$false, ParameterSetName="ApplicationCert", ValueFromPipeline=$true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]
        $ClientCertificate,

        [Parameter(Mandatory=$true, ParameterSetName="ApplicationSecret")]
        [String]
        $ClientSecret,

        [Parameter(Mandatory=$true)]
        [String]
        $TenantId,

        [Parameter(Mandatory=$true)]
        [String]
        $ClientId
    )

    if($ApplicationGrant) {
        #Get a Graph API Toke via Application Grant
        try {
            if($ClientCertificate) {
                #Get token with certificate based auth
                $response = Get-MsalToken -ClientId $ClientID -ClientCertificate $ClientCertificate -TenantId $TenantId -ErrorAction Stop
            } else {
                #Get token with secret based auth
                $response = Get-MsalToken -ClientId $ClientID -ClientSecret $ClientSecret -TenantId $TenantId -ErrorAction Stop
            }
            $authToken = @{
                Authorization = $response.CreateAuthorizationHeader()
                ConsistencyLevel = 'eventual'
            } 
        } catch {
            throw "Error getting application grant token: $_"
        }
    } elseif ($DelegatedGrant) {
        #Get a Graph API Toke via Application Grant
        try {
            $response = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -UserCredential $UserCredential -ErrorAction Stop
            $authToken = @{
                Authorization = $response.CreateAuthorizationHeader()
            }
        
        } catch {
            throw "Error getting delegated grant token: $_"
        }
    } else {
        Write-Error -Message "Missing parameter for permission type! -Delegated or -Application"
    }
    #returning token to function caller
    $script:UserTokenExpiresOnUTC = $response.ExpiresOn.UtcDateTime
    return $authToken
}#endfunction
#Fixing incorrectly formattet phone numbers (Denmark - also works for others)
function update-phoneFormatDK ($phoneNumber) {
    $parsedPhone = $phoneNumber -replace '\s',''
    if ($parsedPhone.Length -eq 8) {
        $parsedPhone = "+45 $parsedPhone"
    } elseif ($parsedPhone.Length -eq 11) {
        $parsedPhone = $parsedPhone.Insert(3," ")
    } elseif ($parsedPhone.Length -eq 13) {
        $parsedPhone = $parsedPhone.Insert(3," ")
    } else {$parsedPhone = $false}
    return $parsedPhone
}#endfunction

#Fixing incorrectly formattet phone numbers (Norway)
function update-phoneFormatNO ($phoneNumber) {
    if ($phoneNumber -match "^(\+47\s)()?[4,9]\d{7}$"){
        $parsedPhone = $phoneNumber
    } else{
        $parsedPhone = $phoneNumber -replace '\s',''
        if ($parsedPhone -match "^[4,9]\d{7}$") {
            $parsedPhone = "+47 $parsedPhone"
        } elseif ($parsedPhone -match "^(47)()?[4,9]\d{7}$") {
            $parsedPhone = "+$parsedPhone" 
            $parsedphone = $parsedPhone.Insert(3," ")
        } elseif ($parsedPhone -match "^(0047)()?[4,9]\d{7}$"){
            $parsedPhone = $parsedPhone.TrimStart("00")
            $parsedPhone = "+$parsedPhone" 
            $parsedphone = $parsedPhone.Insert(3," ")
        } elseif ($parsedPhone -match "^(\+47)()?[4,9]\d{7}$"){
            $parsedphone = $parsedPhone.Insert(3," ")
        } else {$parsedPhone = $false}
    }
    return $parsedPhone
}#endfunction
#endregion functions

#region execute
#region authentication
#Get Application auth token
if($appCert){
    $authTokenApp = get-authToken -ApplicationGrant -ClientCertificate $appCert -ClientId $AppId -TenantId $Tenant
} elseif ($appSecret) {
    $authTokenApp = get-authToken -ApplicationGrant -ClientSecret $appSecret -ClientId $AppId -TenantId $Tenant
} else {
    Write-Error -Message "Neither a Secret or a Client Certificate was provided! Exit Script"
    exit 1
}
#Get Delegated permissions auth token
$authTokenUser = get-authToken -DelegatedGrant -TenantId $Tenant -ClientId $AppId -UserCredential $authenticationCredentials
#endregion authentication

#region get applicable users
#Get all users with a mobile phone in Azure AD
#Check Runstate whether staging group is being used 
if ($UseStagingGroup -eq $true){
    Write-Output "Runstate is using Staging Group: $StagingGroupName"
    $StagingGroup = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Headers $authTokenApp -Resource "groups?filter=displayName eq `'$StagingGroupName`'"
    $stagingGroupId = $StagingGroup.value.id
    $allHasMobileUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Headers $authTokenApp -Resource "groups/$stagingGroupId/transitiveMembers/microsoft.graph.user?count=true&filter=userType ne 'Guest' and mobilePhone ne null"
    $allHasMobileUsersUPN = $allHasMobileUsers.userPrincipalName
} else {
    Write-Output "Runstate is Processing All Users"
    $allHasMobileUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Headers $authTokenApp -Resource "users?count=true&select=userPrincipalName,mobilePhone&filter=userType ne 'Guest' and mobilePhone ne null"
    $allHasMobileUsersUPN = $allHasMobileUsers.userPrincipalName
}

#Get all Non-MFA users principalName
$allNonMFAUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Headers $authTokenApp -Resource "reports/credentialUserRegistrationDetails?`$filter=isMfaRegistered+eq+false" 
$allNonMFAUsersUPN = $allNonMFAUsers.userPrincipalName

#Compare the two results and get only the non-MFA registered users that have a mobile phone in Azure AD so we can update their registration
Write-output "Stats: Numbers of users without MFA is: $($allNonMFAUsersUPN.count)"
if ($allNonMFAUsersUPN.count -eq 0 -or $allHasMobileUsersUPN.count -eq 0){
    Write-Output "There are no eligible users for this run. Exit Script"
    Exit 0
} else {
    $allUsersToRegisterWithMobile = (Compare-Object -ReferenceObject $allNonMFAUsersUPN -DifferenceObject $allHasMobileUsersUPN -Includeequal -ExcludeDifferent).InputObject
}
Write-output "Stats: Numbers of targeted users without MFA and a mobile number in AAD is: $($allUsersToRegisterWithMobile.Count)"
#endregion get applicable users

#region update MFA registration
#Provision users mobile phone number as authentication phone method
Write-Verbose "Updating users MFA registration details one at a time..." -Verbose
$Count = 0 
foreach ($user in $allUsersToRegisterWithMobile) {
    $RefreshTokenTime = ((Get-Date).ToUniversalTime()).AddMinutes(5)
    if ($RefreshTokenTime -ge $UserTokenExpiresOnUTC) {
        Write-Verbose -Message "Refreshing token before expiry.. continue loop script" -Verbose
        $authTokenUser = get-authToken -DelegatedGrant -TenantId $Tenant -ClientId $AppId -UserCredential $authenticationCredentials
    } 
    $userMobilePhone = ($allHasMobileUsers | Where-Object {$_.userPrincipalName -eq "$user"}).mobilePhone
    #fix incorrectly formatted mobile number - supported country mobile numbers are Norway
    $parsedPhone = update-phoneFormatNO -phoneNumber $userMobilePhone
    #If no match for NO try DK
    <#if ($parsedPhone -eq $false){
        $parsedPhone = update-phoneFormatDK -phoneNumber $userMobilePhone
    }#>
    #If still no match - skip (more )
    if ($parsedPhone -eq $false){
        Write-output "Status: Number Format error; User: $($user); Message: MobilePhone $userMobilePhone"
    } else {
        <#if ($parsedPhone -notmatch '((\+[0-9]{1,3}[ ])[0-9]{4,})'){
            Write-output "FORMAT ERROR: Mobile phone is $userMobilePhone - fix it!"
            continue
        }#>
        #sending update via the Graph API using delegated permissions (App permissions are not supported yet in beta)
        #formatting body for post action
        $ObjectBody = @{
            'phoneNumber' = "$parsedPhone"
            'phoneType' = "mobile"
        }
        $JSON = ConvertTo-Json -InputObject $ObjectBody  
        #sending update to graph and report on success
        try {
            $response = Invoke-MSGraphOperation -Post -APIVersion $graphVersion -Headers $authTokenUser -Body $JSON -Resource "users/$user/authentication/phoneMethods" -ErrorAction Stop
            write-output "Status: MFA Phonemethod provisioned successfully; User: $($user); Message: MobilePhone $($parsedPhone)"
            $Count++
            }
        catch {
            write-output "Status: MFA Phonemethod provisioning failed; User: $($user); Message: $($_.Exception.Message)"
        }
    }
}
Write-Output "Stats: $($Count) users has been processed successfully."
Write-Verbose "Execution completed!" -Verbose
#endregion update MFA registration
#endregion execute


