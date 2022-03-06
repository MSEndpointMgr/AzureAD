<#
    .SYNOPSIS
        This script is for getting all non-MFA registered users - and setting their authentication phone for SMS
    .DESCRIPTION
        You need to configure the required strings as Azure Automation variables (EnableStagingGroup and StagingGroupName).
        Create an System Managed MSI that has application permissions: user.read.all reports.read.all groupmember.read.all and UserAuthenticationMethod.ReadWrite.All
        BLOG POST: https://msendpointmgr.com/2020/10/19/prepopulate-mfa-phone-authentication-solution/
    .NOTES
        Authors:     Jan Ketil Skanke, Sandy Zeng, Michael Mardahl
        Contact:     @JankeSkanke + @Sandy_Tsang + @Michael_Mardahl on twitter
        License:     MIT - Leave author details
        Created:     2020-10-07
        Updated:     2022-03-03
        Version history:
        1.0.0 - (2020-05-20) Initial Version
        1.1.0 - (2020-10-22) Ready for blog release
        1.1.1 - (2020-10-30) Bugix + Generic Parser
        2.0.0 - (2021-03-23) Script updated to use App Token only. All references to Delegated auth removed - update by Jan Ketil Skanke
		2.1.0 - (2022-03-03) Script updated to use System Managed Identity in Azure Automation and MSGraphRequest Module - update by Michael Mardahl
    #>    

#requires -module MSAL.PS, MSGraphRequest, Az.Accounts
#region declarations
$script:graphVersion = 'beta'
$UseStagingGroup = Get-AutomationVariable -Name "EnableStagingGroup" #Set this Automation variable to $true or $false to enable use of the Staging/Pilot group.
$StagingGroupName = Get-AutomationVariable -Name "StagingGroupName" #Set this Automation variable to the name of your staging group or comment out if you absolutely never are going to use staging.
$VerbosePreference = "SilentlyContinue"
#endregion declarations

#region functions
#Fixing incorrectly formattet phone numbers (DK, PL, UK - also works for some others)
function update-phoneFormat ($phoneNumber) {
    if ($phoneNumber -notmatch '((\+[0-9]{1,3}[ ])[0-9]{4,})'){
        #the number does not comply with graph requirements, so this will try to fix it.
        $parsedPhone = $phoneNumber -replace '\s',''    #remove all spaces
        if ($parsedPhone.Length -eq 6) {
            $parsedPhone = "+298 $parsedPhone"          #adding DK faroe islands country code
        } elseif ($parsedPhone.Length -eq 8) {
            $parsedPhone = "+45 $parsedPhone"           #adding PL country code
        } elseif ($parsedPhone.Length -eq 9) {
            $parsedPhone = "+48 $parsedPhone"           #adding PL country code
        } elseif ($parsedPhone.Length -eq 10) {
            $parsedPhone = $parsedPhone.Insert(4," ")   #adding missing space after country code for DK (faroe islands)
        } elseif ($parsedPhone.Length -eq 11) {
            $parsedPhone = $parsedPhone.Insert(3," ")   #adding missing space after country code for DK
        } elseif ($parsedPhone.Length -eq 12) {
            $parsedPhone = $parsedPhone.Insert(3," ")   #adding missing space after country code for PL
        } elseif ($parsedPhone.Length -eq 13) {
            $parsedPhone = $parsedPhone.Insert(3," ")   #adding missing space after country code for UK
        } else {
            $parsedPhone = $false
        }
    } else {
        $parsedPhone = $phoneNumber
    }
    return $parsedPhone
}#endfunction

#endregion functions

#region execute
#region authentication

#Connect with Automation Identitiy
Connect-AZAccount -Identity

#Get current access token
$AzAuthToken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com/"
#Convert to web token
$MgAuthToken = @{
	"Content-Type" = "application/json"
	"Authorization" = $AzAuthToken.Token
	"ExpiresOn" = $AzAuthToken.ExpiresOn
	"ocp-client-name" = "AA-MFAPrepopulateScript"
    "ocp-client-version" = "1.0"
	"ConsistencyLevel" = "eventual"
} 
#Set authentication header for future MSGraphRequest module operations
$Global:AuthenticationHeader = $MgAuthToken

#endregion authentication

#region get applicable users
#Get all users with a mobile phone in Azure AD
#Check Runstate whether staging group is being used 
if ($UseStagingGroup -eq $true){
    Write-Output "Runstate is using staging group: $StagingGroupName"
    $StagingGroup = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Resource "groups?filter=displayName eq `'$StagingGroupName`'" -ErrorAction Stop
    $stagingGroupId = $StagingGroup.value.id
    $allHasMobileUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Resource "groups/$stagingGroupId/transitiveMembers/microsoft.graph.user?count=true&filter=userType ne 'Guest' and mobilePhone ne null" -ErrorAction Stop
    $allHasMobileUsersUPN = $allHasMobileUsers.userPrincipalName
} else {
    Write-Output "Runstate is Processing all users"
    $allHasMobileUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Resource "users?count=true&select=userPrincipalName,mobilePhone&filter=userType ne 'Guest' and mobilePhone ne null" -ErrorAction Stop
    $allHasMobileUsersUPN = $allHasMobileUsers.userPrincipalName
}

#Get all Non-MFA users principalName
$allNonMFAUsers = Invoke-MSGraphOperation -Get -APIVersion $graphVersion -Resource "reports/credentialUserRegistrationDetails?`$filter=isMfaRegistered+eq+false"  -ErrorAction Stop
$allNonMFAUsersUPN = $allNonMFAUsers.userPrincipalName

#Compare the two results and get only the non-MFA registered users that have a mobile phone in Azure AD so we can update their registration
Write-output "Stats: Number of users without MFA is: $($allNonMFAUsersUPN.count)"
if ($allNonMFAUsersUPN.count -eq 0 -or $allHasMobileUsersUPN.count -eq 0){
    Write-Output "There are no eligible users for this run. Exit Script"
    Exit 0
} else {
    $allUsersToRegisterWithMobile = (Compare-Object -ReferenceObject $allNonMFAUsersUPN -DifferenceObject $allHasMobileUsersUPN -Includeequal -ExcludeDifferent).InputObject
}
Write-output "Stats: Number of targeted users without MFA and a mobile number in AAD is: $($allUsersToRegisterWithMobile.Count)"
#endregion get applicable users

#region update MFA registration
#Provision users mobile phone number as authentication phone method
Write-Verbose "Updating users MFA registration details one at a time..." -Verbose
$Count = 0 
foreach ($user in $allUsersToRegisterWithMobile) {
    $userMobilePhone = ($allHasMobileUsers | Where-Object {$_.userPrincipalName -eq "$user"}).mobilePhone
    
    #fix incorrectly formatted mobile number
    $parsedPhone = update-phoneFormat -phoneNumber $userMobilePhone
    
    #If still no match - skip
    if ($parsedPhone -eq $false){
        Write-output "Status: Number Format error; User: $($user); Message: MobilePhone $userMobilePhone"
        continue #skipping this iteration of the loop
    } 

    #formatting body for post action
    $ObjectBody = @{
        'phoneNumber' = "$parsedPhone"
        'phoneType' = "mobile"
    }
    $JSON = ConvertTo-Json -InputObject $ObjectBody  
    #sending update to graph and report on success
    try {
        $response = Invoke-MSGraphOperation -Post -APIVersion $graphVersion -Body $JSON -Resource "users/$user/authentication/phoneMethods" -ErrorAction Stop
        write-output "Status: MFA Phonemethod provisioned successfully; User: $($user); Message: MobilePhone $($parsedPhone)"
        $Count++
        }
    catch {
        write-output "Status: MFA Phonemethod provisioning failed; User: $($user); Message: $($_.Exception.Message)"
    }
    
}
Write-Output "Stats: $($Count) users have been processed successfully."
Write-Verbose "Execution completed!" -Verbose
#endregion update MFA registration
#endregion execute
