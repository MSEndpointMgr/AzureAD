<#
.SYNOPSIS
    Automaticly create Azure AD group for each license found in your tenant, and assign the license to this group.     

.DESCRIPTION
    Automaticly create Azure AD group for each license found in your tenant, and assign the license to this group.

.NOTES
    FileName:    Invoke-CreateLicensingGroups.ps1
    Author:      Jan Ketil Skanke
    Contact:     @JankeSkanke
    Created:     2020-15-03
    Updated:     2020-15-03

    Version history:
    1.0.0 - (2020-15-03) Script created
#>
function Get-MSGraphAppToken{
    <#  .SYNOPSIS
        Get an app based authentication token required for interacting with Microsoft Graph API
    .PARAMETER TenantID
        A tenant ID should be provided.
 
    .PARAMETER ClientID
        Application ID for an Azure AD application. Uses by default the Microsoft Intune PowerShell application ID.
 
    .PARAMETER ClientSecret
        Web application client secret.
        
    .EXAMPLE
        # Manually specify username and password to acquire an authentication token:
        Get-MSGraphAppToken -TenantID $TenantID -ClientID $ClientID -ClientSecert = $ClientSecret 
    .NOTES
        Author: Jan Ketil Skanke
        Contact: @JankeSkanke
        Created: 2020-15-03
        Updated: 2020-15-03
 
        Version history:
        1.0.0 - (2020-03-15) Function created      
    #>
[CmdletBinding()]
	param (
		[parameter(Mandatory = $true, HelpMessage = "Your Azure AD Directory ID should be provided")]
		[ValidateNotNullOrEmpty()]
		[string]$TenantID,
		[parameter(Mandatory = $true, HelpMessage = "Application ID for an Azure AD application")]
		[ValidateNotNullOrEmpty()]
		[string]$ClientID,
		[parameter(Mandatory = $true, HelpMessage = "Azure AD Application Client Secret.")]
		[ValidateNotNullOrEmpty()]
		[string]$ClientSecret
	    )
Process {
    $ErrorActionPreference = "Stop"
       
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
        }
    
    try {
        $MyTokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
        $MyToken =($MyTokenRequest.Content | ConvertFrom-Json).access_token
            If(!$MyToken){
                Write-Warning "Failed to get Graph API access token!"
                Exit 1
            }
        $MyHeader = @{"Authorization" = "Bearer $MyToken" }

       }
    catch [System.Exception] {
        Write-Warning "Failed to get Access Token, Error message: $($_.Exception.Message)"; break
    }
    return $MyHeader
    }
}
## ----- Authenticate Section Start ----- ##
#Set Tenant and AzureAD App Information variables 
$ClientID = ""
$ClientSecret = ""
$tenantId = ""

#Get the Auth Token from Azure AD and get the header back for authentication
$Header = Get-MSGraphAppToken -TenantID $tenantId -ClientID $ClientID -ClientSecret $ClientSecret 
## ----- Authenticate Section End ----- ##

#Query all data about Licenses and Groups 
$LicenseQuery = Invoke-RestMethod -Method Get -Uri 'https://graph.microsoft.com/beta/subscribedSkus/' -Headers $Header
$GroupsQuery = Invoke-RestMethod -Method Get -Uri 'https://graph.microsoft.com/beta/groups/' -Headers $Header

#Mapping Table for SKU's to functional Name of License Groups
$Sku = @{
	'MCOMEETADV' = 'LIC_(AUDIO CONFERENCING)'
	'AAD_BASIC' = 'LIC_(AZURE ACTIVE DIRECTORY BASIC)'
	'AAD_PREMIUM' = 'LIC_(AZURE ACTIVE DIRECTORY PREMIUM P1)'
	'AAD_PREMIUM_P2' = 'LIC_(AZURE ACTIVE DIRECTORY PREMIUM P2)'
	'RIGHTSMANAGEMENT' = 'LIC_(AZURE INFORMATION PROTECTION PLAN 1)'
	'DYN365_ENTERPRISE_PLAN1' = 'LIC_(DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN ENTERPRISE EDITION)'
	'DYN365_ENTERPRISE_CUSTOMER_SERVICE' = 'LIC_(DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION)'
	'DYN365_FINANCIALS_BUSINESS_SKU' = 'LIC_(DYNAMICS 365 FOR FINANCIALS BUSINESS EDITION)'
	'DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE' = 'LIC_(DYNAMICS 365 FOR SALES AND CUSTOMER SERVICE ENTERPRISE EDITION)'
	'DYN365_ENTERPRISE_SALES' = 'LIC_(DYNAMICS 365 FOR SALES ENTERPRISE EDITION)'
	'DYN365_ENTERPRISE_TEAM_MEMBERS' = 'LIC_(DYNAMICS 365 FOR TEAM MEMBERS ENTERPRISE EDITION)'
	'Dynamics_365_for_Operations' = 'LIC_(DYNAMICS 365 UNF OPS PLAN ENT EDITION)'
	'EMS' = 'LIC_(ENTERPRISE MOBILITY SECURITY E3)'
	'EMSPREMIUM' = 'LIC_(ENTERPRISE MOBILITY SECURITY E5)'
	'EXCHANGESTANDARD' = 'LIC_(EXCHANGE ONLINE (PLAN 1))'
	'EXCHANGEENTERPRISE' = 'LIC_(EXCHANGE ONLINE (PLAN 2))'
	'EXCHANGEARCHIVE_ADDON' = 'LIC_(EXCHANGE ONLINE ARCHIVING FOR EXCHANGE ONLINE)'
	'EXCHANGEARCHIVE' = 'LIC_(EXCHANGE ONLINE ARCHIVING FOR EXCHANGE SERVER)'
	'EXCHANGEESSENTIALS' = 'LIC_(EXCHANGE ONLINE ESSENTIALS)'
	'EXCHANGE_S_ESSENTIALS' = 'LIC_(EXCHANGE ONLINE ESSENTIALS)'
	'EXCHANGEDESKLESS' = 'LIC_(EXCHANGE ONLINE KIOSK)'
	'EXCHANGETELCO' = 'LIC_(EXCHANGE ONLINE POP)'
	'INTUNE_A' = 'LIC_(INTUNE)'
	'M365EDU_A1' = 'LIC_(Microsoft 365 A1)'
	'M365EDU_A3_FACULTY' = 'LIC_(Microsoft 365 A3 for faculty)'
	'M365EDU_A3_STUDENT' = 'LIC_(Microsoft 365 A3 for students)'
	'M365EDU_A5_FACULTY' = 'LIC_(Microsoft 365 A5 for faculty)'
	'M365EDU_A5_STUDENT' = 'LIC_(Microsoft 365 A5 for students)'
	'SPB' = 'LIC_(MICROSOFT 365 BUSINESS)'
	'SPE_E3' = 'LIC_(MICROSOFT 365 E3)'
	'SPE_E3_USGOV_DOD' = 'LIC_(Microsoft 365 E3_USGOV_DOD)'
	'SPE_E3_USGOV_GCCHIGH' = 'LIC_(Microsoft 365 E3_USGOV_GCCHIGH)'
	'SPE_E5' = 'LIC_(Microsoft 365 E5)'
	'INFORMATION_PROTECTION_COMPLIANCE' = 'LIC_(Microsoft 365 E5 Compliance)'
	'IDENTITY_THREAT_PROTECTION' = 'LIC_(Microsoft 365 E5 Security)'
	'IDENTITY_THREAT_PROTECTION_FOR_EMS_E5' = 'LIC_(Microsoft 365 E5 Security for EMS E5)'
	'SPE_F1' = 'LIC_(Microsoft 365 F1)'
	'WIN_DEF_ATP' = 'LIC_(Microsoft Defender Advanced Threat Protection)'
	'CRMSTANDARD' = 'LIC_(MICROSOFT DYNAMICS CRM ONLINE)'
	'CRMPLAN2' = 'LIC_(MICROSOFT DYNAMICS CRM ONLINE BASIC)'
	'IT_ACADEMY_AD' = 'LIC_(MS IMAGINE ACADEMY)'
	'ENTERPRISEPREMIUM_FACULTY' = 'LIC_(Office 365 A5 for faculty)'
	'ENTERPRISEPREMIUM_STUDENT' = 'LIC_(Office 365 A5 for students)'
	'EQUIVIO_ANALYTICS' = 'LIC_(Office 365 Advanced Compliance)'
	'ATP_ENTERPRISE' = 'LIC_(Office 365 Advanced Threat Protection (Plan 1))'
	'O365_BUSINESS' = 'LIC_(OFFICE 365 BUSINESS)'
	'SMB_BUSINESS' = 'LIC_(OFFICE 365 BUSINESS)'
	'SMB_APPS' = 'LIC_(Business Apps)'	
	'O365_BUSINESS_ESSENTIALS' = 'LIC_(OFFICE 365 BUSINESS ESSENTIALS)'
	'SMB_BUSINESS_ESSENTIALS' = 'LIC_(OFFICE 365 BUSINESS ESSENTIALS)'
	'O365_BUSINESS_PREMIUM' = 'LIC_(OFFICE 365 BUSINESS PREMIUM)'
	'SMB_BUSINESS_PREMIUM' = 'LIC_(OFFICE 365 BUSINESS PREMIUM)'
	'STANDARDPACK' = 'LIC_(OFFICE 365 E1)'
	'STANDARDWOFFPACK' = 'LIC_(OFFICE 365 E2)'
	'ENTERPRISEPACK' = 'LIC_(OFFICE 365 E3)'
	'DEVELOPERPACK' = 'LIC_(OFFICE 365 E3 DEVELOPER)'
	'ENTERPRISEPACK_USGOV_DOD' = 'LIC_(Office 365 E3_USGOV_DOD)'
	'ENTERPRISEPACK_USGOV_GCCHIGH' = 'LIC_(Office 365 E3_USGOV_GCCHIGH)'
	'ENTERPRISEWITHSCAL' = 'LIC_(OFFICE 365 E4)'
	'ENTERPRISEPREMIUM' = 'LIC_(OFFICE 365 E5)'
	'ENTERPRISEPREMIUM_NOPSTNCONF' = 'LIC_(OFFICE 365 E5 WITHOUT AUDIO CONFERENCING)'
	'DESKLESSPACK' = 'LIC_(OFFICE 365 F1)'
	'MIDSIZEPACK' = 'LIC_(OFFICE 365 MIDSIZE BUSINESS)'
	'OFFICESUBSCRIPTION' = 'LIC_(OFFICE 365 PROPLUS)'
	'LITEPACK' = 'LIC_(OFFICE 365 SMALL BUSINESS)'
	'LITEPACK_P2' = 'LIC_(OFFICE 365 SMALL BUSINESS PREMIUM)'
	'WACONEDRIVESTANDARD' = 'LIC_(ONEDRIVE FOR BUSINESS (PLAN 1))'
	'WACONEDRIVEENTERPRISE' = 'LIC_(ONEDRIVE FOR BUSINESS (PLAN 2))'
	'POWERAPPS_PER_USER' = 'LIC_(POWER APPS PER USER PLAN)'
	'POWER_BI_ADDON' = 'LIC_(POWER BI FOR OFFICE 365 ADD-ON)'
	'POWER_BI_PRO' = 'LIC_(POWER BI PRO)'
	'POWER_BI_STANDARD' = 'LIC_(Power BI Free)'
	'PROJECTCLIENT' = 'LIC_(PROJECT FOR OFFICE 365)'
	'PROJECTESSENTIALS' = 'LIC_(PROJECT ONLINE ESSENTIALS)'
	'PROJECTPREMIUM' = 'LIC_(PROJECT ONLINE PREMIUM)'
	'PROJECTONLINE_PLAN_1' = 'LIC_(PROJECT ONLINE PREMIUM WITHOUT PROJECT CLIENT)'
	'PROJECTPROFESSIONAL' = 'LIC_(PROJECT ONLINE PROFESSIONAL)'
	'PROJECTONLINE_PLAN_2' = 'LIC_(PROJECT ONLINE WITH PROJECT FOR OFFICE 365)'
	'SHAREPOINTSTANDARD' = 'LIC_(SHAREPOINT ONLINE (PLAN 1))'
	'SHAREPOINTENTERPRISE' = 'LIC_(SHAREPOINT ONLINE (PLAN 2))'
	'MCOEV' = 'LIC_(SKYPE FOR BUSINESS CLOUD PBX)'
	'MCOIMP' = 'LIC_(SKYPE FOR BUSINESS ONLINE (PLAN 1))'
	'MCOSTANDARD' = 'LIC_(SKYPE FOR BUSINESS ONLINE (PLAN 2))'
	'MCOPSTN2' = 'LIC_(SKYPE FOR BUSINESS PSTN DOMESTIC AND INTERNATIONAL CALLING)'
	'MCOPSTN1' = 'LIC_(SKYPE FOR BUSINESS PSTN DOMESTIC CALLING)'
	'MCOPSTN5' = 'LIC_(SKYPE FOR BUSINESS PSTN DOMESTIC CALLING (120 Minutes))'
	'VISIOONLINE_PLAN1' = 'LIC_(VISIO ONLINE PLAN 1)'
	'VISIOCLIENT' = 'LIC_(VISIO Online Plan 2)'
	'WIN10_PRO_ENT_SUB' = 'LIC_(WINDOWS 10 ENTERPRISE E3)'
	'WIN10_VDA_E5' = 'LIC_(Windows 10 Enterprise E5)'	
	'FLOW_FREE' = "LIC_(Microsoft Power Automate Free)"
}

#Creating Groups and adding licenses based on subscribed SKUs for each SKU found in the tenant
foreach ($lic in $LicenseQuery.value){
	#Check if License is in a good state and assignable to a user/group 
	if ($lic.capabilityStatus -eq "Enabled"){
		if($lic.appliesTo -ne "Company"){
			#Grab the Group name from the mapping table dynamicly based on skuPartNumber
			Write-Output "Subscribed license found: $($lic.skuPartNumber)"
			$DisplayName = $SKU.($lic.skuPartNumber)
			Write-Output $DisplayName
			
			#Check if group exist and create group if missing
			if ($DisplayName  -in ($GroupsQuery.value).displayName) {
				#Group exists
				Write-Output "Group already exists"
				#Get Group Object from Graph 
				$GroupUri = -join("https://graph.microsoft.com/beta/groups?filter=displayname eq '", $DisplayName, "'")
				$Group = (Invoke-RestMethod -Method GET -Uri $GroupUri -ContentType "application/json" -Headers $Header ).value
			}
			else {
				#Group needs to be created. Formatting JSON for Graph request. 
				$newGroupJSONObject = @{
					"description" = "Script Created license group by subscried SKU's"
					"displayName"= $DisplayName
					"mailEnabled" = $false
					"mailNickname" = "none"
					"securityEnabled" = $true
				} | ConvertTo-Json
				#Creating the Group
				$Group = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/beta/groups/' -ContentType "application/json" -Headers $Header -Body $newGroupJSONObject 
				Write-Output "Added Group $($Group.displayName) with GroupID $($Group.id)"
			}
			#Check if group has correct license and add license if not 
			$GroupID = $Group.id
			#Getting the licenses assigned to the Group
			$LicUri = -join ('https://graph.microsoft.com/beta/groups/',$GroupID,'/assignedLicenses')
			$License = Invoke-RestMethod -Method Get -Uri $LicUri -Headers $Header
			#Checking if the license is already assigned. 
			if (($License.value).skuId -match $lic.skuId) {
				Write-Output "License already assigned to the group"
			}	
			else {
				#License not assigned, assigning license to the group. 
				Write-Output "Assigning license to group"
				$URI = -join ('https://graph.microsoft.com/beta/groups/',$Group.id,'/assignLicense')
				#Formatting the licensing JSON to add the SKU to the group. 
				$newLicenseJSONObject = @{
					addLicenses = @(
						@{
							skuId = $lic.skuId
						}
					)
					removeLicenses = @()
				} | ConvertTo-Json -Depth 3
				#Assigning the license
				Invoke-RestMethod -Method POST -Uri $URI -ContentType "application/json" -Headers $Header -Body $newLicenseJSONObject | Out-Null
				Write-Output "Added License  $($lic.skuPartNumber) to $($Group.displayName)"
			}		
		}
	}
}
