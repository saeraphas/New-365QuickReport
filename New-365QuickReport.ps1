<#
.SYNOPSIS
	This script collects data from Exchange Online and Microsoft Graph and builds a report intended for periodic housekeeping, license audits, true-ups, etc.

.DESCRIPTION
	Lists all accounts, mailboxes, and groups in a Microsoft 365 tenant, along with sign-in, license, and role information.

.EXAMPLE
	.\New-365QuickReport.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
    Additional Credit:  Troy Hayes (thayes@nexigen.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>

Param (
	[Parameter(ValueFromPipelineByPropertyName)]
	[switch] $SkipUpdateCheck,
	[switch] $SkipUserReport,
	[switch] $SkipMailboxReport,
	[switch] $SkipGroupReport,
	[switch] $SkipSKUConversion,
	[switch] $GCCHigh,
	[switch] $ClearMGTokenCache
)

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#check for Fortinet EDR since it breaks everything useful
$FortinetEDRPresent = Get-Service -Name "FortiEDR Collector Service" -ErrorAction SilentlyContinue
if ($FortinetEDRPresent) { Write-Warning "Fortinet EDR service present on this machine. Network connections may be blocked." }

#remove cached mggraph tokens to address auth prompts for every query
#https://github.com/microsoftgraph/msgraph-sdk-powershell#known-issues
if ($ClearMGTokenCache) {
	Disconnect-MgGraph | Out-Null
	Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force
	Write-Warning "User profile Graph SDK token cache reset. This script will now exit."
	exit
}

function CheckPrerequisites($PrerequisiteModulesTable) {
	$PrerequisiteModules = $PrerequisiteModulesTable | ConvertFrom-Csv
	$ProgressActivity = "Checking for prerequisite modules."
	ForEach ( $PrerequisiteModule in $PrerequisiteModules ) {
		$moduleName = $($PrerequisiteModule.Name)
		$ProgressOperation = "Checking for module $moduleName."
		Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
		$minimumVersion = $($PrerequisiteModule.minimumversion)
		$installedversion = $(Get-Module -ListAvailable -Name $moduleName | Select-Object -first 1).version
		If (!($installedversion)) {
			try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -MinimumVersion $minimumversion } catch { Write-Error "An error occurred installing $moduleName." }
		}
		elseif ([version]$installedversion -lt [version]$minimumversion) {
			try { Uninstall-Module $moduleName -AllVersions } catch { Write-Error "An error occurred removing $moduleName. You may need to manually remove old versions using admin privileges." }
			try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -MinimumVersion $minimumversion } catch { Write-Error "An error occurred installing $moduleName." }
		}
	}
	Write-Progress -Activity $ProgressActivity -Completed
}

function New-Report() {
	Param($ReportName, $ReportData, $ReportOutput)
	Write-Progress -Id 0 -Activity "Collecting $ReportName data." -CurrentOperation "Building Report."
	if (!$reportdata) {
		$reportdata = @()
		$row = New-Object PSObject
		$row | Add-Member -MemberType NoteProperty -Name "Result" -Value "This report is empty."
		$reportdata += $row
	}
	$reportdata | Export-Excel `
		-Path $ReportOutput `
		-WorkSheetname "$ReportName" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter
	$reportdata = $null
	Write-Progress -Id 0 -Activity "Collecting $ReportName data." -Completed
}

#prerequisite modules and minimum versions as embedded CSV
$PrerequisiteModulesTable = @'
Name,MinimumVersion
Microsoft.Graph,1.25.0
ExchangeOnlineManagement,3.1.0
ImportExcel,7.0.0
'@
CheckPrerequisites($PrerequisiteModulesTable)

$PSNewLine = [System.Environment]::Newline

#download SKU and product name reference information from Microsoft Learn
#https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$MicrosoftDocumentationURI = 'https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference'
$MicrosoftDocumentationCSVURI = ((Invoke-WebRequest -UseBasicParsing -Uri $MicrosoftDocumentationURI).Links | Where-Object { $_.href -like 'http*' } | Where-Object { $_.href -like '*.csv' }).href
$MicrosoftDocumentationDownloadError = "An error occurred while downloading the SKU and Product Name information." + $PSNewLine + "SKU names will not be converted to Product Names."
try { $MicrosoftProducts = Invoke-RestMethod -Uri $MicrosoftDocumentationCSVURI -Method Get | ConvertFrom-CSV | Select-Object String_ID, Product_Display_Name, GUID -Unique } catch { Write-Warning $MicrosoftDocumentationDownloadError; $SkipSKUConversion = $true }

#connect to microsoft services
$ProgressActivity = "Connecting to Microsoft services. You will be prompted twice."
$ProgressOperation = "1 of 2 - Connecting to Exchange Online."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete 0
If ($GCCHigh) { $ExchangeEnvironmentName = "O365USGovGCCHigh" } else { $ExchangeEnvironmentName = "O365Default" }
try { Connect-ExchangeOnline -ExchangeEnvironmentName $ExchangeEnvironmentName -ShowBanner:$false | Out-Null } catch { write-error "Error connecting to Exchange Online. Exiting."; exit }

$ProgressOperation = "2 of 2 - Connecting to Microsoft Graph."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete 50
$Scopes = "Domain.Read.All,User.Read.All,UserAuthenticationMethod.Read.All,RoleManagement.Read.Directory,Group.Read.All,GroupMember.Read.All,OrgContact.Read.All,Device.Read.All"
If ($GCCHigh) { $MSGraphEnvironmentName = "USGovDoD" } else { $MSGraphEnvironmentName = "Global" }
try { Connect-MgGraph -Environment $MSGraphEnvironmentName -Scopes $Scopes | Out-Null } catch { write-error "Not connected to MS Graph. Exiting."; exit }
Write-Progress -Activity $ProgressActivity -Completed

#define variables for file system paths
$DateString = ((get-date).tostring("yyyy-MM-dd"))
#$TenantString = (Get-AcceptedDomain | Where-Object { $_.Default }).name
$TenantString = (Get-MgDomain | Where-Object { $_.isInitial }).Id
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$TenantPath = "$DesktopPath\365QuickReport\$TenantString"
$ReportPath = "$TenantPath\Reports"
$XLSreport = "$ReportPath\$TenantString-report-$DateString.xlsx"

#Get all role definitions from Microsoft Graph (for display names) #Thanks, Troy
$roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition

If ($SkipUserReport) { Write-Verbose "Skipping user report." ; $SkipMailboxReport = $true } else {

	$ProgressActivity = "Retrieving 365 user account data."
	$ProgressOperation = "Retrieving user list."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation

	#Get the 365 user list using Microsoft Graph
	#construct report output object
	$365UserReportObject = @()
	$MGUsers = Get-MGUser -All -Property ID, UserPrincipalName, AccountEnabled, OnPremisesSyncEnabled, DisplayName, Department, JobTitle, Mail, CreatedDateTime, LastPasswordChangeDateTime, AssignedLicenses, Manager | Where-Object { $_.userprincipalname -notmatch '#EXT#@' } | Select-Object ID, UserPrincipalName, AccountEnabled, OnPremisesSyncEnabled, DisplayName, Department, JobTitle, Mail, CreatedDateTime, LastPasswordChangeDateTime, AssignedLicenses, Manager
	$MGUserProgressBarCounter = 0
	Foreach ($MGUser in $MGUsers) {
		$MGUserProgressBarCounter++
		$DisplayName = $MGUser.DisplayName
		$ProgressOperation = "Retrieving user data for $DisplayName."
		$ProgressPercent = ($MGUserProgressBarCounter / $($MGUsers).count) * 100
		Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete $ProgressPercent

		$MGUserEnabled = $null
		if ($MGUser.AccountEnabled -eq $true) { $MGUserEnabled = "allowed" } else { $MGUserEnabled = "blocked" }

		$MGUserPasswordAge = (New-TimeSpan -Start $MGUser.LastPasswordChangeDateTime).Days

		# slow
		# $MGUserLicenses = $(get-mguserlicensedetail -userid $($MGUser).id).SkuPartNumber
		# #convert SKUs to Product Names unless bypassed or downloading the CSV from documentation failed earlier
		# IF ($SkipSKUConversion) { $MGUserLicenseProductNames = $MGUserLicenses -join "," } else {
		#     $MGUserLicenseProductNameArray = @()
		#     if ($MGUserLicenses.count -eq 0) { $MGUserLicenseProductNames = "none" } else {
		#         foreach ($License in $MGUserLicenses) {
		#             $ProductName = $($MicrosoftProducts | Where-Object { $_.String_ID -eq $License }).Product_Display_Name
		#             if (!($ProductName)) { $MGUserLicenseProductNameArray += $License } else { $MGUserLicenseProductNameArray += $ProductName }
		#         }
		#         $MGUserLicenseProductNames = $MGUserLicenseProductNameArray -join ","
		#     }
		# }

		$MGUserLicenseGUIDs = $MGUser.AssignedLicenses
		#convert GUIDs to Product Names
		$MGUserLicenseProductNameArray = @()
		if ($MGUserLicenseGUIDs.count -eq 0) { $MGUserLicenseProductNames = "none" } else {
			foreach ($License in $MGUserLicenseGUIDs) {
				$License = $($License | Select-Object -ExpandProperty SkuId).trim('{}') #the license GUIDs have brackets, but the reference list doesn't
				$ProductName = $($MicrosoftProducts | Where-Object { $_.GUID -eq $License }).Product_Display_Name
				if (!($ProductName)) { $MGUserLicenseProductNameArray += $License } else { $MGUserLicenseProductNameArray += $ProductName }
			}
			$MGUserLicenseProductNames = $MGUserLicenseProductNameArray -join ","
		}

		#Get user's role assignments from Graph API #Thanks, Troy
		$MGUserRoleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalId eq '$($MGUser.ID)'" | Select-Object RoleDefinitionId
		$MGUserRoleArray = @()
		#Match role definition IDs to display names #Thanks, Troy
		if ($MGUserRoleAssignments.count -eq 0) { $MGUserRoles = "none" } else {
			foreach ($RoleAssignment in $MGUserRoleAssignments) {
				$MGUserRoleArray += ($roleDefinitions | Where-Object { $_.Id -eq $RoleAssignment.RoleDefinitionId }).DisplayName
			}
			$MGUserRoles = $MGUserRoleArray -join ","
		}

		$MGUserManager = $null
		$MGUserManager = $(Get-MgUser -UserId $($MGUser).id -ExpandProperty manager | Select-Object @{Name = 'Manager'; Expression = { $_.Manager.AdditionalProperties.displayName } }).Manager

		#Get MFA data
		$MFAStatus = $null
		$MFAPhone = $null
		$MicrosoftAuthenticatorDevice = $null
		$Is3rdPartyAuthenticatorUsed = $null
		[array]$MFAData = Get-MgUserAuthenticationMethod -UserId $($MGUser).id
		$AuthenticationMethod = @()
		$AdditionalDetails = @()

		foreach ($MFA in $MFAData) {
			Switch ($MFA.AdditionalProperties["@odata.type"]) {
				"#microsoft.graph.passwordAuthenticationMethod" {
					$AuthMethod = 'PasswordAuthentication'
					$AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
				}
				"#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
					# Microsoft Authenticator App
					$AuthMethod = 'AuthenticatorApp'
					$AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
					$MicrosoftAuthenticatorDevice = $MFA.AdditionalProperties["displayName"]
				}
				"#microsoft.graph.phoneAuthenticationMethod" {
					# Phone authentication
					$AuthMethod = 'PhoneAuthentication'
					$AuthMethodDetails = $MFA.AdditionalProperties["phoneType", "phoneNumber"] -join ' '
					$MFAPhone = $MFA.AdditionalProperties["phoneNumber"]
				}
				"#microsoft.graph.fido2AuthenticationMethod" {
					# FIDO2 key
					$AuthMethod = 'Fido2'
					$AuthMethodDetails = $MFA.AdditionalProperties["model"]
				}
				"#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
					# Windows Hello
					$AuthMethod = 'WindowsHelloForBusiness'
					$AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
				}
				"#microsoft.graph.emailAuthenticationMethod" {
					# Email Authentication
					$AuthMethod = 'EmailAuthentication'
					$AuthMethodDetails = $MFA.AdditionalProperties["emailAddress"]
				}
				"microsoft.graph.temporaryAccessPassAuthenticationMethod" {
					# Temporary Access pass
					$AuthMethod = 'TemporaryAccessPass'
					$AuthMethodDetails = 'Access pass lifetime (minutes): ' + $MFA.AdditionalProperties["lifetimeInMinutes"]
				}
				"#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod" {
					# Passwordless
					$AuthMethod = 'PasswordlessMSAuthenticator'
					$AuthMethodDetails = $MFA.AdditionalProperties["displayName"]
				}
				"#microsoft.graph.softwareOathAuthenticationMethod" {
					$AuthMethod = 'SoftwareOath'
					$Is3rdPartyAuthenticatorUsed = "True"
				}

			}
			$AuthenticationMethod += $AuthMethod
			if ($null -ne $AuthMethodDetails) {
				$AdditionalDetails += "$AuthMethod : $AuthMethodDetails"
			}
		}
		#To remove duplicate authentication methods
		$AuthenticationMethod = $AuthenticationMethod | Sort-Object | Get-Unique
		#    $AuthenticationMethods = $AuthenticationMethod -join ","
		$AdditionalDetail = $AdditionalDetails -join ", "

		#Determine MFA status
		[array]$MFAMethods = ("Fido2", "PhoneAuthentication", "PasswordlessMSAuthenticator", "AuthenticatorApp", "WindowsHelloForBusiness", "SoftwareOath")
		foreach ($MFAMethod in $MFAMethods) { if ($AuthenticationMethod -contains $MFAMethod) { $MFAStatus = "Enabled"; break } }

		# build result object
		$userHash = $null
		$userHash = @{
			'UserPrincipalName'      = $MGUser.userPrincipalName
			'DisplayName'            = $MGUser.DisplayName
			'Sign-In'                = $MGUserEnabled
			'Synced'                 = $MGUser.OnPremisesSyncEnabled
			'Department'             = $MGUser.Department
			'Title'                  = $MGUser.JobTitle
			'PasswordAge'            = $MGUserPasswordAge
			'Licenses'               = $MGUserLicenseProductNames
			'Roles'                  = $MGUserRoles
			'Manager'                = $MGUserManager
			#        'AuthMethods'        = $AuthenticationMethods
			'MFA_Status'             = $MFAStatus
			'MFA_Phone'              = $MFAPhone
			'MS_Authenticator'       = $MicrosoftAuthenticatorDevice
			'3P_Authenticator'       = $Is3rdPartyAuthenticatorUsed
			'MFA_Additional_Details' = $AdditionalDetail
		}
		$userObject = $null
		$userObject = New-Object PSObject -Property $userHash
		$365UserReportObject += $userObject
	}
	Write-Progress -Activity $ProgressActivity -Completed

	$ProgressActivity = "Building Excel report."
	$ProgressOperation = "Exporting to Excel."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
	$365UserReportObject | Select-Object UserPrincipalName, DisplayName, Sign-In, Synced, Department, Title, PasswordAge, Licenses, Roles, Manager, MFA_Status, MFA_Phone, MS_Authenticator, 3P_Authenticator, MFA_Additional_Details | Sort-Object -Property UserPrincipalName | Export-Excel `
		-Path $XLSreport `
		-WorkSheetname "365 Users" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter `
		-ConditionalText $(
		New-ConditionalText "blocked" -ConditionalTextColor DarkRed -BackgroundColor LightPink
		New-ConditionalText "Never Signed In" -ConditionalTextColor DarkRed -BackgroundColor LightPink
		New-ConditionalText "Global Administrator" -BackgroundColor Yellow
	)
	Write-Progress -Activity $ProgressActivity -Completed
}

#check whether mailbox report skip is set by parameter
If ($SkipMailboxReport) { Write-Verbose "Skipping mailbox report." } else {
	#get 365 mailbox report
	$ProgressActivity = "Retrieving 365 mailbox data."
	$ProgressOperation = "Retrieving mailbox list."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation

	$Mailboxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited -Properties UserPrincipalName, DisplayName, RecipientTypeDetails, WhenMailboxCreated, LitigationHoldEnabled, GUID, Identity | Where-Object { $_.DisplayName -notlike "Discovery Search Mailbox" } | Select-Object -Property UserPrincipalName, DisplayName, RecipientTypeDetails, WhenMailboxCreated, LitigationHoldEnabled, GUID, Identity
	If (!($Mailboxes.count -gt 0)) {
		Write-Verbose "No Mailboxes."
		Write-Progress -Activity $ProgressActivity -Completed
		$SkipMailboxReport = $true
	}
}
If ($SkipMailboxReport) { Write-Verbose "Skipping mailbox report." } else {
	#construct report output object
	$ProgressActivity = "Retrieving 365 mailbox permissions. This may take a while."
	$ProgressOperation = "Retrieving mailbox permissions."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
	$MailboxPermissions = $Mailboxes | Get-EXOMailboxPermission | Where-Object { $_.User -ne 'NT AUTHORITY\SELF' }

	$365MailboxReportObject = @()

	$MailboxProgressBarCounter = 0
	$Mailboxes | ForEach-Object {
		$MailboxProgressBarCounter++
		$DisplayName = $_.DisplayName
		$ProgressOperation = "Retrieving mailbox data for $DisplayName."
		$ProgressPercent = ($MailboxProgressBarCounter / $($Mailboxes).count) * 100
		Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete $ProgressPercent

		#these $Mailbox properties require the ExchangeOnlineManagement module
		$userPrincipalName = $_.UserPrincipalName
		$identity = $_.Identity
		$MailboxCreationDateTime = $_.WhenMailboxCreated
		$MailboxLastLogonDateTime = (Get-EXOMailboxStatistics -Identity $_.guid -Properties lastlogontime).lastlogontime
		$MailboxType = $_.RecipientTypeDetails

		#these properties reference the mailbox permissions
		[array]$Delegates = @()
		$Delegates = $MailboxPermissions | Where-Object { $_.Identity -eq $identity } | Select-Object -ExpandProperty 'User'
		if ( $Delegates.count -ge 1 ) { $Delegates = $Delegates.Split(",") } else { $Delegates = "none" }
		$DelegateString = $Delegates -join (",")
		$DelegateCount = $($Delegates | Where-Object { $_ -ne "none" }).count

		#these properties reference the corresponding 365 user
		$MailboxUser = $365UserReportObject | Where-Object -Property Userprincipalname -eq $userPrincipalName
		#check whether this UPN is blocked for sign-in in the 365 users report
		$MailboxEnabled = $MailboxUser | Select-Object -ExpandProperty Sign-In
		#check whether this UPN has licenses assigned in the 365 users report
		if ($MailboxUser.Licenses -eq "none") { $MailboxLicensed = "no" } else { $MailboxLicensed = "yes" }

		#Retrieve lastlogon time and then calculate days since last use
		if ($null -eq $MailboxLastLogonDateTime) {
			$MailboxLastLogonDateTime = "Never Signed In"
			$MailboxInactiveDays = "-1"
		}
		else {
			$MailboxInactiveDays = (New-TimeSpan -Start $MailboxLastLogonDateTime).Days
		}

		#Retrieve whether or not a litigation hold is enabled on the mailbox
		if ($_.LitigationHoldEnabled -eq $true) { $MailboxLitigationHold = "yes" } else { $MailboxLitigationHold = "no" }

		# build result object
		$mailboxHash = $null
		$mailboxHash = @{
			'UserPrincipalName'   = $userPrincipalName
			'DisplayName'         = $DisplayName
			'Sign-In'             = $MailboxEnabled
			#        'Department'          = $MGUser.Department
			#        'Title'               = $MGUser.JobTitle
			#        'PasswordAge'         = $MGUserPasswordAge
			'MailboxType'         = $MailboxType
			'MailboxCreated'      = $MailboxCreationDateTime
			'MailboxLastLogon'    = $MailboxLastLogonDateTime
			'MailboxInactiveDays' = $MailboxInactiveDays
			'Licensed'            = $MailboxLicensed
			'LitigationHold'      = $MailboxLitigationHold
			#        'Roles'               = $MGUserRoles
			#        'Manager'             = $MGUserManager
			'DelegateCount'       = $DelegateCount
			'Delegates'           = $DelegateString
		}
		$mailboxObject = $null
		$mailboxObject = New-Object PSObject -Property $mailboxHash
		$365MailboxReportObject += $mailboxObject
	}
	Write-Progress -Activity $ProgressActivity -Completed

	$ProgressActivity = "Building Excel report."
	$ProgressOperation = "Exporting to Excel."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
	$365MailboxReportObject | Select-Object UserPrincipalName, DisplayName, Sign-In, Synced, Licensed, MailboxType, MailboxCreated, MailboxLastLogon, MailboxInactiveDays, LitigationHold, DelegateCount, Delegates | Sort-Object -Property UserPrincipalName | Export-Excel `
		-Path $XLSreport `
		-WorkSheetname "365 Mailboxes" `
		-ClearSheet `
		-BoldTopRow `
		-Autosize `
		-FreezePane 2 `
		-Autofilter `
		-ConditionalText $(
		New-ConditionalText "blocked" -ConditionalTextColor DarkRed -BackgroundColor LightPink
		New-ConditionalText "Never Signed In" -ConditionalTextColor DarkRed -BackgroundColor LightPink
		New-ConditionalText "Global Administrator" -BackgroundColor Yellow
	)
	Write-Progress -Activity $ProgressActivity -Completed
}

If ($SkipGroupReport) { Write-Verbose "Skipping group report." } else {
	$Step = "365 Group Memberships"
	# get 365 group report
	$ProgressActivity = "Processing $Step."
	$ProgressOperation = "Retrieving data."
	Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation

	$MGGroupList = Get-MgGroup -all
	If (!($MGGroupList.count -gt 0)) {
		Write-Verbose "No Groups."
		Write-Progress -Activity $ProgressActivity -Completed
		$SkipGroupReport = $true
	}
}
If ($SkipGroupReport) { Write-Verbose "Skipping group report." } else {
	$GroupProgressBarCounter = 0
	$365GroupReportObject = ForEach ($MGGroup in $MGGroupList) {
		$GroupProgressBarCounter++
		$DisplayName = $MGGroup.DisplayName
		$ProgressOperation = "Retrieving group membership data for $DisplayName."
		$ProgressPercent = ($GroupProgressBarCounter / $($MGGroupList).count) * 100
		Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete $ProgressPercent
		$GroupOwner = Get-MgGroupOwner -GroupID $MGGroup.Id | Select-Object -ExpandProperty Id
		if ($null -eq $GroupOwner) { $GroupOwnerUPN = "not set" } else {
			try { $GroupOwnerUPN = Get-MgUser -UserID $GroupOwner | Select-Object -ExpandProperty UserPrincipalName } catch { $GroupOwnerUPN = "other" }
		}
		$GroupSynced = $MGGroup.OnPremisesSyncEnabled
		$GroupDescription = $MGGroup.Description
		Get-MgGroupMember -GroupID $MGGroup.id -all | ForEach-Object {
			[pscustomobject]@{
				GroupName   = $MGGroup.DisplayName
				GroupOwner  = $GroupOwnerUPN
				ADSynced    = $GroupSynced
				Description = $GroupDescription
				MemberName  = $_.additionalproperties['displayName']
				MemberUPN   = $_.additionalproperties['userPrincipalName']
				MemberEmail = $_.additionalproperties['mail']
			}
		}
	}
	Write-Progress -Activity $ProgressActivity -Completed

	New-Report -ReportName $Step -ReportData $365GroupReportObject -ReportOutput $XLSreport
}

#get audit log settings
$Step = "Exchange Online Audit Log"
$ProgressActivity = "Processing $Step. This may take a while."
$ProgressOperation = "Retrieving data."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
$EOAuditLogConfig = Get-AdminAuditLogConfig | Select-Object -Property UnifiedAuditLogIngestionEnabled, AdminAuditLogAgeLimit
New-Report -ReportName $Step -ReportData $EOAuditLogConfig -ReportOutput $XLSreport

#get transport rules (check for external banner)
$Step = "Exchange Online Transport Rules"
$ProgressActivity = "Processing $Step. This may take a while."
$ProgressOperation = "Retrieving data."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
[array]$EOTransportRules = Get-TransportRule | Select-Object -Property Name, State, Priority, Description | Sort-Object -Property Priority
New-Report -ReportName $Step -ReportData $EOTransportRules -ReportOutput $XLSreport

#get mail deliverability records
$Step = "Mail DNS Records"
$ProgressActivity = "Processing $Step. This may take a while."
$ProgressOperation = "Retrieving data."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
$DeliverabilityRecords = Get-MgDomain | ForEach-Object {
	$ErrorActionPreference = 'Stop'
	try { $MX = $(Resolve-DnsName $($_.Id) -Type MX | Select-Object -ExpandProperty NameExchange) -Join "," } catch { $MX = "error" }
	try { $SPF = Resolve-DnsName $($_.Id) -Type TXT | Where-Object { $_.Strings -match "v=spf1" } | Select-Object -ExpandProperty Strings } catch { $SPF = "error" }
	try { $DMARC = Resolve-DnsName _dmarc.$($_.Id) -Type TXT | Select-Object -ExpandProperty Strings } catch { $DMARC = "error" }
	try { $DKIM1 = Resolve-DnsName selector1._domainkey.$($_.Id) -Type CNAME | Select-Object -ExpandProperty NameHost } catch { $DKIM1 = "error" }
	try { $DKIM2 = Resolve-DnsName selector2._domainkey.$($_.Id) -Type CNAME | Select-Object -ExpandProperty NameHost } catch { $DKIM2 = "error" }
	try { $DKIMRecords = Get-DkimSigningConfig -Identity $($_.Id) | Select-Object Selector1CNAME, Selector2CNAME } catch { Write-Warning "No DKIM signing config for $($_.Id)." }
	if ($DKIM1 -eq $($DKIMRecords.Selector1CNAME)) { $DKIM1 = "OK" }
	if ($DKIM2 -eq $($DKIMRecords.Selector2CNAME)) { $DKIM2 = "OK" }
	[pscustomobject]@{
		Domain    = $_.Id
		IsDefault = $_.IsDefault
		IsInitial = $_.IsInitial
		MX        = $MX
		SPF       = $SPF
		DMARC     = $DMARC
		DKIM1     = $DKIM1
		DKIM2     = $DKIM2
	}
}
New-Report -ReportName $Step -ReportData $DeliverabilityRecords -ReportOutput $XLSreport

#get Entra devices
$Step = "Entra Devices"
$ProgressActivity = "Processing $Step. This may take a while."
$ProgressOperation = "Retrieving data."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
$EntraDevices = Get-MgDevice -All -Select "displayName,deviceId,operatingsystem,operatingsystemversion,approximatelastsignindatetime" | Select-Object -Property DisplayName, DeviceID, OperatingSystem, OperatingSystemVersion, ApproximateLastSignInDateTime | Sort-Object -Property ApproximateLastSignInDateTime -Descending
New-Report -ReportName $Step -ReportData $EntraDevices -ReportOutput $XLSreport

#Clean up session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Disconnect-MgGraph | Out-Null

Write-Output "Report output is at $XLSreport."
Write-Output "Finished in $($Stopwatch.Elapsed.TotalSeconds) seconds."
$Stopwatch.Stop()