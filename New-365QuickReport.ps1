<#
.SYNOPSIS
	This script collects data from Exchange Online and Microsoft Graph and builds a quick report intended for periodic housekeeping, license audits, true-ups, etc.

.DESCRIPTION
	Lists all mailboxes in a Microsoft 365 tenant, along with sign-in, license, and role information.  
	
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
    [switch] $SkipSKUConversion
)

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
            try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion } catch { Write-Error "An error occurred installing $moduleName." }
        }
        elseif ([version]$installedversion -lt [version]$minimumversion) {
            try { Uninstall-Module $moduleName -AllVersions } catch { Write-Error "An error occurred removing $moduleName. You may need to manually remove old versions using admin privileges." }
            try { Install-Module $moduleName -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion } catch { Write-Error "An error occurred installing $moduleName." }
        }
    }
}

$PrerequisiteModulesTable = @'
Name,MinimumVersion
Microsoft.Graph,1.25.0
ExchangeOnlineManagement,3.0.0
ImportExcel,7.0.0
'@
CheckPrerequisites($PrerequisiteModulesTable)

$PSNewLine = [System.Environment]::Newline

#download SKU and product name reference information from Microsoft Learn
#https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$MicrosoftDocumentationURI = 'https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference'
$MicrosoftDocumentationCSVURI = ((Invoke-WebRequest -UseBasicParsing -Uri $MicrosoftDocumentationURI).Links | Where-Object { $_.href -like 'http*' } | Where-Object { $_.href -like '*.csv' }).href
$MicrosoftDocumentationDownloadError = "An error occurred while downloading the SKU and Product Name information." + $PSNewLine + "SKU names will not be converted to Product Names."
try { $MicrosoftProducts = Invoke-RestMethod -Uri $MicrosoftDocumentationCSVURI -Method Get | ConvertFrom-CSV | Select-Object String_ID, Product_Display_Name -Unique } catch { Write-Warning $MicrosoftDocumentationDownloadError; $SkipSKUConversion = $true }

#connect to microsoft services
$ProgressActivity = "Connecting to Microsoft services. You will be prompted multiple times."
$ProgressOperation = "1 of 2 - Connecting to Exchange Online."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete 0
try { Connect-ExchangeOnline -ShowBanner:$false | Out-Null } catch { write-error "Error connecting to Exchange Online. Exiting."; exit }

$ProgressOperation = "2 of 2 - Connecting to Microsoft Graph."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete 50
try { Connect-MgGraph -Scopes "User.Read.All,RoleManagement.Read.Directory" | Out-Null } catch { write-error "Not connected to MS Graph. Exiting."; exit } 
Write-Progress -Activity $ProgressActivity -Completed

#define variables for file system paths
$DateString = ((get-date).tostring("yyyy-MM-dd"))
$TenantString = (Get-AcceptedDomain | Where-Object { $_.Default }).name
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$TenantPath = "$DesktopPath\365QuickReport\$TenantString"
$ReportPath = "$TenantPath\Reports"
$XLSreport = "$ReportPath\$TenantString-report-$DateString.xlsx"
#construct report output object
$ResultObject = @()

$ProgressActivity = "Gathering account data."
$ProgressOperation = "Listing Mailboxes."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation

#Get all role definitions from Graph API (for display names) #Thanks, Troy
$roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition

$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.DisplayName -notlike "Discovery Search Mailbox" }
$MailboxProgressBarCounter = 0

$Mailboxes | ForEach-Object {
    $MailboxProgressBarCounter++
    $DisplayName = $_.DisplayName
    $ProgressOperation = "Gathering mailbox data for $DisplayName"
    $ProgressPercent = ($MailboxProgressBarCounter / $($Mailboxes).count) * 100
    Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete $ProgressPercent
    
    #these $Mailbox properties require the ExchangeOnlineManagement module
    $userPrincipalName = $_.UserPrincipalName
    $MailboxCreationDateTime = $_.WhenCreated 
    $MailboxLastLogonDateTime = (Get-MailboxStatistics -Identity $userPrincipalName).lastlogontime
    $MailboxType = $_.RecipientTypeDetails
  
    #Retrieve lastlogon time and then calculate days since last use
    if ($null -eq $MailboxLastLogonDateTime) {
        $MailboxLastLogonDateTime = "Never Signed In"
        $MailboxInactiveDays = "-1"
    }
    else {
        $MailboxInactiveDays = (New-TimeSpan -Start $MailboxLastLogonDateTime).Days
    }

    # these $MGUser properties require the (dreadful,unwieldy) Microsoft Graph module (give me my hair back)
    $MGUser = Get-MGUser -UserID $UserPrincipalName -Property ID, UserPrincipalName, AccountEnabled, DisplayName, Department, JobTitle, Mail, CreatedDateTime, LastPasswordChangeDateTime | Select-Object ID, UserPrincipalName, AccountEnabled, DisplayName, Department, JobTitle, Mail, CreatedDateTime, LastPasswordChangeDateTime

    $MGUserEnabled = $null
    if ($MGUser.AccountEnabled -eq $true) { $MGUserEnabled = "allowed" } else { $MGUserEnabled = "blocked" }

    $MGUserPasswordAge = (New-TimeSpan -Start $MGUser.LastPasswordChangeDateTime).Days 

    $MGUserLicenses = $(get-mguserlicensedetail -userid $($MGUser).id).SkuPartNumber 
    #convert SKUs to Product Names unless bypassed or downloading the CSV from documentation failed earlier
    IF ($SkipSKUConversion) { $MGUserLicenseProductNames = $MGUserLicenses -join "," } else {
        $MGUserLicenseProductNameArray = @()
        if ($MGUserLicenses.count -eq 0) { $MGUserLicenseProductNames = "none" } else {
            foreach ($License in $MGUserLicenses) {
                $ProductName = $($MicrosoftProducts | Where-Object { $_.String_ID -eq $License }).Product_Display_Name
                if (!($ProductName)) { $MGUserLicenseProductNameArray += $License } else { $MGUserLicenseProductNameArray += $ProductName }
            }
            $MGUserLicenseProductNames = $MGUserLicenseProductNameArray -join ","
        }
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
    
    # build result object
    $userHash = $null
    $userHash = @{
        'UserPrincipalName'   = $userPrincipalName
        'DisplayName'         = $MGUser.DisplayName
        'Sign-In'             = $MGUserEnabled
        'Department'          = $MGUser.Department
        'Title'               = $MGUser.JobTitle
        'PasswordAge'         = $MGUserPasswordAge
        'MailboxType'         = $MailboxType
        'MailboxCreated'      = $MailboxCreationDateTime
        'MailboxLastLogon'    = $MailboxLastLogonDateTime
        'MailboxInactiveDays' = $MailboxInactiveDays
        'Licenses'            = $MGUserLicenseProductNames
        'Roles'               = $MGUserRoles
        'Manager'             = $MGUserManager
    }
    $userObject = $null
    $userObject = New-Object PSObject -Property $userHash
    $ResultObject += $userObject
	
}
Write-Progress -Activity $ProgressActivity -Completed

$ProgressActivity = "Building Excel report."
$ProgressOperation = "Exporting to Excel."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation

$ResultObject | Select-Object UserPrincipalName, DisplayName, Sign-In, Department, Title, PasswordAge, MailboxType, MailboxCreated, MailboxLastLogon, MailboxInactiveDays, Licenses, Roles, Manager | Sort-Object -Property UserPrincipalName | Export-Excel `
    -Path $XLSreport `
    -WorkSheetname "365 Quick Report" `
    -ClearSheet `
    -BoldTopRow `
    -Autosize `
    -FreezePane 2 `
    -Autofilter `
    -ConditionalText $(
    New-ConditionalText "blocked" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
    New-ConditionalText "Never Signed In" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
    New-ConditionalText "Global Administrator" -BackgroundColor Yellow
)`
    -Show

Write-Progress -Activity $ProgressActivity -Completed

#Clean up session
Disconnect-ExchangeOnline -Confirm:$false | Out-Null
Disconnect-MgGraph | Out-Null