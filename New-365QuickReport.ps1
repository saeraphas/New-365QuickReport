<#
.SYNOPSIS
	This script collects data from Exchange Online and Azure AD and builds a quick reports intended for customer review as part of periodic housekeeping, license audits, true-ups, etc.

.DESCRIPTION
	Cobbled together from various script snippets, examples, tutorials, and howtos into a good-enough script for my purposes. 
	
.EXAMPLE
	.\New-365QuickReport.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>

#I'm eventually going to write a better prereq checker/installer function and use something like this for input.
$PrerequisitesTable = @'
Name,Repository,Scope,Version
Microsoft.Graph,PSGallery,CurrentUser,1.25.0
ExchangeOnlineManagement,PSGallery,CurrentUser,3.0.0
ImportExcel,PSGallery,CurrentUser,7.0.0
'@
$Prerequisites = $PrerequisitesTable | ConvertFrom-Csv
$Prerequisites #| Out-GridView

#install the necessary modules if they aren't installed already
$ProgressActivity = "Checking Prerequisites."
$ProgressOperation = "Checking for module MSOnline."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
If (!(Get-Module -ListAvailable -Name MSOnline)) { Install-Module MSOnline -scope CurrentUser -Force } 

$ProgressOperation = "Checking for module ExchangeOnlineManagement."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
[version] $moduleversion = $(Get-Module -ListAvailable -Name ExchangeOnlineManagement).version
[version] $minimumversion = "3.0.0"
If (!($moduleversion)) { 
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion
} elseif ($moduleversion -lt $minimumversion) {
    Uninstall-Module ExchangeOnlineManagement #this may fail because i haven't checked for admin
    Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -scope CurrentUser -Force -RequiredVersion $minimumversion
}

$ProgressOperation = "Checking for module ImportExcel."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
If (!(Get-Module -ListAvailable -Name ImportExcel)) { Install-Module ImportExcel -scope CurrentUser -Force } 
Write-Progress -Activity $ProgressActivity -Completed

$ProgressActivity = "Loading Prerequisites."
$ProgressOperation = "Loading module MSOnline."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
import-module MSOnline
$ProgressOperation = "Loading module ExchangeOnlineManagement."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
import-module ExchangeOnlineManagement
$ProgressOperation = "Loading module ImportExcel."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
import-module ImportExcel

Write-Progress -Activity $ProgressActivity -Completed

$FriendlyNameArray = @'
AAD_BASIC = Azure Active Directory Basic
AAD_PREMIUM = Azure Active Directory Premium
AAD_PREMIUM_P1 = Azure Active Directory Premium P1
AAD_PREMIUM_P2 = Azure Active Directory Premium P2
ADALLOM_O365 = Office 365 Advanced Security Management
ADALLOM_STANDALONE = Microsoft Cloud App Security
ADALLOM_S_O365 = POWER BI STANDALONE
ADALLOM_S_STANDALONE = Microsoft Cloud App Security
ATA = Azure Advanced Threat Protection for Users
ATP_ENTERPRISE = Exchange Online Advanced Threat Protection
ATP_ENTERPRISE_FACULTY = Exchange Online Advanced Threat Protection
BI_AZURE_P0 = Power BI (free)
BI_AZURE_P1 = Power BI Reporting and Analytics
BI_AZURE_P2 = Power BI Pro
CCIBOTS_PRIVPREV_VIRAL = Dynamics 365 AI for Customer Service Virtual Agents Viral SKU
CRMINSTANCE = Microsoft Dynamics CRM Online Additional Production Instance (Government Pricing)
CRMIUR = CRM for Partners
CRMPLAN1 = Microsoft Dynamics CRM Online Essential (Government Pricing)
CRMPLAN2 = Dynamics CRM Online Plan 2
CRMSTANDARD = CRM Online
CRMSTORAGE = Microsoft Dynamics CRM Online Additional Storage
CRMTESTINSTANCE = CRM Test Instance
DESKLESS = Microsoft StaffHub
DESKLESSPACK = Office 365 (Plan K1)
DESKLESSPACK_GOV = Microsoft Office 365 (Plan K1) for Government
DESKLESSPACK_YAMMER = Office 365 Enterprise K1 with Yammer
DESKLESSWOFFPACK = Office 365 (Plan K2)
DESKLESSWOFFPACK_GOV = Microsoft Office 365 (Plan K2) for Government
DEVELOPERPACK = Office 365 Enterprise E3 Developer
DEVELOPERPACK_E5 = Microsoft 365 E5 Developer(without Windows and Audio Conferencing)
DMENTERPRISE = Microsoft Dynamics Marketing Online Enterprise
DYN365_ENTERPRISE_CUSTOMER_SERVICE = Dynamics 365 for Customer Service Enterprise Edition
DYN365_ENTERPRISE_P1_IW = Dynamics 365 P1 Trial for Information Workers
DYN365_ENTERPRISE_PLAN1 = Dynamics 365 Plan 1 Enterprise Edition
DYN365_ENTERPRISE_SALES = Dynamics 365 for Sales Enterprise Edition
DYN365_ENTERPRISE_SALES_CUSTOMERSERVICE = Dynamics 365 for Sales and Customer Service Enterprise Edition
DYN365_ENTERPRISE_TEAM_MEMBERS = Dynamics 365 for Team Members Enterprise Edition
DYN365_FINANCIALS_BUSINESS_SKU = Dynamics 365 for Financials Business Edition
DYN365_MARKETING_USER = Dynamics 365 for Marketing USL
DYN365_MARKETING_APP = Dynamics 365 Marketing
DYN365_SALES_INSIGHTS = Dynamics 365 AI for Sales
D365_SALES_PRO = Dynamics 365 for Sales Professional
Dynamics_365_for_Operations = Dynamics 365 Unf Ops Plan Ent Edition
ECAL_SERVICES = ECAL
EMS = Enterprise Mobility + Security E3
EMSPREMIUM = Enterprise Mobility + Security E5
ENTERPRISEPACK = Office 365 Enterprise E3
ENTERPRISEPACKLRG = Office 365 Enterprise E3 LRG
ENTERPRISEPACKWITHOUTPROPLUS = Office 365 Enterprise E3 without ProPlus Add-on
ENTERPRISEPACK_B_PILOT = Office 365 (Enterprise Preview)
ENTERPRISEPACK_FACULTY = Office 365 (Plan A3) for Faculty
ENTERPRISEPACK_GOV = Microsoft Office 365 (Plan G3) for Government
ENTERPRISEPACK_STUDENT = Office 365 (Plan A3) for Students
ENTERPRISEPREMIUM = Enterprise E5 (with Audio Conferencing)
ENTERPRISEPREMIUM_NOPSTNCONF = Enterprise E5 (without Audio Conferencing)
ENTERPRISEWITHSCAL = Office 365 Enterprise E4
ENTERPRISEWITHSCAL_FACULTY = Office 365 (Plan A4) for Faculty
ENTERPRISEWITHSCAL_GOV = Microsoft Office 365 (Plan G4) for Government
ENTERPRISEWITHSCAL_STUDENT = Office 365 (Plan A4) for Students
EOP_ENTERPRISE = Exchange Online Protection
EOP_ENTERPRISE_FACULTY = Exchange Online Protection for Faculty
EQUIVIO_ANALYTICS = Office 365 Advanced Compliance
EQUIVIO_ANALYTICS_FACULTY = Office 365 Advanced Compliance for Faculty
ESKLESSWOFFPACK_GOV = Microsoft Office 365 (Plan K2) for Government
EXCHANGEARCHIVE = Exchange Online Archiving
EXCHANGEARCHIVE_ADDON = Exchange Online Archiving for Exchange Online
EXCHANGEDESKLESS = Exchange Online Kiosk
EXCHANGEENTERPRISE = Exchange Online Plan 2
EXCHANGEENTERPRISE_FACULTY = Exch Online Plan 2 for Faculty
EXCHANGEENTERPRISE_GOV = Microsoft Office 365 Exchange Online (Plan 2) only for Government
EXCHANGEESSENTIALS = Exchange Online Essentials
EXCHANGESTANDARD = Office 365 Exchange Online Only
EXCHANGESTANDARD_GOV = Microsoft Office 365 Exchange Online (Plan 1) only for Government
EXCHANGESTANDARD_STUDENT = Exchange Online (Plan 1) for Students
EXCHANGETELCO = Exchange Online POP
EXCHANGE_ANALYTICS = Microsoft MyAnalytics
EXCHANGE_L_STANDARD = Exchange Online (Plan 1)
EXCHANGE_S_ARCHIVE_ADDON_GOV = Exchange Online Archiving
EXCHANGE_S_DESKLESS = Exchange Online Kiosk
EXCHANGE_S_DESKLESS_GOV = Exchange Kiosk
EXCHANGE_S_ENTERPRISE = Exchange Online (Plan 2) Ent
EXCHANGE_S_ENTERPRISE_GOV = Exchange Plan 2G
EXCHANGE_S_ESSENTIALS = Exchange Online Essentials
EXCHANGE_S_FOUNDATION = Exchange Foundation for certain SKUs
EXCHANGE_S_STANDARD = Exchange Online (Plan 2)
EXCHANGE_S_STANDARD_MIDMARKET = Exchange Online (Plan 1)
FLOW_FREE = Microsoft Flow (Free)
FLOW_O365_P2 = Flow for Office 365
FLOW_O365_P3 = Flow for Office 365
FLOW_P1 = Microsoft Flow Plan 1
FLOW_P2 = Microsoft Flow Plan 2
FORMS_PLAN_E3 = Microsoft Forms (Plan E3)
FORMS_PLAN_E5 = Microsoft Forms (Plan E5)
INFOPROTECTION_P2 = Azure Information Protection Premium P2
INTUNE_A = Windows Intune Plan A
INTUNE_A_VL = Intune (Volume License)
INTUNE_O365 = Mobile Device Management for Office 365
INTUNE_STORAGE = Intune Extra Storage
IT_ACADEMY_AD = Microsoft Imagine Academy
LITEPACK = Office 365 (Plan P1)
LITEPACK_P2 = Office 365 Small Business Premium
LOCKBOX = Customer Lockbox
LOCKBOX_ENTERPRISE = Customer Lockbox
MCOCAP = Command Area Phone
MCOEV = Skype for Business Cloud PBX
MCOIMP = Skype for Business Online (Plan 1)
MCOLITE = Lync Online (Plan 1)
MCOMEETADV = PSTN conferencing
MCOPLUSCAL = Skype for Business Plus CAL
MCOPSTN1 = Skype for Business Pstn Domestic Calling
MCOPSTN2 = Skype for Business Pstn Domestic and International Calling
MCOSTANDARD = Skype for Business Online Standalone Plan 2
MCOSTANDARD_GOV = Lync Plan 2G
MCOSTANDARD_MIDMARKET = Lync Online (Plan 1)
MCVOICECONF = Lync Online (Plan 3)
MDM_SALES_COLLABORATION = Microsoft Dynamics Marketing Sales Collaboration
MEE_FACULTY = Minecraft Education Edition Faculty
MEE_STUDENT = Minecraft Education Edition Student
MEETING_ROOM = Meeting Room
MFA_PREMIUM = Azure Multi-Factor Authentication
MICROSOFT_BUSINESS_CENTER = Microsoft Business Center
MICROSOFT_REMOTE_ASSIST = Dynamics 365 Remote Assist
MIDSIZEPACK = Office 365 Midsize Business
MINECRAFT_EDUCATION_EDITION = Minecraft Education Edition Faculty
MS-AZR-0145P = Azure
MS_TEAMS_IW = Microsoft Teams
NBPOSTS = Microsoft Social Engagement Additional 10k Posts (minimum 100 licenses) (Government Pricing)
NBPROFESSIONALFORCRM = Microsoft Social Listening Professional
O365_BUSINESS = Microsoft 365 Apps for business
O365_BUSINESS_ESSENTIALS = Microsoft 365 Business Basic
O365_BUSINESS_PREMIUM = Microsoft 365 Business Standard
OFFICE365_MULTIGEO = Multi-Geo Capabilities in Office 365
OFFICESUBSCRIPTION = Microsoft 365 Apps for enterprise
OFFICESUBSCRIPTION_FACULTY = Office 365 ProPlus for Faculty
OFFICESUBSCRIPTION_GOV = Office ProPlus
OFFICESUBSCRIPTION_STUDENT = Office ProPlus Student Benefit
OFFICE_FORMS_PLAN_2 = Microsoft Forms (Plan 2)
OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ = Office ProPlus
ONEDRIVESTANDARD = OneDrive
PAM_ENTERPRISE = Exchange Primary Active Manager
PLANNERSTANDALONE = Planner Standalone
POWERAPPS_INDIVIDUAL_USER = Microsoft PowerApps and Logic flows
POWERAPPS_O365_P2 = PowerApps
POWERAPPS_O365_P3 = PowerApps for Office 365
POWERAPPS_VIRAL = PowerApps (Free)
POWERFLOW_P1 = Microsoft PowerApps Plan 1
POWERFLOW_P2 = Microsoft PowerApps Plan 2
POWER_BI_ADDON = Office 365 Power BI Addon
POWER_BI_INDIVIDUAL_USE = Power BI Individual User
POWER_BI_INDIVIDUAL_USER = Power BI for Office 365 Individual
POWER_BI_PRO = Power BI Pro
POWER_BI_STANDALONE = Power BI Standalone
POWER_BI_STANDARD = Power-BI Standard
PREMIUM_ADMINDROID = AdminDroid Office 365 Reporter
PROJECTCLIENT = Project Professional
PROJECTESSENTIALS = Project Lite
PROJECTONLINE_PLAN_1 = Project Online (Plan 1)
PROJECTONLINE_PLAN_1_FACULTY = Project Online for Faculty Plan 1
PROJECTONLINE_PLAN_1_STUDENT = Project Online for Students Plan 1
PROJECTONLINE_PLAN_2 = Project Online and PRO
PROJECTONLINE_PLAN_2_FACULTY = Project Online for Faculty Plan 2
PROJECTONLINE_PLAN_2_STUDENT = Project Online for Students Plan 2
PROJECTPREMIUM = Project Online Premium
PROJECTPROFESSIONAL = Project Online Pro
PROJECTWORKMANAGEMENT = Office 365 Planner Preview
PROJECT_CLIENT_SUBSCRIPTION = Project Pro for Office 365
PROJECT_ESSENTIALS = Project Lite
PROJECT_MADEIRA_PREVIEW_IW_SKU = Dynamics 365 for Financials for IWs
PROJECT_ONLINE_PRO = Project Online Plan 3
RIGHTSMANAGEMENT = Azure Rights Management Premium
RIGHTSMANAGEMENT_ADHOC = Windows Azure Rights Management
RIGHTSMANAGEMENT_STANDARD_FACULTY = Azure Rights Management for faculty
RIGHTSMANAGEMENT_STANDARD_STUDENT = Information Rights Management for Students
RMS_S_ENTERPRISE = Azure Active Directory Rights Management
RMS_S_ENTERPRISE_GOV = Windows Azure Active Directory Rights Management
RMS_S_PREMIUM = Azure Information Protection Plan 1
RMS_S_PREMIUM2 = Azure Information Protection Premium P2
SCHOOL_DATA_SYNC_P1 = School Data Sync (Plan 1)
SHAREPOINTDESKLESS = SharePoint Online Kiosk
SHAREPOINTDESKLESS_GOV = SharePoint Online Kiosk
SHAREPOINTENTERPRISE = SharePoint Online (Plan 2)
SHAREPOINTENTERPRISE_EDU = SharePoint Plan 2 for EDU
SHAREPOINTENTERPRISE_GOV = SharePoint Plan 2G
SHAREPOINTENTERPRISE_MIDMARKET = SharePoint Online (Plan 1)
SHAREPOINTLITE = SharePoint Online (Plan 1)
SHAREPOINTPARTNER = SharePoint Online Partner Access
SHAREPOINTSTANDARD = SharePoint Online Plan 1
SHAREPOINTSTANDARD_EDU = SharePoint Plan 1 for EDU
SHAREPOINTSTORAGE = SharePoint Online Storage
SHAREPOINTWAC = Office Online
SHAREPOINTWAC_EDU = Office Online for Education
SHAREPOINTWAC_GOV = Office Online for Government
SHAREPOINT_PROJECT = SharePoint Online (Plan 2) Project
SHAREPOINT_PROJECT_EDU = Project Online Service for Education
SMB_APPS = Business Apps (free)
SMB_BUSINESS = Office 365 Business
SMB_BUSINESS_ESSENTIALS = Office 365 Business Essentials
SMB_BUSINESS_PREMIUM = Office 365 Business Premium
SPZA IW = Microsoft PowerApps Plan 2 Trial
SPB = Microsoft 365 Business
SPE_E3 = Secure Productive Enterprise E3
SQL_IS_SSIM = Power BI Information Services
STANDARDPACK = Office 365 (Plan E1)
STANDARDPACK_FACULTY = Office 365 (Plan A1) for Faculty
STANDARDPACK_GOV = Microsoft Office 365 (Plan G1) for Government
STANDARDPACK_STUDENT = Office 365 (Plan A1) for Students
STANDARDWOFFPACK = Office 365 (Plan E2)
STANDARDWOFFPACKPACK_FACULTY = Office 365 (Plan A2) for Faculty
STANDARDWOFFPACKPACK_STUDENT = Office 365 (Plan A2) for Students
STANDARDWOFFPACK_FACULTY = Office 365 Education E1 for Faculty
STANDARDWOFFPACK_GOV = Microsoft Office 365 (Plan G2) for Government
STANDARDWOFFPACK_IW_FACULTY = Office 365 Education for Faculty
STANDARDWOFFPACK_IW_STUDENT = Office 365 Education for Students
STANDARDWOFFPACK_STUDENT = Microsoft Office 365 (Plan A2) for Students
STANDARD_B_PILOT = Office 365 (Small Business Preview)
STREAM = Microsoft Stream
STREAM_O365_E3 = Microsoft Stream for O365 E3 SKU
STREAM_O365_E5 = Microsoft Stream for O365 E5 SKU
SWAY = Sway
TEAMS1 = Microsoft Teams
TEAMS_COMMERCIAL_TRIAL = Microsoft Teams Commercial Cloud Trial
THREAT_INTELLIGENCE = Office 365 Threat Intelligence
VIDEO_INTEROP = Skype Meeting Video Interop for Skype for Business
VISIOCLIENT = Visio Online Plan 2
VISIOONLINE_PLAN1 = Visio Online Plan 1
VISIO_CLIENT_SUBSCRIPTION = Visio Pro for Office 365
WACONEDRIVEENTERPRISE = OneDrive for Business (Plan 2)
WACONEDRIVESTANDARD = OneDrive for Business with Office Online
WACSHAREPOINTSTD = Office Online STD
WHITEBOARD_PLAN3 = White Board (Plan 3)
WIN_DEF_ATP = Windows Defender Advanced Threat Protection
WIN10_PRO_ENT_SUB = Windows 10 Enterprise E3
WIN10_VDA_E3 = Windows E3
WIN10_VDA_E5 = Windows E5
WINDOWS_STORE = Windows Store
YAMMER_EDU = Yammer for Academic
YAMMER_ENTERPRISE = Yammer for the Starship Enterprise
YAMMER_ENTERPRISE_STANDALONE = Yammer Enterprise
YAMMER_MIDSIZE = Yammer
'@
$FriendlyNameHash = $FriendlyNameArray | ConvertFrom-StringData

#connect to microsoft services
$ProgressActivity = "Connecting to Microsoft services. You will be prompted twice."
$ProgressOperation = "1 of 2 - Connecting to Exchange Online."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
try { Connect-ExchangeOnline } catch { write-error "Error connecting to Exchange Online. Exiting."; exit }
$ProgressOperation = "2 of 2 - Connecting to MSOL Service."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
try { Connect-MsolService } catch { write-error "Not connected to MSOL service. Exiting."; exit } 
Write-Progress -Activity $ProgressActivity -Completed

#define paths
$datestring = ((get-date).tostring("yyyy-MM-dd"))
$tenant = (Get-AcceptedDomain | Where-Object { $_.Default }).name
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$tenantpath = "$DesktopPath\365QuickReport\$tenant"
$reportspath = "$tenantpath\reports"
$XLSreport = "$reportspath\$tenant-report-$datestring.xlsx"

$ProgressActivity = "Gathering Exchange Online mailbox data."
$ProgressOperation = "Listing Mailboxes."
Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation
#construct report output object
$ResultObject = @()
$MBUserCount = 0
$Mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.DisplayName -notlike "Discovery Search Mailbox" }
$MBUserTotal = $($Mailboxes).count

$Mailboxes | ForEach-Object {
    $MBUserCount++
    $ProgressOperation = "Gathering mailbox data for $DisplayName"
    $ProgressPercent = ($MBUserCount / $MBUserTotal) * 100
    Write-Progress -Activity $ProgressActivity -CurrentOperation $ProgressOperation -PercentComplete $ProgressPercent
    $upn = $_.UserPrincipalName
    $CreationTime = $_.WhenCreated
    $LastLogonTime = (Get-MailboxStatistics -Identity $upn).lastlogontime
    $DisplayName = $_.DisplayName
    $MBType = $_.RecipientTypeDetails
    $RolesAssigned = ""
    
    #Retrieve lastlogon time and then calculate Inactive days
    if ($LastLogonTime -eq $null) {
        $LastLogonTime = "Never Logged In"
        $InactiveDaysOfUser = "-"
    }
    else {
        $InactiveDaysOfUser = (New-TimeSpan -Start $LastLogonTime).Days
    }

    #Get licenses assigned to mailboxes
    $User = (Get-MsolUser -UserPrincipalName $upn)
    $Licenses = $User.Licenses.AccountSkuId
    $Department = $User.Department
    $SignInBlocked = $null
    if ($User.BlockCredential -eq "True") {
        $SignInBlocked = "sign-in blocked"
    }
    else {
        $SignInBlocked = "sign-in allowed"
    }
    $AssignedLicense = ""
    $Count = 0

    #Convert license plan to friendly name
    foreach ($License in $Licenses) {
        $Count++
        $LicenseItem = $License -Split ":" | Select-Object -Last 1
        $EasyName = $FriendlyNameHash[$LicenseItem]
        if (!($EasyName)) { $NamePrint = $LicenseItem } else	{ $NamePrint = $EasyName }
        $AssignedLicense = $AssignedLicense + $NamePrint
        if ($count -lt $licenses.count) { $AssignedLicense = $AssignedLicense + "," }
    }
    if ($Licenses.count -eq 0) { $AssignedLicense = "No License Assigned" }

    #Get roles assigned to user
    $Roles = (Get-MsolUserRole -UserPrincipalName $upn).Name
    if ($Roles.count -eq 0) { 
        $RolesAssigned = "No roles"
    }
    else {
        foreach ($Role in $Roles) {
            $RolesAssigned = $RolesAssigned + $Role
            if ($Roles.indexof($role) -lt (($Roles.count) - 1)) { $RolesAssigned = $RolesAssigned + "," }
        }
    }

    #Add result to output object
	
    $userHash = $null
    $userHash = @{
        'Department'                  = $Department
        'UserPrincipalName'           = $upn
        'DisplayName'                 = $DisplayName
        'LastLogonTime'               = $LastLogonTime
        'CreationTime'                = $CreationTime
        'LastPasswordChangeTimeStamp' = $User.LastPasswordChangeTimeStamp
        'InactiveDays'                = $InactiveDaysOfUser
        "SignInBlocked"               = $SignInBlocked
        'MailboxType'                 = $MBType
        'AssignedLicenses'            = $AssignedLicense
        'Roles'                       = $RolesAssigned
    }
    $userObject = $null
    $userObject = New-Object PSObject -Property $userHash
    $ResultObject += $userObject
	
}
Write-Progress -Activity $ProgressActivity -Completed

$ProgressActivity = "Building Excel report."
$ProgressOperation = "Exporting to Excel."
$ResultObject | Select-Object Department, UserPrincipalName, DisplayName, LastLogonTime, CreationTime, LastPasswordChangeTimeStamp, InactiveDays, SignInBlocked, MailboxType, AssignedLicenses, Roles | Sort-Object -Property Department, MailboxType, UserPrincipalName | Export-Excel `
    -Path $XLSreport `
    -WorkSheetname "365 Quick Report" `
    -ClearSheet `
    -BoldTopRow `
    -Autosize `
    -FreezePane 2 `
    -Autofilter `
    -ConditionalText $(
    New-ConditionalText "sign-in blocked" -ConditionalTextColor DarkRed -BackgroundColor LightPink 
    New-ConditionalText "Company Administrator" -BackgroundColor Yellow
)`
    -Show

Write-Progress -Activity $ProgressActivity -Completed

#Clean up session
Disconnect-ExchangeOnline
[Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()
