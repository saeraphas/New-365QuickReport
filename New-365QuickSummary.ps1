<#
.SYNOPSIS
	Summarizes 365 quick reports. 

.DESCRIPTION
	Imports the specified quick report XLSX and outputs a summary list of common review points.

.EXAMPLE
	.\New-365QuickSummary.ps1

.NOTES
    Author:             Douglas Hammond (douglas@douglashammond.com)
	License: 			This script is distributed under "THE BEER-WARE LICENSE" (Revision 42):
						As long as you retain this notice you can do whatever you want with this stuff.
						If we meet some day, and you think this stuff is worth it, you can buy me a beer in return.
#>

Param (
    [Parameter(ValueFromPipelineByPropertyName)]
    [string] $ReportPath
)

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    $usersdata = Import-Excel $ReportPath -WorkSheetName "365 Users"
    $mailboxdata = Import-Excel $ReportPath -WorkSheetName "365 Mailboxes"

    [array]$EnabledAdminsWithNoMFA = @() #strong typing in case there's exactly 1 result
    $EnabledAdminsWithNoMFA = $usersdata | Where-Object { $_.'Sign-In' -eq "allowed" -and $_.'Roles' -like "*Administrator" -and $_.'MFA_Status' -ne "Enabled" }
    Write-Output "Counted $($EnabledAdminsWithNoMFA.count) admins with MFA not enabled."
    $EnabledAdminsWithNoMFA | Select-Object -Property UserPrincipalName,DisplayName,Roles | Format-Table
    
    [array]$DisabledAdmins = @()
    $DisabledAdmins = $usersdata | Where-Object { $_.'Sign-In' -eq "blocked" -and $_.'Roles' -like "*Administrator" }
    Write-Output "Counted $($DisabledAdmins.count) admins with Sign-In blocked."
    $DisabledAdmins | Select-Object -Property UserPrincipalName,DisplayName,Roles | Format-Table

    [array]$EnabledLicensedUsersWithNoMFA = @()
    $EnabledLicensedUsersWithNoMFA = $usersdata | Where-Object { $_.'Sign-In' -eq "allowed" -and $_.'Licenses' -ne "none" -and $_.'MFA_Status' -ne "Enabled" }
    Write-Output "Counted $($EnabledLicensedUsersWithNoMFA.count) users with license and MFA not enabled."
    $EnabledLicensedUsersWithNoMFA | Select-Object -Property UserPrincipalName,DisplayName | Format-Table

    [array]$DisabledLicensedUsersWithLicenses = @()
    $DisabledLicensedUsersWithLicenses = $usersdata | Where-Object { $_.'Sign-In' -eq "blocked" -and $_.'Licenses' -ne "none" }
    Write-Output "Counted $($DisabledLicensedUsersWithLicenses.count) users with license and Sign-In blocked."
    $DisabledLicensedUsersWithLicenses | Select-Object -Property UserPrincipalName,DisplayName | Format-Table

    [array]$UnlicensedUsers = @()
    $UnlicensedUsers = $usersdata | Where-Object { $_.'Licenses' -eq "none" -and -not ($_.'Roles' -like "*Administrator") }
    Write-Output "Counted $($UnlicensedUsers.count) users with no license."

    [array]$SharedMailboxWithLicense = @()
    $SharedMailboxWithLicense = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'Licensed' -eq "yes" }
    Write-Output "Counted $($SharedMailboxWithLicense.count) shared mailboxes with license."
    $SharedMailboxWithLicense | Select-Object -Property UserPrincipalName,DisplayName | Format-Table

    [array]$SharedMailboxWithNoDelegates = @()
    $SharedMailboxWithNoDelegates = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'Delegates' -eq "none" }
    Write-Output "Counted $($SharedMailboxWithNoDelegates.count) shared mailboxes with no delegates."
    $SharedMailboxWithNoDelegates | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table

    [array]$InactiveSharedMailbox = @()
    $InactiveSharedMailbox = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'MailboxInactiveDays' -ge 30 }
    Write-Output "Counted $($InactiveSharedMailbox.count) shared mailboxes with 30d+ inactivity."
    $InactiveSharedMailbox | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table

    [array]$InactiveLicensedMailbox = @()
    $InactiveLicensedMailbox = $mailboxdata | Where-Object { $_.'MailboxType' -eq "UserMailbox" -and $_.'Licensed' -eq "yes" -and $_.'MailboxInactiveDays' -ge 30 }
    Write-Output "Counted $($InactiveLicensedMailbox.count) licensed mailboxes with 30d+ inactivity."
    $InactiveLicensedMailbox | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table

    Write-Output "Finished."
}