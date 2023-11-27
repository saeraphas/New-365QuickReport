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
    [string] $ReportPath,
    [switch] $ShowTables
)

$reportExists = Test-Path $ReportPath
if (!($reportExists)) { Write-Warning "Specified report file $ReportFile does not exist or could not be read. Exiting."; exit } else {

    $usersdata = Import-Excel $ReportPath -WorkSheetName "365 Users"
    $mailboxdata = Import-Excel $ReportPath -WorkSheetName "365 Mailboxes"

    [array]$EnabledAdminsWithNoMFA = @() #strong typing in case there's exactly 1 result
    $EnabledAdminsWithNoMFA = $usersdata | Where-Object { $_.'Sign-In' -eq "allowed" -and $_.'Roles' -like "*Administrator" -and $_.'MFA_Status' -ne "Enabled" }
    Write-Output "Counted $($EnabledAdminsWithNoMFA.count) admins with MFA not enabled."
    If ($ShowTables) { $EnabledAdminsWithNoMFA | Select-Object -Property UserPrincipalName,DisplayName,Roles | Format-Table }
    
    [array]$DisabledAdmins = @()
    $DisabledAdmins = $usersdata | Where-Object { $_.'Sign-In' -eq "blocked" -and $_.'Roles' -like "*Administrator" }
    Write-Output "Counted $($DisabledAdmins.count) admins with Sign-In blocked."
    If ($ShowTables) { $DisabledAdmins | Select-Object -Property UserPrincipalName,DisplayName,Roles | Format-Table }

    [array]$SyncedUsers = @()
    $SyncedUsers = $usersdata | Where-Object { $_.'Synced' -eq "true" }
    $SyncedUsersPercent = $SyncedUsers.Count / ($usersdata).count * 100
    $SyncedUsersPercent = [math]::Round($SyncedUsersPercent)
    Write-Output "Counted $($SyncedUsers.count)/$(($usersdata).count) ($SyncedUsersPercent%) users synced with an AD domain."
    If ($ShowTables) { $SyncedUsers | Select-Object -Property UserPrincipalName,DisplayName,Synced | Format-Table }

    [array]$EnabledLicensedUsersWithNoMFA = @()
    $EnabledLicensedUsers = $usersdata | Where-Object { $_.'Sign-In' -eq "allowed" -and $_.'Licenses' -ne "none" }
    $EnabledLicensedUsersWithNoMFA = $EnabledLicensedUsers | Where-Object { $_.'MFA_Status' -ne "Enabled" }
    $EnabledLicensedUsersWithNoMFAPercent = $EnabledLicensedUsersWithNoMFA.Count / ($EnabledLicensedUsers).count * 100
    $EnabledLicensedUsersWithNoMFAPercent = [math]::Round($EnabledLicensedUsersWithNoMFAPercent)
    Write-Output "Counted $($EnabledLicensedUsersWithNoMFA.count)/$($EnabledLicensedUsers.count) ($EnabledLicensedUsersWithNoMFAPercent%) users with license and MFA not enabled."
    If ($ShowTables) { $EnabledLicensedUsersWithNoMFA | Select-Object -Property UserPrincipalName,DisplayName | Format-Table }

    [array]$DisabledLicensedUsersWithLicenses = @()
    $DisabledLicensedUsersWithLicenses = $usersdata | Where-Object { $_.'Sign-In' -eq "blocked" -and $_.'Licenses' -ne "none" }
    Write-Output "Counted $($DisabledLicensedUsersWithLicenses.count) users with license and Sign-In blocked."
    If ($ShowTables) { $DisabledLicensedUsersWithLicenses | Select-Object -Property UserPrincipalName,DisplayName | Format-Table }

    [array]$UnlicensedUsers = @()
    $UnlicensedUsers = $usersdata | Where-Object { $_.'Licenses' -eq "none" -and -not ($_.'Roles' -like "*Administrator") }
    Write-Output "Counted $($UnlicensedUsers.count) users with no license."

    [array]$SharedMailboxWithLicense = @()
    $SharedMailboxWithLicense = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'Licensed' -eq "yes" }
    Write-Output "Counted $($SharedMailboxWithLicense.count) shared mailboxes with license."
    If ($ShowTables) { $SharedMailboxWithLicense | Select-Object -Property UserPrincipalName,DisplayName | Format-Table }

    [array]$SharedMailboxWithNoDelegates = @()
    $SharedMailboxWithNoDelegates = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'Delegates' -eq "none" }
    Write-Output "Counted $($SharedMailboxWithNoDelegates.count) shared mailboxes with no delegates."
    If ($ShowTables) { $SharedMailboxWithNoDelegates | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table }

    [array]$InactiveSharedMailbox = @()
    $InactiveSharedMailbox = $mailboxdata | Where-Object { $_.'MailboxType' -eq "SharedMailbox" -and $_.'MailboxInactiveDays' -ge 30 }
    Write-Output "Counted $($InactiveSharedMailbox.count) shared mailboxes with 30d+ inactivity."
    If ($ShowTables) { $InactiveSharedMailbox | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table }

    [array]$InactiveLicensedMailbox = @()
    $InactiveLicensedMailbox = $mailboxdata | Where-Object { $_.'MailboxType' -eq "UserMailbox" -and $_.'Licensed' -eq "yes" -and $_.'MailboxInactiveDays' -ge 30 }
    Write-Output "Counted $($InactiveLicensedMailbox.count) licensed mailboxes with 30d+ inactivity."
    If ($ShowTables) { $InactiveLicensedMailbox | Select-Object -Property UserPrincipalName,DisplayName,MailboxLastLogon | Format-Table }

    [array]$MailboxWithLitigationHold = @()
    $MailboxWithLitigationHold = $mailboxdata | Where-Object { $_.'LitigationHold' -eq "yes" }
    $MailboxWithLitigationHoldPercent = $MailboxWithLitigationHold.Count / ($mailboxdata).count * 100
    $MailboxWithLitigationHoldPercent = [math]::Round($MailboxWithLitigationHoldPercent)
    Write-Output "Counted $($MailboxWithLitigationHold.count)/$(($mailboxdata).count) ($MailboxWithLitigationHoldPercent%) mailboxes with Litigation Holds applied."
    If ($ShowTables) { $MailboxWithLitigationHold | Select-Object -Property UserPrincipalName,DisplayName,LitigationHold | Format-Table }

    Write-Output "Finished."
}