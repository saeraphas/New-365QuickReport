#upn = UserPrincipalName
#department = Department
#title = JobTitle
#displayname = DisplayName
#
#creationtime = CreatedDateTime
#lastpasswordchangetimestamp = ?????
#lastLogonTime = ?????
#InactiveDays = calculated off LastLogonTime
#SigninBlocked = AccountEnabled? 
#MailboxType = ????? userpurpose
#AssignedLicenses = download license table to variable ahead of loop, userlicenses = get mguserlicenses, foreach licenseguid in licenseguids, append licenseguid match description to license array, convert license array to string
#Roles = 