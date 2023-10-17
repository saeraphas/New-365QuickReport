#new-EntraDeviceReport.ps1
#
$computers = Get-MgDevice -Select "displayName,deviceId,operatingsystem,operatingsystemversion,approximatelastsignindatetime" | Select-Object -Property DisplayName,DeviceID,OperatingSystem,OperatingSystemVersion,ApproximateLastSignInDateTime
$computers | Sort-Object -Property ApproximateLastSignInDateTime -Descending | export-csv -path D4M-AADPCs.csv -nti