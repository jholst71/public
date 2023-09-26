### Get-LatestReboot - Get latest Reboot / Shutdown / Restart for logged on server
## Variables
$LastXDays = 30; $LastXHours = 0; # Set number of Days and/or Hours to verify
#$LastXDays = 7; $LastXHours = 0; # Set number of Days and/or Hours to verify
$FileName =  "$($env:USERPROFILE)\Desktop\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)";
## Script
$EventLogStartTime = [DateTime]::Now.AddDays(-$($LastXDays)).AddHours(-$($LastXHours));
$LatestBootupTime = Get-WmiObject win32_operatingsystem | select csname, @{LABEL="LastBootUpTime";EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}};
$LatestBootupEvents = Get-EventLog -LogName System -After $EventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
# Output
$LatestBootupEvents | fl MachineName, TimeGenerated, UserName, Message; $LatestBootupEvents | ft MachineName, TimeGenerated, UserName; $LatestBootupTime; 
#
## Exports
$LatestBootupEvents | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName, Message | Export-CSV "$($FileName).csv" -Delimiter ';' -NoTypeInformation;
Pause; ## Script END
