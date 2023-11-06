### Toolbox, by John Holst
#
Function Get-QueryComputers {  ### Get-QueryComputers - Get Domain Servers names 
  Param( $FQueryComputerSearch, $FQueryComputerExcludeList )
    $FQueryComputers = Foreach ($FComputerSearch in $FQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' | where { $FQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $FComputerSearch)} -ErrorAction Continue)};
    $FQueryComputers = $FQueryComputers | Sort Name;
    Return $FQueryComputers;
};
Function Get-LatestReboot { ### Get-LatestReboot - Get Latest Reboot / Restart / Shutdown for logged on server
  Param(
    $FFileName =  "$([Environment]::GetFolderPath("Desktop"))\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)",
    #$FFileName =  "$($env:USERPROFILE)\Desktop\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)",
    $FExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $FLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $FLastXHours  = ( %{If ( $FLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } })
    );
  ## Script
    $FEventLogStartTime = [DateTime]::Now.AddDays(-$($FLastXDays)).AddHours(-$($FLastXHours));
    Show-Title "Get latest Shutdown / Restart / Reboot for Local Server - Events After: $($FEventLogStartTime)";
    $FLatestBootTime = Get-WmiObject win32_operatingsystem | select csname, @{LABEL="LastBootUpTime";EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}};
    $FLatestBootEvents = Get-EventLog -LogName System -After $FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
  # Output
    # $FLatestBootEvents | Select MachineName, TimeGenerated, UserName, Message | fl; $FLatestBootEvents | Select MachineName, TimeGenerated, UserName | ft -Autosize; $FLatestBootTime;
  ## Exports
    If (($FExport -eq "Y") -or ($FExport -eq "YES")) {$FLatestBootupEvents | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName, Message | Export-CSV "$($FFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation;};
  ## Return
    [hashtable]$Return = @{}
    $Return.LatestBootEventsExtended = $FLatestBootEvents | Select MachineName, TimeGenerated, UserName, Message;
    $Return.LatestBootEvents = $FLatestBootEvents | Select MachineName, TimeGenerated, UserName;
    $Return.LatestBootTime = $FLatestBootTime;
    Return $Return
};
Function Get-LatestRebootDomain {
  Param(
    $FCustomerName  = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $FQueryComputerSearch  = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $FQueryComputerExcludeList  = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $FLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $FLastXHours  = ( %{If ( $FLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    #$FExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $FExport = "Yes",
    $FExportExtended = ("Yes" | %{ If($Entry = Read-Host "  Export Standard & Extended(message included) result to file - ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $FFileName =  "$([Environment]::GetFolderPath("Desktop"))\$($FCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)",
    #$FFileName =  "$($env:USERPROFILE)\Desktop\$($FCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)",
    $FJobNamePrefix = "RegQuery_"
    );
  ## Script
    $FEventLogStartTime = [DateTime]::Now.AddDays(-$($FLastXDays)).AddHours(-$($FLastXHours));
    Show-Title "Get latest Shutdown / Restart / Reboot for multiple Domain Servers - Events After: $($FEventLogStartTime)";
    $FQueryComputers = (Get-QueryComputers -FQueryComputerSearch $FQueryComputerSearch -FQueryComputerExcludeList $FQueryComputerExcludeList).name; # Get Values like .Name, .DNSHostName
    # Foreach ($FComputerSearch in $FQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' | where { $FQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $FComputerSearch)} -ErrorAction Continue).name}; $FQueryComputers = $FQueryComputers | Sort;
    Foreach ($FQueryComputer in $FQueryComputers) {
      Write-Host "Querying Server: $($FQueryComputer)";
      $Block01 = {Get-EventLog -LogName System -After $Using:FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)} };
      IF ($FQueryComputer -eq $Env:COMPUTERNAME) {
        $FLocalHostResult = Get-EventLog -LogName System -After $FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $Block01 -computername $FQueryComputer -JobName "$($FJobNamePrefix)$($FQueryComputer)" -ThrottleLimit 16 -AsJob
      };
    };
    Write-Host "  Waiting for jobs to complete... `n";
    DO { $FStatus = ((Get-Job -State Completed).count/(Get-Job -Name "$($FJobNamePrefix)*").count) * 100;
      Write-Progress -Activity "Waiting for $((Get-Job -State Running).count) job(s) to complete..." -Status "$FStatus % completed" -PercentComplete $FStatus; }
    While ((Get-job -Name "$($FJobNamePrefix)*" | Where State -eq Running));
    $FResult = Foreach ($FJob in (Get-Job -Name "$($FJobNamePrefix)*")) {Receive-Job -id $FJob.ID -Keep}; Get-Job -State Completed | Remove-Job;
    $FResult = $FResult + $FLocalHostResult;
  ## Output
    #$FResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName | FT -autosize;
  ## Exports
    If (($FExport -eq "Y") -or ($FExport -eq "YES")) { $FResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName | Export-CSV "$($FFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
    If (($FExportExtended -eq "Y") -or ($FExportExtended -eq "YES")) { $FResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName, Message | Export-CSV "$($FFileName)_Extended.csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.LatestBootEvents = $FResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName;
    Return $Return;
};
Function Show-Title {
  param ( [string]$Title );
    $host.UI.RawUI.WindowTitle = $Title;
};
Function Show-Menu {
  param (
    [string]$Title = "Toolbox"
    );
    Show-Title $Title;
    Clear-Host;
    Write-Host "`n  ================ $Title ================`n";
    Write-Host "  1: Press '1' for Get-LatestReboot for Local Server.";
    Write-Host "  2: Press '2' for Get-LatestReboot for Domain Servers.";
    Write-Host "  3: Press '3' for this option.";
    Write-Host "  I: Press 'I' for Toolbox Information.";
    Write-Host "  Q: Press 'Q' to quit.";
};
Function ToolboxMenu {
  do {
    Show-Menu
    $selection = Read-Host "`n  Please make a selection"
    switch ($selection){
      "1" { "`n`n  You selected: Get-LatestReboot for Local Server`n"
        $Result = Get-LatestReboot;
        $Result.LatestBootEventsExtended | FL; $result.LatestBootEvents | FT -Autosize; $result.LatestBootTime | FT  -Autosize;
        Pause;
        };
      "2" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LatestRebootDomain;
        $Result.LatestBootEvents | FT -Autosize;
        Pause;
        };
      "3" { "`n`n  You selected: option #3`n"
      Sleep 10;
      };
      "3" { "`n`n  You selected: option #3`n"
      Sleep 10;
	  };
    };
    #Pause;
 } until (($selection -eq "q") -or ($selection -eq "0"));
};
ToolboxMenu;
