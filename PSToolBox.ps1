### Toolbox, by John Holst
#
Function Get-QueryComputers {  ### Get-QueryComputers - Get Domain Servers names 
  Param( $fQueryComputerSearch, $fQueryComputerExcludeList )
  ## Script
    $fQueryComputers = Foreach ($fComputerSearch in $fQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' | where { $fQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $fComputerSearch)} -ErrorAction Continue)};
    $fQueryComputers = $fQueryComputers | Sort Name;
    Return $fQueryComputers;
};
Function Get-LatestReboot { ### Get-LatestReboot - Get Latest Reboot / Restart / Shutdown for logged on server
  Param(
    $fFileName =  "$([Environment]::GetFolderPath("Desktop"))\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)",
    #$fFileName =  "$($env:USERPROFILE)\Desktop\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)",
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours  = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } })
    );
  ## Script
    $fEventLogStartTime = [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
    Show-Title "Get latest Shutdown / Restart / Reboot for Local Server - Events After: $($fEventLogStartTime)";
    $fLatestBootTime = Get-WmiObject win32_operatingsystem | select csname, @{LABEL="LastBootUpTime";EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}};
    $fLatestBootEvents = Get-EventLog -LogName System -After $fEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
  ## Output
    # $fLatestBootEvents | Select MachineName, TimeGenerated, UserName, Message | fl; $fLatestBootEvents | Select MachineName, TimeGenerated, UserName | ft -Autosize; $fLatestBootTime;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) {$fLatestBootupEvents | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName, Message | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation;};
  ## Return
    [hashtable]$Return = @{}
    $Return.LatestBootEventsExtended = $fLatestBootEvents | Select MachineName, TimeGenerated, UserName, Message;
    $Return.LatestBootEvents = $fLatestBootEvents | Select MachineName, TimeGenerated, UserName;
    $Return.LatestBootTime = $fLatestBootTime;
    Return $Return
};
Function Get-LatestRebootDomain { ### Get-LatestReboot - Get Latest Reboot / Restart / Shutdown for multiple Domain servers
  Param(
    $fCustomerName  = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fQueryComputerSearch  = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList  = ("*" | %{ If($Entry = @(((Read-Host "  Enter ExcludeList ServerName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours  = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fExportExtended = ("Yes" | %{ If($Entry = Read-Host "  Export Standard & Extended(message included) result to file - ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName =  "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)",
    #$fFileName =  "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)",
    $fJobNamePrefix = "RegQuery_"
    );
  ## Script
    $fEventLogStartTime = [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
    Show-Title "Get latest Shutdown / Restart / Reboot for multiple Domain Servers - Events After: $($fEventLogStartTime)";
    $fQueryComputers = (Get-QueryComputers -FQueryComputerSearch $fQueryComputerSearch -FQueryComputerExcludeList $fQueryComputerExcludeList).name; # Get Values like .Name, .DNSHostName
    # Foreach ($fComputerSearch in $fQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' | where { $fQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $fComputerSearch)} -ErrorAction Continue).name}; $fQueryComputers = $fQueryComputers | Sort;
    Foreach ($fQueryComputer in $fQueryComputers) {
      Write-Host "Querying Server: $($fQueryComputer)";
      $Block01 = {Get-EventLog -LogName System -After $Using:FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)} };
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Get-EventLog -LogName System -After $fEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
        } ELSE {
        $JobResult = Invoke-Command -scriptblock $Block01 -computername $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
        };
      };
    Write-Host "  Waiting for jobs to complete... `n";
    DO { $fStatus = ((Get-Job -State Completed).count/(Get-Job -Name "$($fJobNamePrefix)*").count) * 100;
      Write-Progress -Activity "Waiting for $((Get-Job -State Running).count) job(s) to complete..." -Status "$fStatus % completed" -PercentComplete $fStatus; }
    While ((Get-job -Name "$($fJobNamePrefix)*" | Where State -eq Running));
    $fResult = Foreach ($fJob in (Get-Job -Name "$($fJobNamePrefix)*")) {Receive-Job -id $fJob.ID -Keep}; Get-Job -State Completed | Remove-Job;
    $fResult = $fResult + $fLocalHostResult;
  ## Output
    #$fResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
    If (($fExportExtended -eq "Y") -or ($fExportExtended -eq "YES")) { $fResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName, Message | Export-CSV "$($fFileName)_Extended.csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.LatestBootEvents = $fResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName;
    Return $Return;
};
Function Get-HotFixInstallDatesDomain { ### Get-HotFixInstallDates for multiple Domain servers
  Param(
    $fCustomerName  = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fQueryComputerSearch  = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList  = ("*" | %{ If($Entry = @(((Read-Host "  Enter ExcludeList ServerName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fHotfixInstallDates = ("3" | %{ If($Entry = Read-Host "  Enter number of Hotfix-install dates per Computer (Default: $_ Days)"){$Entry} Else {$_} }),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fFileName =  "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-HotFixInstallDates_$(get-date -f yyyy-MM-dd_HH.mm)",
    #$fFileName =  "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-HotFixInstallDates_$(get-date -f yyyy-MM-dd_HH.mm)",
    $fJobNamePrefix = "RegQuery_"
    );
  ## Script
    Show-Title "Get latest $($fHotfixInstallDates) HotFix Install Dates multiple Domain Servers";
    $fQueryComputers = (Get-QueryComputers -FQueryComputerSearch $fQueryComputerSearch -FQueryComputerExcludeList $fQueryComputerExcludeList); # Get Values like .Name, .DNSHostName
    $fResult = @(); $fResult = Foreach ($fQueryComputer in $fQueryComputers) {
      Write-Host "  Querying Server: $($fQueryComputer.Name)";
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fInstalledHotfixes = Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $HotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$((Get-ComputerInfo).WindowsProductName)";
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$((Get-NetIPAddress -AddressFamily IPv4 | ? {$_.IPAddress -notlike "127.0.0.1" }).IPAddress)";
        $fInstalledHotfixes; 
      } Else {
        try {
          $fInstalledHotfixes = Get-Hotfix -ComputerName $QueryComputer.Name | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $HotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
          $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$($QueryComputer.OperatingSystem)";
          $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$($QueryComputer.IPv4Address)";
          $fInstalledHotfixes; 
        } catch {
          Write-Host "      An error occurred within the Get-Hotfix command:"
          Write-Host "      $($_.ScriptStackTrace)"
          Write-Host "    Trying with an Invoked Get-Hotfix command: "
          try {
            $fInstalledHotfixes = Invoke-Command -scriptblock { Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $USING:HotfixInstallDates  | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;  } -computername $QueryComputer
            $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$($QueryComputer.OperatingSystem)";
            $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$($QueryComputer.IPv4Address)";
            $fInstalledHotfixes; 
          } catch {
            Write-Host "      An error occurred within the Invoked Get-Hotfix command:"
            Write-Host "      $($_.ScriptStackTrace)`n"
            $fInstalledHotfixes = [pscustomobject]@{
              "PSComputerName" = "$($QueryComputer.Name)"
              "Description" = ""
              "HotFixID" = ""
              "InstalledBy" = ""
              "InstalledOn" = ""
              "IPv4Address" = "$($QueryComputer.IPv4Address)"
              "OperatingSystem" = "$($QueryComputer.OperatingSystem)"};
            $fInstalledHotfixes; 
    }}}};
  ## Output
    #$fResult | sort MachineName, TimeGenerated | Select MachineName, TimeGenerated, UserName | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort PSComputerName | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn, OperatingSystem, IPv4Address | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.HotFixInstallDates = $fResult | sort PSComputerName | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn, OperatingSystem, IPv4Address;
    Return $Return;
};
Function StartSCOMMaintenanceMode {
  param( 
    $fDuration  = ("30" | %{ If($Entry = Read-Host "  Enter MaintenanceMode Duration ( Default: $_ )"){$Entry} Else {$_} }),
    $fComments = "SCOM MaintenanceMode started for $($SCOMMaintenanceModeDuration) minutes from $($env:Computername) by $($Env:USERNAME) at $(Get-Date)"
  );
  ## Script
    Import-Module "C:\Program Files\Microsoft Monitoring Agent\Agent\MaintenanceMode.dll";
    try { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments -Force Y;
      } catch { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments;}
    Write-Host "Request: Start SCOM Maintenance Mode for $($fDuration) minutes";
};
Function Show-Title {
  param ( [string]$Title );
    $host.UI.RawUI.WindowTitle = $Title;
};
Function Show-Menu {
  param (
    [string]$Title = "Progressive Toolbox"
  );
  Show-Title $Title;
  Clear-Host;
  Write-Host "`n  ================ $Title ================`n";
  #Write-Host "  1: Press '1' for Start SCOM MaintenanceMode for Local Server.";
  #Write-Host "  2: Press '2' for Start SCOM MaintenanceMode for Local Server (Script).";
  Write-Host "  5: Press '5' for Get-LatestReboot for Local Server.";
  Write-Host "  6: Press '6' for Get-LatestReboot for Domain Servers.";
  #Write-Host "  7: Press '7' for Get-HotFixInstallDates for Local Server.";
  Write-Host "  8: Press '8' for Get-HotFixInstallDates for Domain Servers.";
  Write-Host "  9: Press '9' for this option.";
  Write-Host "  I: Press 'I' for Toolbox Information.";
  Write-Host "  Q: Press 'Q' to quit.";
};
Function ToolboxMenu {
  do {
    Show-Menu
    $selection = Read-Host "`n  Please make a selection"
    switch ($selection){
      "1" { "`n`n  You selected: Start SCOM MaintenanceMode for Local Server`n"
		StartSCOMMaintenanceMode;
        Sleep 10;
      };
      "2" { "`n`n  You selected: Start SCOM MaintenanceMode for Local Server`n"
        $GitHubRawLink = "https://raw.githubusercontent.com/jholst71/public/main/Start_SCOM_MaintenanceMode.ps1"; IEX ((New-Object System.Net.WebClient).DownloadString($GitHubRawLink));
      };
      "5" { "`n`n  You selected: Get-LatestReboot for Local Server`n"
        $Result = Get-LatestReboot;
        $Result.LatestBootEventsExtended | FL; $result.LatestBootEvents | FT -Autosize; $result.LatestBootTime | FT  -Autosize;
        Pause;
      };
      "6" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LatestRebootDomain;
        $Result.LatestBootEvents | FT -Autosize;
        Pause;
      };
      "7" { "`n`n  You selected: Get-HotFixInstallDates for Locla Server`n"
        $Result = Get-HotFixInstallDates;
        $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      "8" { "`n`n  You selected: Get-HotFixInstallDates for Domain Servers`n"
        $Result = Get-HotFixInstallDatesDomain;
        $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      
      "9" { "`n`n  You selected: option #3`n"
        Sleep 10;
      };
      "I" { "`n`n  You selected: Information option `n"
	  "  Information will be updated later"
        Sleep 10;
	  };
    }; # End Switch
    #Pause;
  } until (($selection -eq "q") -or ($selection -eq "0"));
};
ToolboxMenu;
