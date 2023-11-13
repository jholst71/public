### Progressive Toolbox, by John Holst, Progressive
#
Function Get-LatestRebootLocal { ### Get-LatestReboot - Get Latest Reboot / Restart / Shutdown for logged on server
  Param(
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
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
    $fCustomerName = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fQueryComputerSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Search ServerName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList = ("*" | %{ If($Entry = @(((Read-Host "  Enter ServerName(s) to be Exluded, separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fExportExtended = ("Yes" | %{ If($Entry = Read-Host "  Export Standard & Extended(message included) result to file - ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fJobNamePrefix = "RegQuery_",
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-LatestReboot_$(get-date -f yyyy-MM-dd_HH.mm)"
  );
  ## Script
    $fEventLogStartTime = [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
    Show-Title "Get latest Shutdown / Restart / Reboot for multiple Domain Servers - Events After: $($fEventLogStartTime)";
    $fQueryComputers = (Get-QueryComputers -fQueryComputerSearch $fQueryComputerSearch -fQueryComputerExcludeList $fQueryComputerExcludeList); 
    Foreach ($fQueryComputer in $fQueryComputers.name) { # Get $fQueryComputers-Values like .Name, .DNSHostName, or add them to variables in the scriptblocks/functions
      Write-Host "Querying Server: $($fQueryComputer)";
      $fBlock01 = {Get-EventLog -LogName System -After $Using:FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)} };
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Get-EventLog -LogName System -After $fEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008)};
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $fBlock01 -ComputerName $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
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
Function Get-LoginLogoffLocal { ## Get-LoginLogoff from Logged On Server
  Param(
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\Get-LatestLoginLogoff_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\Get-LatestLoginLogoff_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
  );
  ## Default Variables
    $fEventLogStartTime = [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
    $fUserProperty = @{n="User";e={(New-Object System.Security.Principal.SecurityIdentifier $_.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])}}
    $fTypeProperty = @{n="Action";e={if($_.EventID -eq 7001) {"Logon"} elseif ($_.EventID -eq 7002){"Logoff"} else {"other"}}}
    $fTimeProperty = @{n="Time";e={$_.TimeGenerated}}
    $fMachineNameProperty = @{n="MachinenName";e={$_.MachineName}}
  ## Script
    Show-Title "Get latest Login / Logoff for Local Server - Events After: $($fEventLogStartTime)";
    Write-Host "Querying Computer: $($ENV:Computername)"
    $fResult = Get-EventLog System -Source Microsoft-Windows-Winlogon -after $fEventLogStartTime | select $fUserProperty,$fTypeProperty,$fTimeProperty,$fMachineNameProperty
  ## Output
    #$fResult | sort User, Time | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort User, Time | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.LoginLogoff = $fResult | sort User, Time;
    Return $Return;
};
Function Get-LoginLogoffDomain { ## Get-LoginLogoffDomain (Remote) from Event Log: Microsoft-Windows-Winlogon
  Param(
    $fCustomerName = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fQueryComputerSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Search ServerName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList = ("*" | %{ If($Entry = @(((Read-Host "  Enter ServerName(s) to be Exluded, separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fLastXDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours = ( %{If ( $fLastXDays -gt 0) {0} Else {"12" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-LatestLoginLogoff_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-LatestLoginLogoff_$(get-date -f yyyy-MM-dd_HH.mm)"
  );
  ## Default Variables
    $fQueryComputers = (Get-QueryComputers -fQueryComputerSearch $fQueryComputerSearch -fQueryComputerExcludeList $fQueryComputerExcludeList); 
    $fEventLogStartTime = [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
    $fUserProperty = @{n="User";e={(New-Object System.Security.Principal.SecurityIdentifier $_.ReplacementStrings[1]).Translate([System.Security.Principal.NTAccount])}}
    $fTypeProperty = @{n="Action";e={if($_.EventID -eq 7001) {"Logon"} elseif ($_.EventID -eq 7002){"Logoff"} else {"other"}}}
    $fTimeProperty = @{n="Time";e={$_.TimeGenerated}}
    $fMachineNameProperty = @{n="MachinenName";e={$_.MachineName}}
  ## Script
    Show-Title "Get latest Login / Logoff  for multiple Domain Servers - Events After: $($fEventLogStartTime)";
    $fResult = foreach ($fComputer in $fQueryComputers.name) { # Get Values like .Name, .DNSHostName
      Write-Host "Querying Computer: $($fComputer)"
      Get-EventLog System -Source Microsoft-Windows-Winlogon -ComputerName $fComputer -after $fEventLogStartTime | select $fUserProperty,$fTypeProperty,$fTimeProperty,$fMachineNameProperty
    };
  ## Output
    #$fResult | sort User, Time | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort User, Time | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.LoginLogoff = $fResult | sort User, Time;
    Return $Return;
};
Function Get-HotFixInstallDatesLocal { ### Get-HotFixInstallDates for multiple Domain servers
  Param(
    $fHotfixInstallDates = ("3" | %{ If($Entry = Read-Host "  Enter number of Hotfix-install dates per Computer (Default: $_ Install Dates)"){$Entry} Else {$_} }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\Get-LatestReboot_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
    );
  ## Script
    Show-Title "Get latest $($fHotfixInstallDates) HotFix Install Dates Local Server";
    $fResult = Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $fHotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
    $fResult | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$((Get-ComputerInfo).WindowsProductName)";
    $fResult | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$((Get-NetIPAddress -AddressFamily IPv4 | ? {$_.IPAddress -notlike "127.0.0.1" }).IPAddress)";
  ## Output
    #$fResult | sort MachineName, TimeGenerated | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort MachineName, TimeGenerated | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.HotFixInstallDates = $fResult | sort MachineName, TimeGenerated | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address;
    Return $Return;
};
Function Get-HotFixInstallDatesDomain { ### Get-HotFixInstallDates for multiple Domain servers
  Param(
    $fCustomerName = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fQueryComputerSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Search ServerName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList = ("*" | %{ If($Entry = @(((Read-Host "  Enter ServerName(s) to be Exluded, separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fHotfixInstallDates = ("3" | %{ If($Entry = Read-Host "  Enter number of Hotfix-install dates per Computer (Default: $_ Install Dates)"){$Entry} Else {$_} }),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-HotFixInstallDates_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-HotFixInstallDates_$(get-date -f yyyy-MM-dd_HH.mm)"
    );
  ## Script
    Show-Title "Get latest $($fHotfixInstallDates) HotFix Install Dates multiple Domain Servers";
    $fQueryComputers = (Get-QueryComputers -fQueryComputerSearch $fQueryComputerSearch -fQueryComputerExcludeList $fQueryComputerExcludeList); # Get Values like .Name, .DNSHostName
    $fResult = @(); $fResult = Foreach ($fQueryComputer in $fQueryComputers) {
      Write-Host "  Querying Server: $($fQueryComputer.Name)";
      IF ($fQueryComputer.Name -eq $Env:COMPUTERNAME) {
        $fInstalledHotfixes = Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $fHotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$((Get-ComputerInfo).WindowsProductName)";
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$((Get-NetIPAddress -AddressFamily IPv4 | ? {$_.IPAddress -notlike "127.0.0.1" }).IPAddress)";
        $fInstalledHotfixes; 
      } Else {
        try {
          $fInstalledHotfixes = Get-Hotfix -ComputerName $fQueryComputer.Name | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $fHotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
          $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$($fQueryComputer.OperatingSystem)";
          $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$($fQueryComputer.IPv4Address)";
          $fInstalledHotfixes; 
        } catch {
          Write-Host "      An error occurred within the Get-Hotfix command:"
          Write-Host "      $($_.ScriptStackTrace)"
          Write-Host "    Trying with an Invoked Get-Hotfix command: "
          try {
            $fInstalledHotfixes = Invoke-Command -scriptblock { Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $USING:fHotfixInstallDates  | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;  } -computername $QueryComputer
            $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$($fQueryComputer.OperatingSystem)";
            $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$($fQueryComputer.IPv4Address)";
            $fInstalledHotfixes; 
          } catch {
            Write-Host "      An error occurred within the Invoked Get-Hotfix command:"
            Write-Host "      $($_.ScriptStackTrace)`n"
            $fInstalledHotfixes = [pscustomobject]@{
              "PSComputerName" = "$($fQueryComputer.Name)"
              "Description" = ""
              "HotFixID" = ""
              "InstalledBy" = ""
              "InstalledOn" = ""
              "IPv4Address" = "$($fQueryComputer.IPv4Address)"
              "OperatingSystem" = "$($fQueryComputer.OperatingSystem)"};
            $fInstalledHotfixes; 
    }}}};
  ## Output
    #$fResult | sort MachineName, TimeGenerated | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | sort PSComputerName | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.HotFixInstallDates = $fResult | sort PSComputerName | Select PSComputerName, InstalledOn, InstalledBy, Description, HotFixID, OperatingSystem, IPv4Address;
    Return $Return;
};
Function Get-ExpiredCertificatesLocal {## Get-ExpiredCertificates
  Param(
    $fCertSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Certificate SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fExpiresBeforeDays = ("90" | %{ If($Entry = Read-Host "  Enter number of days before expire (Default: $_ Days)"){$Entry} Else {$_} }),
	$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\Get-Expired_Certificates_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\Get-Expired_Certificates_$($ENV:Computername)_$(get-date -f yyyy-MM-dd_HH.mm)"
  );
  ## Script
    Show-Title "Get Certificates expired or expire within next $($fExpiresBeforeDays) days on Local Server";
	$fExpiresBefore = [DateTime]::Now.AddDays($($fExpiresBeforeDays));
    $fResult = Get-childitem -path "cert:LocalMachine\my" -Recurse | ? {$_.NotAfter -lt "$fExpiresBefore"} | ? {($_.Subject -like $fCertSearch) -or ($_.FriendlyName -like $fCertSearch)} | Select Subject,FriendlyName,NotAfter
  ## Output
    #$fResult | sort NotAfter, FriendlyName | Select NotAfter, FriendlyName, Subject | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult |  sort NotAfter, FriendlyName | Select NotAfter, FriendlyName, Subject | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.ExpiredCertificates = $fResult |  sort NotAfter, FriendlyName | Select NotAfter, FriendlyName, Subject;
    Return $Return;
};
Function Get-ExpiredCertificatesDomain {## Get-Expired_Certificates
  Param(
    $fCustomerName = ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} }),
    $fCertSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Certificate SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList = ("*" | %{ If($Entry = @(((Read-Host "  Enter ServerName(s) to be Exluded, separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fExpiresBeforeDays = ("90" | %{ If($Entry = Read-Host "  Enter number of days before expire (Default: $_ Days)"){$Entry} Else {$_} }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fJobNamePrefix = "RegQuery_",
    $fFileName = "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)_Servers_Get-Expired_Certificates_$(get-date -f yyyy-MM-dd_HH.mm)"
    #$fFileName = "$($env:USERPROFILE)\Desktop\$($fCustomerName)_Servers_Get-Expired_Certificates_$(get-date -f yyyy-MM-dd_HH.mm)"
  );
  ## Script
    Show-Title "Get Certificates expired or expire within next $($fExpiresBeforeDays) days on multiple Domain Servers";
    $fQueryComputers = (Get-QueryComputers -fQueryComputerSearch $fQueryComputerSearch -fQueryComputerExcludeList $fQueryComputerExcludeList); # Get Values like .Name, .DNSHostName
    $fExpiresBefore = [DateTime]::Now.AddDays($($fExpiresBeforeDays));
    $fResult = Foreach ($fQueryComputer in $fQueryComputers.name) { # Get $fQueryComputers-Values like .Name, .DNSHostName, or add them to variables in the scriptblocks/functions
      Write-Host "Querying Server: $($fQueryComputer)";
      $fBlock01 = { Get-childitem -path "cert:LocalMachine\my" -Recurse | ? {$_.NotAfter -lt "$Using:fExpiresBefore"} | ? {($_.Subject -like $Using:fCertSearch) -or ($_.FriendlyName -like $Using:fCertSearch)} | Select Subject,FriendlyName,NotAfter};
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Get-childitem -path "cert:LocalMachine\my" -Recurse | ? {$_.NotAfter -lt "$fExpiresBefore"} | ? {($_.Subject -like $fCertSearch) -or ($_.FriendlyName -like $fCertSearch)} | Select Subject,FriendlyName,NotAfter;
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $fBlock01 -ComputerName $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
      };
    };
    Write-Host "  Waiting for jobs to complete... `n";
    DO { $fStatus = ((Get-Job -State Completed).count/(Get-Job -Name "$($fJobNamePrefix)*").count) * 100;
      Write-Progress -Activity "Waiting for $((Get-Job -State Running).count) job(s) to complete..." -Status "$($fStatus) % completed" -PercentComplete $fStatus; }
    While ((Get-job -Name "$($fJobNamePrefix)*" | Where State -eq Running));
    $fResult = Foreach ($fJob in (Get-Job -Name "$($fJobNamePrefix)*")) {Receive-Job -id $fJob.ID -Keep}; Get-Job -State Completed | Remove-Job;
    $fResult = $fResult + $fLocalHostResult;
  ## Output
    #$fResult | sort NotAfter, NotAfter | Select PSComputerName, NotAfter, FriendlyName, Subject | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult |  sort NotAfter, NotAfter | Select PSComputerName, NotAfter, FriendlyName, Subject | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.ExpiredCertificates = $fResult |  sort NotAfter, NotAfter | Select PSComputerName, NotAfter, FriendlyName, Subject;
    Return $Return;
};
Function StartSCOMMaintenanceMode { ### Start SCOM Maintenance Mode
  param( 
    $fDuration = ("30" | %{ If($Entry = Read-Host "  Enter MaintenanceMode Duration ( Default: $_ )"){$Entry} Else {$_} }),
    $fComments = "SCOM MaintenanceMode started for $($SCOMMaintenanceModeDuration) minutes from $($env:Computername) by $($Env:USERNAME) at $(Get-Date)"
  );
  ## Script
    Show-Title "Start SCOM Maintenance Mode at Local Server";
    Import-Module "C:\Program Files\Microsoft Monitoring Agent\Agent\MaintenanceMode.dll";
    try { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments -Force Y;
      } catch { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments;}
    Write-Host "Request: Start SCOM Maintenance Mode for $($fDuration) minutes";
};
## Shared Functions
Function Get-QueryComputers {  ### Get-QueryComputers - Get Domain Servers names 
  Param( $fQueryComputerSearch, $fQueryComputerExcludeList )
  ## Script
    $fQueryComputers = Foreach ($fComputerSearch in $fQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' -Properties * | where { $fQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $fComputerSearch)} -ErrorAction Continue)};
    $fQueryComputers = $fQueryComputers | Sort Name;
    Return $fQueryComputers;
};
Function Show-Title {
  param ( [string]$Title );
    $host.UI.RawUI.WindowTitle = $Title;
};
Function Show-Help {
  Show-Title "$($Title) Help / Information";
  Clear-Host;
  Write-Host "  Help / Information will be updated later";
};
Function Show-Menu {
  param (
    [string]$Title = "Progressive Toolbox"
  );
  Show-Title $Title;
  Clear-Host;
  Write-Host "`n  ================ $Title ================`n";
  #Write-Host "   1: Press '1'  for Start SCOM MaintenanceMode for Local Server.";
  #Write-Host "   2: Press '2'  for Start SCOM MaintenanceMode for Local Server (Script).";
  Write-Host "   5: Press '5'  for Get-LatestReboot for Local Server.";
  Write-Host "   6: Press '6'  for Get-LatestReboot for Domain Servers.";
  Write-Host "   7: Press '7'  for Get-LoginLogoff for Local Server.";
  Write-Host "   8: Press '8'  for Get-LoginLogoff for Domain Servers.";
  Write-Host "  ";
  Write-Host "  11: Press '11' for Get-HotFixInstallDates for Local Server.";
  Write-Host "  12: Press '12' for Get-HotFixInstallDates for Domain Servers.";
  Write-Host "  13: Press '13' for Get-ExpiredCertificates for Local Server.";
  Write-Host "  14: Press '14' for Get-ExpiredCertificates for Domain Servers.";
  #Write-Host "  99: Press '99' for this option.";
  Write-Host "  ";
  Write-Host "   H: Press 'H'  for Toolbox Help / Information.";
  Write-Host "   Q: Press 'Q'  to quit.";
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
        $Result = Get-LatestRebootLocal;
        $Result.LatestBootEventsExtended | FL; $result.LatestBootEvents | FT -Autosize; $result.LatestBootTime | FT -Autosize;
        Pause;
      };
      "6" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LatestRebootDomain;
        $Result.LatestBootEvents | FT -Autosize;
        Pause;
      };
      "7" { "`n`n  You selected: Get-LatestReboot for Local Server`n"
        $Result = Get-LoginLogoffLocal;
        $Result.LoginLogoff | FT -Autosize;
        Pause;
      };
      "8" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LoginLogoffDomain;
        $Result.LoginLogoff | FT -Autosize;
        Pause;
      };	  
      "11" { "`n`n  You selected: Get-HotFixInstallDates for Local Server`n"
        $Result = Get-HotFixInstallDatesLocal;
        $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      "12" { "`n`n  You selected: Get-HotFixInstallDates for Domain Servers`n"
        $Result = Get-HotFixInstallDatesDomain;
        $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      "13" { "`n`n  You selected: Get-ExpiredCertificates for Local Server`n"
        $Result = Get-ExpiredCertificatesLocal;
        $Result.ExpiredCertificates | FT -Autosize;
        Pause;
      };
      "14" { "`n`n  You selected: Get-ExpiredCertificates for Domain Servers`n"
        $Result = Get-ExpiredCertificatesDomain;
        $Result.ExpiredCertificates | FT -Autosize;
        Pause;
      };
      "99" { "`n`n  You selected: Test option #99`n"
        Sleep 10;
      };
      "H" { "`n`n  You selected: Help / Information option `n"
        Show-Help;
        Pause;
      };
    }; # End Switch
    #Pause;
  } until (($selection -eq "q") -or ($selection -eq "0"));
};
ToolboxMenu;
