### Powershell Toolbox, by John Holst, itm8
#
### Parameter Functions
Function Get-CustomerName { ("CustomerName" | %{ If($Entry = Read-Host "  Enter CustomerName ( Default: $_ )"){$Entry} Else {$_} })};
Function Get-LogStartTime {
  # Add this line to Params: $fEventLogStartTime = (Get-LogStartTime -DefaultDays "7" -DefaultHours "12"),
  Param( $DefaultDays, $DefaultHours,
	$fLastXDays = ("$($DefaultDays)" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fLastXHours = ( %{If ( $fLastXDays -gt 0) {0} Else {"$($DefaultHours)" | %{ If($Entry = Read-Host "  Enter number of hours in searchscope (Default: $_ Hours)"){$Entry} Else {$_} } } })
	);
  ## Script
    Return [DateTime]::Now.AddDays(-$($fLastXDays)).AddHours(-$($fLastXHours));
};
Function Get-QueryComputers {  ### Get-QueryComputers - Get Domain Servers names
  # Add this line to Params: $fQueryComputers = $(Get-QueryComputers),
  Param(
    $fQueryComputerSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputerExcludeList = ($Entry = @(((Read-Host "  Enter ServerName(s) to be Exluded, separated by comma ").Split(",")).Trim()))
	);
  ## Script
    $fQueryComputers = Foreach ($fComputerSearch in $fQueryComputerSearch) {(Get-ADComputer -Filter 'operatingsystem -like "*server*" -and enabled -eq "true"' -Properties * | where { $fQueryComputerExcludeList -notcontains $_.name} -ErrorAction Continue | where { ($_.name -like $fComputerSearch)} -ErrorAction Continue)};
    $fQueryComputers = $fQueryComputers | Sort Name;
    Return $fQueryComputers;
};
Function Get-Filename { Param ( $fFileNameText, $fCustomerName ); ##
  # Add this line to Params: $fFileName = (Get-FileName -fFileNameText "<FILENAMETEXT>" -fCustomerName $fCustomerName)
  Return "$([Environment]::GetFolderPath("Desktop"))\$($fCustomerName)$($fFileNameText)_$(get-date -f yyyy-MM-dd_HH.mm)";
  #Return "$($env:USERPROFILE)\Desktop\$($fCustomerName)$($fFileNameText)_$(get-date -f yyyy-MM-dd_HH.mm)";
};
#
### Functions
Function Get-LatestRebootLocal { ### Get-LatestReboot - Get Latest Reboot / Restart / Shutdown for logged on server
  Param(
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
	$fEventLogStartTime = (Get-LogStartTime -DefaultDays "7" -DefaultHours "12"),
    $fFileName = (Get-FileName -fFileNameText "Get-LatestReboot_$($ENV:Computername)" -fCustomerName $fCustomerName)	
  );
  ## Script
    Show-Title "Get latest Shutdown / Restart / Reboot for Local Server - Events After: $($fEventLogStartTime)";
    $fLatestBootTime = Get-WmiObject win32_operatingsystem | select csname, @{LABEL="LastBootUpTime";EXPRESSION={$_.ConverttoDateTime($_.lastbootuptime)}};
    $fLatestBootEvents = Get-EventLog -LogName System -After $fEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008) -or ($_.EventID -eq 41)};
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
    $fCustomerName = $(Get-CustomerName),
    $fQueryComputers = $(Get-QueryComputers),
    $fEventLogStartTime = $(Get-LogStartTime -DefaultDays "7" -DefaultHours "12"),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fExportExtended = ("Yes" | %{ If($Entry = Read-Host "  Export Standard & Extended(message included) result to file - ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fJobNamePrefix = "LatestReboot_",
    $fFileName = (Get-FileName -fFileNameText "Servers_Get-LatestReboot" -fCustomerName $fCustomerName)
  );
  ## Script
    Show-Title "Get latest Shutdown / Restart / Reboot for multiple Domain Servers - Events After: $($fEventLogStartTime)";
    Foreach ($fQueryComputer in $fQueryComputers.name) { # Get $fQueryComputers-Values like .Name, .DNSHostName, or add them to variables in the scriptblocks/functions
      Write-Host "Querying Server: $($fQueryComputer)";
      $fBlock01 = {Get-EventLog -LogName System -After $Using:FEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008) -or ($_.EventID -eq 41) } };
      $fLocalBlock01 = {Get-EventLog -LogName System -After $fEventLogStartTime | Where-Object {($_.EventID -eq 1074) -or ($_.EventID -eq 6008) -or ($_.EventID -eq 41) }};
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Invoke-Command -scriptblock $fLocalBlock01;
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $fBlock01 -ComputerName $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
      };
    };
    Write-Host "  Waiting for jobs to complete... `n";
    Show-JobStatus $fJobNamePrefix;
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
    $fEventLogStartTime = $(Get-LogStartTime -DefaultDays "7" -DefaultHours "12"),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "Get-LatestLoginLogoff_$($ENV:Computername)" -fCustomerName $fCustomerName)
  );
  ## Default Variables
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
    $fCustomerName = $(Get-CustomerName),
    $fQueryComputers = $(Get-QueryComputers),
    $fEventLogStartTime = $(Get-LogStartTime -DefaultDays "7" -DefaultHours "12"),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "Servers_Get-LatestLoginLogoff" -fCustomerName $fCustomerName)
  );
  ## Default Variables
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
Function Get-InavtiveADUsers {## Get inactive AD Users / Latest Logon more than eg 90 days
  Param(
    $fCustomerName = $(Get-CustomerName),
    $fDaysInactive = ("90" | %{ If($Entry = Read-Host "  Enter number of inactive days (Default: $_ Days)"){$Entry} Else {$_} }),
	$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "Inactive_ADUsers_last_$($fDaysInactive)_days" -fCustomerName $fCustomerName)
  );
  ## Script
    Show-Title "Get AD Users Latest Logon / inactive more than $($fDaysInactive) days";
	$fDaysInactiveTimestamp = [DateTime]::Now.AddDays(-$($fDaysInactive));
    $fResult = Get-Aduser -Filter {(LastLogonTimeStamp -lt $fDaysInactiveTimestamp) -or (LastLogonTimeStamp -notlike "*")} -Properties *  | Sort-Object -Property samaccountname | Select CN,DisplayName,Samaccountname,@{n="LastLogonDate";e={[datetime]::FromFileTime($_.lastLogonTimestamp)}},Enabled,PasswordNeverExpires,@{Name='PwdLastSet';Expression={[DateTime]::FromFileTime($_.PwdLastSet)}},Description;
  ## Output
    #$fResult | Sort DisplayName | Select CN,DisplayName,Samaccountname,LastLogonDate,Enabled,PasswordNeverExpires,PwdLastSet,Description;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | Sort DisplayName | Select CN,DisplayName,Samaccountname,LastLogonDate,Enabled,PasswordNeverExpires,PwdLastSet,Description | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.InavtiveADUsers = $fResult | Sort DisplayName | Select CN,DisplayName,Samaccountname,LastLogonDate,Enabled,PasswordNeverExpires,PwdLastSet,Description;
    Return $Return;
};
Function Get-HotFixInstallDatesLocal { ### Get-HotFixInstallDates for multiple Domain servers
  Param(
    $fHotfixInstallDates = ("3" | %{ If($Entry = Read-Host "  Enter number of Hotfix-install dates per Computer (Default: $_ Install Dates)"){$Entry} Else {$_} }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "Get-LatestReboot_$($ENV:Computername)" -fCustomerName $fCustomerName)
    );
  ## Script
    Show-Title "Get latest $($fHotfixInstallDates) HotFix Install Dates Local Server";
    $fResult = Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $fHotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
    $fResult | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$((Get-ComputerInfo).WindowsProductName)";
    $fResult | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$((Get-NetIPAddress -AddressFamily IPv4 | ? {$_.IPAddress -notlike '127.0.0.1' }).IPAddress)";
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
    $fCustomerName = $(Get-CustomerName),
    $fQueryComputers = $(Get-QueryComputers),
    $fHotfixInstallDates = ("3" | %{ If($Entry = Read-Host "  Enter number of Hotfix-install dates per Computer (Default: $_ Install Dates)"){$Entry} Else {$_} }),
    #$fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fExport = "Yes",
    $fFileName = (Get-FileName -fFileNameText "Servers_Get-HotFixInstallDates" -fCustomerName $fCustomerName)
    );
  ## Script
    Show-Title "Get latest $($fHotfixInstallDates) HotFix Install Dates multiple Domain Servers";
    $fResult = @(); $fResult = Foreach ($fQueryComputer in $fQueryComputers) {
      Write-Host "  Querying Server: $($fQueryComputer.Name)";
      IF ($fQueryComputer.Name -eq $Env:COMPUTERNAME) {
        $fInstalledHotfixes = Get-Hotfix | sort InstalledOn -Descending -Unique -ErrorAction SilentlyContinue | Select -First $fHotfixInstallDates | Select PSComputerName, Description, HotFixID, InstalledBy, InstalledOn;
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "OperatingSystem" -Value "$((Get-ComputerInfo).WindowsProductName)";
        $fInstalledHotfixes | Add-Member -MemberType NoteProperty -Name "IPv4Address" -Value "$((Get-NetIPAddress -AddressFamily IPv4 | ? {$_.IPAddress -notlike '127.0.0.1' }).IPAddress)";
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
    $fFileName = (Get-FileName -fFileNameText "Get-Expired_Certificates" -fCustomerName $fCustomerName)
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
    $fCustomerName = $(Get-CustomerName),
    $fCertSearch = ("*" | %{ If($Entry = @(((Read-Host "  Enter Certificate SearchName(s), separated by comma ( Default: $_ )").Split(",")).Trim())){$Entry} Else {$_} }),
    $fQueryComputers = $(Get-QueryComputers),
    $fExpiresBeforeDays = ("90" | %{ If($Entry = Read-Host "  Enter number of days before expire (Default: $_ Days)"){$Entry} Else {$_} }),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fJobNamePrefix = "ExpiredCertificates_",
    $fFileName = (Get-FileName -fFileNameText "Servers_Get-Expired_Certificates" -fCustomerName $fCustomerName)
  );
  ## Script
    Show-Title "Get Certificates expired or expire within next $($fExpiresBeforeDays) days on multiple Domain Servers";
    $fExpiresBefore = [DateTime]::Now.AddDays($($fExpiresBeforeDays));
    $fResult = Foreach ($fQueryComputer in $fQueryComputers.name) { # Get $fQueryComputers-Values like .Name, .DNSHostName, or add them to variables in the scriptblocks/functions
      Write-Host "Querying Server: $($fQueryComputer)";
      $fBlock01 = {Get-childitem -path "cert:LocalMachine\my" -Recurse | ? {$_.NotAfter -lt "$Using:fExpiresBefore"} | ? {($_.Subject -like $Using:fCertSearch) -or ($_.FriendlyName -like $Using:fCertSearch)} | Select Subject,FriendlyName,NotAfter};
      $fLocalBlock01 = {Get-childitem -path "cert:LocalMachine\my" -Recurse | ? {$_.NotAfter -lt "$fExpiresBefore"} | ? {($_.Subject -like $fCertSearch) -or ($_.FriendlyName -like $fCertSearch)} | Select Subject,FriendlyName,NotAfter;};
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Invoke-Command -scriptblock $fLocalBlock01;
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $fBlock01 -ComputerName $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
      };
    };
    Write-Host "  Waiting for jobs to complete... `n";
    Show-JobStatus $fJobNamePrefix;
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
Function Get-DateTimeStatusDomain {## Get Date & Time Status - need an AD Server or Server with RSAT
  Param(
    $fCustomerName = $(Get-CustomerName),
    $fQueryComputers = $(Get-QueryComputers),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fJobNamePrefix = "DateTimeStatus_",
    $fFileName = (Get-FileName -fFileNameText "DateTimeStatus" -fCustomerName $fCustomerName)
	);
  ## Script
    Show-Title "Get Date and Time status from Domain Servers";
    Foreach ($fQueryComputer in $fQueryComputers.name) { # Get $fQueryComputers-Values like .Name, .DNSHostName, or add them to variables in the scriptblocks/functions
      Write-Host "Querying Server: $($fQueryComputer)";
      $fBlock01 = {
        $fInternetTime = Try {(Invoke-RestMethod -Uri "https://timeapi.io/api/Time/current/zone?timeZone=Europe/Copenhagen")} Catch {"Not Available"};
        $fLocalTime = (Get-Date -f "yyyy-MM-dd HH:mm:ss");
        New-Object psobject -Property ([ordered]@{
          InternetTime = If ($fInternetTime -ne "Not Available") {$fInternetTime.dateTime.replace("T"," ").split(".")[0]} Else {$fInternetTime};
          LocalTime = $fLocalTime;
          LocalNTPServer = (w32tm /query /source);
          LocalCulture = Get-Culture;
          LocalTimeZone = Try {Get-TimeZone} Catch {(Get-WMIObject -Class Win32_TimeZone).Caption};
          InternetTimeZone = If ($fInternetTime -ne "Not Available") {$fInternetTime.timeZone} Else {$fInternetTime};
        });
      };
      $fLocalBlock01 = {
        $fInternetTime = Try {(Invoke-RestMethod -Uri "https://timeapi.io/api/Time/current/zone?timeZone=Europe/Copenhagen")} Catch {"Not Available"};
        $fLocalTime = (Get-Date -f "yyyy-MM-dd HH:mm:ss");
        New-Object psobject -Property ([ordered]@{
          PSComputerName = $Env:COMPUTERNAME;
          InternetTime = If ($fInternetTime -ne "Not Available") {$fInternetTime.dateTime.replace("T"," ").split(".")[0]} Else {$fInternetTime};
          LocalTime = $fLocalTime;
          LocalNTPServer = (w32tm /query /source);
          LocalCulture = Get-Culture;
          LocalTimeZone = Try {Get-TimeZone} Catch {(Get-WMIObject -Class Win32_TimeZone).Caption};
          InternetTimeZone = If ($fInternetTime -ne "Not Available") {$fInternetTime.timeZone} Else {$fInternetTime};
        });
      };
      IF ($fQueryComputer -eq $Env:COMPUTERNAME) {
        $fLocalHostResult = Invoke-Command -scriptblock $fLocalBlock01 
      } ELSE {
        $JobResult = Invoke-Command -scriptblock $fBlock01 -ComputerName $fQueryComputer -JobName "$($fJobNamePrefix)$($fQueryComputer)" -ThrottleLimit 16 -AsJob
      };
    };
    Write-Host "  Waiting for jobs to complete... `n";
	Show-JobStatus $fJobNamePrefix;
	$fResult = Foreach ($fJob in (Get-Job -Name "$($fJobNamePrefix)*")) {Receive-Job -id $fJob.ID -Keep}; Get-Job -State Completed | Remove-Job;
    $fResult = $fResult + $fLocalHostResult;
  ## Output
    #$fResult | Sort PSComputerName | Select PSComputerName, InternetTime, LocalTime, LocalNTPServer, LocalCulture, LocalTimeZone, InternetTimeZone;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | Sort PSComputerName | Select PSComputerName, InternetTime, LocalTime, LocalNTPServer, LocalCulture, LocalTimeZone, InternetTimeZone | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.DateTimeStatus = $fResult | Sort PSComputerName | Select PSComputerName, InternetTime, LocalTime, LocalNTPServer, LocalCulture, LocalTimeZone, InternetTimeZone;
    Return $Return;
};
Function Get-FSLogixErrorsDomain {## Get FSLogix Errors - need an AD Server or Server with RSAT
  Param(
    $fCustomerName = $(Get-CustomerName),
    $fQueryComputers = $(Get-QueryComputers),
    $fLogDays = ("7" | %{ If($Entry = Read-Host "  Enter number of days in searchscope (Default: $_ Days)"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "FSLogix Errors" -fCustomerName $fCustomerName),
    $fErrorCodes1 = @("00000079", "0000001f", "00000020"),
    $fErrorCodes2 = @("00000079", "0000001f"),
    $fErrorCodeList = "  Internal Error Code Description:
  00000005 Access is denied
  00000020 Operation 'OpenVirtualDisk' failed / Failed to open virtual disk / The process cannot access the file because it is being used by another process.
  00000091 The directory is not empty
  00000079 Failed to attach VHD / LoadProfile failed / AttachVirtualDisk error for user
  0000001f Error (A device attached to the system is not functioning.)
  00000091 Error removing directory (The directory is not empty.)
  000003f8 Restoring registry key (An I/O operation initiated by the registry failed unrecoverable...)
  0000078f FindFile failed for path: / LoadProfile failed.
  00000490 Failed to remove RECYCLE.BIN redirect (Element not found.)
  00001771 Failed to restore credentials. Unable to decrypt value from BlobDpApi attribute (The specified file could not be decrypted.)
  0000a418 Unable to get the supported size or compact the disk, Message: Cannot shrink a partition containing a volume with errors
  80070003 Failed to save installed AppxPackages (The system cannot find the path specified.)
  80070490 Error removing Rule (Element not found)"
  );
  ## Script
    Show-Title "Get FSLogix Errors for past $($fLogDays) days";
    $fExportAllErrors = "$FALSE" ; # Select "$TRUE" or "$FALSE"
    ## ErrorCode Selection
      Clear-Host;
      Write-Host "`n  ================ Select FSLogix ErrorCodes ================`n";
      Write-Host "  Press '1'  for FSLogix ErrorCodes $($fErrorCodes1).";
      Write-Host "  Press '2'  for FSLogix ErrorCodes $($fErrorCodes2).";
      Write-Host "  Press 'M'  for entering FSLogix ErrorCodes manually.";
      Write-Host "  Press 'A'  for All FSLogix ErrorCodes.";
    $ErrorCodeSelection = Read-Host "`n  Please make a selection"
    switch ($ErrorCodeSelection){
      "1" {$fErrorCodes = $fErrorCodes1;} # @("ERROR:", "WARN:") @("00000079", "0000001f", "00000020")
      "2" {$fErrorCodes = $fErrorCodes2;} # @("ERROR:", "WARN:") @("00000079", "0000001f");} 
      "m" {$fErrorCodes = ($Entry = @(((Read-Host "  Enter FXLogix ErrorCode(s), to search for, separated by comma").Split(",")).Trim()));}
      "a" {$fExportAllErrors = "$TRUE"} ; # Select "$TRUE" or "$FALSE"
    };
    $fLogText = Foreach ( $fQueryComputer in $fQueryComputers.name) {
      Write-Host "Querying Computer: $($fQueryComputer)";
      Foreach ($fProfilePath in (gi \\$fQueryComputer\C$\ProgramData\FSLogix\Logs\Profile\Profile-*.log)[-$($fLogDays)..-1]) {
        Get-Content -Path $fProfilePath | Where-Object {($_ -like "*ERROR:*") -or ($_ -like "*WARN:*")} |Foreach  {
        New-Object psobject -Property @{
          Servername = $fQueryComputer
          Date = ($fProfilePath | Select -ExpandProperty CreationTime) | Get-Date -f "yyyy-MM-dd"
          Time = $_.split("]")[0].trim("[")
          tid = $_.split("]")[1].trim("[")
          Error = $_.split("]")[2].trim("[")
          LogText = $_.split("]")[3].trim("  ")
      }}};
      Foreach ($fProfilePath in (gi \\$fQueryComputer\C$\ProgramData\FSLogix\Logs\ODFC\ODFC-*.log)[-$($fLogDays)..-1]) {
        Get-Content -Path $fProfilePath | Where-Object {($_ -like "*ERROR:*") -or ($_ -like "*WARN:*")} |Foreach  {
        New-Object psobject -Property @{
          Servername = $fQueryComputer
          Date = ($fProfilePath | Select -ExpandProperty CreationTime) | Get-Date -f "yyyy-MM-dd"
          Time = $_.split("]")[0].trim("[")
          tid = $_.split("]")[1].trim("[")
          Error = $_.split("]")[2].trim("[")
          LogText = $_.split("]")[3].trim("  ")
      }}};
    };
    $fResult = Foreach ($fErrorCode in $fErrorCodes) {$fLogText | Where-Object { $_ -like "*$($fErrorCode)*" }};
  ## Output
    #$fResult | Sort DisplayName | Select CN,DisplayName,Samaccountname,LastLogonDate,Enabled,PasswordNeverExpires,PwdLastSet,Description;
    #If ($fExportAllErrors -ne $true) { $fResult | sort Servername, Date, Time | FT Servername, Date, Time, Error, tid, LogText; Write-Host "   Number of errorcodes listed: $($fResult.count)`n"; };
    If ($fExportAllErrors -ne $true) { Write-Host "`n  Number of errorcodes listed: $($fResult.count)`n"; } else { Write-Host "`n  Number of errorcodes listed: $($fLogText.count)`n" };
  ## Exports
    #If (($fExport -eq "Y") -or ($fExport -eq "YES")) { $fResult | Sort Servername, Date, Time | Select Servername, Date, Time, Error, tid, LogText | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
    If ($fExportAllErrors -ne $true) { 
      $fResult | sort Servername, Date, Time | Select Servername, Date, Time, Error, tid, LogText | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation;
	} else {
      $fLogText | sort Servername, Date, Time | Select Servername, Date, Time, Error, tid, LogText | Export-CSV "$($fFileName)_All_Errors.csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation;
	};
  ## Return
    [hashtable]$Return = @{}
    $Return.FSLogixErrors = $fResult | Sort Servername, Date, Time | Select Servername, Date, Time, Error, tid, LogText;
    Return $Return;
};
Function Get-FolderPermissionLocal { ## 
  Param(
    $fCustomerName = $(Get-CustomerName),
    $fFolderPaths = ($Entry = @(((Read-Host "  Enter FolderPath(s) to get Permission list, separated by comma ").Split(",")).Trim())),
    $fExport = ("Yes" | %{ If($Entry = Read-Host "  Export result to file ( Y/N - Default: $_ )"){$Entry} Else {$_} }),
    $fFileName = (Get-FileName -fFileNameText "Get-FolderPermission_$($ENV:Computername)" -fCustomerName $($fCustomerName))
  );
  ## Script
    Show-Title "Get Certificates expired or expire within next $($fExpiresBeforeDays) days on Local Server";
    $fResult = ForEach ($fFolderPath in $fFolderPaths) {
      $fFolders = Get-ChildItem -Directory -Path "$($fFolderPath)" -Recurse -Force;
      ForEach ($fFolder in $fFolders) {
        $fAcl = Get-Acl -Path $fFolder.FullName;
        ForEach ($fAccess in $fAcl.Access) {
          New-Object PSObject -Property ([ordered]@{
            'FolderName'=$fFolder.FullName;
            'Group/User'=$fAccess.IdentityReference;
            'Permissions'= $fAccess.FileSystemRights;
            'Inherited'=$fAccess.IsInherited;
      });};};};
  ## Output
    #$fResult | Sort FolderName, "Group/User" | Select FolderName, "Group/User", Permissions, Inherited | FT -autosize;
  ## Exports
    If (($fExport -eq "Y") -or ($fExport -eq "YES")) {$fResult | Sort FolderName, "Group/User" | Select FolderName, "Group/User", Permissions, Inherited | Export-CSV "$($fFileName).csv" -Delimiter ';' -Encoding UTF8 -NoTypeInformation; };
  ## Return
    [hashtable]$Return = @{}
    $Return.FolderPermission = $fResult | Sort FolderName, "Group/User" | Select FolderName, "Group/User", Permissions, Inherited;
    Return $Return;
};
## Shared Functions
Function Show-Title {
  param ( [string]$Title );
    $host.UI.RawUI.WindowTitle = $Title;
};
Function Show-JobStatus { Param ($fJobNamePrefix)
    DO { IF ((Get-Job -Name "$($fJobNamePrefix)*").count -ge 1) {$fStatus = ((Get-Job -State Completed).count/(Get-Job -Name "$($fJobNamePrefix)*").count) * 100;
      Write-Progress -Activity "Waiting for $((Get-Job -State Running).count) job(s) to complete..." -Status "$($fStatus) % completed" -PercentComplete $fStatus; }; }
    While ((Get-job -Name "$($fJobNamePrefix)*" | Where State -eq Running));
};
Function Show-Help {
  Show-Title "$($Title) Help / Information";
  Clear-Host;
  Write-Host "  Help / Information will be updated later";
};
Function Show-Menu {
  param (
    [string]$Title = "PSToolbox"
  );
  Show-Title $Title;
  Clear-Host;
  Write-Host "`n  ================ $Title ================`n";
  Write-Host "  Press '1'  for Get-LatestReboot for Local Server.";
  Write-Host "  Press '2'  for Get-LatestReboot for Domain Servers.";
  Write-Host "  Press '3'  for Get-LoginLogoff for Local Server.";
  Write-Host "  Press '4'  for Get-LoginLogoff for Domain Servers.";
  Write-Host "  Press '5'  for Get inactive AD Users / last logon more than eg 90 days.";
  #Write-Host "  Press '6'  for Start SCOM MaintenanceMode for Local Server (Script).";
  Write-Host "  "
  Write-Host "  Press '11' for Get-HotFixInstallDates for Local Server.";
  Write-Host "  Press '12' for Get-HotFixInstallDates for Domain Servers.";
  Write-Host "  Press '13' for Get-ExpiredCertificates for Local Server.";
  Write-Host "  Press '14' for Get-ExpiredCertificates for Domain Servers.";
  Write-Host "  Press '15' for Get-FolderPermission for Local Server.";
  Write-Host "  Press '16' for Get-DateTimeStatus for Domain Servers.";
  Write-Host "  Press '17' for Get-FSLogixErrors for Domain Servers.";
  #Write-Host "  Press '99' for this option.";
  Write-Host "  ";
  Write-Host "   Press 'H'  for Toolbox Help / Information.";
  Write-Host "   Press 'Q'  to quit.";
};
Function ToolboxMenu {
  do {
    Show-Menu
    $selection = Read-Host "`n  Please make a selection"
    switch ($selection){
      "1" { "`n`n  You selected: Get-LatestReboot for Local Server`n"
        $Result = Get-LatestRebootLocal; $Result.LatestBootEventsExtended | FL; $result.LatestBootEvents | FT -Autosize; $result.LatestBootTime | FT -Autosize;
        Pause;
      };
      "2" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LatestRebootDomain; $Result.LatestBootEvents | FT -Autosize;
        Pause;
      };
      "3" { "`n`n  You selected: Get-LatestReboot for Local Server`n"
        $Result = Get-LoginLogoffLocal; $Result.LoginLogoff | FT -Autosize;
        Pause;
      };
      "4" { "`n`n  You selected: Get-LatestReboot for Domain Servers`n"
        $Result = Get-LoginLogoffDomain; $Result.LoginLogoff | FT -Autosize;
        Pause;
      };	  
      "5" { "`n`n  You selected: Get inactive AD Users / last logon more than eg 90 days`n"
        $Result = Get-InavtiveADUsers; $Result.InavtiveADUsers | FT -Autosize;
        Pause;
      };
      "6" { "`n`n  You selected: Start SCOM MaintenanceMode for Local Server`n"
        
      };
      "11" { "`n`n  You selected: Get-HotFixInstallDates for Local Server`n"
        $Result = Get-HotFixInstallDatesLocal; $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      "12" { "`n`n  You selected: Get-HotFixInstallDates for Domain Servers`n"
        $Result = Get-HotFixInstallDatesDomain; $Result.HotFixInstallDates | FT -Autosize;
        Pause;
      };
      "13" { "`n`n  You selected: Get-ExpiredCertificates for Local Server`n"
        $Result = Get-ExpiredCertificatesLocal; $Result.ExpiredCertificates | FT -Autosize;
        Pause;
      };
      "14" { "`n`n  You selected: Get-ExpiredCertificates for Domain Servers`n"
        $Result = Get-ExpiredCertificatesDomain; $Result.ExpiredCertificates | FT -Autosize;
        Pause;
      };
      "15" { "`n`n  You selected: Get-FolderPermission `n"
        $Result = Get-FolderPermissionLocal; $Result.FolderPermission | FT -Autosize;
        Pause;
      };
      "16" { "`n`n  You selected: Get-DateTimeStatus for Domain Servers`n"
        $Result = Get-DateTimeStatusDomain; $Result.DateTimeStatus | FT -Autosize;
        Pause;
      };
      "17" { "`n`n  You selected: Get-FSLogixErrors for Domain Servers`n"
        $Result = Get-FSLogixErrorsDomain; $Result.FSLogixErrors | FT -Autosize;
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
## Start Menu
ToolboxMenu;
