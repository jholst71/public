### Name: Start_SCOM_MaintenanceMode.ps1
# Rightclick and Run with PowerShell
  ## Verify and Elevated Script
  If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy bypass;$arguments = "& '" + $myinvocation.mycommand.definition + "'";
    Start-Process powershell -Verb runAs -ArgumentList $arguments; Break;}
## Functions
Function F-StartSCOMMaintenanceMode {
  param( 
    $fDuration = ("30" | %{ If($Entry = Read-Host "  Enter MaintenanceMode Duration ( Default: $_ )"){$Entry} Else {$_} }), 
    $fComments = "SCOM MaintenanceMode started for $($SCOMMaintenanceModeDuration) minutes from $($env:Computername) by $($Env:USERNAME) at $(Get-Date)"
  );
  ## Script Begin
  Import-Module "C:\Program Files\Microsoft Monitoring Agent\Agent\MaintenanceMode.dll";
  try { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments -Force Y;
      } catch { Start-SCOMAgentMaintenanceMode -Reason "PlannedOther" -Duration $fDuration -Comment $fComments;}
  Write-Host "Request: Start SCOM Maintenance Mode for $($fDuration) minutes";
  };
## Script
F-StartSCOMMaintenanceMode
Sleep 15;
## Script End
