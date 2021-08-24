
$logo = @('


    ____  __           ________              __            
   / __ )/ /_  _____  / ____/ /_  ___  _____/ /_____  _____
  / __  / / / / / _ \/ /   / __ \/ _ \/ ___/ //_/ _ \/ ___/
 / /_/ / / /_/ /  __/ /___/ / / /  __/ /__/ ,< /  __/ /    
/_____/_/\__,_/\___/\____/_/ /_/\___/\___/_/|_|\___/_/  

Creator: Securethelogs.com     |      @Securethelogs


')

$logo


$computername = hostname
$version = $PSVersionTable.PSVersion.Major 

$v2check = Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2Root -ErrorAction SilentlyContinue
if ($v2check.state -ne "Enabled" -and $v2check.state -ne "disabled"){$v2check = Get-WindowsFeature PowerShell-V2}

$winrmstatus = (Get-Service WinRm).status

$status = @()

$Status += "Hostname: " + $computername
$Status += "PSVersion: " + $version
$Status += "PowerShell V2 Status: " + $v2check.state
$Status += "WinRM Service: " + $winrmstatus

Write-Output ""
$status
Write-Output ""


# Check History

$CheckPSReadLine = Test-Path -Path "C:\Program Files\WindowsPowerShell\Modules\PSReadline"

if ($CheckPSReadLine -eq "True") {
$downgradedcheck = @((Select-String -Pattern '-version 2','-v 2' -Path "$env:APPDATA\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt").Line)

}

if ($downgradedcheck -ne $null){$dwngrdhistory = $true}else{$dwngrdhistory = $false}
if ($dwngrdhistory -eq $false){$downgradedcheck = "No Evidence Found"}


Write-Output "Checking auditing levels...."


# Check auditing


$modkey = ""
$modscript = ""

$CheckIfModule = Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging -erroraction SilentlyContinue
$CheckIfScript = Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging -erroraction SilentlyContinue
$CheckIfCmd = Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Audit -erroraction SilentlyContinue



# ------------- Module Logging --------------------

if ($CheckIfModule -eq $null) {

$CheckIfModuleWoW = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell\ModuleLogging -erroraction SilentlyContinue

if ($CheckIfModuleWoW -eq $null){$modkey = $false}else{$modkey = "2"}

}else{$modkey = "1"}


if ($modkey -eq "1") {$EML = (Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging).EnableModuleLogging}

if ($modkey -eq "2") {$EML = (Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell\ModuleLogging).EnableModuleLogging}


if ($modkey -eq $false) {$moduleresults = "<p>  Module Logging: EnableModuleLogging : Disabled </p>"} else{

$moduleresults = "<p> Module Logging: EnableModuleLogging : Enabled </p>"

}




# ------------ BlockScripting ------------


if ($CheckIfScript -eq $null) {$CheckIfScriptWoW = Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging -erroraction SilentlyContinue

if ($CheckIfScriptWoW -eq $null){$modscript = $false}else{$modscript = "2"}

}else {$modscript = "1"}


if ($modscript -eq "1"){$EBS = (Get-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging).EnableScriptBlockLogging}

if ($modscript -eq "2"){$EBS = (Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging).EnableScriptBlockLogging}


if ($modscript -eq $false){$EBS = "<p> ScriptBlockLogging: Disabled </p>"}else{

$EBS = "<p>ScriptBlockLogging: EnableScriptBlockLogging is: Enabled </p>"

}


# -------------- Get ProcessCreationIncludeCmdLine_Enabled Result and Return -----------------------

if ($CheckIfCmd -eq $null) {$PCIC = "<p> ProcessCreationLogging: Disabled </p>"}else{

$PCIC = (Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Audit).ProcessCreationIncludeCmdLine_Enabled

$PCIC = "<p> ProcessCreationLogging: ProcessCreationIncludeCmdLine is set to: " + $PCIC + "</p>"

}



Write-Output "Checking for common exploits and keywords...."


# ------------ Check History For Common Exploit Scripts ------------------------


if ($CheckPSReadLine -eq "True"){$Checkforkeywords = @(Select-String -Pattern 'nishang','powersploit','mimikatz','mimidogz','mimiyakz','-nop','(New-Object Net.WebClient).DownloadString','–ExecutionPolicy Bypass' -Path "$env:APPDATA\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt")}

if ($Checkforkeywords -eq $null){$Checkforkeywords = "No suspicious keywords found"}




Write-Output "Checking firewall rules for Powershell...."




# ---------------------------- Check For Firewall ---------------------------

###################improve
$psfirewall = @(Get-NetFirewallApplicationFilter -Program "*powershell.exe" | Get-NetFirewallRule | Select-Object DisplayName, Enabled, Profiles, Direction, Action)


if ($psfirewall -eq $null){$psfirewall = "No firewall rules found"}






Write-Output "Checking for events........."





# ------------------ Check Events Logs ---------------------

# Module Logging EventID 4103
#Script Block EventID 4105 and 4106



$4103 = Get-WinEvent -LogName 'Microsoft-Windows-Powershell/Operational'| Where-Object {$_.ID -eq 4103} | Select-Object -First 1
$410456 = Get-WinEvent -LogName 'Microsoft-Windows-Powershell/Operational'| Where-Object {$_.ID -eq 4104 -or $_.ID -eq 4105 -or $_.ID -eq 4106} | Select-Object -First 1

if ($4103 -ne $null){$4103 = "Module logging events found (EventID: 4103)"}else{$4103 = "Module logging events not found (EventID: 4103)"}
if ($410456 -ne $null){$410456 = "ScriptBlock logging events found (EventID: 4104, 4105, 4106)"}else{$410456 = "ScriptBlock Logging events not found (EventID: 4104, 4105, 4106)"}

$eventfind = @()
$eventfind += $4103
$eventfind += $410456



Write-Output "creating the report...."



# HTML

$report = "C:\temp\"

 Remove-Item -Path "C:\temp\report.html" -Force
 New-Item -Path $report -Name "report.html"

 $report = "C:\temp\report.html"

   $htmlstart = @('

<!DOCTYPE html>
<html lang="en">
<head>
<title>PSWatcher Report</title>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
/* Style the body */
body {
  font-family: Arial;
  margin: 0;
}

/* Header/Logo Title */
.header {
  padding: 60px;
  text-align: center;
  background: #2AABE5;
  color: white;
  font-size: 30px;
}

/* Page Content */
.content {padding:20px;}

.scan {
    padding:20px;
    line-height: normal;
}

</style>
</head>
<body>

<div class="header">
  <h1>BlueChecker Report</h1>

</div>

<div class="content">
  <h1>Your Results</h1>
  <p><b>Checking Client Status:</b></p>
  

  <p>

    
   ')

   

   Add-Content -Path $report -Value $htmlstart

   
   foreach ($l in $status){$l = "<br>"+ $l
   Add-Content -Path $report -Value $l}

   $header = @('
   
   <br><br>
  <p><b>Checking For Downgrade:</b></p>

   ')

   Add-Content -Path $report -Value $header

   if ($dwngrdhistory -eq $true){

   Add-Content -Path $report -Value "<p> Evidence Found: </p>"

   foreach ($l in $downgradedcheck){$l = "<br>"+ $l
   Add-Content -Path $report -Value $l}
   
   }else{Add-Content -Path $report -Value $downgradedcheck}







   $header = @('
   
   <br><br>
  <p><b>Checking Auditing Settings : </b></p>

   ')

   Add-Content -Path $report -Value $header
   
   Add-Content -Path $report -Value $moduleresults

   Add-Content -path $report -Value $EBS

   Add-Content -path $report -Value $PCIC








     $header = @('
   
   <br>
  <p><b>Checking history for suspicious commands : </b></p>

   ')

   Add-Content -Path $report -Value $header


   foreach ($l in $Checkforkeywords){$l = "<br>"+ $l
   Add-Content -Path $report -Value $l}




   
      $header = @('
   
   <br><br>
  <p><b>Checking Firewall Rules : </b></p>

   ')

   Add-Content -Path $report -Value $header


   foreach ($l in $psfirewall){$l = "<br>"+ $l
   Add-Content -Path $report -Value $l}

   





      $header = @('
   
   <br><br>
  <p><b>Checking Events : </b></p>

   ')

   Add-Content -Path $report -Value $header

  
   foreach ($l in $eventfind){$l = "<br>"+ $l
   Add-Content -Path $report -Value $l}

   

   








    
    $htmlend = @('
    </p>

    <p>For more information, visit: <a href="https://securethelogs.com">Securethelogs.com</a></p>

    </div>
    
    
    </body>
    </html>

    ')
    
    Add-Content -Path $report -Value $htmlend

    Invoke-Item $report
