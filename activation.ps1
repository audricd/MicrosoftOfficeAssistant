$date = Get-Date -format dd.MM.yyy-HH.mm

If(-not(Test-Path -Path logs))
 {
     Write-Output "Creating logs folder"
	 New-Item -Path logs -Type Directory -Force | Out-Null
  }
Start-Transcript -Path logs\log$date.txt  | Out-Null

Write-Output "Microsoft Office Assistant Activator part v 0.2.
--------------------------------------------
This tool only works if registry keys for office are found on 
HKLM:\SOFTWARE\Microsoft\Office\YourVersionOfOffice\Common\InstallRoot\Path (MSI) or
HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\InstallationPath (C2R)

Detection for other setups are planned to be added in the future.
--------------------------------------------
Consider this:
Office has two version numbers, the long/year and the short/version. They go as follow:
2007 = 13.0
2010 = 14.0
2013 = 15.0
2016 = 16.0
"
$ErrorActionPreference= 'silentlycontinue'
$Oversion = Read-Host -Prompt 'Which version would you like to activate? Write the SHORT version number'
$Oversionintegrer = $Oversion.TrimEnd(".0")

$installrootmsikey = "HKLM:\SOFTWARE\Microsoft\Office\$Oversion\Common\InstallRoot"
$installrootc2rkey = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"


$installrootmsikeytp = Test-Path "HKLM:\SOFTWARE\Microsoft\Office\$Oversion\Common\InstallRoot"
$installrootc2rkeytp = Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"


$installrootmsi = (gp "HKLM:\SOFTWARE\Microsoft\Office\$Oversion\Common\InstallRoot").Path.TrimEnd("\")
$installrootc2r = (gp "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration").InstallationPath

if ($installrootmsikeytp -eq "True") {$installdir = "$installrootmsi"}
if ($installrootc2rkeytp -eq "True") {$installdir = "$installrootc2r" + "\" + "Office$Oversionintegrer"}


$ospp = "$installdir" + "\" + "ospp.vbs"
$findospp = Test-Path $ospp
try{
if ($findospp -eq "True"){}
	}
catch{
	$installdir = Read-Host -Prompt "ospp was not found in. Please specify the location:"
	}
	finally{
		Write-Output "found $ospp"
	}



Set-Location "$installdir"
$cmd = "cscript"
$status = "/dstatus"
do {
  [int]$userMenuChoice = 0
  while ( $userMenuChoice -lt 1 -or $userMenuChoice -gt 8) {
	Write-Host "1. Check status"
	Write-Host "2. Install key"
	Write-Host "3. Remove key"
	Write-Host "4. Force activation"
	Write-Host "5. Show computer CMID"
	Write-Host "6. Set KMS host (fqdn)"
	Write-Host "7. Set KMS port (no need to change if KMS server is running on default port)"
	Write-Host "8. Exit"

    [int]$userMenuChoice = Read-Host "Please choose an option"

    switch ($userMenuChoice) {
	  1{& $cmd $ospp $status}
	  2{$installkey = Read-Host -Prompt "Input the key you would like to install" | & $cmd $ospp  /inpkey:$installkey}
	  3{$removekey = Read-Host -Prompt "Input the last 5 characters of the key you want to remove" | & $cmd $ospp  /unpkey:$removekey}
	  4{& $cmd $ospp /act}
	  5{& $cmd $ospp /dcmid}
	  6{$kmsfqdn = Read-Host -Prompt "Input the FQDN of your KMS, example: kms.domain.com" | & $cmd $ospp /sethst:$kmsfqdn}
	  7{$kmsport = Read-Host -Prompt "Input the port used by your KMS. Default is 1688" | & $cmd $ospp /setprt:$kmsport}
}
}
	}

 while	 ( $userMenuChoice -ne 8 )
 Stop-Transcript