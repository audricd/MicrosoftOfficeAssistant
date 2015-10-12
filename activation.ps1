$date = Get-Date -format dd.MM.yyy-HH.mm

If(-not(Test-Path -Path logs))
 {
     Write-Output "Creating logs folder"
	 New-Item -Path logs -Type Directory -Force | Out-Null
  }
Start-Transcript -Path logs\log$date.txt  | Out-Null

Write-Output "Microsoft Office Assistant Activator part v 0.1.
--------------------------------------------
Consider this:
Office has two version numbers, the long/year and the short/version. They go as follow:
2007 = 13.0
2010 = 14.0
2013 = 15.0
2016 = 16.0
"
$Oversion = Read-Host -Prompt 'Which version would you like to activate? Write the SHORT version number'
$installroot = (gp "HKLM:\SOFTWARE\Microsoft\Office\$Oversion\Common\InstallRoot").Path
$osppautomode = Test-Path "$installroot\ospp.vbs"
if ($osppautomode -eq "True"){Write-Output "ospp.vbs was found at $installroot"}
else {$installroot = Read-Host -Prompt "ospp was not found in. Please specify the location:"}
Set-Location $installroot
$ospp = "ospp.vbs"
$cmd = "cscript"
$status = "/dstatus"
do {
  [int]$userMenuChoice = 0
  while ( $userMenuChoice -lt 1 -or $userMenuChoice -gt 2) {
	Write-Host "1. Check status"
    Write-Host "2. Exit"

    [int]$userMenuChoice = Read-Host "Please choose an option"

    switch ($userMenuChoice) {
	  1{invoke-expression "$cmd $ospp $status"}
}
}
	}

 while	 ( $userMenuChoice -ne 2 )
 Stop-Transcript