Write-Output "Welcome to Microsoft Office Assistant v0.1, this part is for diagnotics and removals of Office installation.
ROIScan is a script that will gather all the information about your Office Installation. After it ran, a notepad will open. Please save it.
OffScrub are complete uninstalls of the selected product. If you have side by side installations issues for 2013(Office 15), run option 5 and 6. Then install just one version, either MSI or Click2Run."
$location = “$Server $FolderName”
$roiscan = “cmd /C cscript $PSScriptRoot\scripts\roiscan.vbs $location”
$offscrub03 = “cmd /C cscript $PSScriptRoot\scripts\OffScrub03.vbs $location”
$offscrub07 = “cmd /C cscript $PSScriptRoot\scripts\OffScrub07.vbs $location”
$offscrub10 = “cmd /C cscript $PSScriptRoot\scripts\OffScrub10.vbs $location”
$offscrubO15msi = “cmd /C cscript $PSScriptRoot\scripts\OffScrub_O15msi.vbs $location”
$offscrubO15c2r = “cmd /C cscript $PSScriptRoot\scripts\OffScrubC2R.vbs $location”
$offscrubO16msi = “cmd /C cscript $PSScriptRoot\scripts\OffScrub_O16msi.vbs $location”
$offscrubO16c2r = “cmd /C cscript $PSScriptRoot\scripts\OffScrubO16C2R.vbs $location”
	

do {
  [int]$userMenuChoice = 0
  while ( $userMenuChoice -lt 1 -or $userMenuChoice -gt 9) {
	Write-Host "------------Diagnostics-----------"
    Write-Host "1. Run ROIScan"
	Write-Host "-------------OffScrubs------------"
    Write-Host "2. OffScrub Office 2003"
    Write-Host "3. OffScrub Office 2007"
	Write-Host "4. OffScrub Office 2010"
	Write-Host "5. OffScrub Office 2013 MSI"
	Write-Host "6. OffScrub Office 2013 Click2Run"
	Write-Host "7. OffScrub Office 2016 MSI"
	Write-Host "8. OffScrub Office 2016 Click2Run"
	Write-Host "----------------------------------"
	Write-Host "9. Close and exit"

    [int]$userMenuChoice = Read-Host "Please choose an option"

    switch ($userMenuChoice) {
	  1{invoke-expression $roiscan
	  Write-Output "Save the notepad file that just opened"}
	  2{Invoke-Expression $offscrub03}
	  3{Invoke-Expression $offscrub07}
	  4{Invoke-Expression $offscrub10}
	  5{Invoke-Expression $offscrubO15msi}
	  6{Invoke-Expression $offscrubO15c2r}
	  7{Invoke-Expression $offscrubO16msi}
	  8{Invoke-Expression $offscrubO16c2r}
}
}
	}
 while	 ( $userMenuChoice -ne 9 )
	