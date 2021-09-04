<#
	Appname:	Microsoft Update Downloader
	Author:		NEONFLOPPY
	Website:	http://neonfloppy.sytes.net
	Version:	Release 1.00 - Published: 12 Feb 2021
	License:    Unlicense (https://unlicense.org)
#>

$param = @($args[0],$args[1],$args[2],$args[3],$args[4],$args[5])

function MSUpdate-DL-Batch
{
	$list = Get-Content .\MSUpdate-DL-Batch-List.txt
	$elements = $list.length
	$count = 0
	
	# show brief introductory help about this program
	foreach ($parameter in $param) {
		if (($parameter -eq "-help")) {
			Write-Host "Use the following syntax:" -ForegroundColor Cyan
			Write-Host "MSUpdate-DL-Batch.ps1 [OPTIONS]`r`n" -ForegroundColor Magenta
			Write-Host "Valid option arguments are the following:"
			Write-Host "    -details       (download details for each update in html format)" -ForegroundColor White
			Write-Host "    -architecture  (filter updates by processor/os architecture)" -ForegroundColor White
			Write-Host "    -ntversion     (filter updates by operating system version)" -ForegroundColor White
			Write-Host "    -language      (filter updates by language)" -ForegroundColor White
			Write-Host "    -output        (specify download directory)" -ForegroundColor White
			Write-Host "    -version       (display program version)" -ForegroundColor White
			Write-Host "    -help          (display help introduction)" -ForegroundColor White
			Write-Host "    -log           (enable debug file logging)" -ForegroundColor White
			Write-Host "For a complete operation guide, refer to documentation.htm`r`n" -ForegroundColor White
			return
		}
		if ($parameter -eq "-version") {
			Write-Host "Microsoft Update Downloader" -ForegroundColor Cyan
			Write-Host "Release 1.00 - 12 Feb 2021" -ForegroundColor Magenta
			Write-Host "by NEONFLOPPY. http://neonfloppy.sytes.net`r`n" -ForegroundColor White
			Write-Host "This project has been published under the Unlincense license. https://unlicense.org" -ForegroundColor White
			Write-Host "This is free software: you are free to change and redistribute it`r`n" -ForegroundColor White
			return
		}
		if ($parameter -match "-timeout=") {
			$time = $parameter.Substring(9)
			$timeout = $true
		}
	}
	
	foreach ($line in $list) {
		$count++
		Write-Host "[INFO] Update $count of $elements" -ForegroundColor Green -BackgroundColor Black
		.\MSUpdate-DL.ps1 $line $param[0] $param[1] $param[2] $param[3] $param[4] $param[5]
		
		if ($timeout -eq $true) {
			Write-Host "[INFO] Timeout for $time seconds" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s $time
		}
	}
}
MSUpdate-DL-Batch