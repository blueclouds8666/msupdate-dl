<#
	Appname:	Microsoft Update Downloader
	Author:		NEONFLOPPY
	Website:	http://neonfloppy.sytes.net
	Version:	Release 1.03 - Published: 26 Jan 2024
	License:    Unlicense (https://unlicense.org)
#>

$param = @($args[0],$args[1],$args[2],$args[3],$args[4],$args[5],$args[6])

function MSUpdate-DL-Batch
{
	$list = Get-Content .\MSUpdate-DL-Batch-List.txt
	$elements = ($list | Measure-Object -Line).Lines
	$count = 0
	
	# show brief introductory help about this program
	foreach ($parameter in $param) {
		if (($parameter -eq "--help")) {
			Write-Host "Missing arguments. Use the following syntax:" -ForegroundColor Cyan -BackgroundColor Black
			Write-Host "MSUpdate-DL.ps1 [REQUEST] [OPTIONS]" -ForegroundColor Magenta -BackgroundColor Black
			Write-Host "`r`nValid option arguments are the following:"
			Write-Host "    --details         (download details for each update in html format)" -ForegroundColor White
			Write-Host "    --architecture    (filter updates by processor/os architecture)" -ForegroundColor White
			Write-Host "    --ntversion       (filter updates by operating system version)" -ForegroundColor White
			Write-Host "    --language        (filter updates by language)" -ForegroundColor White
			Write-Host "    --double-check    (prevent incorrect languages from downloading)" -ForegroundColor White
			Write-Host "    --retry-policy    (specify the policy for retrying)" -ForegroundColor White
			Write-Host "    --retry-number    (specify the number of retries)" -ForegroundColor White
			Write-Host "    --retry-timeout   (specify the timeout between retries)" -ForegroundColor White
			Write-Host "    --output          (specify download directory)" -ForegroundColor White
			Write-Host "    --version         (display program version)" -ForegroundColor White
			Write-Host "    --help            (display help introduction)" -ForegroundColor White
			Write-Host "    --log             (enable debug file logging)" -ForegroundColor White
			Write-Host "    --timeout         (enable timeout between downloads)" -ForegroundColor White
			Write-Host "For a complete operation guide, refer to documentation.htm`r`n" -ForegroundColor White
			return
		}
		if ($parameter -eq "--version") {
			Write-Host "Microsoft Update Downloader" -ForegroundColor Cyan -BackgroundColor Black
			Write-Host "Release 1.03 - 26 Jan 2024" -ForegroundColor Magenta -BackgroundColor Black
			Write-Host "by NEONFLOPPY. http://neonfloppy.sytes.net`r`n" -ForegroundColor White
			Write-Host "This project has been published under the Unlincense license. https://unlicense.org" -ForegroundColor White
			Write-Host "This is free software: you are free to change and redistribute it`r`n" -ForegroundColor White
			return
		}
		if ($parameter -match "--timeout=") {
			$time = $parameter.Substring(9)
			$timeout = $true
		}
	}
	
	foreach ($line in $list) {
		$count++
		Write-Host "[INFO] Update $count of $elements" -ForegroundColor Green -BackgroundColor Black
		.\MSUpdate-DL.ps1 $line $param[0] $param[1] $param[2] $param[3] $param[4] $param[5] $param[6]
		
		if ($timeout -eq $true) {
			Write-Host "[INFO] Timeout between updates for $time seconds" -ForegroundColor Green -BackgroundColor Black
			Start-Sleep -s $time
		}
	}
}
MSUpdate-DL-Batch