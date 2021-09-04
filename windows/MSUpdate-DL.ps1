<#
	Appname:	Microsoft Update Downloader
	Author:		NEONFLOPPY
	Website:	http://neonfloppy.sytes.net
	Version:	Release 1.00 - Published: 12 Feb 2021
	License:    Unlicense (https://unlicense.org)
#>

$param = @($args[0],$args[1],$args[2],$args[3],$args[4],$args[5],$args[6])

function MSUpdate-DL
{
	$request = $param[0]
	$basedir = "C:\msupdate-dl\"
	$outdir = -join($basedir, $request, "\")
	$logging = $false
	$details = $false
	$arch32 = $true
	$arch64 = $true
	$itanium = $true
	$ntversion = "all"
	$langs = "all"
	
	# check and recognize all specified arguments
	foreach ($parameter in $param) {
		if ($parameter -match "-details") {
			$details = $true
		}
		
		if ($parameter -match "-architecture=") {
			$arch32 = $false 
			$arch64 = $false 
			$itanium = $false
			
			if ($parameter -match "32") {
				$arch32 = $true
			}
			if ($parameter -match "64") {
				$arch64 = $true
			}
			if ($parameter -match "itanium") {
				$itanium = $true
			}
			if ($parameter -match "all") {
				$arch32 = $true 
				$arch64 = $true 
				$itanium = $true
			}
			if (($arch32 -eq $false) -and ($arch64 -eq $false) -and ($itanium -eq $false)) {
				Write-Host "[ERROR] Arguments were specified incorrectly" -ForegroundColor Red -BackgroundColor Black
				return
			}
		}
		
		if ($parameter -match "-ntversion=") {
			if ($parameter -match "-ntversion=all") {
				$ntversion = "all"
			} else {
				$ntversion = $parameter
			}
		}
		
		if ($parameter -match "-language=") {
			if ($parameter -match "-language=all") {
				$langs = "all"
			} else {
				$langs = $parameter
			}
		}
		
		if ($parameter -match "-output=") {
			$basedir = $parameter.substring(8)
			$outdir = -join($basedir, $request, "\")
		}
		
		if ($parameter -match "-log") {
			$now = (Get-Date -Format "(dd-MM-yyyy)").ToString()
			$logfile = -join("logs\log-", $now, ".txt") 
			$logging = $true
			
			# logs are stored on a different folder. create the log folder and file (if needed)
			if (!(Test-Path "logs")) {
				New-Item -ItemType Directory -Force -Path "logs" | Out-Null
			}
			if (!(Test-Path $logfile)) {
				New-Item -Name $logfile -ItemType File | Out-Null
			}
		}
	}
	
	# show brief introductory help about this program
	if (($request -eq $null) -or ($request -eq "-help")) {
		Write-Host "Missing arguments. Use the following syntax:" -ForegroundColor Cyan
		Write-Host "MSUpdate-DL.ps1 [REQUEST] [OPTIONS]`r`n" -ForegroundColor Magenta
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
	
	# show version information
	if ($request -eq "-version") {
		Write-Host "Microsoft Update Downloader" -ForegroundColor Cyan
		Write-Host "Release 1.00 - 12 Feb 2021" -ForegroundColor Magenta
		Write-Host "by NEONFLOPPY. http://neonfloppy.sytes.net`r`n" -ForegroundColor White
		Write-Host "This project has been published under the Unlincense license. https://unlicense.org" -ForegroundColor White
		Write-Host "This is free software: you are free to change and redistribute it`r`n" -ForegroundColor White
		return
	}
	
	# this block shows a screen message relative to the process and stores it to the log file if logging is enabled by the user
	$msg = "[INFO] Downloading $request"
	Write-Host $msg -ForegroundColor Green -BackgroundColor Black
	if ($logging -eq $true) {
		$out = -join((Get-Date).ToString(), "  ", $msg)
		Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
	}
	
	# send query to the update catalog website for specified request
	$kbObj = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$request"
	$Available_KBIDs = $kbObj.InputFields |
	Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
	Select-Object -ExpandProperty ID
	
	# check if the query returns any results
	if ($Available_KBIDs -eq $null) {
		$msg = "[INFO] $request is not present on the Microsoft Update Catalog"
		Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
		if ($logging -eq $true) {
			$out = -join((Get-Date).ToString(), "  ", $msg)
			Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
		}
		return
		
	} else {
		# create the directories where the downloaded files will be placed (if needed)
		if (!(Test-Path $basedir)) {
			New-Item -ItemType Directory -Force -Path $basedir | Out-Null
		}
		if (!(Test-Path $outdir)) {
			New-Item -ItemType Directory -Force -Path $outdir | Out-Null
		}
	}
	
	# download a copy of the query page in case the user specified the details argument
	if ($details -eq $true) {
		$infodir = -join($outdir, "info\")
		$infofile = -join($infodir, "Microsoft Update Catalog.htm")
		
		# html details are stored on a different folder. create the folder (if needed)
		if (!(Test-Path $infodir)) {
			New-Item -ItemType Directory -Force -Path $infodir | Out-Null
		}
		if ($logging -eq $true) {
			$out = -join((Get-Date).ToString(), "  ", "[INFO] Downloading html details page: https://www.catalog.update.microsoft.com/Search.aspx?q=", $request)
			Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
		}
		
		# call the HTTP retriever. wget used in this project
		wget\wget.exe https://www.catalog.update.microsoft.com/Search.aspx?q=$request --user-agent="Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; InfoPath.2)" --output-document $infofile
		
		# resources inside the html file point to external websites. we want a fully offline copy. resource addresses are changed to local ones
		(Get-Content -Path $infofile) -replace 'href="Style/catalog.css', 'href="html_files/styles/catalog.css' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="/ScriptResource.axd', 'src="html_files/scripts/ScriptResource.js' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="SiteConstants.', 'src="html_files/scripts/SiteConstants.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Script/CommonTypes.', 'src="html_files/scripts/CommonTypes.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Script/DownloadBasket.', 'src="html_files/scripts/DownloadBasket.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Script/MasterComponents.', 'src="html_files/scripts/MasterComponents.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Script/ScopedView.', 'src="html_files/scripts/ScopedView.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/decor_BigInformation.', 'src="html_files/images/decor_BigInformation.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/button_Xml.', 'src="html_files/images/button_Xml.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/button_PreviousArrow', 'src="html_files/images/button_PreviousArrow' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/button_NextArrow', 'src="html_files/images/button_NextArrow' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/bg_SearchGlow', 'src="html_files/images/bg_SearchGlow' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/spacer.', 'src="html_files/images/spacer.' | 
		Out-File -encoding UTF8 $infofile
		(Get-Content -Path $infofile) -replace 'src="Images/button_Sort', 'src="html_files/images/button_Sort' | 
		Out-File -encoding UTF8 $infofile
		
		# hotfix for the download buttons (style shows incorrectly. this error is also present in the online version)
		(Get-Content -Path $infofile) -replace 'flatBlueButtonDownload', 'flatLightBlueButton' | 
		Out-File -encoding UTF8 $infofile
		
		# hotfix specifically for IE6 (table header shows incorrectly. this error is also present in the online version)
		(Get-Content -Path $infofile) -replace 'ResultsHeaderTD resultsBorderRight resultsDateWidth', 'ResultsHeaderTD resultsBorderRight' | 
		Out-File -encoding UTF8 $infofile
		
		# copy the needed resources to target folder
		$htmldir = -join($infodir, "html_files\")
		if (!(Test-Path $htmldir)) {
			Copy-Item "html_files\" -Destination $infodir -Recurse
		}
	}

	# analyze the query and extract element identifier codes
	$kbGUIDs = $kbObj.Links |
	Where-Object ID -match '_link' |
	ForEach-Object { $_.id.replace('_link', '') } |
	Where-Object { $_ -in $Available_KBIDs }

	# go across every element of the query
	$kbNumber = 0
	foreach ($kbGUID in $kbGUIDs) {
		# download a copy of the element page in case the user specified the details argument
		if ($details -eq $true) {
			$kbNumber++
			$infofile = -join($infodir, "Update Details ", $kbNumber, ".htm")
			
			# on rare occasions, downloaded page file is erroneous. the following loop downloads the page and checks if it's correct. it will retry the download if it isn't
			$tries = 0
			do {
				$tries++
				if ($tries -ge 2) {	
					$msg = "[INFO] Details page download has failed. Retrying..."
					Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
					if ($logging -eq $true) {
						$out = -join((Get-Date).ToString(), "  ", $msg)
						Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
					}
				}
				
				if ($logging -eq $true) {
					$out = -join((Get-Date).ToString(), "  ", "[INFO] Downloading html details page: https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=", $kbGUID)
					Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
				}
				
				# call the HTTP retriever. wget used in this project
				wget\wget.exe https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$kbGUID --user-agent="Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.1; WOW64; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; Media Center PC 6.0; .NET4.0C; InfoPath.2)" --output-document $infofile
			} while (Select-String -Path $infofile -Pattern "The website has encountered a problem")
			
			if ($tries -ge 2) {	
				$msg = "[INFO] Details page finally downloaded correctly"
				Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
				if ($logging -eq $true) {
					$out = -join((Get-Date).ToString(), "  ", $msg)
					Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
				}
			}
			$tries = 0

			# resources inside the html file point to external websites. we want a fully offline copy. resource addresses are changed to local ones
			(Get-Content -Path $infofile) -replace 'href="Style/catalog.css', 'href="html_files/styles/catalog.css' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="/ScriptResource.axd', 'src="html_files/scripts/ScriptResource.js' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="SiteConstants.', 'src="html_files/scripts/SiteConstants.' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="Script/CommonTypes.', 'src="html_files/scripts/CommonTypes.' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="Script/DownloadBasket.', 'src="html_files/scripts/DownloadBasket.' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="Script/MasterComponents.', 'src="html_files/scripts/MasterComponents.' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="Script/ScopedView.', 'src="html_files/scripts/ScopedView.' | 
			Out-File -encoding UTF8 $infofile
			(Get-Content -Path $infofile) -replace 'src="Images/decor_BigInformation.', 'src="html_files/images/decor_BigInformation.' | 
			Out-File -encoding UTF8 $infofile
			
			# one script redirects to the error page if the html file is not in the right location. disabling it
			(Get-Content -Path $infofile) -replace 'window.location.href.toLowerCase', 'void' | 
			Out-File -encoding UTF8 $infofile
		}
		
		# do a query to the download dialog page with current element identifier code
		$Post = @{ size = 0; updateID = $kbGUID; uidInfo = $kbGUID } | ConvertTo-Json -Compress
		$PostBody = @{ updateIDs = "[$Post]" }
		Invoke-WebRequest -Uri 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $postBody |
		Select-Object -ExpandProperty Content |
		# select all the URLs present on query result
		Select-String -AllMatches -Pattern "(http[s]?)(:\/\/)([^\s,]+)(?=')" |
		Select-Object -Unique |
		# checks will be done with each and every URL found
		ForEach-Object { 
			foreach ($line in $_.matches.value) {
				# check language of URL referenced file. the condition is met for each language specified by the user (all by default)
				$condition1 = $true
				if ($langs -ne "all") {
					$condition1 = $false
					switch ($line) {
						{$_ -match '-jpn_' } {
							if ($langs -match "jpn") { $condition1 = $true }		# japanase
						}
						{$_ -match '-chs_' } {
							if ($langs -match "chs") { $condition1 = $true }		# chinese (trad)
						} 
						{$_ -match '-cht_' } {
							if ($langs -match "cht") { $condition1 = $true }		# chinese (simp)
						} 
						{$_ -match '-deu_' } {
							if ($langs -match "deu") { $condition1 = $true }		# german
						} 
						{$_ -match '-kor_' } {
							if ($langs -match "kor") { $condition1 = $true }		# korean
						} 
						{$_ -match '-ara_' } {
							if ($langs -match "ara") { $condition1 = $true }		# arabic
						} 
						{$_ -match '-fin_' } {
							if ($langs -match "fin") { $condition1 = $true }		# finnish
						} 
						{$_ -match '-plk_' } {
							if ($langs -match "plk") { $condition1 = $true }		# polish
						} 
						{$_ -match '-nor_' } {
							if ($langs -match "nor") { $condition1 = $true }		# norwegian
						} 
						{$_ -match '-ptb_' } {
							if ($langs -match "ptb") { $condition1 = $true }		# portuguese (por)
						} 
						{$_ -match '-ptg_' } {
							if ($langs -match "ptg") { $condition1 = $true }		# portuguese (bra)
						} 
						{$_ -match '-fra_' } {
							if ($langs -match "fra") { $condition1 = $true }		# french
						} 
						{$_ -match '-dan_' } {
							if ($langs -match "dan") { $condition1 = $true }		# danish
						} 
						{$_ -match '-ell_' } {
							if ($langs -match "ell") { $condition1 = $true }		# greek
						} 
						{$_ -match '-csy_' } {
							if ($langs -match "csy") { $condition1 = $true }		# czech
						}
						{$_ -match '-trk_' } {
							if ($langs -match "trk") { $condition1 = $true }		# turkish
						} 
						{$_ -match '-rus_' } {
							if ($langs -match "rus") { $condition1 = $true }		# russian
						}
						{$_ -match '-esn_' } {
							if ($langs -match "esn") { $condition1 = $true }		# spanish
						}
						{$_ -match '-ita_' } {
							if ($langs -match "ita") { $condition1 = $true }		# italian
						} 
						{$_ -match '-hun_' } {
							if ($langs -match "hun") { $condition1 = $true }		# hungarian
						}
						{$_ -match '-nld_' } {
							if ($langs -match "nld") { $condition1 = $true }		# dutch
						}
						{$_ -match '-heb_' } {
							if ($langs -match "heb") { $condition1 = $true }		# hebrew
						} 
						{$_ -match '-sve_' } {
							if ($langs -match "sve") { $condition1 = $true }		# swedish
						} 
						{$_ -match '-enu_' } {
							if ($langs -match "enu") { $condition1 = $true }		# english
						} 
						{$_ -match '-eng_' } {
							if ($langs -match "enu") { $condition1 = $true }		# english
						} 
						Default {
							$condition1 = $true		# language independent files
						}
					}
				}
				
				# check operating system version of URL referenced file. the condition is met for each NT version specified by the user (all by default)
				$condition2 = $true
				if ($ntversion -ne "all") {
					$condition2 = $false
					switch ($line) {
						{$_ -match 'windows2000' } {
							if ($ntversion -match "50") { $condition2 = $true }		# NT 5.0 (2000)
						}
						{$_ -match 'windowsxp' } {
							if ($ntversion -match "51") { $condition2 = $true }		# NT 5.1 (xp)
						}
						{$_ -match 'windowsserver2003' } {
							if ($ntversion -match "52") { $condition2 = $true }		# NT 5.2 (2003)
						}
						{$_ -match 'windows6.0' } {
							if ($ntversion -match "60") { $condition2 = $true }		# NT 6.0 (vista)
						}
						{$_ -match 'windows6.1' } {
							if ($ntversion -match "61") { $condition2 = $true }		# NT 6.1 (7)
						}
						{$_ -match 'windows6.2' } {
							if ($ntversion -match "62") { $condition2 = $true }		# NT 6.2 (8)
						}
						{$_ -match 'windows6.3' } {
							if ($ntversion -match "63") { $condition2 = $true }		# NT 6.3 (8.1)
						}
						Default {
							$condition2 = $true		# version independent files
						}
					}
				}
				
				# check architecture of URL referenced file. the condition is met for each architecture specified by the user (all by default)
				$condition3 = $true
				if (($arch32 -eq $false) -or ($arch64 -eq $false) -or ($itanium -eq $false)) {
					$condition3 = $false
					switch ($line) {
						{$_ -match '-x64' } {
							if ($arch64 -eq $true) { $condition3 = $true }			# 64-bit
						}
						{$_ -match '-x86' } {
							if ($arch32 -eq $true) { $condition3 = $true }			# 32-bit
						}
						{$_ -match '-ia64' } {
							if ($itanium -eq $true) { $condition3 = $true }			# itanium
						}
						Default {
							$condition3 = $true		# architecture independent files
						}
					}
				}
				
				# if all the aforementioned conditions are met, proceed with the file download. otherwise, the file is skipped
				if (($condition1 -eq $true) -and ($condition2 -eq $true) -and ($condition3 -eq $true)) {
					if ($logging -eq $true) {
						$out = -join((Get-Date).ToString(), "  ", "[INFO] Downloading file: ", $line)
						Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
					}
					# call the HTTP retriever. wget used in this project
					wget\wget.exe $line --no-clobber --directory-prefix=$outdir
				}
			}
		}
	}
}
MSUpdate-DL