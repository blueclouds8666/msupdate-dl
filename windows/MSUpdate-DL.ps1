<#
	Appname:	Microsoft Update Downloader
	Author:		NEONFLOPPY
	Website:	http://neonfloppy.sytes.net
	Version:	Release 1.03 - Published: 26 Jan 2024
	License:    Unlicense (https://unlicense.org)
#>

# force enable TLS 1.2 connections in PowerShell 3.0 and later
[System.Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([System.Net.SecurityProtocolType], 3072);

$param = @($args[0],$args[1],$args[2],$args[3],$args[4],$args[5],$args[6],$args[7])

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
	$retryPolicy = "normal"
	$retryNumber = "infinity"
	$retryTimeout = 2
	$retryActive = $false
	$retryTries = 0
	$languageAny = $true
	$languageCheck = $false
	$languageCodeISO = New-Object System.Collections.ArrayList
	$languageCodeMicrosoft = New-Object System.Collections.ArrayList
	
	# check and recognize all specified arguments
	foreach ($parameter in $param) {
		if ($parameter -match "--details") {
			$details = $true
		}
		
		if ($parameter -match "--double-check") {
			$languageCheck = $true
		}
		
		if ($parameter -match "--retry-policy=") {
			if ($parameter -match "normal") {
				$retryPolicy = "normal"
			}
			if ($parameter -match "aggressive") {
				$retryPolicy = "aggressive"
			}
		}
		
		if ($parameter -match "--retry-number=") {
			if ($parameter -match '\d+') {
				$retryNumber = $Matches[0]
			}
			if ($parameter -match "infinity") {
				$retryNumber = "infinity"
			}
		}
		
		if ($parameter -match "--retry-timeout=") {
			if ($parameter -match '\d+') {
				$retryTimeout = $Matches[0]
			}
		}
		
		if ($parameter -match "--architecture=") {
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
		
		if ($parameter -match "--ntversion=") {
			if ($parameter -match "-ntversion=all") {
				$ntversion = "all"
			} else {
				$ntversion = $parameter
			}
		}
		
		if ($parameter -match "--language=") {
			$languageAny = $false
			if (($parameter -match "all")) {
				$languageAny = $true  }
			if (($parameter -match "jpn") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("jpn");  [void]$languageCodeISO.Add("ja, ja-JP")  }		# 1.  japanase
			if (($parameter -match "chs") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("chs");  [void]$languageCodeISO.Add("zh, zh-CN")  }		# 2.  chinese (simp)
			if (($parameter -match "cht") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("cht");  [void]$languageCodeISO.Add("zh, zh-TW")  }		# 3.  chinese (trad)
			if (($parameter -match "deu") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("deu");  [void]$languageCodeISO.Add("de, de-DE")  }		# 4.  german
			if (($parameter -match "kor") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("kor");  [void]$languageCodeISO.Add("ko, ko-KR")  }		# 5.  korean
			if (($parameter -match "ara") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("ara");  [void]$languageCodeISO.Add("ar, ar-AE")  }		# 6.  arabic
			if (($parameter -match "fin") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("fin");  [void]$languageCodeISO.Add("fi, fi-FI")  }		# 7.  finnish
			if (($parameter -match "plk") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("plk");  [void]$languageCodeISO.Add("pl, pl-PL")  }		# 8.  polish
			if (($parameter -match "nor") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("nor");  [void]$languageCodeISO.Add("no, no-NO")  }		# 9.  norwegian
			if (($parameter -match "ptb") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("ptb");  [void]$languageCodeISO.Add("pt, pt-BR")  }		# 10. portuguese (bra)
			if (($parameter -match "ptg") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("ptg");  [void]$languageCodeISO.Add("pt, pt-PT")  }		# 11. portuguese (por)
			if (($parameter -match "fra") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("fra");  [void]$languageCodeISO.Add("fr, fr-FR")  }		# 12. french
			if (($parameter -match "dan") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("dan");  [void]$languageCodeISO.Add("da, da-DK")  }		# 13. danish
			if (($parameter -match "ell") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("ell");  [void]$languageCodeISO.Add("el, el-GR")  }		# 14. greek
			if (($parameter -match "csy") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("csy");  [void]$languageCodeISO.Add("cs, cs-CZ")  }		# 15. czech
			if (($parameter -match "trk") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("trk");  [void]$languageCodeISO.Add("tr, tr-TR")  }		# 16. turkish
			if (($parameter -match "rus") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("rus");  [void]$languageCodeISO.Add("ru, ru-RU")  }		# 17. russian
			if (($parameter -match "esn") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("esn");  [void]$languageCodeISO.Add("es, es-ES")  }		# 18. spanish
			if (($parameter -match "ita") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("ita");  [void]$languageCodeISO.Add("it, it-IT")  }		# 19. italian
			if (($parameter -match "hun") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("hun");  [void]$languageCodeISO.Add("hu, hu-HU")  }		# 20. hungarian
			if (($parameter -match "nld") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("nld");  [void]$languageCodeISO.Add("nl, nl-NL")  }		# 21. dutch
			if (($parameter -match "heb") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("heb");  [void]$languageCodeISO.Add("he, he-IL")  }		# 22. hebrew
			if (($parameter -match "sve") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("sve");  [void]$languageCodeISO.Add("sv, sv-SE")  }		# 23. swedish
			if (($parameter -match "enu") -or ($parameter -match "all")) {
				[void]$languageCodeMicrosoft.Add("enu");  [void]$languageCodeISO.Add("en, en-US")  }		# 24. english
		}
		
		if ($parameter -match "--output=") {
			$basedir = $parameter.substring(9)
			
			# check if the path ends with a slash
			if ($string -match '\\$') {
				$outdir = -join($basedir, $request, "\")
			
			# if it doesn't, we need to add one
			} else {
				$outdir = -join($basedir, "\", $request, "\")
			}
		}
		
		if ($parameter -match "--log") {
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
	if (($request -eq $null) -or ($request -eq "--help")) {
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
		Write-Host "For a complete operation guide, refer to documentation.htm`r`n" -ForegroundColor White
		return
	}
	
	# show version information
	if ($request -eq "--version") {
		Write-Host "Microsoft Update Downloader" -ForegroundColor Cyan -BackgroundColor Black
		Write-Host "Release 1.03 - 26 Jan 2024" -ForegroundColor Magenta -BackgroundColor Black
		Write-Host "by NEONFLOPPY. http://neonfloppy.sytes.net`r`n" -ForegroundColor White
		Write-Host "This project has been published under the Unlincense license. https://unlicense.org" -ForegroundColor White
		Write-Host "This is free software: you are free to change and redistribute it`r`n" -ForegroundColor White
		return
	}
	
	# show a screen message relative to the process and stores it to the log file if logging is enabled by the user
	$msg = "[INFO] Downloading $request"
	Write-Host $msg -ForegroundColor Green -BackgroundColor Black
	if ($logging -eq $true) {
		$out = -join((Get-Date).ToString(), "  ", $msg)
		Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
	}
	
	# default to english language if it was not explicitly specified
	if ($languageCodeISO.Count -eq 0) {
		$msg = "[INFO] No language(s) specified. Script will download English and Universal only"
		Write-Host $msg -ForegroundColor Green -BackgroundColor Black
		[void]$languageCodeMicrosoft.Add("enu")
		[void]$languageCodeISO.Add("en, en-US")
	}
	
	$retryTries = 1
	$retryActive = $true
	while ($retryActive -eq $true) {
		# send query to the update catalog website for specified request
		$kbObj = Invoke-WebRequest -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$request"
		$Available_KBIDs = $kbObj.InputFields |
		Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
		Select-Object -ExpandProperty ID
		
		# get the html contents of the query, in order to check the actual textual response
		$doc = $kbObj.ParsedHtml
		
		# check if the html page includes kb update results
		if ([regex]::matches($kbObj, "ctl00_catalogBody_noResults").count -ne 0) {
			$div = $true
		} else {
			$div = $null
		}
		
		# check if the html contains an item not found message
		if ($div -ne $null) {
			if ($Available_KBIDs -eq $null) {
				$msg = "[INFO] The server reports that $request is not present on the Microsoft Update Catalog"
				Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
				if ($logging -eq $true) {
					$out = -join((Get-Date).ToString(), "  ", $msg)
					Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
				}
				return
				
			} else {
				$msg = "[INFO] Unexpected server response. This update might not be available on the Microsoft Update Catalog"
				Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
				if ($logging -eq $true) {
					$out = -join((Get-Date).ToString(), "  ", $msg)
					Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
				}
				return
			}
			
		} else {
			if ($Available_KBIDs -eq $null) {
				if (($retryNumber -ne "infinity") -and ($retryTries -ge $retryNumber)) {
					$retryActive = $false
					$msg = "[INFO] The server refuses to report a valid answer. Could not retrieve the update from the Microsoft Update Catalog"
					Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
					if ($logging -eq $true) {
						$out = -join((Get-Date).ToString(), "  ", $msg)
						Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
					}
					
				} else {
					$retryActive = $true
					$msg = "[INFO] The server refuses to report a valid answer. Retry $retryTries of $retryNumber"
					Write-Host $msg -ForegroundColor Yellow -BackgroundColor Black
					if ($logging -eq $true) {
						$out = -join((Get-Date).ToString(), "  ", $msg)
						Add-Content -Force -Encoding UTF8 -Path $logfile -Value $out
					}
					
					if ($retryPolicy -eq "normal") {
						Start-Sleep -Seconds $retryTimeout
					}
					
					$retryTries = $retryTries + 1
				}
			} else {
				$retryActive = $false
			}
		}
	}
	
	# create the directories where the downloaded files will be placed (if needed)
	if (!(Test-Path $basedir)) {
		New-Item -ItemType Directory -Force -Path $basedir | Out-Null
	}
	if (!(Test-Path $outdir)) {
		New-Item -ItemType Directory -Force -Path $outdir | Out-Null
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
	
		# go across every language locale specified so the server can give us the damn files
		for ($i = 0; $i -lt $languageCodeISO.Count; $i++) {
	
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
			
			# prepare the required attributes to do a query to the download dialog page
			$Post = @{ size = 0; updateID = $kbGUID; uidInfo = $kbGUID } | ConvertTo-Json -Compress
			$PostBody = @{ updateIDs = "[$Post]" }		
			$Headers = @{ "Accept-Language" = $languageCodeISO[$i] }
			
			# do a query to the download dialog page with current element identifier code
			Invoke-WebRequest -Uri 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Headers $Headers -Body $postBody |
			Select-Object -ExpandProperty Content |
			# select all the URLs present on query result
			Select-String -AllMatches -Pattern "(http[s]?)(:\/\/)([^\s,]+)(?=')" |
			Select-Object -Unique |
			# checks will be done with each and every URL found
			ForEach-Object {
				foreach ($line in $_.matches.value) {
					# check language of URL referenced file. the condition is met for each language specified by the user (all by default)
					# given that the server will now only give you only the files for only your language, this is now optional
					if (($languageAny -eq $false) -and ($languageCheck -eq $true)) {
						$condition1 = $false
						switch ($line) {
							{$_ -match '-jpn_' } {
								if ($languageCodeMicrosoft[$i] -match "jpn") { $condition1 = $true }  }		# 1.  japanase
							{$_ -match '-chs_' } {
								if ($languageCodeMicrosoft[$i] -match "chs") { $condition1 = $true }  }		# 2.  chinese (simp)
							{$_ -match '-cht_' } {
								if ($languageCodeMicrosoft[$i] -match "cht") { $condition1 = $true }  }		# 3.  chinese (trad)
							{$_ -match '-deu_' } {
								if ($languageCodeMicrosoft[$i] -match "deu") { $condition1 = $true }  }		# 4.  german
							{$_ -match '-kor_' } {
								if ($languageCodeMicrosoft[$i] -match "kor") { $condition1 = $true }  }		# 5.  korean
							{$_ -match '-ara_' } {
								if ($languageCodeMicrosoft[$i] -match "ara") { $condition1 = $true }  }		# 6.  arabic
							{$_ -match '-fin_' } {
								if ($languageCodeMicrosoft[$i] -match "fin") { $condition1 = $true }  }		# 7.  finnish
							{$_ -match '-plk_' } {
								if ($languageCodeMicrosoft[$i] -match "plk") { $condition1 = $true }  }		# 8.  polish
							{$_ -match '-nor_' } {
								if ($languageCodeMicrosoft[$i] -match "nor") { $condition1 = $true }  }		# 9.  norwegian
							{$_ -match '-ptb_' } {
								if ($languageCodeMicrosoft[$i] -match "ptb") { $condition1 = $true }  }		# 10. portuguese (bra)
							{$_ -match '-ptg_' } {
								if ($languageCodeMicrosoft[$i] -match "ptg") { $condition1 = $true }  }		# 11. portuguese (por)
							{$_ -match '-fra_' } {
								if ($languageCodeMicrosoft[$i] -match "fra") { $condition1 = $true }  }		# 12. french
							{$_ -match '-dan_' } {
								if ($languageCodeMicrosoft[$i] -match "dan") { $condition1 = $true }  }		# 13. danish 
							{$_ -match '-ell_' } {
								if ($languageCodeMicrosoft[$i] -match "ell") { $condition1 = $true }  }		# 14. greek 
							{$_ -match '-csy_' } {
								if ($languageCodeMicrosoft[$i] -match "csy") { $condition1 = $true }  }		# 15. czech
							{$_ -match '-trk_' } {
								if ($languageCodeMicrosoft[$i] -match "trk") { $condition1 = $true }  }		# 16. turkish
							{$_ -match '-rus_' } {
								if ($languageCodeMicrosoft[$i] -match "rus") { $condition1 = $true }  }		# 17. russian
							{$_ -match '-esn_' } {
								if ($languageCodeMicrosoft[$i] -match "esn") { $condition1 = $true }  }		# 18. spanish
							{$_ -match '-ita_' } {
								if ($languageCodeMicrosoft[$i] -match "ita") { $condition1 = $true }  }		# 19. italian
							{$_ -match '-hun_' } {
								if ($languageCodeMicrosoft[$i] -match "hun") { $condition1 = $true }  }		# 20. hungarian
							{$_ -match '-nld_' } {
								if ($languageCodeMicrosoft[$i] -match "nld") { $condition1 = $true }  }		# 21. dutch
							{$_ -match '-heb_' } {
								if ($languageCodeMicrosoft[$i] -match "heb") { $condition1 = $true }  }		# 22. hebrew
							{$_ -match '-sve_' } {
								if ($languageCodeMicrosoft[$i] -match "sve") { $condition1 = $true }  }		# 23. swedish
							{$_ -match '-enu_' } {
								if ($languageCodeMicrosoft[$i] -match "enu") { $condition1 = $true }  }		# 24. english
							{$_ -match '-eng_' } {
								if ($languageCodeMicrosoft[$i] -match "enu") { $condition1 = $true }  }		# english (alt)
							Default {
								$condition1 = $true	 }		# language independent files
						}
					} else {
						$condition1 = $true
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
	
	# if the output directory is empty, we get rid of it
	if ((Get-ChildItem -Path $outdir).Count -eq 0) {
		$msg = "[INFO] No update files found matching the specified filters"
		Write-Host $msg -ForegroundColor Green -BackgroundColor Black
		Remove-Item -Path $outdir
	}
}
MSUpdate-DL