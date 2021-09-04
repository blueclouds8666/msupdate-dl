# Microsoft Update Downloader

Microsoft Update Downloader is a PowerShell script designed to allow easy file download for updates and hotfixes from the Microsoft Update Catalog. Advanced functionality allows to batch download multiple updates at once, as well as filtering the files by language, Windows version and architecture.
<br />
<br />

## System Requirements

#### Windows users
- Windows 7 with SP1 (32 and 64 bits) or later
- PowerShell 3.0 or later
- NET Framework 4.5
- DigiCert Global Root G2 Certificate
- GNU Wget (already included with release versions)

#### Linux users
- Linux kernel 3.13 (AMD64, ARM32 and ARM64) or later
- PowerShell 6.0 or later
- DigiCert Global Root G2 Certificate
- GNU Wget
<br />

## Getting started
#### Script execution:
Once your system meets mentioned requirements and has the MSUpdate-DL scripts downloaded and ready, you can run it by using the following syntax:
> MSUpdate-DL.ps1 [REQUEST] [OPTIONS]

A second script allows for a batch download, reading lines from MSUpdate-DL-Batch-List.txt. Syntax for this script is as follows:
> MSUpdate-DL-Batch.ps1 [OPTIONS]

#### Valid arguments:
Argument | Description
--- | ---
REQUEST&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| KB update number to download. This field is required and must always come first.
-details | download a details sheet for each update in html format
-architecture="VALUE" | filter files by OS architecture. valid values are: all, 32, 64, itanium
-ntversion="VALUE" | filter files by Windows NT version. valid values are: 50, 51, 52, 60, 61...
-language="VALUE" | filter files by language. valid values are: enu, fra, ita, jpn, rus...
-output="VALUE" | specify a different output directory (default is C:\msupdate-dl\ on Windows, and \home\username\msupdate-dl\ on Linux)
-version | show script version
-help | show brief help information
-log | enable logging to a text file

#### Complete documentation:
Please refer to documentation.htm for a complete operation guide. Make sure to read and understand the documentation before using the script. This markdown section only serves as a brief introduction to this script.
<br />
<br />


## Use examples

Download the english version file for the KB951376 update, and log the process:
> MSUpdate-DL.ps1 KB951376 -language="enu" -log

Batch download updates to an external drive with the details files:
> MSUpdate-DL-Batch.ps1 -output="E:\backup\updates\" -details

Batch download updates only those in english, german or french, only for Windows Vista x64, and log the process:
> MSUpdate-DL-Batch.ps1 -language="enu,deu,fra" -ntversion="60" -architecture="64" -log

<br />


## Download resources

- [MSUpdate-DL latest release for Win32](https://github.com/blueclouds8666/msupdate-dl/releases/download/1.00/msupdate-dl-windows-i686.7z)
- [MSUpdate-DL latest release for Win64](https://github.com/blueclouds8666/msupdate-dl/releases/download/1.00/msupdate-dl-windows-AMD64.7z)
- [MSUpdate-DL latest release for Linux](https://github.com/blueclouds8666/msupdate-dl/releases/download/1.00/msupdate-dl-linux.7z)
- [DigiCert Global Root G2 Certificate](https://cacerts.digicert.com/DigiCertGlobalRootG2.crt)
- [PowerShell 3.0](https://www.microsoft.com/en-us/download/details.aspx?id=34595)
- [PowerShell 6.0](https://github.com/PowerShell/PowerShell/releases/tag/v6.0.0)
- [NET Framework 4.5](https://www.microsoft.com/en-US/download/details.aspx?id=40779)
- [GNU Wget for Windows](https://eternallybored.org/misc/wget/)
- [GNU Wget for Linux](https://www.gnu.org/software/wget/)

<br />


## Dev team

GitHub name | Email address
--- | ---
blueclouds8666 | blueclouds8666@mail.com 

MSUpdate-DL has been brought to you by the NEONFLOPPY Team. Check our website at [neonfloppy.sytes.net](http://neonfloppy.sytes.net)

<br />
 
