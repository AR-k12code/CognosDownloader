# CognosDownloader


## Suggested Install Process
> This install process is recommended so you can easily update to newer versions. You can download git at https://git-scm.com/download/win.

> **Note:** You do not have to use this installation method as you can download the raw file and save it where you want.
````
mkdir \Scripts
cd \Scripts
git clone https://github.com/AR-k12code/CognosDownloader.git
New-Item -Path C:\scripts\CognosDownload-New.ps1 -ItemType SymbolicLink -Value C:\scripts\CognosDownloader\CognosDownload.ps1 -Force
````

## To download latest updates
> This process only works if you used the suggested install process above.
````
cd \scripts\CognosDownloader
git pull
````

## CognosDefaults.ps1
>The file CognosDefaults.ps1 has variables that are commented out. Any uncommented variables in this file will override anything you specify on the command line OR modify in the CognosDownload.ps1 file. This file must be in the same folder as the CognosDownload.ps1 file to work.
````
#$username = '0000username'
#$passwordfile = 'c:\scripts\mysavedpassword.txt'
#$espdsn = 'schoolsms'
#$savepath = 'c:\scripts\files'
````

## Additional help can be found
````
Get-Help .\CognosDownload.ps1
Get-Help .\CognosDownload.ps1 -Examples
````

## Sample Command Line
> Basic Syntax
````
.\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report "Active Students" -savepath "c:\scripts\clever\files"
````

> Show Details of the above Report without Downloading
````
.\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report "Active Students" -savepath "c:\scripts\clever\files" -ShowReportDetails -SkipDownloadingFile
````

## SSO Password Change
> You MUST change the password under the same Windows user account used to run your downloads.
````
.\ResetPassword.ps1
````