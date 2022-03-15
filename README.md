# CognosDownloader

**These scripts come without warranty of any kind. Use them at your own risk. I assume no liability for the accuracy, correctness, completeness, or usefulness of any information provided by this site nor for any sort of damages using these scripts may cause.**

## Suggested Install Process
> This install process is recommended so you can easily update to newer versions. You can download git at https://git-scm.com/download/win.

> **Note:** You do not have to use this installation method as you can download the raw file and save it where you want.
````
mkdir \Scripts
cd \Scripts
git clone https://github.com/AR-k12code/CognosDownloader.git
New-Item -Path C:\scripts\CognosDownload.ps1 -ItemType SymbolicLink -Value C:\scripts\CognosDownloader\CognosDownload.ps1 -Force
````

## To download latest updates
> This process only works if you used the suggested install process above.
````
cd \scripts\CognosDownloader
git pull
````

### Error Updating?
> If you see errors like:
````
error: Your local changes to the following files would be overwritten by merge:
CognosDefaults.ps1
Please commit your changes or stash them before you merge.
Aborting
Updating a47e288..862d533
````
>Then you have modified files in the CognosDownloader folder. The quickest option is to delete the c:\scripts\CognosDownloader folder and run the suggested install procedure again.

## CognosDefaults.ps1
>The file CognosDefaults.ps1 has variables that are commented out. Any uncommented variables in this file will override anything you specify on the command line. This file must be in the same folder as the CognosDownload.ps1 file to work. You should save your modified file to c:\scripts\CognosDefaults.ps1 and call the Cognos Downloader from c:\scripts\CognosDownload.ps1.
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

## Exit Codes
````
1 = Save Path does not exist
2 = Unable to Authenticate
3 = Unable to find report with the supplied path
4 = Unable to retrieve output formats from the Cognos RDS
5 = Report has required Prompts you must answer
6 = Failed to download file
7 = Could not find file after it downloaded
8 = Failed CSV header count. This can be because of a formatting issue returned from Cognos.
9 = Failed CSV line count. Not enough data returned or blank.
10 = General verification error. Could not convert to a format to verify.
12 = Failed to download file from Cognos server after attempting 3 times.
30 = Failed to send email on download error.
50 = Timeout. The report did not finish processing on the Cognos server within the timeout period. Default 5 minutes.
````

You can check the $LASTEXITCODE variable to ensure your download finished correctly.

## eFinance Username
> It is possible that your eFinance username is not the same as your SSO username. If your account came from the old APSCN GUI days you'll need to specify both -username and -efpuser separately. For example my username for eFinance is not the same as my SSO username. We did not have the LEA at the front to begin with. The $efpuser can also be specified in the CognosDefaults.ps1.
````
.\CognosDownload.ps1 -username 0401cmillsap -efpuser cmillsap -efpdsn gentryfms -eFinance -report openpos -savepath "c:\scripts"
````

## Running on Linux
> You will need to install Powershell 7 and this dependency on Ubuntu.
````
apt-get install -y --no-install-recommends gss-ntlmssp
````