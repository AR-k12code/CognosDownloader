# CognosDownloader

## Suggested Install Process
````
mkdir \Scripts
cd \Scripts
git clone https://github.com/AR-k12code/CognosDownloader.git
copy CognosDownloader\CognosDownload.ps1 c:\scripts\CognosDownload.ps1
````

## To pull updates
````
cd \scripts\CognosDownloader
git pull
copy CognosDownloader\CognosDownload.ps1 c:\scripts\CognosDownload.ps1
````

## To set defaults
The file CognosDefaults.ps1 have variables that are commented out. Any uncommented variables in this file will override anything you specify on the command line OR modify in the CognosDownload.ps1 file.