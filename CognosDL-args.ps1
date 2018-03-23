# Setup a report with specific name (MyReportName) to run scheduled to save with which format you want then schedule this script to download it.
# Run from command line or batch script: powershell.exe -executionpolicy bypass -file C:\ImportFiles\Scripts\CognosDownload.ps1 MyReportName
# Change the below variables
# Use lines 5 to create and save your password in a obfuscated text file.!!!!! Line 5 must be ran on the machine and USER context that runs the script!!!!
# Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File "C:\ImportFiles\scripts\apscnpw.txt" 
# Run line 5 each time you update your School Password to re-export the updated file

$username = "0000yourname" #APSCN/SSO username
$password = Get-Content "C:\scripts\apscnpw.txt" | ConvertTo-SecureString #Should Match the export location from line 5
$userdomain = "APSCN"
# dsnname is the database name for your district
# To obtain this you need to log in to the eSchool Cognos site using and view the source code of the overall frameset.
# The dsn is displayed in the second <frame> tag like so where the ****** is: src="https://adecognos.arkansas.gov/ibmcognos/cgi-bin/cognosisapi.dll?dsn=******
$dsnname = "yourschoolsms" 
# The file path where you want the file placed
$savepath = $args[1]
# extension for report format: csv, xlsx
$extension = "csv"
# end of variables to change
$report = $args[0]
$type = $args[2] #Required #Specificies if the download is a report or query studio file
$nestfold = $args[3] #Optional for nested folders under My Folder

# Revisions:
# 2014-07-23: Brian Johnson: Updated URL string to include dsn parameters necessary for eSchool and re-enabled CredentialCache setting to login
# 2018-03-15: Charles Weber added $args[2] to support switching between report and query studio's witht he same download file
# 2018-03-16: Charles Weber added $args[3] to support nested folder downloads, you can now download from one MyFolders and one nested folder.

# server location for Cognos
$baseURL = "https://adecognos.arkansas.gov"
$cWebDir = "ibmcognos"
$camName = "esp"

[System.Net.CredentialCache]$MyCredentialCache
[System.Net.HttpWebRequest]$request
[System.Net.HttpWebResponse]$response
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Create url path with default path being in your My Folders using your username and password from above
IF($args[3] -eq $null){
$url = "$baseURL/$cWebDir/cgi-bin/cognos.cgi?dsn=$dsnname&CAM_action=logonAs&CAMNamespace=$camName&CAMUsername=$username&CAMPassword=$password&b_action=cognosViewer&ui.action=view&ui.object=defaultOutput(CAMID(%22$camName%3aa%3a$username%22)%2ffolder%5b%40name%3d%27My%20Folders%27%5d%2freport%5b%40name%3d%27$report%27%5d)&ui.name=$report&ui.format=CSV"
}
Else{
 $url = "$baseURL/$cWebDir/cgi-bin/cognos.cgi?dsn=$dsnname&CAM_action=logonAs&CAMNamespace=$camName&CAMUsername=$username&CAMPassword=$password&b_action=cognosViewer&ui.action=view&ui.object=defaultOutput(CAMID(%22$camName%3aa%3a$username%22)%2ffolder%5b%40name%3d%27My%20Folders%27%5d%2ffolder%5b%40name%3d%27$nestfold%27%5d%2f$type%5b%40name%3d%27$report%27%5d)&ui.name=$report&ui.format=CSV"
#        ********/********/cgi-bin/cognos.cgi?dsn=********&CAM_action=logonAs&CAMNamespace=********&CAMUsername=*********&CAMPassword=*********&b_action=cognosViewer&ui.action=view&ui.object=defaultOutput(CAMID(%22********%3aa%3a*********%22)%2ffolder%5b%40name%3d%27My%20Folders%27%5d%2freport%5b%40name%3d%27*******%27%5d)&ui.name=$report&ui.format=CSV
}
$fullfilepath = "$savepath\$report.$extension"

If (!(Test-Path ($args[1]))) {
	write-host("Specified folder does not exist! [$fullfilepath]") -ForeGroundColor Yellow
	exit 1
}

trap
{
    write-output $_
    exit 9
}

if(!(Split-Path -parent $savepath) -or !(Test-Path -pathType Container (Split-Path -parent $savepath))) {
  $savepath = Join-Path $pwd (Split-Path -leaf $savepath)
}

$request = [System.Net.HttpWebRequest]::Create($url)

#Set the Credentials
write-host("Setting the Credentials now..") -ForeGroundColor Yellow
$request.UseDefaultCredentials = $true
$request.PreAuthenticate = $true
[System.Net.NetworkCredential]$NetworkCredential = New-Object System.Net.NetworkCredential($username, $password, $userdomain)
$MyCredentialCache = New-Object System.Net.CredentialCache
$MyCredentialCache.Add($url, "Basic", $NetworkCredential)
$request.Credentials = $NetworkCredential

$cookieJar = new-object "System.Net.CookieContainer"
$request.Method = "GET"
$request.CookieContainer = $cookieJar

$response = $request.GetResponse()
if ($response.StatusCode -ne 200)
{
    $result = "Error : " + $response.StatusCode + " : " + $response.StatusDescription
    $result
}
else
{
    $sr = New-Object System.IO.StreamReader($response.GetResponseStream())
    $HTMLDataString = $sr.ReadToEnd()

    write-host("Downloaded HTML to retrieve report url.") -ForeGroundColor Yellow

    $regex = [regex]"var sURL = '(.*?)'"
    if ($HTMLDataString -notmatch $regex)
    {
        if ($HTMLDataString -match [regex]"CAM_PASSPORT_ERROR") #this error is in the output of HTMLDataString
        {
            write-output "Found 'CAM_PASSPORT_ERROR': Please check the password used for script"
            exit 10
        }
        elseif ($HTMLDataString -match [regex]"AAA-AUT-0011") #this error is in the output of HTMLDataString for Invalid Namespace
        {
            write-host($HTMLDataString) -ForeGroundColor White
            write-output "Found 'AAA-AUT-0011': Invalid Namespace error"
            exit 11
        }
        write-host($HTMLDataString) -ForeGroundColor White
        throw "'var sURL' not found"
        
        exit 1
    }
    $urlMatch = $regex.Matches($HTMLDataString)
    write-host("Found URL in data to download report.") -ForeGroundColor Yellow
    $fileURLString = $urlMatch[0].Value.Replace("var sURL = '", "").Replace("'", "")

    if ($uiAction -match "run") #run the report live for the data
    {
        # temp to show data from the HTMLDataString for the sURL
        write-host($HTMLDataString) -ForeGroundColor White
    }

    # Append beginning part of url
    $fileURLString = "$baseURL$fileURLString"

    [System.Net.HttpWebRequest] $fileRequest = [System.Net.HttpWebRequest] [System.Net.WebRequest]::Create($fileURLString)
    $fileRequest.CookieContainer = $request.CookieContainer
    $fileRequest.AllowWriteStreamBuffering = $false
    $fileRequest.Credentials = $NetworkCredential
    $fileResponse = [System.Net.HttpWebResponse] $fileRequest.GetResponse()

    write-host("Downloading $report...") -ForeGroundColor Yellow
    [System.IO.Stream]$st = $fileResponse.GetResponseStream()
    # write to disk
    $mode = [System.IO.FileMode]::Create
    $fs = New-Object System.IO.FileStream $fullfilepath, $mode
    $read = New-Object byte[] 256
    [int] $count = $st.Read($read, 0, $read.Length)
    while ($count -gt 0)
    {
        $fs.Write($read, 0, $count)
        $count = $st.Read($read, 0, $read.Length)
    }
    $fs.Close()
    $st.Close()
    $fileResponse.Close()
    write-host("File [$fullfilepath] downloaded") -ForeGroundColor Yellow
	exit 0
}
$response.Close()
