#Requires -Version 5.1
#Get-Help .\CognosDownload.ps1
#Get-Help .\CognosDownload.ps1 -Examples
#  ___ _____ ___  ___   ___   ___    _  _  ___ _____         
# / __|_   _/ _ \| _ \ |   \ / _ \  | \| |/ _ \_   _|        
# \__ \ | || (_) |  _/ | |) | (_) | | .` | (_) || |          
# |___/ |_| \___/|_|   |___/ \___/  |_|\_|\___/ |_|

#  ___ ___ ___ _____   _____ _  _ ___ ___   ___ ___ _    ___ 
# | __|   \_ _|_   _| |_   _| || |_ _/ __| | __|_ _| |  | __|
# | _|| |) | |  | |     | | | __ || |\__ \ | _| | || |__| _| 
# |___|___/___| |_|     |_| |_||_|___|___/ |_| |___|____|___|
#                                                           
# Please see the https://www.github.com/AR-K12code/CognosDownload to see how to use the CognosDefaults.ps1 file.

# This version is NOT complete. Please check back over the next few weeks for updates!
# Script Contributors - Brian Johnson, Charlie Weber, Scott Organ, Joshua Reed, Craig Millsap, and Michael Hayes.

<#
  .SYNOPSIS
  This script is used to download reports from the Arkansas Cognos 11 using your SSO credentails.

  .DESCRIPTION
  CognosDownload.ps1 invoked with the proper parameters will download a report in the desired format.

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report students

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap  -espdsn gentrysms -report sections -reportparams "p_year=2021"
  This provides a simple solution to answer a single page prompt.

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report sections -XMLParameters "CustomPromptAnswers.xml"
  This provides for answering more complex and multipage prompt pages. Script will automatically use an XML file named the the Report ID with an extension of .xml
  
  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report activities -cognosfolder "_Share Temporarily Between Districts/Gentry/automation" -TeamContent

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report "APSCN Virtual AR Student File" -savepath .\ -TeamContent -cognosfolder "Demographics/Demographic Download Files" -SavePrompts
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report "APSCN Virtual AR Student File" -savepath .\ -ShowReportDetails -TeamContent -cognosfolder "Demographics/Demographic Download Files" -XMLParameters i4C884862DFD8470ABFF2571CB47F01EA.xml -extension pdf 
  For reports with complex paramters you can capture and save the prompts for reuse by specifying the -SavePrompts paramter.

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report students -SendMail -mailto "technology@gentrypioneers.com" -mailfrom noreply@gentrypioneers.com

  .EXAMPLE
  PS> .\CognosDownload.ps1 -username 0401cmillsap -espdsn gentrysms -report students -ShowReportDetails -SkipDownloadingFile

  .PARAMETER ShowReportDetails
  Print details about the report including Title, Owner, Owner Email, Location in Cognos, and the ID.
  
  .PARAMETER SkipDownloadingFile
  All other steps except actually downloading the final file to the specified save path.


#>

Param(
    [parameter(Mandatory=$true,HelpMessage="Give the name of the report you want to download.")]
        [string]$report,
    [parameter(Mandatory=$false,HelpMessage="Give a specific folder to download the report into.")]
        [string]$savepath="C:\scripts",
    [parameter(Position=2,Mandatory=$false,HelpMessage="Format you want to download report as.")]
        [string]$extension="csv",
    [parameter(Mandatory=$false,HelpMessage="eSchool SSO username to use.")]
        [string]$username="", #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
    [parameter(Mandatory=$false,HelpMessage="File for ADE SSO Password")]
        [string]$passwordfile="C:\Scripts\apscnpw.txt", # Override where the script should find the password for the user specified with -username.
    [parameter(Mandatory=$false,HelpMessage="eSchool DSN location.")]
        [string]$espdsn="", #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
    [parameter(Mandatory=$false,HelpMessage="eFinance username to use.")]
        [string]$efpuser="", #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
    [parameter(Mandatory=$false,HelpMessage="eFinance DSN location.")]
        [string]$efpdsn="", #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
    [parameter(Mandatory=$false,HelpMessage="Cognos Folder Structure.")]
        [string]$cognosfolder="My Folders", #Cognos Folder "Folder 1/Sub Folder 2/Sub Folder 3" NO TRAILING SLASH
    [parameter(Mandatory=$false,HelpMessage="Report Parameters")]
        [string]$reportparams="", #If a report requires parameters you can specifiy them here. Example:"p_year=2017&p_school=Middle School"
    [parameter(Mandatory=$false,HelpMessage="Get the report from eFinance.")]
        [switch]$eFinance,
    [parameter(Mandatory=$false,HelpMessage="Send an email on failure.")]
        [switch]$SendMail,
    [parameter(Mandatory=$false,HelpMessage="SMTP Auth Required.")]
        [switch]$smtpauth,
    [parameter(Mandatory=$false,HelpMessage="SMTP Server")]
        [string]$smtpserver="smtp-relay.gmail.com", #--- VARIABLE --- change for your email server
    [parameter(Mandatory=$false,HelpMessage="SMTP Server Port")]
        [int]$smtpport="587", #--- VARIABLE --- change for your email server
    [parameter(Mandatory=$false,HelpMessage="SMTP eMail From")]
        [string]$mailfrom="noreply@yourdomain.com", #--- VARIABLE --- change for your email from address
    [parameter(Mandatory=$false,HelpMessage="File for SMTP eMail Password")]
        [string]$smtppasswordfile="C:\Scripts\emailpw.txt", #--- VARIABLE --- change to a file path for email server password
    [parameter(Mandatory=$false,HelpMessage="Send eMail to")]
        [string]$mailto="technology@yourdomain.com", #--- VARIABLE --- change for your email to address
    [parameter(Mandatory=$false,HelpMessage="Minimum line count required for CSVs")]
        [int]$requiredlinecount=3, #This should be the ABSOLUTE minimum you expect to see. Think schools.csv for smaller districts.
    [parameter(Mandatory=$false)]
        [switch]$ShowReportDetails, #Print report details to terminal.
    [parameter(Mandatory=$false)]
        [switch]$SkipDownloadingFile, #Do not actually download the file.
    [parameter(Mandatory=$false)]
        [switch]$dev, #use the development URL dev.adecognos.arkansas.gov
    [parameter(Mandatory=$false)]
        [switch]$TeamContent, #Report is in the Team Content folder. You will also need to have specified the -cognosfolder parameter with the path.
    [parameter(Mandatory=$false)]
        [string]$XMLParameters, #Path to XML for answering prompts.
    [parameter(Mandatory=$false)]
        [switch]$SavePrompts,
    [parameter(Mandatory=$false)]
        [string]$Encoding="utf8",
    [parameter(Mandatory=$false)] #Verification now includes checking CSV,XLSX and XML.
        [switch]$DisableCSVVerification,
    [parameter(Mandatory=$false)]
        [int]$reportwait = 15,
    [parameter(Mandatory=$false)] #not used anymore. here for backwards compatibility
        [switch]$RunReport,
    [parameter(Mandatory=$false)] #not used anymore. here for backwards compatibility
        [switch]$ReportStudio,
    [parameter(Mandatory=$false)] #This is to establish a session for subsequent calls.
        [switch]$EstablishSessionOnly,
    [parameter(Mandatory=$false)] #This is if you're going to run subsequent sessions aftewards.
        [switch]$SessionEstablished,
    [parameter(Mandatory=$false)] #Remove Spaces in CSV files. This requires Powershell 7.1+
        [switch]$TrimCSVWhiteSpace,
    [parameter(Mandatory=$false)] #If you Trim CSV White Space do you want to wrap everything in quotes?
        [switch]$CSVUseQuotes,
    [parameter(Mandatory=$false)] #If you want to override the default saved filename. You need to include the file extension.
        [string]$FileName,
    [parameter(Mandatory=$false)] #How long in minutes are you willing to let CognosDownloader run for said report? 5 mins is default and gives us a way to error control.
        [int]$Timeout = 5,
    [parameter(Mandatory=$false)] #If you need to download the same report multiple times but with different parameters we have to use a random temp file so they don't conflict.
        [switch]$RandomTempFile
)

$version = [version]"22.3.15"

Add-Type -AssemblyName System.Web

$startTime = Get-Date

#powershell.exe -executionpolicy bypass -file C:\Scripts\CognosDownload.ps1 -username 0000username -espdns schoolsms -report MyReportName -cognosfolder "subfolder" -savepath "c:\scripts\downloads" 

# When the password expires, just delete the specific file (c:\scripts\apscnpw.txt) and run the script to re-create.
#Example for the Team Content folder:
#https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/wsil/path/Team%20Content%2FStudent%20Management%20System%2F_Share%20Temporarily%20Between%20Districts%2FGentry%2Fautomation
#/content/folder[@name='Student Management System']/folder[@name='_Share Temporarily Between Districts']/folder[@name='Gentry']/folder[@name='automation']/query[@name='activities']
#https://dev.adecognos.arkansas.gov/ibmcognos/bi/v1/disp/rds/atom/path/Team Content/Student Management System/_Share Temporarily Between Districts/Gentry/automation/activities
#CAMID("esp:a:0401cmillsap")/folder[@name='My Folders']/folder[@name='automation']/query[@name='activities']

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
if (Test-Path $currentPath\CognosDefaults.ps1) {
    . $currentPath\CognosDefaults.ps1
}

#version check, continue on failure as if nothing happened at all.
try {
    $versioncheck = Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/AR-k12code/CognosDownloader/master/version.json'
    if ($version -lt [version]($versioncheck.version)) {
        Write-Host "`r`nInfo: There is a new version of this script available at https://www.github.com/AR-K12code/CognosDownloader"
        Write-Host "Info: Version $($versioncheck.version) is available. Description: $($versioncheck.description)"
    }
    if ($versioncheck.versions) {
        $versioncheck.versions | ForEach-Object { $PSItem.version = [version]$PSitem.version }
        $versioncheck.versions | Where-Object { $PSItem.version -gt $version } | ForEach-Object {
            Write-Host "Info: Version $($($PSItem.version).ToString()) is available. Description: $($PSItem.description)"
        }
    }
    Write-Host "`r`n"
} catch {} #Do and show nothing if we don't get a response.

#send mail on failure.
$mailsubject = "[CognosDownloader]"
function Send-Email([string]$failurereason,[string]$errormessage) {
    if ($SendMail) {
        $msg = New-Object Net.Mail.MailMessage
        $smtp = New-Object Net.Mail.SmtpClient($smtpserver, $smtpport)
        #port 25 is likely non-ssl (for internal restricted relays), maybe change to switch option?
        if ($smtpport -eq 25) {$smtp.EnableSSL = $False} else { $smtp.EnableSSL = $True }
        #If authentication is required.
        if ($smtpauth) { $smtp.Credentials = New-Object System.Net.NetworkCredential($mailfrom,$mailfrompassword) }
        $msg.From = $mailfrom
        $msg.To.Add($mailto)
        #Include date so emails don't group in a thread.
        $msg.subject =  $mailsubject + $failurereason + "[$(Get-Date -format MM/dd/y)]" + '[' + $report + ']'
        $msg.Body = "The report " + $report  + " failed to download properly.`r`n"
        if ($errormessage) {
            $msg.Body += "$errormessage`r`n"
        }
        $msg.Body += $url
        
        try {
            $smtp.send($msg)
        } catch {
            Write-Host("Failed to send email: $_") -ForeGroundColor Red
            exit(30)
        }
    }
}

function Reset-DownloadedFile([string]$fullfilepath) {
    $PrevOldFileExists = Test-Path ($fullfilepath + ".old")
    if ($PrevOldFileExists -eq $True) {
        Write-Host -NoNewline "Deleting old $report..." -ForeGroundColor Yellow
        Remove-Item -Path $fullfilepath -Force -ErrorAction SilentlyContinue
        Rename-Item -Path ($fullfilepath + ".old") -newname ($fullfilepath)
    } else {
        Remove-Item -Path $fullfilepath -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Reversing old $($report)." -ForeGroundColor Red
}

# URL for Cognos
if ($dev) {
    $baseURL = "https://dev.adecognos.arkansas.gov"
} else {
    $baseURL = "https://adecognos.arkansas.gov"
}

If ($eFinance) {
    $camName = "efp"    #efp for eFinance
    $dsnparam = "spi_db_name"
    $dsnname = $efpdsn
    $camid = "CAMID(""efp_x003Aa_x003A$($efpuser)"")"
} else {
    $camName = "esp"    #esp for eSchool
    $dsnparam = "dsn"
    $dsnname = $espdsn
    $camid = "CAMID(""esp_x003Aa_x003A$($username)"")"
}

#Script to create a password file for Cognos download Directory
#This script MUST BE RAN LOCALLY to work properly! Run it on the same machine doing the cognos downloads, this does not work remotely!

if ((Test-Path ($passwordfile))) {
    $password = Get-Content $passwordfile | ConvertTo-SecureString
} else {
    Write-Host("Password file does not exist! [$passwordfile]. Please enter a password to be saved on this computer for scripts") -ForeGroundColor Yellow
    Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $passwordfile
    $password = Get-Content $passwordfile | ConvertTo-SecureString
}

$creds = New-Object System.Management.Automation.PSCredential $username,$password

If ($smtpauth) {
    if ((Test-Path ($smtppasswordfile))) {
        $smtppassword = Get-Content $smtppasswordfile | ConvertTo-SecureString
    } else {
        Write-Host("SMTP Password file does not exist! [$smtppasswordfile]. Please enter a password to be saved on this computer for emails") -ForeGroundColor Yellow
        Read-Host "Enter Password" -AsSecureString |  ConvertFrom-SecureString | Out-File $smtppasswordfile
        $mailfrompassword = Get-Content $smtppasswordfile | ConvertTo-SecureString
    }
}

#https://www.ibm.com/docs/en/cognos-analytics/11.1.0?topic=sets-outputformatenum
switch ($extension) {
    "pdf" { $fileformat = "PDF" }
    "csv" { $fileformat = "CSV" }
    "xlsx" { $fileformat = "spreadsheetML" }
    "xml" { $fileformat = "XML" }
    #DEFAULT { $fileformat = "CSV" }
}

if ($FileName) {
    $fullfilepath = "$savepath\$FileName"
} else {
    $fullfilepath = "$savepath\$report.$extension"
}

If (!(Test-Path ($savepath))) {
    Write-Host("Specified save folder does not exist! [$fullfilepath]") -ForeGroundColor Yellow
    Send-Email("[Failure][Save Path Missing]","Missing path $fullfilepath")
    exit(1) #specified save folder does not exist
}

if(!(Split-Path -parent $savepath) -or !(Test-Path -pathType Container (Split-Path -parent $savepath))) {
  $savepath = Join-Path $pwd (Split-Path -leaf $savepath)
}

$FileExists = Test-Path $fullfilepath
If ($FileExists -eq $True) {
    #replace datetime for if-modified-since header from existing file
    $filetimestamp = (Get-Item $fullfilepath).LastWriteTime
}

#submit login and switch to site.
if (-Not($SessionEstablished)) {
    $failedlogin = 0
    do {
        try {
            Write-Host "Authenticating and switching to $dsnname... " -ForegroundColor Yellow -NoNewline
            $response1 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/login" -SessionVariable session `
            -Method "POST" `
            -ContentType "application/json; charset=UTF-8" `
            -Credential $creds `
            -Body "{`"parameters`":[{`"name`":`"h_CAM_action`",`"value`":`"logonAs`"},{`"name`":`"CAMNamespace`",`"value`":`"$camName`"},{`"name`":`"$dsnparam`",`"value`":`"$dsnname`"}]}" 

            Write-Host "Success." -ForegroundColor Yellow
        } catch {
            $failedlogin++            
            if ($failedlogin -ge 2) {
                Write-Host "Unable to authenticate and switch into $dsnname. $($_)" -ForegroundColor Red
                Send-Email("[Failure][Authentication]","$($_)")
                exit(2)
            } else {
                #Unfortuantely we are still having an issue authenticating to Cognos. So we need to make another attemp after a random number of seconds.
                Write-Host "Failed to authenticate. Attempting again..." -ForegroundColor Red
                Remove-Variable -Name session
                Start-Sleep -Seconds (Get-Random -Maximum 15 -Minimum 5)
            }
        }
    } until ($session)
} else {
    $session = $incomingsession
}

#Stop here for Established Session.
if ($EstablishSessionOnly) { exit(0) }

#No subfolder specified.
if ($cognosfolder -eq "My Folders") {
    #$cognosfolder = ([System.Web.HttpUtility]::UrlEncode("My Folders")).Replace('+','%20')
    $cognosfolder = "$($camid)/My Folders".Replace(' ','%20')
} elseif ($TeamContent) {
    if ($eFinance) {
        $cognosfolder = "Team Content/Financial Management System/$($cognosfolder)".Replace(' ','%20')
    } else {
        $cognosfolder = "Team Content/Student Management System/$($cognosfolder)".Replace(' ','%20')
    }
} else {
    #$cognosfolder = ([System.Web.HttpUtility]::UrlEncode("My Folders/$($cognosfolder)")).Replace('+','%20')
    $cognosfolder = "$($camid)/My Folders/$($cognosfolder)".Replace(' ','%20')
}

#Get the Atom feed
try {
    Write-Host "Attempting to retrieve report details for $($report)... " -ForegroundColor Yellow -NoNewline
    $response2 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/atom/path/$($cognosfolder)/$($report)" -WebSession $session
    $reportDetails = $response2.feed
    $reportID = $reportDetails.entry.storeID
    Write-Host "Success." -ForegroundColor Yellow
    
} catch {
    Write-Host "Unable to retrieve report details. Please check the supplied report name and cognosfolder. $($_)" -ForegroundColor Red
    Send-Email("[Failure][Missing Path]","$($_)")
    exit(3)
}

#Get the possible outputformats.
try {
    Write-Host -NoNewline "Retrieving possible formats... " -ForegroundColor Yellow
    $response3 = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/outputFormats/path/$($cognosfolder)/$($report)" -WebSession $session
    Write-Host "Success." -ForegroundColor Yellow -NoNewline

    if ($response3.GetOutputFormatsResponse.supportedFormats.outputFormatName) {
        Write-Host " - $($report) ($($reportID)) can be exported in the following formats:" $($($response3.GetOutputFormatsResponse.supportedFormats.outputFormatName) -join ',') -ForegroundColor Yellow
    
        #This is case sensitive. So we need to retrieve the value from the response and match to the possible incorrect case provided to script.
        if ($($response3.GetOutputFormatsResponse.supportedFormats.outputFormatName) -contains $fileformat) {
            $possibleFormats = $($response3.GetOutputFormatsResponse.supportedFormats.outputFormatName)
            $validExtension = $possibleFormats[$($possibleFormats.ToLower().IndexOf($fileformat.ToLower()))]
        } else {
            Write-Host "You have requested an invalid extension type for this report."
            Throw "Invalid extension requested."
        }
    } else {
        throw
    }

} catch {
    Write-Host "Failed to retrieve output formats for the supplied report. $($_)" -ForegroundColor Red
    Send-Email("[Failure][Report Details Missing]","$($_)")
    exit(4)
}

#Print Additional Details to Terminal
if ($ShowReportDetails) {
    $details = $reportDetails | Select-Object -Property title,owner,ownerEmail,location,id
    $details.id = $reportID
    $details | Format-List
}

if (-Not($SkipDownloadingFile)) {
    Try {
        
        #Move Previous File
        $PrevFileExists = Test-Path $fullfilepath
        If ($PrevFileExists -eq $True) {
            $PrevOldFileExists = Test-Path ($fullfilepath + ".old")
            If ($PrevOldFileExists -eq $True) {
                Write-Host("Deleting old $report...") -ForeGroundColor Yellow
                Remove-Item -Path ($fullfilepath + ".old")
            }
            try {
                Write-Host "Renaming old $report... " -ForeGroundColor Yellow -NoNewline
                Rename-Item -Path $fullfilepath -newname ($fullfilepath + ".old")
                Write-Host "Success." -ForegroundColor Yellow
            } catch {
                Write-Host "Failed to rename old report." -ForegroundColor Red
            }
        }

        $downloadURL = "$($baseURL)/ibmcognos/bi/v1/disp/rds/outputFormat/path/$($cognosfolder)/$($report)/$($validExtension)?v=3&async=MANUAL"
                
        #https://www.ibm.com/support/knowledgecenter/SSEP7J_11.1.0/com.ibm.swg.ba.cognos.ca_dg_cms.doc/c_dg_raas_run_rep_prmpt.html#dg_raas_run_rep_prmpt
        #I think this should be a path as well to the xmlData so you can save it to a text file and pull in when needed to run.
        #Maybe if the prompts return a Test-Path $True then import and use the xmlData field instead. This should allow for more complex prompts.

        if ($reportparams -ne '') {
            $downloadURL = $downloadURL + '&' + $reportparams
        }

        try {
            if ($XMLParameters -ne '') {
                if (Test-Path "$XMLParameters") {
                    Write-Host "Info: Using """$XMLParameters""" in current directory for report prompts." -ForegroundColor Yellow
                    $reportParamXML = (Get-Content "$XMLParameters") -replace ' xmlns:rds="http://www.ibm.com/xmlns/prod/cognos/rds/types/201310"','' -replace 'rds:','' -replace '<','%3C' -replace '>','%3E' -replace '/','%2F'
                    $promptXML = [xml]((Get-Content "$XMLParameters") -replace ' xmlns:rds="http://www.ibm.com/xmlns/prod/cognos/rds/types/201310"','' -replace 'rds:','')
                    $downloadURL = $downloadURL + '&xmlData=' + $reportParamXML
                }
            } elseif (Test-Path "$($reportID).xml") {
                Write-Host "Info: Found ""$($reportID).xml"" in current directory. Using saved report prompts." -ForegroundColor Yellow
                $reportParamXML = (Get-Content "$($reportID).xml") -replace ' xmlns:rds="http://www.ibm.com/xmlns/prod/cognos/rds/types/201310"','' -replace 'rds:','' -replace '<','%3C' -replace '>','%3E' -replace '/','%2F'
                $promptXML = [xml]((Get-Content "$($reportID).xml") -replace ' xmlns:rds="http://www.ibm.com/xmlns/prod/cognos/rds/types/201310"','' -replace 'rds:','')
                $downloadURL = $downloadURL + '&xmlData=' + $reportParamXML
            }

            if ($promptXML) {
                Write-Host "Info: You can customize your prompts by changing any of the following fields and using the -reportparams parameter."
                $promptXML.promptAnswers.promptValues | ForEach-Object {
                    $promptname = $PSItem.name
                    $PSItem.values.item.SimplePValue.useValue | ForEach-Object {
                        Write-Host ("&p_$($promptname)=$($PSItem)").Trim() -NoNewline
                    }
                }
                Write-Host "`r`n"
            }

        } catch {}

        Write-Host "Downloading Report to ""$($fullfilepath)""... " -ForegroundColor Yellow -NoNewline
        
        #Response4 should always be a ticket. Then we move forward with the conversationID.
        $response4 = Invoke-RestMethod -Uri $downloadURL -WebSession $session

        if ($response4.receipt.status -eq "working") {

            #At this point we have our conversationID that we can use to query for if the report is done or not. If it is still running it will return an XML response with reciept.status = working.
            #The problem now is that Cognos decides to either reply with the actual file or another receipt. Since we can't decipher which one prior to our next request we need to save the output
            #to the disk. I have tried keeping this in memory but the way Invoke-WebRequest and Invoke-RestMethod move it from memory often leads to incorrect encoding. 

            #We need a filename to save to that won't conflict with the reportID.xml we already use for saved parameters. Lets hash the reportID for a predictable yet nonconflicting name.
            #With the new espDatabase project use the same report for any table. So the hash needs to be randomized.
            if ($RandomTempFile) {
                $reportIDHash = (New-Guid).Guid.ToString()
            } else {   
                $reportIDHash = (Get-FileHash -InputStream ([System.IO.MemoryStream]::New([System.Text.Encoding]::ASCII.GetBytes($reportID)))).Hash
            }
            #We want to make sure we always keep data together so we don't leave confidential information somewhere else on the system. DO NOT USE TEMP!
            $reportIDHashFilePath = "$(Split-Path $fullfilepath)\$($reportIDHash)"

            #Attempt first download.
            Invoke-WebRequest -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/sessionOutput/conversationID/$($response4.receipt.conversationID)?v=3&async=MANUAL" -WebSession $session -OutFile "$reportIDHashFilePath" -ErrorAction STOP

            #Now we need to test if we got a ticket or not. If not then it should be the actual data and we can rename.
            try {
                $fileContents = [XML](Get-Content $reportIDHashFilePath)
            } catch {}

            #This would indicate a generic failure or a prompt failure.
            if ($fileContents.error) {
                $errorResponse = $fileContents.error
                Write-Host "Error detected in downloaded file. $($errorResponse.message)" -ForegroundColor Red

                if ($errorResponse.promptID) {
                    $promptid = $errorResponse.promptID
                    #Expecting prompts. Lets see if we can find them.
                    $promptsConversation = Invoke-RestMethod -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/reportPrompts/report/$($reportID)?v=3&async=MANUAL" -WebSession $session
                    $prompts = Invoke-WebRequest -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/sessionOutput/conversationID/$($promptsConversation.receipt.conversationID)?v=3&async=MANUAL" -WebSession $session
                    Write-Host "`r`nError: This report expects the following prompts:" -ForegroundColor RED

                    Select-Xml -Xml ([xml]$prompts.Content) -XPath '//x:pname' -Namespace @{ x = "http://www.ibm.com/xmlns/prod/cognos/layoutData/201310" } | ForEach-Object {
                        
                        $promptname = $PSItem.Node.'#text'
                        Write-Host "p_$($promptname)="

                        if (Select-Xml -Xml ([xml]$prompts.Content) -XPath '//x:p_value' -Namespace @{ x = "http://www.ibm.com/xmlns/prod/cognos/layoutData/200904" }) {
                            $promptvalues = Select-Xml -Xml ([xml]$prompts.Content) -XPath '//x:p_value' -Namespace @{ x = "http://www.ibm.com/xmlns/prod/cognos/layoutData/200904" } | Where-Object { $PSItem.Node.pname -eq $promptname }
                            if ($promptvalues.Node.selOptions.sval) {
                                $promptvalues.Node.selOptions.sval
                            }
                        }

                    }

                    Write-Host "Info: If you want to save prompts please run the script again with the -SavePrompts switch."

                    if ($SavePrompts) {
                        
                        Write-Host "`r`nInfo: For complex prompts you can submit your prompts at the following URL. You must have a browser window open and signed into Cognos for this URL to work." -ForegroundColor Yellow
                        Write-Host ("$($baseURL)" + ([uri]$errorResponse.url).PathAndQuery) + "`r`n"
                        
                        $promptAnswers = Read-Host -Prompt "After you have followed the link above and finish the prompts, would you like to download the responses for later use? (y/n)"

                        if (@('Y','y') -contains $promptAnswers) {
                            Write-Host "Info: Saving Report Responses to $($reportID).xml to be used later." -ForegroundColor Yellow
                            Invoke-WebRequest -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/promptAnswers/conversationID/$($promptid)?v=3&async=OFF" -WebSession $session -OutFile "$($reportID).xml"
                            Write-Host "Info: You will need to rerun this script to download the report using the saved prompts." -ForegroundColor Yellow

                            $promptXML = [xml]((Get-Content "$($reportID).xml") -replace ' xmlns:rds="http://www.ibm.com/xmlns/prod/cognos/rds/types/201310"','' -replace 'rds:','')
                            $promptXML.promptAnswers.promptValues | ForEach-Object {
                                $promptname = $PSItem.name
                                $PSItem.values.item.SimplePValue.useValue | ForEach-Object {
                                    Write-Host "&p_$($promptname)=$($PSItem)"
                                }
                            }
                            
                        }
                    }
                }

                Send-Email("[Failure][Prompts]","Report $report requires prompts to run properly.")
                exit(5)

            } elseif ($fileContents.receipt) { #task is still in a working status
                
                Write-Host "`r`nInfo: Report is still working."
                Start-Sleep -Seconds 1 #Cognos is stupid fast sometimes but not so fast that we can make another query immediately.
                
                #The Cognos Server has started randomly timing out, 502 bad gateway, or TLS errors. We need to allow at least 3 errors becuase its not consistent.
                $errorResponse = 0
                do {

                    if ((Get-Date) -gt $startTime.AddMinutes($Timeout)) {
                        Write-Host "Error: Timeout of $Timeout met. Exiting." -ForegroundColor Red
                        Send-Email("[Failure][Download Timeout]","Failed to download file in alloted time of $Timeout. $($_)")
                        Reset-DownloadedFile($fullfilepath)
                        exit(50)
                    }

                    try {
                        Invoke-WebRequest -Uri "$($baseURL)/ibmcognos/bi/v1/disp/rds/sessionOutput/conversationID/$($response4.receipt.conversationID)?v=3&async=AUTO" -WebSession $session -OutFile "$reportIDHashFilePath" -ErrorAction STOP
                        $errorResponse = 0 #reset error response counter. We want three in a row, not three total.
                    } catch {
                        #on failure $response7 is not overwritten.
                        $errorResponse++ #increment error response counter.
                        if ($errorResponse -ge 3) {
                            Write-Host "Failed to download file. $($_)" -ForegroundColor Red
                            Send-Email("[Failure][Download Failed]","Failed to download file. $($_)")
                            Reset-DownloadedFile($fullfilepath)
                            exit(12) #Encountered 3 errors trying to download the file from the conversationID link. Its likely it won't ever succeed.
                        }
                        Start-Sleep -Seconds ($reportwait + 10) #Lets wait just a bit longer to see if its a timing issue.
                    }

                    #Now we need to test if we got a ticket or not. If not then it should be the actual data and we can rename.
                    try {
                        $fileContents = [XML](Get-Content $reportIDHashFilePath)
                    } catch {
                        #If we can't convert it to XML then the format is different and we should be ready to move forward.
                        #However, $fileContents is not overwritten if the XML conversion fails.
                        $fileContents = $null
                        $Error.Clear()
                    }

                    if ($fileContents.receipt.status -eq "working") {
                        Write-Host '.' -NoNewline
                        Start-Sleep -Seconds $reportwait
                    }

                } until ($fileContents.receipt.status -ne "working")

                #We sould have downloaded a file to the $reportIDHashFilePath and it didn't open as XML with a receipt or error.
                #This should be our file so lets rename to the actual full filepath.
                Move-Item $reportIDHashFilePath $fullfilepath -Force

            } else {
                #We sould have downloaded a file to the $reportIDHashFilePath and it didn't open as XML with a receipt or error.
                #This should be our file so lets rename to the actual full filepath.
                Move-Item $reportIDHashFilePath $fullfilepath -Force
            }
        }
        
        Write-Host "Success." -ForegroundColor Yellow
    } catch {
        Write-Host "Failed to download file. $($_)" -ForegroundColor Red
        Send-Email("[Failure][Download Failed]","Failed to download file. $($_)")
        Reset-DownloadedFile($fullfilepath)
        exit(6)
    }
} else {
    #Just showing report details. No reason to continue.
    Write-Host "Skip downloading file specified. Exiting..." -ForegroundColor Yellow
    exit(0)
}

# check file for proper format of formats with rows.
if (-Not($DisableCSVVerification)) {
    if (@('csv','xlsx','xml') -contains $extension) {
        $FileExists = Test-Path $fullfilepath
        if ($FileExists -eq $False) {
            Write-Host("File does not exist:" + $fullfilepath)
            Send-Email("[Failure][Output]","File Did not download to expected path.")
            exit(7) #File didn't download to expected path
        }
        
        try {
            
            if ($extension -eq "csv") {
                $fileContents = Import-CSV $fullfilepath
            } elseif ($extension -eq "xlsx") {    
                #Verify XLSX file.
                if (Get-Command Import-Excel -ErrorAction SilentlyContinue) {
                    $fileContents = Import-Excel $fullfilepath
                } else {
                    Write-Host "Notify: You did not specify the parameter -DisableCSVVerification but you don't have the Import-Excel module to verify this xlsx file."
                }
            } elseif ($extension -eq "xml") {
                #Verify XML file.
                $fileContents = ([xml](Get-Content $fullfilepath)).dataset.data.row
                if ($fileContents.Count -ge $requiredlinecount) {
                    #XML does not have NoteProperty. We have to trust that returned rows count as data.
                    #we could check $fileContents.dataset.data.row[0].value | Measure-Command but honestly since it doesn't have named values it seems kinda useless.
                    exit(0)
                }
            }
            
            $headercount = ($fileContents | Get-Member | Where-Object { $PSItem.MemberType -eq 'NoteProperty' } | Select-Object -ExpandProperty Name | Measure-Object).Count
            if ($headercount -gt 1) {
                Write-Host("Passed CSV header check with $headercount headers...") -ForeGroundColor Yellow
            } else {
                Write-Host("Failed CSV header check with only $headercount headers...") -ForeGroundColor Yellow
                Reset-DownloadedFile($fullfilepath)
                Send-Email("[Failure][Verify]","Only $headercount header found in CSV.")
                exit(8)
            }

            $linecount = ($fileContents | Measure-Object).Count
            if ($linecount -ge $requiredlinecount) { #Think schools.csv for smaller districts with only 3 campuses.
                Write-Host("Passed CSV line count with $linecount lines...") -ForeGroundColor Yellow
            } else {
                Write-Host("Failed CSV line count with only $linecount lines...") -ForeGroundColor Yellow
                Reset-DownloadedFile($fullfilepath)
                Send-Email("[Failure][Verify]","Only $linecount lines found in CSV.")
                exit(9)
            }

            if ($extension -eq "csv" -and $TrimCSVWhiteSpace) {
                if ($PSVersionTable.PSVersion -lt [version]"7.1.0") {
                    Write-Host "Error: You specified you wanted to remove the CSV Whitespaces but his requires Powershell 7.1. Not modifying downloaded file." -ForegroundColor RED
                } else {

                    Write-Host "Info: Cleaning up white spaces in CSV."
                    $fileContents | Foreach-Object {  
                        $_.PSObject.Properties | Foreach-Object {
                            $_.Value = $_.Value.Trim()
                        }
                    }

                    if ($CSVUseQuotes) {
                        Write-Host "Info: Exporting CSV using quotes."
                        $fileContents | Export-Csv -UseQuotes Always -Path $fullfilepath -Force
                    } else {
                        $fileContents | Export-Csv -UseQuotes AsNeeded -Path $fullfilepath -Force
                    }

                }
            }

        } catch {
            Write-Host "Error: Unable to verify downloaded file. Is it possible the $extension file is empty? $($_)"
            Send-Email("[Failure][Verify]","Error: Unable to verify downloaded file. Is it possible the $extension file is empty? $_")
            Reset-DownloadedFile($fullfilepath)
            exit(10) #General Verification Failure
        }
    }

}

#need a valid exit here so this script can be put into a loop in case a file fails to download on first try
exit(0)