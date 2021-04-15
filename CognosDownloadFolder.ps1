#Requires -Version 7.1
<#

Craig Millsap - 4/15/2021

This script only works on folders in the Team Content Folder.

This is an example ONLY! You will need to customize this for your own needs.

You must include any possible prompt for all reports in the folder or it will fail.  Each report either needs to take the same parameter or if you need to pass different parameters
then the prompt will need to be named something else.

HIGHLY RECOMMEND using the CognosDefaults.ps1 file to set your username, espdsn, and passwordfile path.

.\CognosDownloadFolder.ps1 -username 0403cmillsap -espdsn gentrysms -cognosfolder "_Share Temporarily Between Districts/Gentry/automation" -reportparams "p_year=2021&p_anythingelse=12345" 

#>

Param(
    [parameter(Mandatory=$false,HelpMessage="Give a specific folder to download the report into.")]
        [string]$savepath="C:\scripts",
    [parameter(Mandatory=$true,HelpMessage="Cognos Folder Structure.")]
        [string]$cognosfolder, #Cognos Folder "Folder 1/Sub Folder 2/Sub Folder 3" NO TRAILING SLASH
    [parameter(Mandatory=$false,HelpMessage="Report Parameters")]
        [string]$reportparams, #If a report requires parameters you can specifiy them here. Example:"p_year=2017&p_school=Middle School"
    [parameter(Mandatory=$false,HelpMessage="eSchool SSO username to use.")]
        [string]$username, #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
    [parameter(Mandatory=$false,HelpMessage="File for ADE SSO Password")]
        [string]$passwordfile="C:\Scripts\apscnpw.txt", # Override where the script should find the password for the user specified with -username.
    [parameter(Mandatory=$false,HelpMessage="eSchool DSN location.")]
        [string]$espdsn #YOU SHOULD NOT MODIFY THIS. USE THE PARAMETER. CONSIDER USING THE CognosDefaults.ps1 OVERRIDES.
)

#Establish Session Only. Report parameter is required but we can provide a fake one for authentication only.
. .\CognosDownload.ps1 -username $username -espdsn $espdsn -report FAKE -EstablishSessionOnly -cognosfolder "$cognosfolder" -TeamContent -savepath "$savepath" -reportparams "$reportparams"

# $cognosfolderEncoded = "Team Content/Student Management System/$($cognosfolder)".Replace(' ','%20')

try {
    $foldercontents = Invoke-WebRequest -Uri "$baseURL/ibmcognos/bi/v1/disp/rds/wsil/path/Team Content/Student Management System/$cognosfolder" -WebSession $session
    $files = Select-Xml -Xml ([xml]$foldercontents.Content) -XPath '//x:service' -Namespace @{ x = "http://schemas.xmlsoap.org/ws/2001/10/inspection/" }
} catch {
    $PSItem
    exit(1)
}

$results = $files | ForEach-Object -Parallel {

    $incomingsession = $using:session
    $username = $using:username
    $espdsn = $using:espdsn
    $cognosfolder = $using:cognosfolder
    $savepath = $using:savepath
    $reportparams = $using:reportparams

    $reportname = ($PSItem.Node).name
    
    .\CognosDownload.ps1 -username $username -espdsn $espdsn -report "$reportname" -cognosfolder "$cognosfolder" -TeamContent -SessionEstablished -savepath "$savepath" -reportparams "$reportparams" -ShowReportDetails -TrimCSVWhiteSpace
    
    if ($LASTEXITCODE -ne 0) { throw }

} -AsJob -ThrottleLimit 5 | Wait-Job #Please don't overload the Cognos Server.

$results.ChildJobs | Where-Object { $PSItem.State -eq "Completed" } | Receive-Job

#Output any failed jobs information.
$failedJobs = $results.ChildJobs | Where-Object { $PSItem.State -ne "Completed" }
$failedJobs | ForEach-Object {
    $PSItem | Receive-Job
}

if (($failedJobs | Measure-Object).count -ge 1) {
    Write-Host "Failed running", (($failedJobs | Measure-Object).count), "jobs." -ForegroundColor RED
}

#profit.