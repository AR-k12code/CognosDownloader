#Requires -Version 7.1
#This requires Powershell 7.1 to run the parrallel ForEach-Object.

#This is an example ONLY! You will need to customize this for your own needs.

#These variables will be overwritten if you use the CognosDefaults.ps1 file.
$username = '0403cmillsap'
$espdsn = 'gentrysms'

$reports = @{
    'enrollments' = @{ 'parameters' = ''; 'folder' = 'automation'; 'savepath' = 'c:\scripts\' }
    'schools' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'sections' = @{ 'parameters' = "p_year=2021"; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'students' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'teachers' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'students_extras' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'contacts' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'facultyids' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'activities' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
    'transportationx' = @{ 'parameters' = ''; 'folder' = 'automation';  'savepath' = 'c:\scripts\' }
}

#Establish Session Only. Report paramter is required but we can provide a fake one for authentication only.
. .\CognosDownload.ps1 -username $username -espdsn $espdsn -report FAKE -EstablishSessionOnly

#Look throught he hash table, pull in session, use established session.

$results = $reports.Keys | ForEach-Object -Parallel  {
    #report title
    #$PSitem
    
    #pull in session to script block
    $incomingsession = $using:session
    
    #Pull in properties for each hashtable key.
    $options = ($using:reports).$PSItem

    #Run Cognos Download using incoming options.
    .\CognosDownload.ps1 -username 0403cmillsap -espdsn gentrysms -report $PSItem -cognosfolder "$($options.folder)" -SessionEstablished -savepath "$($options.savepath)" -reportparams "$($options.parameters)" -SkipDownloadingFile -ShowReportDetails

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