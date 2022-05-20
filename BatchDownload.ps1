param (
    [parameter(Mandatory=$false,HelpMessage="CSV file containing Cognos report names and parameters")][string]$csvFile=".\BatchDownload.csv"
)

$switchParameters = @(
    "ShowReportDetails",
    "SkipDownloadingFile",
    "dev",
    "TeamContent",
    "SavePrompts",
    "DisableCSVVerification",
    "RunReport",
    "ReportStudio",
    "EstablishSessionOnly",
    "SessionEstablished",
    "TrimCSVWiteSpace",
    "CSVUseQuotes"
)

$ignoreParameters = @(
    "EstablishSessionOnly",
    "SessionEstablished"
)

$lineArray = @()

if (Test-Path $csvFile) {
    $csv = Import-Csv -Path $csvFile
} else {
    Throw "Could not find 'BatchDownload.csv' file in the current directory.  Please provide the name of a valid CSV file using the -csvFile command line parameter"
}

foreach ($row in $csv) {
    if (![string]::IsNullOrWhiteSpace($row.report)) {
        $commandLine = ".\CognosDownload.ps1"
        foreach ($property in $row.PSObject.Properties) {
            if ($switchParameters.Contains($property.Name)) {
                if (![string]::IsNullOrWhiteSpace($property.Value) -and !$ignoreParameters.Contains($property.Name)) {
                    $commandLine += " -$($property.Name)"
                }
            } else {
                if (![string]::IsNullOrWhiteSpace($property.Value)) {
                    if ($property.Value -notmatch "^[a-z0-9]+$") {$property.Value = """$($property.Value)"""}
                    $commandLine += " -$($property.Name) $($property.Value)"
                }
            }
        }
    }
    $lineArray += "$commandLine -SessionEstablished"
}

function Get-Asynchronously($commands) {

    . .\CognosDownload.ps1 -EstablishSessionOnly -report none -savepath .\

    Write-Host "Preparing $($commands.Count) reports - Please wait..." -ForegroundColor Green
    
    $commands | ForEach-Object -Parallel {
        $incomingSession = $using:session
        Invoke-Expression -Command $_
    } -AsJob -ThrottleLimit 5 | Wait-Job | Out-Null
    
    Get-Job | Receive-Job
    Remove-Job -State Completed
}

function Get-Synchronously($commands) {

    Write-Host "Downloading all reports synchronously.  For better performance, conside upgrading to PowerShell version 7 or later." -ForegroundColor Green

    . .\CognosDownload.ps1 -EstablishSessionOnly -report none -savepath .\

    foreach ($line in $commands) {
        $incomingSession = $session
        Invoke-Expression -Command $line
    }
}

if ($Host.Version -lt [Version]'7.0') {Get-Synchronously($lineArray)} else {Get-Asynchronously($lineArray)}