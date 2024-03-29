Param(
    [parameter(Mandatory=$false,HelpMessage="File for ADE SSO Password")]
    [string]$passwordfile="C:\Scripts\apscnpw.txt"
)

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
if (Test-Path $currentPath\CognosDefaults.ps1) {
    . $currentPath\CognosDefaults.ps1
} elseif ((Test-Path $currentPath\..\CognosDefaults.ps1)) {
    . $currentPath\..\CognosDefaults.ps1
}

try {
    #Remove old password file. If this fails to remove the file we don't have NTFS permissions and script will stop.    
    If ((Test-Path ($passwordfile))) {
        Remove-Item $passwordfile -Force
    }

    Read-Host "Enter new password" -AsSecureString | ConvertFrom-SecureString | Out-File $passwordfile

    Write-Host "You have successfully updated the saved password to ""$($passwordfile)""."
} catch {
    Write-Host "Failed to change password. Please make sure you have permissions to delete/create ""$($passwordfile)""." -ForegroundColor RED
}


