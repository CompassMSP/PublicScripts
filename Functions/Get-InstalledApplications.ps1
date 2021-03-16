Function Get-InstalledApplications {
    $InstalledApplications = @()
    $UninstallKeys = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'

    #Check for x86 software on x64 Windows
    if (Test-Path -Path 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
        $UninstallKeys += Get-ChildItem 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    }

    Foreach ($SubKey in $UninstallKeys) {

        $DisplayName = (Get-ItemProperty -Path "Registry::$($SubKey.Name)" -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
        if ([string]::IsNullOrEmpty($DisplayName)) {
        }
        else {
            $InstalledApplications += [PSCustomObject]@{
                DisplayName = $DisplayName
            }
        }

    }
    Return $InstalledApplications | Sort-Object displayName
}