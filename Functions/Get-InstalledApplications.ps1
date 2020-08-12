Function Get-InstalledApplications {
    $InstalledApplications = @()
    $UninstallKeys = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    Foreach ($SubKey in $UninstallKeys) {

        $DisplayName = (Get-ItemProperty -Path "Registry::$($SubKey.Name)" -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
        if ([string]::IsNullOrEmpty($DisplayName)) {
        }
        else{
            $InstalledApplications += [PSCustomObject]@{
                DisplayName = $DisplayName
            }
        }

    }
    Return $InstalledApplications
}