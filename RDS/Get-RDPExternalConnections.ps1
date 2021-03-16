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

#Don't do anything if DUO is installed
if ((Get-InstalledApplications).displayName -Contains 'Duo Authentication for Windows Logon x64') {
    Write-Output 'DUO installed'
}
else{
    $Events = Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational','Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin'

    #Filter by event ID
    $FilteredEvents = @()

    foreach ($evt in $events) {
        if ($evt.id -eq '1149' -or $evt.id -eq '1158') {
            $FilteredEvents += $evt
        }
    }

    #Find events that contain public IPs
    $ExternalEvents = @()

    $rfc1918regex = '(192\.168\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(172\.([1][6-9]|[2][0-9]|[3][0-1])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(10\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))|(127\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5])\.([0-9]|[0-9][0-9]|[0-2][0-5][0-5]))'

    foreach ($fEvent in $FilteredEvents) {
        if ($fEvent.Message -notmatch $rfc1918regex) {
            $ExternalEvents += $fEvent
        }
    }

    if($ExternalEvents.Count -gt 0){
        Write-Output "External RDP events found on $($env:COMPUTERNAME). Most Recent:"
        Write-Output $FilteredEvents.Message[0..5]
    }
    else{
        Write-Output 'No events found'
    }
}