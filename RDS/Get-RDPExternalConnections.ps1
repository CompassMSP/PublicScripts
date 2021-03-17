<#
.LINK
https://dfironthemountain.wordpress.com/2019/02/15/rdp-event-log-dfir/
#>
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
function Test-RegistryValue {
    <#
    Checks if a reg key/value exists

    #Modified version of the function below
    #https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html

    Andy Morales
    #>

    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            Position = 1,
            HelpMessage = 'HKEY_LOCAL_MACHINE\SYSTEM')]
        [ValidatePattern('Registry::.*|HKEY_')]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [parameter(Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [parameter(Position = 3)]
        $ValueData
    )

    Set-StrictMode -Version 2.0

    #Add RegDrive if it is not present
    if ($Path -notmatch 'Registry::.*') {
        $Path = 'Registry::' + $Path
    }

    try {
        #Reg key with value
        if ($ValueData) {
            if ((Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop) -eq $ValueData) {
                return $true
            }
            else {
                return $false
            }
        }
        #Key key without value
        else {
            $RegKeyCheck = Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop
            if ($null -eq $RegKeyCheck) {
                #if the Key Check returns null then it probably means that the key does not exist.
                return $false
            }
            else {
                return $true
            }
        }
    }
    catch {
        return $false
    }
}

#Don't do anything if DUO is installed
if ((Get-InstalledApplications).displayName -Contains 'Duo Authentication for Windows Logon x64') {
    Write-Output 'DUO installed'
}
else {
    $Events = Get-WinEvent -LogName 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational', 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin'

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



    if ($ExternalEvents.Count -gt 0) {
        $OutputText = "External RDP events found on $($env:COMPUTERNAME).`n`n"

        #Check for NLA setting
        $NLASetting = Test-RegistryValue -Path 'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name 'UserAuthentication' -ValueData 1

        $NLAPolicy = Test-RegistryValue -Path 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' -Name 'UserAuthentication' -ValueData 1

        if ($NLAPolicy) {
            #NLA enabled
        }
        elseif ($NLASetting) {
            #NLA Enabled
        }
        else {
            $OutputText += "NLA is not enabled`n`n"
        }

        $OutputText += "Most Recent:`n"

        $OutputText += @($FilteredEvents.Message)[0..3] | Out-String

        Write-Output $OutputText
    }
    else {
        Write-Output 'No events found'
    }
}