<#
.DESCRIPTION
This script looks through RDP logs in order to find logins from public IPs. If one is found
then it is likely that the computer has RDP exposed to the internet.

The script will not alert if DUO is installed.

Events older than $LogThreshold will be ignored

.LINK
https://dfironthemountain.wordpress.com/2019/02/15/rdp-event-log-dfir/

Andy Morales
#>

$LogThreshold = (Get-Date).AddDays(-90)

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

    $ExternalEvents = @()

    $rfc1918regex = '^(127(?:\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$)|(10(?:\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$)|(192\.168(?:\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){2}$)|(172\.(?:1[6-9]|2\d|3[0-1])(?:\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){2}$)'
    #simple regex that just looks for ':' in IPs
    $IPv6regex = ':'

    #region remoteConnectionAdminEvents
    $RemoteConnectionAdminLogFilter = @{
        LogName   = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Admin'
        ID        = 1158
        StartTime = $LogThreshold
    }

    $RemoteConnectionAdminEvents += @(Get-WinEvent -FilterHashtable $RemoteConnectionAdminLogFilter -ErrorAction SilentlyContinue)

    #Find events that contain public IPs
    if ($RemoteConnectionAdminEvents.count -gt 0) {
        foreach ($rcaEvent in $RemoteConnectionAdminEvents) {

            [xml]$rcaEntry = $rcaEvent.ToXml()
            if ([string]::IsNullOrEmpty($rcaEntry.Event.UserData.EventXML.Param1)) {
                #discard empty param
            }
            elseif ($rcaEntry.Event.UserData.EventXML.Param1 -match $IPv6regex) {
                #discard IPv6
            }
            elseif ($rcaEntry.Event.UserData.EventXML.Param1 -notmatch $rfc1918regex) {
                $ExternalEvents += $rcaEvent
            }
        }
    }
    #endregion remoteConnectionAdminEvents

    #region remoteConnectionOperationalEvents
    $RemoteConnectionOperationalLogFilter = @{
        LogName   = 'Microsoft-Windows-TerminalServices-RemoteConnectionManager/Operational'
        ID        = 1149
        StartTime = $LogThreshold
    }

    $RemoteConnectionOperationalEvents += @(Get-WinEvent -FilterHashtable $RemoteConnectionOperationalLogFilter -ErrorAction SilentlyContinue)

    #Find events that contain public IPs
    if ($RemoteConnectionOperationalEvents.count -gt 0) {
        foreach ($rcoEvent in $RemoteConnectionOperationalEvents) {

            [xml]$rcoEntry = $rcoEvent.ToXml()
            if ([string]::IsNullOrEmpty($rcoEntry.Event.UserData.EventXML.Param3)) {
                #discard empty param
            }
            elseif ($rcoEntry.Event.UserData.EventXML.Param3 -match $IPv6regex) {
                #discard IPv6
            }
            elseif ($rcoEntry.Event.UserData.EventXML.Param3 -notmatch $rfc1918regex) {
                $ExternalEvents += $rcoEvent
            }
        }
    }
    #endregion remoteConnectionOperationalEvents

    #output results
    if ($ExternalEvents.Count -gt 0) {
        $OutputText = "External RDP events found on $($env:COMPUTERNAME). "

        #Get Computer Type
        switch ((Get-WmiObject Win32_ComputerSystem ).domainRole) {
            0 {
                $OutputText += "StandaloneWorkstation`n`n"
            }
            1 {
                $OutputText += "DomainWorkstation`n`n"
            }
            2 {
                if ((Get-WmiObject -Class Win32_TerminalServiceSetting -Namespace 'root\cimv2\TerminalServices').terminalServerMode -eq 1) {
                    $OutputText += "RDSH`n`n"
                }
                else {
                    $OutputText += "StandardServer`n`n"
                }
            }
            3 {
                if ((Get-WmiObject -Class Win32_TerminalServiceSetting -Namespace 'root\cimv2\TerminalServices').terminalServerMode -eq 1) {
                    $OutputText += "RDSH`n`n"
                }
                else {
                    $OutputText += "StandardServer`n`n"
                }
            }
            4 {
                $OutputText += "DomainController`n`n"
            }
            5 {
                $OutputText += "DomainController`n`n"
            }
        }

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

        $OutputText += @($ExternalEvents.Message)[0..3] | Out-String

        $OutputText += "`n"

        #region sessionManagerLogs
        $SessionManagerLogFilter = @{
            LogName   = 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational'
            ID        = 21, 22, 25
            StartTime = $LogThreshold
        }

        $SessionManagerEvents = Get-WinEvent -FilterHashtable $SessionManagerLogFilter -ErrorAction SilentlyContinue

        $ExternalSmEvents = @()
        foreach ($smEvent in $SessionManagerEvents) {

            [xml]$smEntry = $smEvent.ToXml()
            if ([string]::IsNullOrEmpty($smEntry.Event.UserData.EventXML.Address)) {
                #discard empty param
            }
            if ($smEntry.Event.UserData.EventXML.Address -match $IPv6regex) {
                #Discard IPv6
            }
            elseif ($smEntry.Event.UserData.EventXML.Address -notmatch $rfc1918regex) {
                [array]$ExternalSmEvents += [PSCustomObject]@{
                    TimeCreated = $smEvent.TimeCreated
                    User        = $smEntry.Event.UserData.EventXML.User
                    IPAddress   = $smEntry.Event.UserData.EventXML.Address
                }
            }
        }

        if ($ExternalSmEvents.count -gt 0) {
            $OutputText += "Recent Logins"

            $OutputText += @($ExternalSmEvents)[0..3] | Out-String
        }
        #endregion sessionManagerLogs

        Write-Output $OutputText
    }
    else {
        Write-Output 'No events found'
    }
}