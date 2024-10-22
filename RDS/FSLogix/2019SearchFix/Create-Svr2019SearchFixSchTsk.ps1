<#
Creates a scheduled task that restarts the search index service every time a user logs out.

Andy Morales
#>
#https://jkindon.com/2020/03/15/windows-search-in-server-2019-and-multi-session-windows-10/

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

    #Add Regdrive if it is not present
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

$CreateTask = $false
if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles' -Name 'Enabled' -ValueData '1') {
    #FSL Profiles
    $CreateTask = $true
}

if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings' -Name 'UvhdEnabled' -ValueData '1') {
    #UPD
    $CreateTask = $true
}

if ($CreateTask) {
    [string]$XMLTask = (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/RDS/FSLogix/2019SearchFix/2019ResetSearchOnLogoff.xml') | Out-String
    Register-ScheduledTask -XML $XMLTask -TaskName 'Reset Search on Logoff'
}
