<#
This script will identify UPDs and FSL disks that are low on space. The threshold can be controlled through $LowSpaceThreshold.

The script should be configured to run 2-3 times on an RDS server. The RMM should be configured to send an alert if the output contains "LowSpaceProfile"

Andy Morales
#>
$LowSpaceThreshold = '5GB'

function Test-RegistryValue {
    #Modified version of the function below
    #https://www.jonathanmedd.net/2014/02/testing-for-the-presence-of-a-registry-key-and-value.html
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true,
            Position = 1,
            HelpMessage = 'Registry::HKEY_LOCAL_MACHINE\SYSTEM')]
        [ValidatePattern('Registry::.*')]
        [ValidateNotNullOrEmpty()]
        [String]$Path,

        [parameter(Mandatory = $true,
            Position = 2)]
        [ValidateNotNullOrEmpty()]
        [string]$Name,

        [parameter(Position = 3)]
        [ValidateNotNullOrEmpty()]$ValueData
    )
    try {
        if ($ValueData) {
            if ((Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop) -eq $ValueData) {
                return $true
            }
            else {
                return $false
            }
        }
        else {
            $RegKeyCheck = Get-ItemProperty -Path $Path -ErrorAction Stop | Select-Object -ExpandProperty $Name -ErrorAction Stop | Out-Null
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

#Check to see if FSL or UPDs are enabled
If (
    (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\FSLogix\Profiles' -Name 'Enabled' -ValueData '1') -or
    (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings' -Name 'UvhdEnabled' -ValueData '1')){

    $AttachedProfileDisks = Get-CimInstance -Class win32_volume -Filter "Label like 'O365-%' OR Label like 'Profile-%' OR Label like 'User Disk'"

    [array]$LowDiskSpaceDisks = $AttachedProfileDisks | Where-Object { $_.FreeSpace -lt $LowSpaceThreshold }

    if ($LowDiskSpaceDisks.count -gt 0) {

        $AllLowSpaceDisks = @()

        Foreach ($Disk in $LowDiskSpaceDisks) {
            #UPDs
            if ($Disk.Name -like 'C:\Users\*') {
                $AllLowSpaceDisks += $Disk.Name.ToString()
            }
            #FSLogix
            elseif (($Disk.Label -like 'Profile-*') -or ($Disk.Label -like 'O365-*')) {
                $AllLowSpaceDisks += $Disk.label.ToString()

            }
        }
        Return "LowSpaceProfile: $Env:COMPUTERNAME $($AllLowSpaceDisks -join ',')"
    }

}
