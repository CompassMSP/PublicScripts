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
        [ValidateNotNullOrEmpty()]$ValueData
    )

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

#Checks to see if Duo GW is set to enabled
$DuoGatewayInstalled = Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Duo Security\DuoTsg' -Name EnableDuo -ValueData 1

#This checks for the DUO DLL GUID. I did some testsing and it was identical across computers in the same domain
#Some testing should be done to see if it's the same across other domains
$DuoForRdpInstalled = Test-Path -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{44E2ED41-48C7-4712-A3C3-250C5E6D5D84}'

if($DuoGatewayInstalled -and $DuoForRdpInstalled){
    Return "DUO GW + RDP"
}
elseif($DuoGatewayInstalled){
    Return "DUO GW"
}
elseif($DuoForRdpInstalled){
    Return "DUO RDP"
}
else{
    Write-Verbose "DUO not detected"
    Return ""
}