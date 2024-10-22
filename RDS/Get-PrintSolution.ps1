<#
This script will identify which (if any) printing solution is installed on the machine.

The script will return a blank space if nothing is found. The goal of this is to set the computer property to blank in the event that the application is removed.

Andy Morales
#>

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
    if ($Path -notmatch 'Registry::.*'){
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

$PrintSolution = @()

If (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Tricerat\Simplify Console\External Tools' -Name 'Menu0' -ValueData RegDiff) {
    #Simplify Print Console
    $PrintSolution += 'Simplify Print Console'
}

if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Tricerat\Simplify Printing\ScrewDrivers Print Server v6' -Name Port) {
    #Simplify Print Server
    $PrintSolution += 'Simplify Print Server'
}

if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Tricerat\Simplify Printing' -Name dwProviderAvailable -ValueData 1) {
    #Simplify Printing
    $PrintSolution += 'Simplify Printing'
}

if (Test-RegistryValue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Tricerat\Simplify Printing\ScrewDrivers Server v6' -Name StandAlone -ValueData 1) {
    #ScrewDrivers Redirection
    $PrintSolution += 'ScrewDrivers Redirection'
}

if (Test-Path "$env:ProgramFiles\PaperCut MF Client\pc-client.exe") {
    #Papercut
    $PrintSolution += 'PaperCut'
}

Return $PrintSolution -join ' '
