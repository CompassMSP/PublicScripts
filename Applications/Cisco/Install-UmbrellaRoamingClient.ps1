Function Install-UmbrellaRoamingClient {
    <#
    This script downloads the latest version of Umbrella and installs it.

    Andy Morales
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [String]$ORG_ID,

        [parameter(Mandatory = $true)]
        [String]$ORG_FINGERPRINT,

        [parameter(Mandatory = $true)]
        [String]$USER_ID,

        [parameter(Mandatory = $false)]
        [switch]$HIDE_UI
    )

    #Check to see if the computer is a domain controller
    if ((Get-WmiObject Win32_ComputerSystem).DomainRole -le 3) {
        $MSIDestination = 'C:\Windows\Temp\UmbrellaSetup.msi'

        #Download MSI
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        (New-Object System.Net.WebClient).DownloadFile('http://shared.opendns.com/roaming/enterprise/release/win/production/Setup.msi', "$MSIDestination")

        #Install the application
        $MSIParams = "/i $($MSIDestination) /qn ORG_ID=$($ORG_ID) ORG_FINGERPRINT=$($ORG_FINGERPRINT) USER_ID=$($USER_ID)"

        if ($HIDE_UI) {
            $MSIParams += 'HIDE_UI=1'
        }

        Start-Process msiexec.exe -Wait -ArgumentList $MSIParams -PassThru

        #Install root CA
        Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/Applications/Cisco/Install-UmbrellaRootCA.ps1')
    }
    else {
        Write-Output 'Computer is a domain controller. The script will exit.'
    }
}