<#
This script installs the Cisco Umbrella root CA certificate on the computer store

Andy Morales
#>
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

(New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/Applications/Cisco/Cisco_Umbrella_Root_CA.cer', 'C:\Windows\Temp\Cisco_Umbrella_Root_CA.cer')

CERTUTIL -addStore -enterprise -f -v root "C:\Windows\Temp\Cisco_Umbrella_Root_CA.cer"