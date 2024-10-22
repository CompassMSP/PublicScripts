<#
This script installs the Cisco Umbrella root CA certificate on the computer store.

Andy Morales
#>
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

(New-Object System.Net.WebClient).DownloadFile('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/Applications/Cisco/Cisco_Umbrella_Root_CA.cer', 'C:\Windows\Temp\Cisco_Umbrella_Root_CA.cer')

CERTUTIL -addStore -enterprise -f -v root "C:\Windows\Temp\Cisco_Umbrella_Root_CA.cer"

#Force Firefox to use local cert store
REG  add  "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Mozilla\Firefox\Certificates"  /v  "ImportEnterpriseRoots" /t  REG_DWORD  /d  1 /f