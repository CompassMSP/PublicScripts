#region CreateSecureFolder
#It is a best practice to wipe the folder once you are done running the script that relies on it.
$FolderDirectory = 'C:\Temp\LDAPS'

New-Item -Path $FolderDirectory -ItemType Directory -Force | Out-Null

#Remove all explicit permissions
ICACLS ("$FolderDirectory") /reset | Out-Null

#Add SYSTEM permission
ICACLS ("$FolderDirectory") /grant ("SYSTEM" + ':(OI)(CI)F') | Out-Null

#Give Administrators Full Control
ICACLS ("$FolderDirectory") /grant ("Administrators" + ':(OI)(CI)F') | Out-Null

#Disable Inheritance on the Folder. This is done last to avoid permission errors.
ICACLS ("$FolderDirectory") /inheritance:r | Out-Null
#endregion CreateSecureFolder

$ServerFQDN = "$($ENV:COMPUTERNAME).$((Get-WmiObject Win32_ComputerSystem).Domain)"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/request.inf') -replace "`n", "`r`n") -replace 'SERVER.example.com', "$ServerFQDN" |  Out-File -FilePath 'C:\Temp\LDAPS\request.inf'
((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/v3ext.txt') -replace "`n", "`r`n") -replace 'SERVER.example.com', "$ServerFQDN"  | Out-File -FilePath 'C:\Temp\LDAPS\v3ext.txt'
(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/enable_ldaps.txt') -replace "`n", "`r`n" | Out-File -FilePath 'C:\Temp\LDAPS\enable_ldaps.txt'

#Import the root CA
Import-Certificate -FilePath 'C:\Temp\LDAPS\ca.crt' -CertStoreLocation 'Cert:\LocalMachine\Root' -Verbose

#Request cert
certreq -new request.inf ad.csr


###TO DO
#Approve the cert request
openssl x509 -req -days 825 -in 'C:\Temp\LDAPS\ad.csr' -CA 'C:\Temp\LDAPS\ca.crt' -CAkey 'C:\Temp\LDAPS\ca.key' -extfile 'C:\Temp\LDAPS\v3ext.txt' -set_serial 01 -out 'C:\Temp\LDAPS\ad_ldaps_cert.crt'
