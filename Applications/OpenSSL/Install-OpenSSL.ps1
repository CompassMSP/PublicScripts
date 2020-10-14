<#
This script installs openSSL.

Andy Morales
#>

$OpenSslUrl = 'https://slproweb.com/download/Win64OpenSSL_Light-1_1_1h.msi'
$OpenSslMsiPath = 'C:\Windows\Temp\OpenSSL.msi'

#Download Application
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

(New-Object System.Net.WebClient).DownloadFile("$OpenSslUrl", "$OpenSSLMSI")

Start-Process msiexec.exe -Wait -ArgumentList "/i $($OpenSslMsiPath) /qn" -PassThru

#Configure ENV

[System.Environment]::SetEnvironmentVariable('OPENSSL_CONF', 'C:\Program Files\OpenSSL-Win64\bin\openssl.cfg', [System.EnvironmentVariableTarget]::Machine)

[System.Environment]::SetEnvironmentVariable('Path', "$($env:Path);C:\Program Files\OpenSSL-Win64\bin", [System.EnvironmentVariableTarget]::Machine)