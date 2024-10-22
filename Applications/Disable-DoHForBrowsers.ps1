<#
This script disables DNS over HTTPS on all major browsers.

Andy Morales
#>

#Chrome
REG ADD "HKLM\SOFTWARE\Policies\Google\Chrome" /v "DnsOverHttpsMode" /t REG_SZ /d 'off' /f

#Edge
REG ADD "HKLM\SOFTWARE\Policies\Microsoft\Edge" /v "DnsOverHttpsMode" /t REG_SZ /d 'off' /f

#FireFox
REG ADD "HKLM\SOFTWARE\Policies\Mozilla\Firefox\DNSOverHTTPS" /v "Enabled" /t REG_DWORD /d '0' /f
REG ADD "HKLM\SOFTWARE\Policies\Mozilla\Firefox\DNSOverHTTPS" /v "Locked" /t REG_DWORD /d '1' /f