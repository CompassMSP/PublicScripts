<#
Checks to see if AD Connect is installed on a computer

Andy Morales
#>
if(Test-Path -Path "C:\Program Files\Microsoft Azure Active Directory Connect\AzureADConnect.exe"){
    RETURN 'AD Connect Installed'
}