Function New-LDAPSCert {
    <#
    This script requests and installs a certificate for LDAPS on a machine.

    Requirements:
        OpenSSL must be installed (use Install-OpenSSL)
        The CA.cer file must be copied to C:\Temp\LDAPS
        The ca.key file must be copied to C:\Temp\LDAPS

    Andy Morales
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$KeyPass, #ideally this should be a secure string, but the RMM won't accept it

        [parameter(Mandatory = $false)]
        [switch]$InstallOpenSSL
    )

    $LDAPWorkingDirectory = 'C:\Temp\LDAPS'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if ($InstallOpenSSL){
        Invoke-Expression(New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/Applications/OpenSSL/Install-OpenSSL.ps1')
    }

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

    $FilesToDownload = @(
        'https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/request.inf',
        'https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/v3ext.txt',
        'https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/LDAPS/enable_ldaps.txt'
    )

    Foreach ($File in $FilesToDownload) {
        $String = ((New-Object Net.WebClient).DownloadString("$File") -replace "`n", "`r`n") -replace 'SERVER.example.com', "$ServerFQDN"

        $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
        [System.IO.File]::WriteAllLines("$LDAPWorkingDirectory\$(($File -split '/')[-1])", $String, $Utf8NoBomEncoding)
    }

    #Import the root CA
    Import-Certificate -FilePath "$LDAPWorkingDirectory\ca.crt" -CertStoreLocation 'Cert:\LocalMachine\Root' -Verbose

    #Request cert
    certreq -f -new "$LDAPWorkingDirectory\request.inf" "$LDAPWorkingDirectory\ad.csr"

    #Approve the cert request
    & "C:\Program Files\OpenSSL-Win64\bin\openssl.exe" x509 -req -days 825 -in "$LDAPWorkingDirectory\ad.csr" -CA "$LDAPWorkingDirectory\ca.crt" -CAkey "$LDAPWorkingDirectory\ca.key" -extfile "$LDAPWorkingDirectory\v3ext.txt" -set_serial 01 -out "$LDAPWorkingDirectory\ad_ldaps_cert.crt" -passin "pass:$($KeyPass)"

    #Install the certificate
    certreq -accept "$LDAPWorkingDirectory\ad_ldaps_cert.crt"

    #Tell AD to start using LDAPS
    ldifde -i -f "$LDAPWorkingDirectory\enable_ldaps.txt"

    #Delete local files
    Remove-Item -Path $LDAPWorkingDirectory -Recurse -Force
    Remove-Variable KeyPass -Force
}
