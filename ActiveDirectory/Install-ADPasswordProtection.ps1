
#https://haveibeenpwned.com/Passwords
$HIBPHashesLink = 'https://downloads.pwnedpasswords.com/passwords/pwned-passwords-ntlm-ordered-by-hash-v6.7z'
$HIBPHashesCompressedFile = 'C:\Temp\pwned-passwords-ntlm-ordered-by-hash.7z'
$HIBPHashesExtractDir = 'C:\Temp\HIBP'
$PasswordProtectionMSIFile = 'C:\Windows\Temp\Lithnet.ActiveDirectory.PasswordProtection.msi'

#Clean up any old files
Function Remove-OldFiles {
    $ItemsToDelete = @(
        $HIBPHashesCompressedFile,
        $HIBPHashesExtractDir
    )

    foreach ($item in $ItemsToDelete) {
        Remove-Item $item -Force -Recurse -ErrorAction SilentlyContinue
    }
}

Remove-OldFiles

if ((Get-PSDrive C).free -gt 30GB) {
    #Download and install MSI
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    (New-Object System.Net.WebClient).DownloadFile('https://github.com/lithnet/ad-password-protection/releases/latest/download/Lithnet.ActiveDirectory.PasswordProtection.msi', "$PasswordProtectionMSIFile")

    msiexec.exe /i 'C:\Windows\Temp\Lithnet.ActiveDirectory.PasswordProtection.msi' /qn

    #Install 7zip module
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Set-PSRepository -Name 'PSGallery' -SourceLocation "https://www.powershellgallery.com/api/v2" -InstallationPolicy Trusted
    Install-Module -Name 7Zip4PowerShell -Force

    #Download HIBP Hashes
    (New-Object System.Net.WebClient).DownloadFile("$HIBPHashesLink", "$HIBPHashesCompressedFile")

    #Extract HIBP Hashes
    Expand-7Zip -ArchiveFileName $HIBPHashesCompressedFile -TargetPath $HIBPHashesExtractDir

    Remove-Item $HIBPHashesCompressedFile -Force

    Import-Module LithnetPasswordProtection
    Open-Store 'C:\Program Files\Lithnet\Active Directory Password Protection\Store'
    Import-CompromisedPasswordHashes -Filename (Get-ChildItem -Path $HIBPHashesExtractDir).fullname

    #region Copy ADM files to central store
    if (Test-Path 'C:\Windows\SYSVOL') {
        $SYSVOLPath = 'C:\Windows\SYSVOL'
    }
    elseif (Test-Path 'C:\Windows\SYSVOL_DFSR') {
        $SYSVOLPath = 'C:\Windows\SYSVOL_DFSR'
    }

    $FilesToCopy = @(
        'C:\Windows\PolicyDefinitions\lithnet.admx',
        'C:\Windows\PolicyDefinitions\lithnet.activedirectory.passwordfilter.admx',
        'C:\Windows\PolicyDefinitions\en-US\lithnet.activedirectory.passwordfilter.adml',
        'C:\Windows\PolicyDefinitions\en-US\lithnet.adml'
    )

    Foreach ($File in $FilesToCopy) {
        if($File -like '*.adml'){
            $Destination = "$($SYSVOLPath)\domain\Policies\PolicyDefinitions\en-US"
        }
        elseif ($File -like '*.admx') {
            $Destination = "$($SYSVOLPath)\domain\Policies\PolicyDefinitions"
        }

        Move-Item -Path $File -Destination $Destination -Force
    }
    #endregion Copy ADM files to central store

}
else {
    Write-Error 'Not enough free space on the disk'
}

#Cleanup
Remove-OldFiles