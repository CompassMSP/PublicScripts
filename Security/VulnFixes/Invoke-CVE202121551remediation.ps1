<#
Removal script for CVE-2021-21551

.LINK
https://www.dell.com/support/kbdoc/en-us/000186020/additional-information-regarding-dsa-2021-088-dell-driver-insufficient-access-control-vulnerability
https://www.dell.com/support/kbdoc/en-us/000186019/dsa-2021-088-dell-client-platform-security-update-for-dell-driver-insufficient-access-control-vulnerability

Andy Morales
#>

$PossibleFiles = @(Get-ChildItem -Path 'C:\Users\*\AppData\Local\Temp', 'C:\Windows\Temp' -Force -Recurse -Include 'dbutil_2_3.sys' -ErrorAction SilentlyContinue)

if ($PossibleFiles.count -gt 0) {

    $BadHashes = @(
        '0296E2CE999E67C76352613A718E11516FE1B0EFC3FFDB8918FC999DD76A73A5',
        '87E38E7AEAAAA96EFE1A74F59FCA8371DE93544B7AF22862EB0E574CEC49C7C3'
    )

    Foreach ($file in $PossibleFiles) {
        #try to verify the file by using the file hash if PSVersion is 4+. If an earlier PS is installed (Get-fileHash is not available), just remove the file.
        if ($PSVersionTable.PSVersion.Major -ge 4) {
            $FileHash = $File | Get-FileHash -Algorithm SHA256

            #only remove the file if it matches the bad hashes
            if ($BadHashes -contains $FileHash.Hash) {
                $file | Remove-Item -Force
            }
        }
        else {
            #just remove the file on earlier PS builds
            $file | Remove-Item -Force
        }
    }
}