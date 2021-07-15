<#
Enumerates members of the "Pre-Windows 2000 Compatible Access" to make sure that "Everyone" and "Anonymous Logon" are not members.

.LINK
https://www.stigviewer.com/stig/active_directory_domain/2016-02-19/finding/V-8547

Andy Morales
#>

if ((Get-WmiObject -Class Win32_OperatingSystem).productType -eq '2') {

    Import-Module -Name ActiveDirectory

    $BadGroupMembers = @(
        'S-1-1-0',
        'S-1-5-7'
    )

    $GroupMembers = Get-ADGroupMember -Identity 'S-1-5-32-554'

    $BadMembersFound = $false

    Foreach ($group in $GroupMembers) {
        if ($BadGroupMembers -contains $group.SID) {
            $BadMembersFound = $true

            BREAK
        }
    }

    if ($BadMembersFound) {
        Write-Output 'BadGroupMembersFound'
    }
}
else{
    Write-Output 'ComputerNotDC'
}
