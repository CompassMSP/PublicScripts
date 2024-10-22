<#
This script disables all computer and User objects located in Disabled OUs. OU names can be added to the "DisabledOUNames" list. Usernames can be excluded under "WhiteListedNames"

Andy Morales
#>

#This OU list will be used for all clients. Avoid using names that might be too broad.
$DisabledOUNames = @(
    'Disabled Users',
    'Disabled Computers',
    'Disabled'
)

#The SamAccountNames(computers and Users) below will be excluded
$WhiteListedSamNames = @(
    'AZUREADSSOACC$',
    #The "Administrator" account should not be disabled.
    'Administrator',
    'DiscoverySearchMailbox'
)

$DisabledAccountsOUs = @()

#Get All of the OUs that should have disabled Users
Foreach ($OuName in $DisabledOUNames) {
    $DisabledAccountsOUs += Get-ADOrganizationalUnit -Filter * | Where-Object { $_.Name -like "*$OuName*" }
}

Foreach ($OU in $DisabledAccountsOUs) {
    Get-ADUser -Filter 'Enabled -eq "True"' -SearchBase $OU.DistinguishedName | Where-Object { ($WhiteListedSamNames -NotContains $_.SamAccountName) -and ($_.SID -notLike '*-500') } | Set-ADUser -Enabled $false -Description "Disabled by Script on $(Get-Date -Format 'FileDate')"
    Get-ADComputer -Filter 'Enabled -eq "True"' -SearchBase $OU.DistinguishedName | Where-Object { $WhiteListedSamNames -NotContains $_.SamAccountName } | Set-ADComputer -Enabled $false -Description "Disabled by Script on $(Get-Date -Format 'FileDate')"
}