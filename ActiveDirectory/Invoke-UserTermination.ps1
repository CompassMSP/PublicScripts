<#
This script handles most of the Office 365/AD tasks during user termination.

All required modules must be installed in order for the script to execute successfully.

Andy Morales

#>
#requires -Modules activeDirectory,ExchangeOnlineManagement,AzureAD,ADSync,MSOnline -RunAsAdministrator

[cmdletbinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$user
)

#region pre-check
Write-Output "Attempting to find $($user) in Active Directory"

try {
    $UserFromAD = Get-ADUser -Identity $User -Properties MemberOf -ErrorAction Stop
}
catch {
    Write-Output "Could not find user $($User) in Active Directory"
    exit
}

Write-Output "Attempting to find Disabled users OU"

$DisabledOUs = @(Get-ADOrganizationalUnit -Filter 'Name -like "*disabled*"')

if ($DisabledOUs.count -gt 0) {
    #set the destination OU to the first one found, but try to find a better one(user specific)
    $DestinationOU = $DisabledOUs[0].DistinguishedName

    #try to find user specific OU
    foreach ($OU in $DisabledOUs) {
        if ($OU.DistinguishedName -like '*user*') {
            $DestinationOU = $OU.DistinguishedName
        }
    }
}
else {
    Write-Output "Could not find disabled OU in Active Directory"
    exit
}
#endregion pre-check

Write-Output "Logging into 365 services. You might get 3 prompts."

Import-Module ExchangeOnlineManagement, AzureAD, MSOnline

Connect-ExchangeOnline
Connect-AzureAD
Connect-MsolService

Write-Output "Attempting to find $($UserFromAD.UserPrincipalName) in azure/365"

try {
    $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
    $AZUser = Get-AzureADUser -ObjectId $UserFromAD.UserPrincipalName -ErrorAction Stop
}
catch {
    Write-Output "Could not find user $($UserFromAD.UserPrincipalName) in Office 365"
    exit
}

$Confirmation = Read-Host -Prompt "The user below will be disabled:`n
Display Name = $($UserFromAD.Name)
UserPrincipalName = $($UserFromAD.UserPrincipalName)
Office 365 Mailbox name =  $($365Mailbox.DisplayName)
Azure name = $($AZUser.DisplayName)
Destination OU = $($DestinationOU)`n
(Y/N)`n"

if ($Confirmation -ne 'y') {
    Write-Output 'User did not enter "Y"'
    exit
}

#region ActiveDirectory

#Modify the AD user account
Write-Output "Performing Active Directory Steps"

$SetADUserParams = @{
    Identity    = $UserFromAD.SamAccountName
    Description = "Disabled on $(Get-Date -Format 'FileDate')"
    Enabled     = $False
    Replace = @{msExchHideFromAddressLists=$true}
}

Set-ADUser @SetADUserParams

#remove user from all AD groups
Foreach ($group in $UserFromAD.MemberOf) {
    Remove-ADGroupMember -Identity $group -Members $UserFromAD.SamAccountName -Confirm:$false
}

#Move user to disabled OU
$UserFromAD | Move-ADObject -TargetPath $DestinationOU
#endregion ActiveDirectory

#region Office365
Write-Output "Performing Office 365 Steps"

#Revoke all sessions
Revoke-AzureADUserAllRefreshToken -ObjectId $AZUser.ObjectId

#Remove devices
Get-MobileDevice -Mailbox $UserFromAD.UserPrincipalName | ForEach-Object { Remove-MobileDevice $_.DeviceID -Confirm:$false } 

#reset MFA
Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $UserFromAD.UserPrincipalName

#Change mailbox to shared
$365Mailbox | Set-Mailbox -Type Shared

# Set Mailbox forwarding address 
$UserFwdConfirmation = Read-Host -Prompt "Would you like to forward users email? (Y/N)"

if ($UserFwdConfirmation -eq 'y') {

    $UserFWD = Read-Host -Prompt "Enter the email address of forward recipient"
    try { 
        $GetFWDUser = get-mailbox $UserFWD -ErrorAction Stop 
    }
    catch { 
	Write-Output "User mailbox $UserFWD not found. Skipping mailbox forward"
	$GetFWDUser = 'no'
	}
    
} Else {
    Write-Output "Skipping mailbox forwarding"
}

if ($GetFWDUser -ne 'no') { Set-Mailbox $UserFromAD.UserPrincipalName -ForwardingAddress $UserFWD -DeliverToMailboxAndForward $False }

#Find 365 only groups
$All365Groups = Get-AzureADUserMembership -ObjectId $UserFromAD.UserPrincipalName | Where-Object { $_.OnPremisesSecurityIdentifier -eq $null -and $_.DisplayName -ne 'All Users'}

#Remove user from all groups
Foreach ($365Group in $All365Groups) {
    try{
        Remove-AzureADGroupMember -ObjectId $365Group.ObjectId -MemberId $AZUser.ObjectId -ErrorAction Stop
    }
    catch{
        Remove-DistributionGroupMember -Identity $365Group.ObjectId -Member $UserFromAD.UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false
    }
}

#Disable user
Set-AzureADUser -ObjectId $UserFromAD.UserPrincipalName -AccountEnabled $false

#Remove Licenses
(Get-MsolUser -UserPrincipalName $UserFromAD.UserPrincipalName).licenses.AccountSkuId | ForEach-Object { Set-MsolUserLicense -UserPrincipalName $UserFromAD.UserPrincipalName -RemoveLicenses $_ }

#endregion Office365

#Start AD Sync cycle
Start-ADSyncSyncCycle -PolicyType Delta

Write-Output "User $($user) should now be disabled unless any errors occurred during the process."
