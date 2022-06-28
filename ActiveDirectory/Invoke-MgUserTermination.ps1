#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#Author       : Chris Williams
# Creation Date: 12-20-2021
# Usage        : This script handles most of the Office 365/AD tasks during user termination.

#********************************************************************************
# Date                        Version       Changes
#------------------------------------------------------------------------
# 12-20-2021                    1.0         Initial Version
# 03-15-2022                    1.1         Added exports of groups and licenses
# 06-27-2022                    1.2         Fixes for Remove-MgGroupMemberByRef and Revoke-MgUserSign
# 06-28-2022                    1.3         Add removal of manager from disabled user and optimization changes
#
#
#********************************************************************************
# Run from the Primary Domain Controller with AD Connect installed
#
# The following modules must be installed
# Install-Module ExchangeOnlineManagement
# Install-Module Microsoft.Graph
#>


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

Write-Output "Logging into 365 services. You should get 2 prompts."

Connect-ExchangeOnline
Select-MgProfile Beta
Connect-Graph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Directory.AccessAsUser.All", "Group.ReadWrite.All", "GroupMember.Read.All"

Write-Output "Attempting to find $($UserFromAD.UserPrincipalName) in azure/365"

try {
    $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
    $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -ErrorAction Stop
}
catch {
    Write-Output "Could not find user $($UserFromAD.UserPrincipalName) in Office 365"
    exit
}

$Confirmation = Read-Host -Prompt "The user below will be disabled:`n
Display Name = $($UserFromAD.Name)
UserPrincipalName = $($UserFromAD.UserPrincipalName)
Office 365 Mailbox name =  $($365Mailbox.DisplayName)
Azure name = $($MgUser.DisplayName)
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
    Manager = $NULL
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
Revoke-MgUserSign -UserId $MgUser.UserPrincipalName
#Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$($MgUser.Id)/microsoft.graph.revokeSignInSessions" -Method POST -Body @{}

Get-MobileDevice -Mailbox $UserFromAD.UserPrincipalName | ForEach-Object { Remove-MobileDevice $_.DeviceID -Confirm:$false -ErrorAction SilentlyContinue } 

#Change mailbox to shared
$365Mailbox | Set-Mailbox -Type Shared

# Grant User FullAccess to Mailbox
$UserAccessConfirmation = Read-Host -Prompt "Would you like to add FullAccess permissions to mailbox to $($UserFromAD.UserPrincipalName)? (Y/N)"

if ($UserAccessConfirmation -eq 'y') {

    $UserAccess = Read-Host -Prompt "Enter the email address of FullAccess recipient"
    try { 
        $GetAccessUser = get-mailbox $UserAccess -ErrorAction Stop
        $GetAccessUserCheck = 'yes'
    }
    catch { 
	Write-Output "User mailbox $UserAccess not found. Skipping access rights setup"
	$GetAccessUserCheck = 'no'
	}   
} Else {
    Write-Output "Skipping access rights setup"
}

if ($GetAccessUserCheck -eq 'yes') { 
    Write-Output "Adding Full Access permissions for $($GetAccessUser.PrimarySmtpAddress) to $($UserFromAD.UserPrincipalName)"
    Add-MailboxPermission -Identity $UserFromAD.UserPrincipalName -User $UserAccess -AccessRights FullAccess -InheritanceType All -AutoMapping $false }

# Set Mailbox forwarding address 
$UserFwdConfirmation = Read-Host -Prompt "Would you like to forward users email? (Y/N)"

if ($UserFwdConfirmation -eq 'y') {

    $UserFWD = Read-Host -Prompt "Enter the email address of forward recipient"
    try { 
        $GetFWDUser = get-mailbox $UserFWD -ErrorAction Stop 
        $GetFWDUserCheck = 'yes'
        Write-Output "Applying forward from $($UserFromAD.UserPrincipalName) to $($GetFWDUser.PrimarySmtpAddress)"
    }
    catch { 
	Write-Output "User mailbox $UserFWD not found. Skipping mailbox forward"
	$GetFWDUserCheck = 'no'
	}
    
} Else {
    Write-Output "Skipping mailbox forwarding"
}

if ($GetFWDUserCheck -eq 'yes') { Set-Mailbox $UserFromAD.UserPrincipalName -ForwardingAddress $UserFWD -DeliverToMailboxAndForward $False }

#Find 365 only groups

#$All365Groups = (Get-MgUserMemberOf -UserId $MgUser.UserPrincipalName).Id | Where-Object {$_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' }

$All365Groups = (Get-MgUserMemberOf -UserId $MgUser.UserPrincipalName).Id  | Where-Object {$_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole'} | `
        ForEach-Object { @{ GroupId=$_.Id}} | Get-MgGroup | Where-Object {$_.OnPremisesSyncEnabled -eq $NULL} | Select-Object DisplayName, SecurityEnabled, Mail, Id

$Localpath = 'C:\Temp'

$UserGroupsBackupConfirmation = Read-Host -Prompt "Would you like to backup user groups? (Y/N)"

if ($UserGroupsBackupConfirmation -eq 'y') {

    if((Test-Path $Localpath) -eq $false) {
        Write-Host `
            -ForegroundColor Cyan `
            -BackgroundColor Black `
            "Creating temp directory for user group export"
        New-Item -Path $Localpath -ItemType Directory
    }
    
    Write-Output "Checking to see if User Group export exists"
    
    if ( Get-ChildItem -Path c:\temp | Where-Object {$_.Name -like 'User_Groups_Id.csv'} ) { 
        Write-Output "Previous export exists. Please backup and then confirm removal."
        Remove-Item -Path C:\temp\User_Groups_Id.csv -Confirm}
    
    $All365Groups | Export-Csv c:\temp\User_Groups_Id.csv -NoTypeInformation
    
    Write-Output "Export User Groups Completed. Path: C:\temp\User_Groups_Id.csv"

}

#Remove user from all groups
Foreach ($365Group in $All365Groups) {
    try {
        Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
        #Invoke-GraphRequest -Method 'Delete' -Uri "https://graph.microsoft.com/v1.0/groups/$($365Group)/members/$($mgUser.Id)/`$ref"
    } catch {
        Remove-DistributionGroupMember -Identity $365Group.Mail -Member $MgUser.UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false
    }
}

#Get user licenses 
$AllLicenses = Get-MgUserLicenseDetail -UserId $MgUser.Id

$UserLicensesBackupConfirmation = Read-Host -Prompt "Would you like to backup user licenses? (Y/N)"

if ($UserLicensesBackupConfirmation -eq 'y') {

    if((Test-Path $Localpath) -eq $false) {
        Write-Host `
            -ForegroundColor Cyan `
            -BackgroundColor Black `
            "Creating temp directory for user group export"
        New-Item -Path $Localpath -ItemType Directory
    }
    
    Write-Output "Checking to see if User license export exists"
    
    if ( Get-ChildItem -Path c:\temp | Where-Object {$_.Name -like 'User_License_Id.csv'} ) { 
        Write-Output "Previous export exists. Please backup and then confirm removal."
        Remove-Item -Path C:\temp\User_License_Id.csv -Confirm}
    
    $AllLicenses | Export-Csv c:\temp\User_License_Id.csv -NoTypeInformation
    
    Write-Output "Export User Licenses Completed. Path: C:\temp\User_License_Id.csv"

}

#Remove Licenses
Write-Output "Starting removal of user licenses."

Get-MgUserLicenseDetail -UserId $MgUser.Id | Where-Object `
   {($_.SkuPartNumber -ne "O365_BUSINESS_ESSENTIALS" -and $_.SkuPartNumber -ne "SPE_E3" -and $_.SkuPartNumber -ne "SPB" -and $_.SkuPartNumber -ne "EXCHANGESTANDARD") } `
   | ForEach-Object { Set-MgUserLicense -UserId $MgUser.Id -AddLicenses @() -RemoveLicenses $_.SkuId -ErrorAction Stop }

Get-MgUserLicenseDetail -UserId $MgUser.Id | ForEach-Object { Set-MgUserLicense -UserId $MgUser.Id -AddLicenses @() -RemoveLicenses $_.SkuId }

Write-Output "Removal of user licenses completed."

#endregion Office365

#Start AD Sync cycle
Start-ADSyncSyncCycle -PolicyType Delta

Write-Output "User $($user) should now be disabled unless any errors occurred during the process."
