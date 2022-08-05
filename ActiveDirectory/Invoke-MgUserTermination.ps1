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
# 07-06-2022                    1.4         Improved readability and export for user groups
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

$Localpath = 'C:\Temp'

if((Test-Path $Localpath) -eq $false) {
    Write-Host "Creating temp directory for user group export" 
    New-Item -Path $Localpath -ItemType Directory
}

#region pre-check
Write-Host "Attempting to find $($user) in Active Directory" 

try {
    $UserFromAD = Get-ADUser -Identity $User -Properties MemberOf -ErrorAction Stop
}
catch {
    Write-Host "Could not find user $($User) in Active Directory" -ForegroundColor Red -BackgroundColor Black
    exit
}

Write-Host "Attempting to find Disabled users OU"

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
    Write-Host "Could not find disabled OU in Active Directory" -ForegroundColor Red -BackgroundColor Black
    exit
}
#endregion pre-check

Write-Host "Logging into Azure services. You should get 2 prompts." 

Connect-ExchangeOnline
Select-MgProfile Beta
Connect-Graph -Scopes "Directory.ReadWrite.All", "User.ReadWrite.All", "Directory.AccessAsUser.All", "Group.ReadWrite.All", "GroupMember.Read.All"

Write-Host "Attempting to find $($UserFromAD.UserPrincipalName) in Azure" 

try {
    $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
    $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -ErrorAction Stop
}
catch {
    Write-Host "Could not find user $($UserFromAD.UserPrincipalName) in Azure" -ForegroundColor Red -BackgroundColor Black
    exit
}

$Confirmation = Read-Host -Prompt "The user below will be disabled:`n
Display Name = $($UserFromAD.Name)
UserPrincipalName = $($UserFromAD.UserPrincipalName)
Mailbox name =  $($365Mailbox.DisplayName)
Azure name = $($MgUser.DisplayName)
Destination OU = $($DestinationOU)`n
(Y/N)`n"

if ($Confirmation -ne 'y') {
    Write-Host 'User did not enter "Y"' -ForegroundColor Red -BackgroundColor Black
    exit
}

#region ActiveDirectory

#Modify the AD user account
Write-Host "Performing Active Directory Steps" 

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

#region Azure
Write-Host "Performing Azure Steps" 

#Revoke all sessions
Revoke-MgUserSign -UserId $MgUser.UserPrincipalName

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
	Write-Host "User mailbox $UserAccess not found. Skipping access rights setup" -ForegroundColor Red -BackgroundColor Black
	$GetAccessUserCheck = 'no'
	}   
} Else {
    Write-Host "Skipping access rights setup"
}

if ($GetAccessUserCheck -eq 'yes') { 
    Write-Host "Adding Full Access permissions for $($GetAccessUser.PrimarySmtpAddress) to $($UserFromAD.UserPrincipalName)" 
    Add-MailboxPermission -Identity $UserFromAD.UserPrincipalName -User $UserAccess -AccessRights FullAccess -InheritanceType All -AutoMapping $false }

# Set Mailbox forwarding address 
$UserFwdConfirmation = Read-Host -Prompt "Would you like to forward users email? (Y/N)"

if ($UserFwdConfirmation -eq 'y') {

    $UserFWD = Read-Host -Prompt "Enter the email address of forward recipient"
    try { 
        $GetFWDUser = get-mailbox $UserFWD -ErrorAction Stop 
        $GetFWDUserCheck = 'yes'
        Write-Host "Applying forward from $($UserFromAD.UserPrincipalName) to $($GetFWDUser.PrimarySmtpAddress)" 
    }
    catch { 
	Write-Host "User mailbox $UserFWD not found. Skipping mailbox forward" -ForegroundColor Red -BackgroundColor Black
	$GetFWDUserCheck = 'no'
	}
    
} Else {
    Write-Host "Skipping mailbox forwarding" 
}

if ($GetFWDUserCheck -eq 'yes') { Set-Mailbox $UserFromAD.UserPrincipalName -ForwardingAddress $UserFWD -DeliverToMailboxAndForward $False }

#Find Azure only groups

$AllAzureGroups = Get-MgUserMemberOf -UserId $MgUser.UserPrincipalName  | Where-Object {$_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.Id -ne '3e08099a-4cc4-42fb-aa37-e4c988ea8eff'} | `
        ForEach-Object { @{ GroupId=$_.Id}} | Get-MgGroup | Where-Object {$_.OnPremisesSyncEnabled -eq $NULL} | Select-Object DisplayName, SecurityEnabled, Mail, Id

$AllAzureGroups | Export-Csv c:\temp\$($user)_Groups_Id.csv -NoTypeInformation
    
Write-Host "Export User Groups Completed. Path: C:\temp\$($user)_Groups_Id.csv" 

#Remove user from all groups
Foreach ($365Group in $AllAzureGroups) {
    try {
        Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
    } catch {
        Remove-DistributionGroupMember -Identity $365Group.Id -Member $MgUser.UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction SilentlyContinue
    }
}

#Export user licenses 
Get-MgUserLicenseDetail -UserId $MgUser.Id | Select-Object SkuPartNumber, SkuId, Id | Export-Csv c:\temp\$($user)_License_Id.csv -NoTypeInformation
    
Write-Host "Export User Licenses Completed. Path: C:\temp\$($user)_License_Id.csv" 

#Remove Licenses
Write-Host "Starting removal of user licenses." 

Get-MgUserLicenseDetail -UserId $MgUser.Id | Where-Object `
   {($_.SkuPartNumber -ne "O365_BUSINESS_ESSENTIALS" -and $_.SkuPartNumber -ne "SPE_E3" -and $_.SkuPartNumber -ne "SPB" -and $_.SkuPartNumber -ne "EXCHANGESTANDARD") } `
   | ForEach-Object { Set-MgUserLicense -UserId $MgUser.Id -AddLicenses @() -RemoveLicenses $_.SkuId -ErrorAction Stop }

Get-MgUserLicenseDetail -UserId $MgUser.Id | ForEach-Object { Set-MgUserLicense -UserId $MgUser.Id -AddLicenses @() -RemoveLicenses $_.SkuId }

Write-Host "Removal of user licenses completed." 

#endregion Office365

#Start AD Sync cycle
Start-ADSyncSyncCycle -PolicyType Delta

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-Graph

Write-Host "User $($user) should now be disabled unless any errors occurred during the process." 
