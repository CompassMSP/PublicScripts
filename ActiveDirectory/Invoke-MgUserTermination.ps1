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
# 08-02-2023                    1.5         Added OneDrive access grant
# 02-12-2024                    1.6         Add AppRoleAssignment for KnowBe4 SCIM App
# 02-14-2024                    1.7         Fix issues with copy groups function and code cleanup
# 02-19-2024                    1.8         Changes to Get-MgUserMemberOf function
# 03-08-2024                    1.9         Cleaned up licenses select display output
# 05-08-2024                    2.0         Add input box for Variables
# 05-09-2024                    2.1         Remove user from directory roles
# 05-13-2024                    2.2         Fixed AppRoleAssignment and added Term User to accept SAM or UPN
# 05-15-2024                    2.3         Set OneDrive as Readonly
# 10-15-2024                    2.4         Remove AppRoleAssignment for KnowBe4 SCIM App
# 10-22-2024                    2.5         Add KB4 offboarding email delivery to SecurePath
#********************************************************************************
# Run from the Primary Domain Controller with AD Connect installed
#
#
# The following modules must be installed
# Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, PnP.PowerShell
#>

#Import-Module adsync -UseWindowsPowerShell

#Add-Type -AssemblyName PresentationFramework -UseWindowsPowerShell
#[System.Windows.MessageBox]::Show('For all fields please enter users email address', 'Compass Termination Request')

function CompassUserTermination {
    param (
        [Parameter(Mandatory)]
        [string]$UserToTerm,
        [string]$GrantOneDriveAccessTo,
        [string]$GrantMailboxFullControlTo,
        [string]$FowardMailboxTo

    )
    [pscustomobject]@{
        InputUserToTerm         = $UserToTerm
        InputUserFullControl    = $GrantMailboxFullControlTo
        InputUserFWD            = $FowardMailboxTo
        InputUserOneDriveAccess = $GrantOneDriveAccessTo
    }
}

$result = Invoke-Expression (Show-Command CompassUserTermination -PassThru)

$User = $result.InputUserToTerm
$GrantUserFullControl = $result.InputUserFullControl
$SetUserMailFWD = $result.InputUserFWD
$GrantUserOneDriveAccess = $result.InputUserOneDriveAccess

if (!$result.InputUserFullControl) { $UserAccessConfirmation = 'n' } else { $UserAccessConfirmation = 'y' }
if (!$result.InputUserFWD) { $UserFwdConfirmation = 'n' } else { $UserFwdConfirmation = 'y' }
if (!$result.InputUserOneDriveAccess) { $SPOAccessConfirmation = 'n' } else { $SPOAccessConfirmation = 'y' }

$Localpath = 'C:\Temp'

if ((Test-Path $Localpath) -eq $false) {
    Write-Host "Creating temp directory for user group export" 
    New-Item -Path $Localpath -ItemType Directory
}

#region pre-check
Write-Host "Attempting to find $($user) in Active Directory" 

try {
    $UserFromAD = Get-ADUser -Filter "userPrincipalName -eq '$($User)'" -Properties MemberOf -ErrorAction Stop
} catch {
    Write-Host "Could not find user $($User) in Active Directory" -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
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
} else {
    Write-Host "Could not find disabled OU in Active Directory" -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    exit
}
#endregion pre-check

Write-Host "Logging into Azure services. You should get 2 prompts." 
$AppId = "432beb65-bc40-4b40-9366-1c5a768ee717"
$tenantID = "02e68a77-717b-48c1-881a-acc8f67c291a"
$Certificate = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=Graph PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
Connect-Graph -TenantId $TenantId -AppId $AppId -Certificate $Certificate -NoWelcome
Connect-ExchangeOnline -ShowBanner:$false

Write-Host "Attempting to find $($UserFromAD.UserPrincipalName) in Azure" 

try {
    $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
    $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -ErrorAction Stop
} catch {
    Write-Host "Could not find user $($UserFromAD.UserPrincipalName) in Azure" -ForegroundColor Red -BackgroundColor Black
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
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
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    exit
}

#region ActiveDirectory

#Modify the AD user account
Write-Host "Performing Active Directory Steps" 

$SetADUserParams = @{
    Identity    = $UserFromAD.SamAccountName
    Description = "Disabled on $(Get-Date -Format 'FileDate')"
    Enabled     = $False
    Replace     = @{msExchHideFromAddressLists = $true }
    Clear       = @(
        'company',
        'Title',
        'physicalDeliveryOfficeName',
        'Department',
        'facsimileTelephoneNumber',
        'mobile',
        'telephoneNumber',
        'l', # l is for Location because Microsoft AD attribues are stupid
        'Manager',
        'extensionAttribute1',
        'extensionAttribute2',
        'extensionAttribute3',
        'extensionAttribute4',
        'extensionAttribute5',
        'extensionAttribute6',
        'extensionAttribute15'
    )
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
Revoke-MgUserSignInSession -UserId $UserFromAD.UserPrincipalName -ErrorAction SilentlyContinue

#Remove Mobile Device
Get-MobileDevice -Mailbox $UserFromAD.UserPrincipalName | ForEach-Object { Remove-MobileDevice $_.DeviceID -Confirm:$false -ErrorAction SilentlyContinue } 

#Disable AzureAD registered devices
$termUserDeviceId = Get-MgUserRegisteredDevice -UserId $UserFromAD.UserPrincipalName

$termUserDeviceId | ForEach-Object {
    $params = @{
        AccountEnabled = $false
    }
    Update-MgDevice -DeviceId $_.Id -BodyParameter $params
}

$termUserDeviceId | ForEach-Object { Get-MgDevice -DeviceId $_.Id | Select-Object Id, DisplayName, ApproximateLastSignInDateTime, AccountEnabled } 

# Disabled mailbox forwarding
$365Mailbox | Set-Mailbox -ForwardingAddress $null -ForwardingSmtpAddress $null 

# Change mailbox to shared
$365Mailbox | Set-Mailbox -Type Shared

# Grant User FullAccess to Mailbox

if ($UserAccessConfirmation -eq 'y') {

    try { 
        $GetAccessUser = Get-Mailbox $GrantUserFullControl -ErrorAction Stop
        $GetAccessUserCheck = 'yes'
    } catch { 
        Write-Host "User mailbox $GrantUserFullControl not found. Skipping access rights setup" -ForegroundColor Red -BackgroundColor Black
        $GetAccessUserCheck = 'no'
    }

} Else {
    Write-Host "Skipping access rights setup"
}

if ($GetAccessUserCheck -eq 'yes') { 
    Write-Host "Adding Full Access permissions for $($GetAccessUser.PrimarySmtpAddress) to $($UserFromAD.UserPrincipalName)" 
    Add-MailboxPermission -Identity $UserFromAD.UserPrincipalName -User $GrantUserFullControl -AccessRights FullAccess -InheritanceType All -AutoMapping $true 
}

# Set Mailbox forwarding address 

if ($UserFwdConfirmation -eq 'y') {

    try { 
        $GetFWDUser = Get-Mailbox $SetUserMailFWD -ErrorAction Stop 
        $GetFWDUserCheck = 'yes'
        Write-Host "Applying forward from $($UserFromAD.UserPrincipalName) to $($GetFWDUser.PrimarySmtpAddress)" 
    } catch { 
        Write-Host "User mailbox $SetUserMailFWD not found. Skipping mailbox forward" -ForegroundColor Red -BackgroundColor Black
        $GetFWDUserCheck = 'no'
    }
    
} Else {
    Write-Host "Skipping mailbox forwarding" 
}

if ($GetFWDUserCheck -eq 'yes') { Set-Mailbox $UserFromAD.UserPrincipalName -ForwardingAddress $SetUserMailFWD -DeliverToMailboxAndForward $False }

#Find user directory roles
$AllDirectoryRoles = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserFromAD.UserPrincipalName).Id | `
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' } | `
    Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

#Remove user from directory roles
if (!$AllDirectoryRoles) { Write-Host "Skipping removal of directory roles as user is not assigned." } else {
    Foreach ($DirectoryRole in $AllDirectoryRoles) {
        Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $DirectoryRole.Id -DirectoryObjectId $MgUser.Id
    }
}
#Find Azure only groups
$AllAzureGroups = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserFromAD.UserPrincipalName).Id | `
    Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.AdditionalProperties.membershipRule -eq $NULL -and $_.onPremisesSyncEnabled -ne 'False' } | `
    Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

$AllAzureGroups | Export-Csv c:\temp\terminated_users_exports\$($user)_Groups_Id.csv -NoTypeInformation
    
Write-Host "Export User Groups Completed. Path: C:\temp\terminated_users_exports\$($user)_Groups_Id.csv" 

#Remove user from groups
Foreach ($365Group in $AllAzureGroups) {
    try {
        Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $MgUser.Id -ErrorAction Stop
    } catch {
        Remove-DistributionGroupMember -Identity $365Group.Id -Member $UserFromAD.UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction SilentlyContinue
    }
}

#Export user licenses 
Get-MgUserLicenseDetail -UserId $UserFromAD.UserPrincipalName | Select-Object SkuPartNumber, SkuId, Id | Export-Csv c:\temp\terminated_users_exports\$($user)_License_Id.csv -NoTypeInformation
    
Write-Host "Export User Licenses Completed. Path: C:\temp\terminated_users_exports\$($user)_License_Id.csv" 

#Remove Licenses
Write-Host "Starting removal of user licenses." 

Get-MgUserLicenseDetail -UserId $UserFromAD.UserPrincipalName | Where-Object `
{ ($_.SkuPartNumber -ne "O365_BUSINESS_ESSENTIALS" -and $_.SkuPartNumber -ne "SPE_E3" -and $_.SkuPartNumber -ne "SPB" -and $_.SkuPartNumber -ne "EXCHANGESTANDARD") } `
| ForEach-Object { Set-MgUserLicense -UserId $UserFromAD.UserPrincipalName -AddLicenses @() -RemoveLicenses $_.SkuId -ErrorAction Stop }

Get-MgUserLicenseDetail -UserId $UserFromAD.UserPrincipalName | ForEach-Object { Set-MgUserLicense -UserId $UserFromAD.UserPrincipalName -AddLicenses @() -RemoveLicenses $_.SkuId }

Write-Host "Removal of user licenses completed."

## Sends email to SecurePath Team (soc@compassmsp.com) with the offboarding user information.
$MsgFrom = 'noreply@compassmsp.com'

$params = @{
    message         = @{
        subject      = "KB4 – Remove User"
        body         = @{
            contentType = "HTML"
            content     = "The following user need to be removed to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
        }
        toRecipients = @(
            @{
                emailAddress = @{
                    address = "soc@compassmsp.com"
                }
            }
        )
    }
    saveToSentItems = "false"
}

Send-MgUserMail -UserId $MsgFrom -BodyParameter $params

#Disconnect from Exchange and Graph
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-Graph

# Set OneDrive as Read Only

Connect-PnPOnline -Url compassmsp-admin.sharepoint.com -ClientId '24e3c6ad-9658-4a0d-b85f-82d67d148449' -Tenant compassmsp.onmicrosoft.com -Thumbprint '3b51fcc465d26593303453c8a636b13587e0dc81'

$UserOneDriveURL = (Get-PnPUserProfileProperty -Account $UserFromAD.UserPrincipalName -Properties PersonalUrl).PersonalUrl

Set-PnPTenantSite -Url $UserOneDriveURL -LockState ReadOnly
        
# Set OneDrive grant access
if ($SPOAccessConfirmation -eq 'y') {

    try {
        $GetUserOneDriveAccess = Get-Mailbox $GrantUserOneDriveAccess -ErrorAction Stop 
        $GetUserOneDriveAccessCheck = 'yes'
        Write-Host "Granting OneDrive access rights to $($GetUserOneDriveAccess.PrimarySmtpAddress)" 
    } catch { 
        Write-Host "User $GrantUserOneDriveAccess not found. Skipping OneDrive access grant" -ForegroundColor Red -BackgroundColor Black
        $GetUserOneDriveAccessCheck = 'no'
    }

} Else {
    Write-Host "Skipping OneDrive access grant" 
}

if ($GetUserOneDriveAccessCheck -eq 'yes') { 
    Set-PnPTenantSite -Url $UserOneDriveURL -Owners $GrantUserOneDriveAccess
    $UserOneDriveURL
    Read-Host 'Please copy the OneDrive URL. Press any key to continue'
}

Disconnect-PnPOnline

#endregion Office365

#Start AD Sync
powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta

Write-Host "User $($User) should now be disabled unless any errors occurred during the process."