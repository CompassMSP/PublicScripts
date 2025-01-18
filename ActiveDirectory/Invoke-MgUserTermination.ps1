#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#
.SYNOPSIS
    Handles Office 365/AD tasks during user termination.

.DESCRIPTION
    This script automates the termination process by handling both Active Directory
    and Microsoft 365 tasks including group removal, license removal, and mailbox management.

    IMPORTANT: This script must be run from the Primary Domain Controller with AD Connect installed.

.PARAMETER User
    The email address of the user to terminate.

.EXAMPLE
    .\Invoke-MgTerminateUser.ps1 -User "john.smith@domain.com"

.NOTES
    Author: Chris Williams
    Created: 2021-12-20
    Last Modified: 2025-01-17

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    2.8        2025-01-17  Added status messaging system with improved error handling and progress tracking
    2.7        2025-01-10  Add function to disable QuickEdit and InsertMode
    2.6        2024-11-08  Added better UI boxes for variables
    2.5        2024-10-22  Add KB4 offboarding email delivery to SecurePath
    2.4        2024-10-15  Remove AppRoleAssignment for KnowBe4 SCIM App
    2.3        2024-05-15  Set OneDrive as Readonly
    2.2        2024-05-13  Fixed AppRoleAssignment and added Term User to accept SAM or UPN
    2.1        2024-05-09  Remove user from directory roles
    2.0        2024-05-08  Add input box for Variables
    1.9        2024-03-08  Cleaned up licenses select display output
    1.8        2024-02-19  Changes to Get-MgUserMemberOf function
    1.7        2024-02-14  Fix issues with copy groups function and code cleanup
    1.6        2024-02-12  Add AppRoleAssignment for KnowBe4 SCIM App
    1.5        2023-08-02  Added OneDrive access grant
    1.4        2022-07-06  Improved readability and export for user groups
    1.3        2022-06-28  Add removal of manager from disabled user and optimization changes
    1.2        2022-06-27  Fixes for Remove-MgGroupMemberByRef and Revoke-MgUserSign
    1.1        2022-03-15  Added exports of groups and licenses
    1.0        2021-12-20  Initial Version
    ------------------------------------------------------------------------------
#>

#Import-Module adsync -UseWindowsPowerShell

$QuickEditCodeSnippet=@"
using System;
using System.Runtime.InteropServices;

public static class ConsoleModeSettings
{
    const uint ENABLE_QUICK_EDIT = 0x0040;
    const uint ENABLE_INSERT_MODE = 0x0020;

    const int STD_INPUT_HANDLE = -10;

    [DllImport("kernel32.dll", SetLastError = true)]
    static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll")]
    static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport("kernel32.dll")]
    static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    public static void EnableQuickEditMode()
    {
        SetConsoleFlag(ENABLE_QUICK_EDIT, true);
    }

    public static void DisableQuickEditMode()
    {
        SetConsoleFlag(ENABLE_QUICK_EDIT, false);
    }

    public static void EnableInsertMode()
    {
        SetConsoleFlag(ENABLE_INSERT_MODE, true);
    }

    public static void DisableInsertMode()
    {
        SetConsoleFlag(ENABLE_INSERT_MODE, false);
    }

    private static void SetConsoleFlag(uint modeFlag, bool enable)
    {
        IntPtr consoleHandle = GetStdHandle(STD_INPUT_HANDLE);
        uint consoleMode;
        if (GetConsoleMode(consoleHandle, out consoleMode))
        {
            if (enable)
                consoleMode |= modeFlag;
            else
                consoleMode &= ~modeFlag;

            SetConsoleMode(consoleHandle, consoleMode);
        }
    }
}

"@

Add-Type -TypeDefinition $QuickEditCodeSnippet -Language CSharp

function Set-ConsoleProperties()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [switch]$EnableQuickEditMode=$false,

        [Parameter(Mandatory=$false)]
        [switch]$DisableQuickEditMode=$false,

        [Parameter(Mandatory=$false)]
        [switch]$EnableInsertMode=$false,

        [Parameter(Mandatory=$false)]
        [switch]$DisableInsertMode=$false
    )

    if ($PSBoundParameters.Count -eq 0)
    {
        [ConsoleModeSettings]::EnableQuickEditMode()
        [ConsoleModeSettings]::EnableInsertMode()
        Write-Output "All settings have been enabled"
        return
    }

    if ($EnableQuickEditMode)
    {
        [ConsoleModeSettings]::EnableQuickEditMode()
        Write-Output "QuickEditMode has been enabled"
    }

    if ($DisableQuickEditMode)
    {
        [ConsoleModeSettings]::DisableQuickEditMode()
        Write-Output "QuickEditMode has been disabled"
    }

    if ($EnableInsertMode)
    {
        [ConsoleModeSettings]::EnableInsertMode()
        Write-Output "InsertMode has been enabled"
    }

    if ($DisableInsertMode)
    {
        [ConsoleModeSettings]::DisableInsertMode()
        Write-Output "InsertMode has been disabled"
    }
}

Set-ConsoleProperties -DisableQuickEditMode -DisableInsertMode

Add-Type -AssemblyName PresentationFramework

# Add status message functions near the start
function Write-StatusMessage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [string]$Status = "INFO",

        [Parameter(Mandatory=$false)]
        [ConsoleColor]$Color = "White"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $statusPadded = $Status.PadRight(7)
    Write-Host "[$timestamp] [$statusPadded] $Message" -ForegroundColor $Color
}

function Write-SuccessMessage {
    param([string]$Message)
    Write-StatusMessage -Message $Message -Status "OK" -Color Green
}

function Write-ErrorMessage {
    param([string]$Message)
    Write-StatusMessage -Message $Message -Status "ERROR" -Color Red
}

function Write-WarningMessage {
    param([string]$Message)
    Write-StatusMessage -Message $Message -Status "WARN" -Color Yellow
}

# Function to validate email addresses
function Test-EmailAddress {
    param (
        [string]$Email
    )
    return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'  # Basic email regex
}

# Function to create and show a custom WPF window for user termination
function Show-CustomTerminationWindow {
    # Create a new WPF window
    $window = New-Object System.Windows.Window
    $window.Title = "User Termination Request"
    $window.Width = 500
    $window.Height = 310
    $window.WindowStartupLocation = 'CenterScreen'

    # Create a StackPanel to hold the controls
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = '10'
    $window.Content = $stackPanel

    # Create label and textbox for User to Terminate
    $lblUserToTerm = New-Object System.Windows.Controls.Label
    $lblUserToTerm.Content = "User to Terminate (Email):"
    $stackPanel.Children.Add($lblUserToTerm)

    $txtUserToTerm = New-Object System.Windows.Controls.TextBox
    $txtUserToTerm.Margin = '0,0,0,4'
    $stackPanel.Children.Add($txtUserToTerm)

    # Create label and textbox for OneDrive Access
    $lblOneDriveAccess = New-Object System.Windows.Controls.Label
    $lblOneDriveAccess.Content = "Grant OneDrive Access To (Email):"
    $stackPanel.Children.Add($lblOneDriveAccess)

    $txtOneDriveAccess = New-Object System.Windows.Controls.TextBox
    $txtOneDriveAccess.Margin = '0,0,0,4'
    $stackPanel.Children.Add($txtOneDriveAccess)

    # Create label and textbox for Mailbox Full Control
    $lblMailboxControl = New-Object System.Windows.Controls.Label
    $lblMailboxControl.Content = "Grant Mailbox Full Control To (Email):"
    $stackPanel.Children.Add($lblMailboxControl)

    $txtMailboxControl = New-Object System.Windows.Controls.TextBox
    $txtMailboxControl.Margin = '0,0,0,4'
    $stackPanel.Children.Add($txtMailboxControl)

    # Create label and textbox for Forward Mailbox
    $lblForwardMailbox = New-Object System.Windows.Controls.Label
    $lblForwardMailbox.Content = "Forward Mailbox To (Email):"
    $stackPanel.Children.Add($lblForwardMailbox)

    $txtForwardMailbox = New-Object System.Windows.Controls.TextBox
    $txtForwardMailbox.Margin = '0,0,0,4'
    $stackPanel.Children.Add($txtForwardMailbox)

    # Create and add OK and Cancel buttons
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = '0,10,0,0'

    $okButton = New-Object System.Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Margin = '0,0,10,0'
    $okButton.Add_Click({
        # Validate user termination input
        if (-not $txtUserToTerm.Text) {
            [System.Windows.MessageBox]::Show("User to Terminate is a mandatory field. Please enter a email address.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            return
        }

        # Validate optional email inputs
        $emailInputs = @{
            "Grant OneDrive Access To (Email):"      = $txtOneDriveAccess.Text
            "Grant Mailbox Full Control To (Email):" = $txtMailboxControl.Text
            "Forward Mailbox To (Email):"            = $txtForwardMailbox.Text
        }

        foreach ($input in $emailInputs.GetEnumerator()) {
            if ($input.Value -and -not (Test-EmailAddress -Email $input.Value)) {
                [System.Windows.MessageBox]::Show("Invalid email format for: $($input.Key). Please enter a valid email address.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
        }

        # Set the DialogResult to true and close the window
        $window.DialogResult = $true
        $window.Close()
    })
    $buttonPanel.Children.Add($okButton)

    $cancelButton = New-Object System.Windows.Controls.Button
    $cancelButton.Content = "Cancel"
    $cancelButton.Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })
    $buttonPanel.Children.Add($cancelButton)

    $stackPanel.Children.Add($buttonPanel)

    # Show the window
    $result = $window.ShowDialog()

    if ($result -eq $true) {
        return @{
            InputUserToTerm         = $txtUserToTerm.Text
            InputUserFullControl    = $txtMailboxControl.Text
            InputUserFWD            = $txtForwardMailbox.Text
            InputUserOneDriveAccess = $txtOneDriveAccess.Text
        }
    } else {
        return $null
    }
}

# Call the custom input window function
$result = Show-CustomTerminationWindow
if ($null -eq $result) {
    Write-StatusMessage "Operation cancelled by user"
    exit
}

$User = $result.InputUserToTerm
$GrantUserFullControl = $result.InputUserFullControl
$SetUserMailFWD = $result.InputUserFWD
$GrantUserOneDriveAccess = $result.InputUserOneDriveAccess

if (!$result.InputUserFullControl) { $UserAccessConfirmation = 'n' } else { $UserAccessConfirmation = 'y' }
if (!$result.InputUserFWD) { $UserFwdConfirmation = 'n' } else { $UserFwdConfirmation = 'y' }
if (!$result.InputUserOneDriveAccess) { $SPOAccessConfirmation = 'n' } else { $SPOAccessConfirmation = 'y' }

$Localpath = 'C:\Temp'

if ((Test-Path $Localpath) -eq $false) {
    Write-StatusMessage "Creating temp directory for user group export"
    New-Item -Path $Localpath -ItemType Directory
}

#region pre-check
Write-StatusMessage "Attempting to find $($user) in Active Directory"

try {
    $UserFromAD = Get-ADUser -Filter "userPrincipalName -eq '$($User)'" -Properties MemberOf -ErrorAction Stop
} catch {
    Write-ErrorMessage "Could not find user $($User) in Active Directory"
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

Write-StatusMessage "Attempting to find Disabled users OU"

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
    Write-SuccessMessage "Found disabled users OU: $DestinationOU"
} else {
    Write-ErrorMessage "Could not find disabled OU in Active Directory"
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
#endregion pre-check

#Connect-ExchangeOnline
$ExOAppId = "baa3f5d9-3bb4-44d8-b10a-7564207ddccd"
$Org = "compassmsp.onmicrosoft.com"
$ExOCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=ExO PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $ExOCert) {
    Write-ErrorMessage "No valid ExO PowerShell certificates found in the LocalMachine\My store"
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Write-StatusMessage "Connecting to Exchange Online..."
Connect-ExchangeOnline -AppId $ExOAppId -Organization $Org -CertificateThumbprint $($ExOCert.Thumbprint) -ShowBanner:$false
Write-SuccessMessage "Connected to Exchange Online"

#Connect-Graph
Write-StatusMessage "Connecting to Microsoft Graph..."
$GraphAppId = "432beb65-bc40-4b40-9366-1c5a768ee717"
$tenantID = "02e68a77-717b-48c1-881a-acc8f67c291a"
$GraphCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=Graph PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $GraphCert) {
    Write-Host "No valid Graph PowerShell certificates found in the LocalMachine\My store. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Connect-Graph -TenantId $TenantId -AppId $GraphAppId -Certificate $GraphCert -NoWelcome

Write-Host "Attempting to find $($UserFromAD.UserPrincipalName) in Azure"

try {
    $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
    $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -ErrorAction Stop
} catch {
    Write-Host "Could not find user $($UserFromAD.UserPrincipalName) in Azure" -ForegroundColor Red -BackgroundColor Black
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
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
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
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
        'l', # l is for Location because Microsoft AD attributes are stupid
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

# Before removing from AD groups
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = Join-Path $Localpath "AD_Groups_Backup_${User}_${timestamp}.csv"
$UserFromAD.MemberOf | ForEach-Object {
    Get-ADGroup $_ | Select-Object Name, DistinguishedName
} | Export-Csv -Path $backupPath -NoTypeInformation
Write-StatusMessage "AD group memberships backed up to: $backupPath"

#remove user from all AD groups
Foreach ($group in $UserFromAD.MemberOf) {
    Remove-ADGroupMember -Identity $group -Members $UserFromAD.SamAccountName -Confirm:$false
}

#Move user to disabled OU
$UserFromAD | Move-ADObject -TargetPath $DestinationOU
#endregion ActiveDirectory

#region Azure
Write-StatusMessage "Performing Azure Steps"

#Revoke all sessions
Revoke-MgUserSignInSession -UserId $UserFromAD.UserPrincipalName -ErrorAction SilentlyContinue

#Remove Mobile Device
Get-MobileDevice -Mailbox $UserFromAD.UserPrincipalName | ForEach-Object { Remove-MobileDevice $_.DeviceID -Confirm:$false -ErrorAction SilentlyContinue }

#Disable AzureAD registered devices
$termUserDeviceId = Get-MgUserRegisteredDevice -UserId $UserFromAD.UserPrincipalName

$termUserDeviceId | ForEach-Object {
    $MgDeviceparams = @{
        AccountEnabled = $false
    }
    Update-MgDevice -DeviceId $_.Id -BodyParameter $MgDeviceparams
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
        Write-ErrorMessage "User mailbox $GrantUserFullControl not found. Skipping access rights setup"
        $GetAccessUserCheck = 'no'
    }

} Else {
    Write-StatusMessage "Skipping access rights setup"
}

if ($GetAccessUserCheck -eq 'yes') {
    Write-StatusMessage "Adding Full Access permissions for $($GetAccessUser.PrimarySmtpAddress) to $($UserFromAD.UserPrincipalName)"
    Add-MailboxPermission -Identity $UserFromAD.UserPrincipalName -User $GrantUserFullControl -AccessRights FullAccess -InheritanceType All -AutoMapping $true
}

# Set Mailbox forwarding address

if ($UserFwdConfirmation -eq 'y') {
    try {
        $GetFWDUser = Get-Mailbox $SetUserMailFWD -ErrorAction Stop
        $GetFWDUserCheck = 'yes'
        Write-StatusMessage "Applying forward from $($UserFromAD.UserPrincipalName) to $($GetFWDUser.PrimarySmtpAddress)"
    } catch {
        Write-ErrorMessage "User mailbox $SetUserMailFWD not found. Skipping mailbox forward"
        $GetFWDUserCheck = 'no'
    }
} Else {
    Write-StatusMessage "Skipping mailbox forwarding"
}

if ($GetFWDUserCheck -eq 'yes') { Set-Mailbox $UserFromAD.UserPrincipalName -ForwardingAddress $SetUserMailFWD -DeliverToMailboxAndForward $False }

#Find user directory roles
$AllDirectoryRoles = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserFromAD.UserPrincipalName).Id | `
    Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' } | `
    Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

#Remove user from directory roles
if (!$AllDirectoryRoles) {
    Write-StatusMessage "Skipping removal of directory roles as user is not assigned."
} else {
    Foreach ($DirectoryRole in $AllDirectoryRoles) {
        Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $DirectoryRole.Id -DirectoryObjectId $MgUser.Id
    }
}
#Find Azure only groups
$AllAzureGroups = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserFromAD.UserPrincipalName).Id | `
    Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.AdditionalProperties.membershipRule -eq $NULL -and $_.onPremisesSyncEnabled -ne 'False' } | `
    Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

$AllAzureGroups | Export-Csv c:\temp\terminated_users_exports\$($user)_Groups_Id.csv -NoTypeInformation

Write-SuccessMessage "Export User Groups Completed. Path: C:\temp\terminated_users_exports\$($user)_Groups_Id.csv"

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

Write-SuccessMessage "Export User Licenses Completed. Path: C:\temp\terminated_users_exports\$($user)_License_Id.csv"

#Remove Licenses
Write-StatusMessage "Starting removal of user licenses."

Get-MgUserLicenseDetail -UserId $UserFromAD.UserPrincipalName | Where-Object `
{ ($_.SkuPartNumber -ne "O365_BUSINESS_ESSENTIALS" -and $_.SkuPartNumber -ne "SPE_E3" -and $_.SkuPartNumber -ne "SPB" -and $_.SkuPartNumber -ne "EXCHANGESTANDARD") } `
| ForEach-Object { Set-MgUserLicense -UserId $UserFromAD.UserPrincipalName -AddLicenses @() -RemoveLicenses $_.SkuId -ErrorAction Stop }

Get-MgUserLicenseDetail -UserId $UserFromAD.UserPrincipalName | ForEach-Object { Set-MgUserLicense -UserId $UserFromAD.UserPrincipalName -AddLicenses @() -RemoveLicenses $_.SkuId }

Write-SuccessMessage "Removal of user licenses completed."

## Sends email to SecurePath Team (soc@compassmsp.com) with the offboarding user information.
$MsgFrom = 'noreply@compassmsp.com'

$Emailparams = @{
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

Send-MgUserMail -UserId $MsgFrom -BodyParameter $Emailparams

## Remove user from Zoom SCIM App

$ZoomSSO = Get-MgUserAppRoleAssignment -UserId $MgUser.Id | Where-Object { $_.ResourceDisplayName -eq 'Zoom Workplace Phones' }

if ($ZoomSSO) {
    Remove-MgUserAppRoleAssignment -AppRoleAssignmentId $ZoomSSO.Id -UserId $MgUser.Id
    Write-SuccessMessage "User has been removed from Zoom Workplace Phones"
} else {
    Write-StatusMessage "User is not assigned to Zoom Workplace Phones"
}

#Disconnect from Exchange and Graph
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-Graph

## Connect to PnP PowerShell
$PnPAppId = "24e3c6ad-9658-4a0d-b85f-82d67d148449"
$Org = "compassmsp.onmicrosoft.com"
$PnPCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=PnP PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $PnPCert) {
    Write-ErrorMessage "No valid PnP PowerShell certificates found in the LocalMachine\My store"
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Connect-PnPOnline -Url compassmsp-admin.sharepoint.com -ClientId $PnPAppId -Tenant $Org -Thumbprint $($PNPCert.Thumbprint)

## Set OneDrive as Read Only
$UserOneDriveURL = (Get-PnPUserProfileProperty -Account $UserFromAD.UserPrincipalName -Properties PersonalUrl).PersonalUrl
Set-PnPTenantSite -Url $UserOneDriveURL -LockState ReadOnly

# Set OneDrive grant access
if ($SPOAccessConfirmation -eq 'y') {
    try {
        $GetUserOneDriveAccess = Get-Mailbox $GrantUserOneDriveAccess -ErrorAction Stop
        $GetUserOneDriveAccessCheck = 'yes'
        Write-StatusMessage "Granting OneDrive access rights to $($GetUserOneDriveAccess.PrimarySmtpAddress)"
    } catch {
        Write-ErrorMessage "User $GrantUserOneDriveAccess not found. Skipping OneDrive access grant"
        $GetUserOneDriveAccessCheck = 'no'
    }
} Else {
    Write-StatusMessage "Skipping OneDrive access grant"
}

if ($GetUserOneDriveAccessCheck -eq 'yes') {
    Set-PnPTenantSite -Url $UserOneDriveURL -Owners $GrantUserOneDriveAccess
    Write-StatusMessage "OneDrive URL: $UserOneDriveURL"
    Write-Host "`nPlease copy the OneDrive URL above if needed." -ForegroundColor Yellow
    Read-Host "Press Enter to continue"
}

Disconnect-PnPOnline

#endregion Office365

#Start AD Sync
Write-StatusMessage "Starting AD sync cycle..."
try {
    powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta
    Write-SuccessMessage "AD sync cycle started successfully"
} catch {
    Write-ErrorMessage "Failed to start AD sync cycle: $($_.Exception.Message)"
}

Write-StatusMessage "`nSummary of Actions:"
Write-StatusMessage "----------------------------------------"
Write-StatusMessage "User disabled: $($UserFromAD.UserPrincipalName)"
Write-StatusMessage "Moved to OU: $DestinationOU"
if ($GetAccessUserCheck -eq 'yes') { Write-StatusMessage "Mailbox access granted to: $GrantUserFullControl" }
if ($GetFWDUserCheck -eq 'yes') { Write-StatusMessage "Mail forwarded to: $SetUserMailFWD" }
if ($GetUserOneDriveAccessCheck -eq 'yes') { Write-StatusMessage "OneDrive access granted to: $GrantUserOneDriveAccess" }
Write-StatusMessage "Exports saved to: $Localpath"
Write-StatusMessage "----------------------------------------`n"

Write-SuccessMessage "User $($User) should now be disabled unless any errors occurred during the process."