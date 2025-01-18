#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#
.SYNOPSIS
    Creates a new user based on a template user in Active Directory and Microsoft 365.

.DESCRIPTION
    This script creates a new user account by copying attributes and group memberships
    from an existing template user. It handles both on-premises AD and Microsoft 365 setup.

    IMPORTANT: This script must be run from the Primary Domain Controller with AD Connect installed.

.PARAMETER NewUser
    The display name of the new user in "FirstName LastName" format.

.PARAMETER UserToCopy
    The display name of the template user to copy from in "FirstName LastName" format.

.PARAMETER Phone
    The mobile phone number for the new user in "(XXX) XXX-XXXX" format.

.EXAMPLE
    .\Invoke-MgNewUserRequest.ps1 -UserToCopy "John Smith" -NewUser "Jane Doe" -Phone "(555) 555-5555"

.NOTES
    Author: Chris Williams
    Created: 2022-03-02
    Last Modified: 2025-01-17

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    3.1        2025-01-17  Added status messaging system with improved error handling and progress tracking
    3.0        2025-01-14  BookWithMeId validation and AD Sync loop changes
    2.9        2025-01-13  Rework custom WPF window
    2.8        2025-01-10  Add function to disable QuickEdit and InsertMode
    2.7        2025-01-03  Added check for duplicate SMTP Address
    2.6        2024-11-11  Added checkbox for EntraID P2 license
    2.5        2024-11-08  Added better UI boxes for variables
    2.4        2024-10-22  Add KB4 offboarding email delivery to SecurePath
    2.3        2024-10-21  Add MeetWithMeId and AD User properties
    2.2        2024-10-15  Remove AppRoleAssignment for KnowBe4 SCIM App
    2.1        2024-05-21  Added stop for if UserToCopy cannot be found
    2.0        2024-05-08  Add input box for Variables
    1.9        2024-03-08  Cleaned up licenses select display output
    1.8        2024-02-19  Changes to Get-MgUserMemberOf function
    1.7        2024-02-14  Fix issues with copy groups function and code cleanup
    1.6        2024-02-12  Add AppRoleAssignment for KnowBe4 SCIM App
    1.5        2022-10-07  Add check for duplicate SamAccountName attributes
    1.4        2022-09-29  Add fax attributes copy
    1.3        2022-06-27  Change Group Lookup and Member Add
    1.2        2022-03-06  Add Check Loop for AD Sync
    1.1        2022-03-04  Add Checks For Duplicate Attributes
    1.0        2022-03-02  Initial Version
    ------------------------------------------------------------------------------
#>

$QuickEditCodeSnippet = @"
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

function Set-ConsoleProperties() {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$EnableQuickEditMode = $false,

        [Parameter(Mandatory = $false)]
        [switch]$DisableQuickEditMode = $false,

        [Parameter(Mandatory = $false)]
        [switch]$EnableInsertMode = $false,

        [Parameter(Mandatory = $false)]
        [switch]$DisableInsertMode = $false
    )

    if ($PSBoundParameters.Count -eq 0) {
        [ConsoleModeSettings]::EnableQuickEditMode()
        [ConsoleModeSettings]::EnableInsertMode()
        Write-StatusMessage "All console settings have been enabled"
        return
    }

    if ($EnableQuickEditMode) {
        [ConsoleModeSettings]::EnableQuickEditMode()
        Write-StatusMessage "QuickEditMode has been enabled"
    }

    if ($DisableQuickEditMode) {
        [ConsoleModeSettings]::DisableQuickEditMode()
        Write-StatusMessage "QuickEditMode has been disabled"
    }

    if ($EnableInsertMode) {
        [ConsoleModeSettings]::EnableInsertMode()
        Write-StatusMessage "InsertMode has been enabled"
    }

    if ($DisableInsertMode) {
        [ConsoleModeSettings]::DisableInsertMode()
        Write-StatusMessage "InsertMode has been disabled"
    }
}

Set-ConsoleProperties -DisableQuickEditMode -DisableInsertMode

#Add these functions at the beginning after the QuickEdit code

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

#Connect-ExchangeOnline
$ExOAppId = "baa3f5d9-3bb4-44d8-b10a-7564207ddccd"
$Org = "compassmsp.onmicrosoft.com"
$ExOCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=ExO PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $ExOCert) {
    Write-ErrorMessage "No valid ExO PowerShell certificates found in the LocalMachine\My store."
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
    Write-ErrorMessage "No valid Graph PowerShell certificates found in the LocalMachine\My store."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Connect-Graph -TenantId $TenantId -AppId $GraphAppId -Certificate $GraphCert -NoWelcome
Write-SuccessMessage "Connected to Microsoft Graph"

# Build out UI for user input
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Function to create and show a custom WPF window
function Show-NewUserRequestWindow {

    # Function to validate display names
    function Test-DisplayName {
        param (
            [string]$DisplayName
        )
        return $DisplayName -match '^[A-Za-z]+ [A-Za-z]+$'  # Regex to check for "First Last"
    }

    # Function to validate and format mobile numbers
    function Format-MobileNumber {
        param (
            [string]$MobileNumber
        )
        # Remove all non-digit characters
        $digits = -join ($MobileNumber -replace '\D', '')

        # Check if we have 10 digits
        if ($digits.Length -eq 10) {
            return "($($digits.Substring(0, 3))) $($digits.Substring(3, 3))-$($digits.Substring(6, 4))"
        }
    }

    # Get available licenses from tenant
    $skus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{
        Name = 'PrepaidUnits'; Expression = { $_.PrepaidUnits.Enabled }
    }

    # Create license display information
    $licenseInfo = $skus | ForEach-Object {
        $available = $_.PrepaidUnits - $_.ConsumedUnits
        $SkuDisplayName = switch -Regex ($_.SkuPartNumber) {
            "POWERAUTOMATE_ATTENDED_RPA" { "Power Automate Premium" ; break }
            "PROJECT_MADEIRA_PREVIEW_IW_SKU" { "Dynamics 365 Business Central for IWs" ; break }
            "PROJECT_PLAN3_DEPT" { "Project Plan 3 (for Department)" ; break }
            "FLOW_FREE" { "Microsoft Power Automate Free" ; break }
            "WINDOWS_STORE" { "Windows Store for Business" ; break }
            "RMSBASIC" { "Rights Management Service Basic Content Protection" ; break }
            "RIGHTSMANAGEMENT_ADHOC" { "Rights Management Adhoc" ; break }
            "POWERAPPS_VIRAL" { "Microsoft Power Apps Plan 2 Trial" ; break }
            "POWERAPPS_PER_USER" { "Power Apps Premium" ; break }
            "POWERAPPS_DEV" { "Microsoft PowerApps for Developer" ; break }
            "PHONESYSTEM_VIRTUALUSER" { "Microsoft Teams Phone Resource Account" ; break }
            "MICROSOFT_BUSINESS_CENTER" { "Microsoft Business Center" ; break }
            "MCOPSTNC" { "Communications Credits" ; break }
            "MCOPSTN1" { "Skype for Business PSTN Domestic Calling" ; break }
            "MEETING_ROOM" { "Microsoft Teams Rooms Standard" ; break }
            "MCOMEETADV" { "Microsoft 365 Audio Conferencing" ; break }
            "CCIBOTS_PRIVPREV_VIRAL" { "Power Virtual Agents Viral Trial" ; break }
            "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" ; break }
            "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" ; break }
            "SPE_E3" { "Microsoft 365 E3" ; break }
            "SPB" { "Microsoft 365 Business Premium" ; break }
            "ENTERPRISEPACK" { "Office 365 E3" ; break }
            "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" ; break }
            "PROJECT_P1" { "Project Plan 1" ; break }
            "PROJECTPROFESSIONAL" { "Project Plan 3" ; break }
            "VISIOCLIENT" { "Visio Plan 2" ; break }
            "Microsoft_Teams_Audio_Conferencing_select_dial_out" { "Microsoft Teams Audio Conferencing with dial-out to USA/CAN" ; break }
            "POWER_BI_PRO" { "Power BI Pro" ; break }
            "Microsoft_365_Copilot" { "Microsoft 365 Copilot" ; break }
            "Microsoft_Teams_Premium" { "Microsoft Teams Premium" ; break }
            "MCOEV" { "Microsoft Teams Phone Standard" ; break }
            "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" ; break }
            "POWER_BI_STANDARD" { "Power BI Standard" ; break }
            "Microsoft365_Lighthouse" { "Microsoft 365 Lighthouse" ; break }
            default { $_.SkuPartNumber }
        }

        # If $SkuDisplayName is null or empty, use the original SkuPartNumber
        if ([string]::IsNullOrEmpty($SkuDisplayName)) {
            $SkuDisplayName = $_.SkuPartNumber
        }

        @{
            DisplayName = "$($SkuDisplayName) (Available: $available)"
            SkuId       = $_.SkuId
            SortName    = $SkuDisplayName  # Add this for sorting
        }
    } | Sort-Object { $_.SortName }  # Sort by the display name without the "Available" count

    # Create a new WPF window
    $window = New-Object System.Windows.Window
    $window.Title = "New User Request"
    $window.Width = 500  # Set a wider fixed width for the window
    $window.Height = 560  # Set the ideal height for the window
    $window.WindowStartupLocation = 'CenterScreen'

    # Create a StackPanel to hold the controls
    $mainPanel = New-Object System.Windows.Controls.StackPanel
    $mainPanel.Margin = '3'  # Add margin around the stack panel
    $window.Content = $mainPanel

    # Create a StackPanel for the new user input and checkbox
    $newuserPanel = New-Object System.Windows.Controls.StackPanel
    $newuserPanel.Margin = '0,0,0,3'  # Add margin below the new user panel

    # Create and add a label for the new user
    $newuserLabel = New-Object System.Windows.Controls.Label
    $newuserLabel.Content = "New User (First Last):"
    $newuserLabel.Margin = '0,0,0,4'  # Add margin below the label
    $newuserPanel.Children.Add($newuserLabel)

    # Create and add TextBox for New User
    $newUserTextBox = New-Object System.Windows.Controls.TextBox
    $newUserTextBox.Margin = '0,0,0,3'  # Add margin below the text box
    $newuserPanel.Children.Add($newUserTextBox)  # Add to new user panel

    # Add the new user panel to the stack panel
    $mainPanel.Children.Add($newuserPanel)

    # Create a StackPanel for the copy user input and checkbox
    $copyuserPanel = New-Object System.Windows.Controls.StackPanel
    $copyuserPanel.Margin = '0,0,0,3'  # Add margin below the copy user panel

    # Create and add a label for the copy user
    $copyuserLabel = New-Object System.Windows.Controls.Label
    $copyuserLabel.Content = "User To Copy (First Last):"
    $copyuserLabel.Margin = '0,0,0,4'  # Add margin below the label
    $copyuserPanel.Children.Add($copyuserLabel)

    # Create and add TextBox for User To Copy
    $userToCopyTextBox = New-Object System.Windows.Controls.TextBox
    $userToCopyTextBox.Margin = '0,0,0,3'  # Add margin below the text box
    $copyuserPanel.Children.Add($userToCopyTextBox)  # Add to copy user panel

    # Add the copy user panel to the stack panel
    $mainPanel.Children.Add($copyuserPanel)

    # Create a StackPanel for the mobile number input and checkbox
    $mobilePanel = New-Object System.Windows.Controls.StackPanel
    $mobilePanel.Margin = '0,0,0,3'  # Add margin below the mobile panel

    # Create and add a label for the mobile number
    $mobileLabel = New-Object System.Windows.Controls.Label
    $mobileLabel.Content = "Mobile Number:"
    $mobileLabel.Margin = '0,0,0,4'  # Add margin below the label
    $mobilePanel.Children.Add($mobileLabel)

    # Create and add the mobile number text box
    $mobileTextBox = New-Object System.Windows.Controls.TextBox
    $mobileTextBox.Margin = '0,0,0,5'  # Add margin below the text box
    $mobilePanel.Children.Add($mobileTextBox)

    # Create a CheckBox for bypassing mobile number formatting
    $bypassFormattingCheckBox = New-Object System.Windows.Controls.CheckBox
    $bypassFormattingCheckBox.Content = "Bypass Mobile Number Formatting"
    $bypassFormattingCheckBox.Margin = '0,0,0,3'  # Add margin below the checkbox
    $mobilePanel.Children.Add($bypassFormattingCheckBox)

    # Add the mobile panel to the stack panel
    $mainPanel.Children.Add($mobilePanel)

    # Modify the licenses section to create two separate controls
    $requiredLicenses = @(
        "Exchange Online (Plan 1)",
        "Office 365 E3",
        "Microsoft 365 Business Basic",
        "Microsoft 365 E3",
        "Microsoft 365 Business Premium"
    )

    $ignoredLicenses = @(
        "Microsoft Teams Rooms Standard",
        "Microsoft Teams Phone Standard",
        "Power Automate Premium",
        "Power Apps Premium",
        "Power BI Pro",
        "Power BI Standard",
        "Microsoft 365 Lighthouse",
        "Rights Management Service Basic Content Protection",
        "Communications Credits",
        "Rights Management Adhoc",
        "Power Virtual Agents Viral Trial",
        "Windows Store for Business",
        "Skype for Business PSTN Domestic Calling",
        "Microsoft Business Center",
        "Microsoft Teams Phone Resource Account"
        "Microsoft PowerApps for Developer",
        "Microsoft Power Apps Plan 2 Trial",
        "Microsoft Power Automate Free",
        "Microsoft_Copilot_for_Finance_trial",
        "STREAM",
        "Project Plan 3 (for Department)",
        "Dynamics 365 Business Central for IWs"
    )

    # Required License ComboBox Section
    $requiredGroup = New-Object System.Windows.Controls.GroupBox
    $requiredGroup.Header = "Required License (Select One)"
    $requiredGroup.Margin = "10"

    $requiredComboBox = New-Object System.Windows.Controls.ComboBox
    $requiredComboBox.Margin = "5"

    # Create a custom object for each required license
    foreach ($license in $licenseInfo) {
        foreach ($reqLicense in $requiredLicenses) {
            if ($license.DisplayName -like "*$reqLicense*") {
                $item = [PSCustomObject]@{
                    DisplayName = $license.DisplayName
                    SkuId       = $license.SkuId
                }
                $requiredComboBox.Items.Add($item)
            }
        }
    }

    # Set the DisplayMemberPath to show the DisplayName
    $requiredComboBox.DisplayMemberPath = "DisplayName"
    $requiredGroup.Content = $requiredComboBox
    $mainPanel.Children.Add($requiredGroup)

    # Ancillary Licenses Section
    $ancillaryGroup = New-Object System.Windows.Controls.GroupBox
    $ancillaryGroup.Header = "Ancillary Licenses"
    $ancillaryGroup.Margin = "10"

    # Create ScrollViewer for ancillary licenses
    $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
    $scrollViewer.VerticalScrollBarVisibility = "Auto"
    $scrollViewer.MaxHeight = 200

    $licensesStack = New-Object System.Windows.Controls.StackPanel
    $scrollViewer.Content = $licensesStack
    $ancillaryGroup.Content = $scrollViewer

    $SkuCheckBoxes = @()
    foreach ($license in $licenseInfo) {
        # Skip licenses that are in the required licenses list
        $isRequired = $false
        foreach ($reqLicense in $requiredLicenses) {
            if ($license.DisplayName -like "*$reqLicense*") {
                $isRequired = $true
                break
            }
        }
        $isIgnored = $false
        foreach ($ignoredLicense in $ignoredLicenses) {
            if ($license.DisplayName -like "*$ignoredLicense*") {
                $isIgnored = $true
                break
            }
        }
        if (-not $isRequired -and -not $isIgnored) {
            $skucb = New-Object System.Windows.Controls.CheckBox
            $skucb.Content = $license.DisplayName
            $skucb.Tag = $license.SkuId
            $skucb.Margin = "5,5,5,5"
            if ($license.DisplayName -like "*Microsoft Entra ID P2*") {
                $skucb.IsChecked = $true
            }
            $SkuCheckBoxes += $skucb
            $licensesStack.Children.Add($skucb)
        }
    }
    $mainPanel.Children.Add($ancillaryGroup)

    # Create and add OK and Cancel buttons
    $buttonPanel = New-Object System.Windows.Controls.StackPanel
    $buttonPanel.Orientation = 'Horizontal'
    $buttonPanel.HorizontalAlignment = 'Right'
    $buttonPanel.Margin = '0,10,0,0'  # Add margin above the button panel

    $okButton = New-Object System.Windows.Controls.Button
    $okButton.Content = "OK"
    $okButton.Margin = '0,0,10,0'  # Add margin to the right of the OK button
    $okButton.Add_Click({
            # Validate New User input
            if (-not $newUserTextBox.Text) {
                [System.Windows.MessageBox]::Show("New User is a mandatory field. Please enter a valid Display Name.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
            if (-not (Test-DisplayName $newUserTextBox.Text)) {
                [System.Windows.MessageBox]::Show("Invalid format for New User. Please use 'First Last' name format.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }

            # Validate User To Copy input
            if (-not $userToCopyTextBox.Text) {
                [System.Windows.MessageBox]::Show("User To Copy is a mandatory field. Please enter a Display Name.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }
            if (-not (Test-DisplayName $userToCopyTextBox.Text)) {
                [System.Windows.MessageBox]::Show("Invalid format for User To Copy. Please use 'First Last' name format.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
            }

            # Validate Mobile Number
            if (-not $bypassFormattingCheckBox.IsChecked) {
                $unformattedMobile = $mobileTextBox.Text
                $digits = -join ($unformattedMobile -replace '\D', '')  # Remove non-digit characters
                if ($digits.Length -ne 10) {
                    [System.Windows.MessageBox]::Show("Invalid mobile number format. Please enter a valid 10-digit mobile number.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                    return
                }
            }

            # Validate required license selection
            if ($null -eq $requiredComboBox.SelectedItem) {
                [System.Windows.MessageBox]::Show(
                    "Please select a required license.",
                    "Required License Missing",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning)
                return
            }

            # Get selected licenses (both required and ancillary)
            $script:selectedLicenses = @()
            $script:selectedLicenses += $requiredComboBox.SelectedItem.SkuId
            $script:selectedLicenses += ($SkuCheckBoxes | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag })

            # Check for available licenses
            if ($requiredComboBox.SelectedItem.DisplayName -match "Available: (\d+)") {
                $availableCount = [int]$Matches[1]
                if ($availableCount -eq 0) {
                    $licenseName = $requiredComboBox.SelectedItem.DisplayName -replace ' \(Available: \d+\)$', ''
                    [System.Windows.MessageBox]::Show(
                        "$licenseName has no licenses available.",
                        "No Available Licenses",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning)
                    return
                }
            }

            # Check ancillary licenses availability
            $selectedCheckboxes = $SkuCheckBoxes | Where-Object { $_.IsChecked }
            foreach ($cb in $selectedCheckboxes) {
                if ($cb.Content -match "Available: (\d+)") {
                    $availableCount = [int]$Matches[1]
                    if ($availableCount -eq 0) {
                        $licenseName = $cb.Content -replace ' \(Available: \d+\)$', ''
                        [System.Windows.MessageBox]::Show(
                            "$licenseName has no licenses available.",
                            "No Available Licenses",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning)
                        return
                    }
                }
            }

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

    $mainPanel.Children.Add($buttonPanel)

    # Show the window
    $result = $window.ShowDialog()

    # Initialize formattedMobile variable
    $formattedMobile = $null  # Initialize to null
    if ($null -ne $mobileTextBox.Text) {
        # Check if the checkbox is checked
        if (-not $bypassFormattingCheckBox.IsChecked) {
            $formattedMobile = Format-MobileNumber $mobileTextBox.Text
        } else {
            $formattedMobile = $mobileTextBox.Text  # Use the unformatted mobile number
        }
    }

    if ($result -eq $true) {
        return @{
            InputNewUser           = $newUserTextBox.Text
            InputNewMobile         = $formattedMobile
            InputUserToCopy        = $userToCopyTextBox.Text
            InputRequiredLicense   = @{
                SkuId       = $requiredComboBox.SelectedItem.SkuId
                DisplayName = $requiredComboBox.SelectedItem.DisplayName
            }
            InputAncillaryLicenses = ($SkuCheckBoxes |
                Where-Object { $_.IsChecked } |
                ForEach-Object {
                    @{
                        SkuId       = $_.Tag
                        DisplayName = $_.Content
                    }
                })
        }
    } else {
        return $null
    }
}

# Call the custom input window function
$result = Show-NewUserRequestWindow

# Setting variables from window function result
$NewUser = $result.InputNewUser
$Phone = $result.InputNewMobile
$UserToCopy = $result.InputUserToCopy

$RequiredLicense = $result.InputRequiredLicense.SkuId
$AncillaryLicenses = $result.InputAncillaryLicenses.SkuId

$UserToCopyUPN = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -Properties Title, Fax, wWWHomePage, physicalDeliveryOfficeName, Office, Manager, Description, Department, Company

## Check for duuplicate DisplayName in AD for selected UserToCopy
if ($UserToCopyUPN.Count -gt 1) {
    Write-Host "UserToCopy has multiple values. Please check AD for accounts with duplicate DisplayName attributes. Press any key to exit script." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
} elseif ($NULL -eq $UserToCopyUPN) {
    Write-Host "Could not find user $($UserToCopy) in AD to copy from. Press any key to exit script." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

## Building out new user variables
$Domain = $($UserToCopyUPN.UserPrincipalName -replace '.+?(?=@)')
$NewUserFirstName = $($NewUser.split(' ')[-2])
$NewUserLastName = $($NewUser -replace '.+\s')
$NewUserSamAccountName = $(($NewUser -replace '(?<=.{1}).+') + ($NewUser -replace '.+\s')).ToLower()
$NewUserEmail = $($NewUserSamAccountName + $Domain).ToLower()

$CheckNewUserUPN = $(try { Get-ADUser -Identity $NewUserSamAccountName } catch { $null })
if ($null -ne $CheckNewUserUPN) {
    Write-Host "SamAccountName exist for user $NewUser. Please check AD for accounts with duplicate SamAccountName attributes. Press any key to exit script." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

$CheckNewUserEmail = $(try { Get-MgUser -UserId $NewUserEmail } catch { $null })
if ($null -ne $CheckNewUserEmail) {
    Write-Host "Account in 365 exist for user $NewUser. Please check 365 for accounts with duplicate SMTP Address. Press any key to exit script." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

## New user creation in AD
function Get-NewPassword { -join ('abcdefghkmnrstuvwxyzABCDEFGHKLMNPRSTUVWXYZ23456789$%&*#'.ToCharArray() | Get-Random -Count 16) }

$Password = Get-NewPassword

$Confirmation = Read-Host -Prompt "The user below will be created:`n
Display Name = $($NewUser)
Email Address = $($NewUserEmail)
Password = $($Password)
First Name = $($NewUserFirstName)
Last Name = $($NewUserLastName)
SamAccountName = $($NewUserSamAccountName)
Destination OU = $($UserToCopyUPN.DistinguishedName.split(",",2)[1])`n
Template User to Copy = $($UserToCopy)`n
Continue? (Y/N)`n"

if ($Confirmation -ne 'y') {
    Write-WarningMessage 'User cancelled operation'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

try {
    Write-StatusMessage "Creating new AD user: $NewUser"
    New-ADUser -Name $NewUser `
        -SamAccountName $NewUserSamAccountName `
        -UserPrincipalName $NewUserEmail `
        -DisplayName $NewUser `
        -GivenName $NewUserFirstName `
        -Surname $NewUserLastName  `
        -MobilePhone $Phone `
        -EmailAddress $NewUserEmail `
        -OtherAttributes @{ 'proxyAddresses' = "SMTP:$($NewUserEmail)" } `
        -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
        -Path $($UserToCopyUPN.DistinguishedName.split(",", 2)[1]) `
        -Instance $UserToCopyUPN `
        -Enabled $True
    Write-SuccessMessage "AD User created successfully"
} catch {
    Write-ErrorMessage "Failed to create AD User: $_"
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

Write-StatusMessage "Adding AD Groups to new user"
$CopyFromUser = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -prop MemberOf
$CopyToUser = Get-ADUser -Filter "DisplayName -eq '$($NewUser)'" -prop MemberOf
$groupCount = 0
$CopyFromUser.MemberOf | Where-Object { $CopyToUser.MemberOf -notcontains $_ } | ForEach-Object {
    try {
        Add-ADGroupMember -Identity $_ -Members $CopyToUser
        $groupCount++
    } catch {
        Write-WarningMessage "Failed to add user to group $_"
    }
}
Write-SuccessMessage "Added user to $groupCount AD groups"

Write-StatusMessage "Starting AD Sync"

try {
    Import-Module -UseWindowsPowerShell -Name ADSync
    Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
} catch {
    powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta
}

Write-StatusMessage "Waiting for AD Connect sync..."
$progress = 0
1..30 | ForEach-Object {
    Write-Progress -Activity "Waiting for AD Connect Sync" -Status "$($progress)% Complete" -PercentComplete $progress
    Start-Sleep -Seconds 1
    $progress += (100/30)
}
Write-Progress -Activity "Waiting for AD Connect Sync" -Completed

## Check if AD User has synced to Azure loop
$Stoploop = $false
[int]$Retrycount = "0"

do {
    try {
        $MgUser = Get-MgUser -UserId $NewUserEmail -Property Id, Mail, displayName, Department | Select-Object Id, Mail, displayName, Department -ErrorAction Stop
        Write-SuccessMessage "User $NewUser has synced to Azure"
        $Stoploop = $true
        $ADSyncCompleteYesorExit = 'yes'
    } catch {
        if ($Retrycount -gt 3) {
            Write-ErrorMessage "Could not sync AD User to 365 after 3 retries"
            $Stoploop = $true
        } else {
            Write-WarningMessage "Could not sync AD User to 365 retrying in 30 seconds..."
            Start-Sleep -Seconds 30
            $Retrycount = $Retrycount + 1
        }
    }
} while ($Stoploop -eq $false)

if (!$MgUser) {
    $ADSyncCompleteYesorExit = Read-Host -Prompt 'AD Sync has not completed within allotted time frame. Please wait for AD sync. To resume type yes or exit'
} while ("yes", "exit" -notcontains $ADSyncCompleteYesorExit ) {
    $ADSyncCompleteYesorExit = Read-Host "Please enter your response (yes/exit)"
}

if ($ADSyncCompleteYesorExit -eq 'exit') {
    Write-ErrorMessage 'You will need to set the license and add Office 365 groups via the portal'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

if ($ADSyncCompleteYesorExit -eq 'yes') {

    $MgUserCopy = Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName

    if (!$MgUser) {
        try {
            $MgUser = Get-MgUser -UserId $NewUserEmail -Property Id, Mail, displayName, Department | Select-Object Id, Mail, displayName, Department -ErrorAction Stop
        } catch {
            Write-ErrorMessage 'Script cannot find new user. You will need to set the license and add Office 365 groups via the portal'
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            exit
        }
    }

    Write-StatusMessage 'Setting Usage Location for new user'
    Update-MgUser -UserId $MgUser.Id -UsageLocation US
    Write-SuccessMessage 'Usage Location set to US'

    function Set-UserLicenses {
        param(
            [Parameter(Mandatory = $true)]
            [string]$UserId,

            [Parameter(Mandatory = $true)]
            [string[]]$License
        )

        try {
            foreach ($sku in $License) {
                Write-StatusMessage "Assigning license $sku to user: $UserId"
                Set-MgUserLicense -UserId $UserId -AddLicenses @{SkuId = $sku } -RemoveLicenses @() -ErrorAction Stop | Out-Null
                Write-SuccessMessage "Successfully assigned license $sku"
            }
        } catch {
            Write-ErrorMessage "Failed to assign license: $_"
        }
    }

    Set-UserLicenses -UserId $MgUser.Id -License $RequiredLicense

    Start-Sleep -Seconds 30

    if ($null -ne $AncillaryLicenses) {
        Set-UserLicenses -UserId $MgUser.Id -License $AncillaryLicenses
    }

    ## Add BookWithMeId to the extensionAttribute15 property of the new user.
    Write-StatusMessage "Setting up BookWithMeId attribute"
    try {
        $NewUserExchGuid = (Get-Mailbox -Identity $NewUserEmail).ExchangeGuid.Guid -replace "-" -replace ""
        $extAttr15 = $NewUserExchGuid + '@compassmsp.com?anonymous&ep=plink'

        if ($extAttr15 -eq '@compassmsp.com?anonymous&ep=plink') {
            Write-ErrorMessage "$NewUserSamAccountName BookWithMeId missing ExchangeGuidId. Please add the attribute in AD manually."
        } else {
            Set-ADUser -Identity $NewUserSamAccountName -Add @{extensionAttribute15 = "$extAttr15" }
            Write-SuccessMessage "BookWithMeId attribute set successfully"
        }
    } catch {
        Write-ErrorMessage "Failed to set BookWithMeId attribute: $_"
    }

    ## Provision New Users OneDrive
    Get-MgUserDefaultDrive -UserId $MgUser.Id

    Write-StatusMessage 'Adding Office 365 Groups to new user.'

    ## Copy groups to new user from old user
    $All365Groups = Get-MgUserMemberOf -UserId $MgUserCopy.Id | `
        Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.AdditionalProperties.membershipRule -eq $NULL -and $_.onPremisesSyncEnabled -ne 'False' } | `
        Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

    $groupProgress = 0
    $totalGroups = $All365Groups.Count

    Foreach ($365Group in $All365Groups) {
        $groupProgress++
        Write-Progress -Activity "Adding Office 365 Groups" -Status "Processing group $($365Group.DisplayName)" `
            -PercentComplete (($groupProgress / $totalGroups) * 100)
        try {
            New-MgGroupMember -GroupId $365Group.Id -DirectoryObjectId $MgUser.Id -ErrorAction Stop
            Write-StatusMessage "Added to 365 group: $($365Group.DisplayName)"
        } catch {
            Add-DistributionGroupMember -Identity $365Group.DisplayName -Member $NewUserEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction 'SilentlyContinue'
            Write-StatusMessage "Added to Distribution group: $($365Group.DisplayName)"
        }
    }

    $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $MgUserCopy.Id).Count
    $NewUserGroupCount = (Get-MgUserMemberOf -UserId $MgUser.Id).Count

    function Add-UserToZoom {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
            $MgUser
        )

        try {
            Write-StatusMessage "Setting up Zoom for user: $($MgUser.DisplayName)"

            # Determine Zoom role based on department
            if ($MgUser.Department -eq 'Reactive') {
                $zoom_app_role_name = "Basic"
            } else {
                $zoom_app_role_name = "Licensed"
            }

            $zoom_app_name = "Zoom Workplace Phones"

            try {
                Write-StatusMessage "Getting Zoom service principal"
                $zoom_ServicePrincipal = Get-MgServicePrincipal -Filter "displayName eq '$zoom_app_name'"
                if (-not $zoom_ServicePrincipal) {
                    Write-ErrorMessage "Zoom Service Principal not found. Skipping Zoom setup."
                    return
                }

                $zoom_synchronizationJob = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $zoom_ServicePrincipal.Id
                $zoom_synchronizationJobRuleId = (Get-MgServicePrincipalSynchronizationJobSchema -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_synchronizationJob.Id).SynchronizationRules.Id
                Write-SuccessMessage "Retrieved Zoom app details"
            }
            catch {
                Write-ErrorMessage "Failed to get Zoom app details: $_"
                throw
            }

            try {
                Write-StatusMessage "Assigning user to Zoom app with role: $zoom_app_role_name"
                $params = @{
                    "PrincipalId" = $MgUser.Id
                    "ResourceId"  = $zoom_ServicePrincipal.Id
                    "AppRoleId"   = ($zoom_ServicePrincipal.AppRoles | Where-Object { $_.DisplayName -eq $zoom_app_role_name }).Id
                }

                New-MgUserAppRoleAssignment -UserId $MgUser.Id -BodyParameter $params |
                    Format-List Id, AppRoleId, CreationTime, PrincipalDisplayName, PrincipalId, PrincipalType, ResourceDisplayName, ResourceId

                Write-SuccessMessage "User assigned to Zoom app"
                Write-StatusMessage "Waiting 30 seconds before syncing..."
                Start-Sleep -Seconds 30
            }
            catch {
                Write-ErrorMessage "Failed to assign user to Zoom app: $_"
                throw
            }

            try {
                Write-StatusMessage "Initiating Zoom sync"
                $params = @{
                    parameters = @(
                        @{
                            subjects = @(
                                @{
                                    objectId       = "$($MgUser.Id)"
                                    objectTypeName = "User"
                                }
                            )
                            ruleId   = $zoom_synchronizationJobRuleId
                        }
                    )
                }

                New-MgServicePrincipalSynchronizationJobOnDemand -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_synchronizationJob.Id -BodyParameter $params
                Write-SuccessMessage "Zoom sync initiated successfully"
            }
            catch {
                Write-ErrorMessage "Failed to sync user to Zoom: $_"
                throw
            }
        }
        catch {
            Write-ErrorMessage "Failed to process Zoom app assignment: $_"
            throw
        }
    }

    # After creating the new user
    Add-UserToZoom -MgUser $MgUser

    Write-StatusMessage "Sending notification email to SecurePath team"
    ## Sends email to SecurePath Team (soc@compassmsp.com) with the new user information.
    $MsgFrom = 'noreply@compassmsp.com'

    $params = @{
        message         = @{
            subject      = "KB4 – New User"
            body         = @{
                contentType = "HTML"
                content     = "The following user need to be added to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
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
    Write-SuccessMessage "Notification email sent"

    Write-StatusMessage "Disconnecting from services"
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-Graph
    Write-SuccessMessage "Disconnected from all services"

    # At the end, show a summary
    Write-StatusMessage "=== New User Creation Summary ===" -Color Cyan
    Write-StatusMessage "User: $NewUser" -Color Cyan
    Write-StatusMessage "Email: $NewUserEmail" -Color Cyan
    Write-StatusMessage "Password: $Password" -Color Cyan
    Write-StatusMessage "Template User Groups: $CopyUserGroupCount" -Color Cyan
    Write-StatusMessage "New User Groups: $NewUserGroupCount" -Color Cyan
    if ($CopyUserGroupCount -ne $NewUserGroupCount) {
        Write-WarningMessage "Group count mismatch - please verify group assignments"
    }
    Write-StatusMessage "Process completed" -Color Cyan

}