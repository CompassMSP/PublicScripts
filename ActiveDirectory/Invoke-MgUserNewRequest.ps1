#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#Author       : Chris Williams
# Creation Date: 03-02-2022
# Usage        : Copies user template and creates new user with groups and licenses

#********************************************************************************
# Date                     Version      Changes
#--------------------------------------------------------------------------------
# 03-02-2022                    1.0         Initial Version
# 03-04-2022                    1.1         Add Checks For Duplicate Attributes
# 03-06-2022                    1.2         Add Check Loop for AD Sync
# 06-27-2022                    1.3         Change Group Lookup and Member Add
# 09-29-2022                    1.4         Add fax attributes copy
# 10-07-2022                    1.5         Add check for duplicate SamAccountName attributes
# 02-12-2024                    1.6         Add AppRoleAssignment for KnowBe4 SCIM App
# 02-14-2024                    1.7         Fix issues with copy groups function and code cleanup
# 02-19-2024                    1.8         Changes to Get-MgUserMemberOf function
# 03-08-2024                    1.9         Cleaned up licenses select display output
# 05-08-2024                    2.0         Add input box for Variables
# 05-21-2024                    2.1         Added stop for if UserToCopy cannot be found
# 10-15-2024                    2.2         Remove AppRoleAssignment for KnowBe4 SCIM App
# 10-21-2024                    2.3         Add MeetWithMeId and AD User properties
# 10-22-2024                    2.4         Add KB4 offboarding email delivery to SecurePath
# 11-08-2024                    2.5         Added better UI boxes for variables
# 11-11-2024                    2.6         Added added checkbox for EntraID P2 license
# 01-03-2025                    2.7         Added added check for duplicate SMTP Address
#********************************************************************************
#
# Run from the Primary Domain Controller with AD Connect installed
#
# The following modules must be installed
# Install-Module ExchangeOnlineManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, PnP.PowerShell
#
# Azure licenses Sku - Selected Sku must have free licenses available. This MUST be set in the portal before running the script
#
# Exchange Online = EXCHANGESTANDARD
# Microsoft 365 Business Basic = O365_BUSINESS_ESSENTIALS
# Microsoft 365 E3 = SPE_E3
# Microsoft 365 Business Premium = SPB
# Office 365 E3 = ENTERPRISEPACK
#
# .\Invoke-MgNewUserRequest.ps1 -UserToCopy "Copy User" -NewUser "Chris Williams" -Phone "555-555-5555"
#>

#Import-Module adsync -UseWindowsPowerShell

#Connect-ExchangeOnline
$ExOAppId = "baa3f5d9-3bb4-44d8-b10a-7564207ddccd"
$Org = "compassmsp.onmicrosoft.com"
$ExOCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=ExO PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $ExOCert) {
    Write-Host "No valid ExO PowerShell certificates found in the LocalMachine\My store. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Connect-ExchangeOnline -AppId $ExOAppId -Organization $Org -CertificateThumbprint $($ExOCert.Thumbprint) -ShowBanner:$false

#Connect-Graph
Write-Host "Logging into Azure services."
$GraphAppId = "432beb65-bc40-4b40-9366-1c5a768ee717"
$tenantID = "02e68a77-717b-48c1-881a-acc8f67c291a"
$GraphCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=Graph PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
if ($NULL -eq $GraphCert) {
    Write-Host "No valid Graph PowerShell certificates found in the LocalMachine\My store. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}
Connect-Graph -TenantId $TenantId -AppId $GraphAppId -Certificate $GraphCert -NoWelcome

# Build out UI for user input
Add-Type -AssemblyName PresentationFramework

# Define the properties to select from the subscribed SKUs
$SelectObjectPropertyList = @(
    "SkuPartNumber"
    "SkuId"
    @{
        n = "ActiveUnits"
        e = { ($_.PrepaidUnits).Enabled }
    }
    "ConsumedUnits"
)

# Define the filter for the SKUs
$WhereObjectFilter = {
    ($_.SkuPartNumber -eq 'EXCHANGESTANDARD') -or
    ($_.SkuPartNumber -eq 'O365_BUSINESS_ESSENTIALS') -or
    ($_.SkuPartNumber -eq 'SPE_E3') -or
    ($_.SkuPartNumber -eq 'SPB') -or
    ($_.SkuPartNumber -eq 'ENTERPRISEPACK') -or
    ($_.SkuPartNumber -eq "AAD_PREMIUM_P2")
}

# Retrieve the available licenses
$selectLicense = Get-MgSubscribedSku |
Select-Object $SelectObjectPropertyList |
Where-Object -FilterScript $WhereObjectFilter |
ForEach-Object {
    [PSCustomObject]@{
        Available = ($_.ActiveUnits - $_.ConsumedUnits)
        SkuName   = switch -Regex ($_.SkuPartNumber) {
            "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
            "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
            "SPE_E3" { "Microsoft 365 E3" }
            "SPB" { "Microsoft 365 Business Premium" }
            "ENTERPRISEPACK" { "Office 365 E3" }
            "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
        }
    }
} | Sort-Object SkuName

# Format the license information for display
$licenseInfo = "Available Licenses:`n"
foreach ($license in $selectLicense) {
    $licenseInfo += "$($license.SkuName): $($license.Available) available`n"  # Simple list format
}

# Show a pop-up message with available licenses before the input window
[System.Windows.MessageBox]::Show("Please check the Microsoft 365 portal for available licenses.`n`n$licenseInfo", "License Check", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

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

# Function to create and show a custom WPF window
function Show-CustomNewUserRequestWindow {
    # Create a new WPF window
    $window = New-Object System.Windows.Window
    $window.Title = "New User Request"
    $window.Width = 500  # Set a wider fixed width for the window
    $window.Height = 330  # Set the ideal height for the window
    $window.WindowStartupLocation = 'CenterScreen'

    # Create a StackPanel to hold the controls
    $stackPanel = New-Object System.Windows.Controls.StackPanel
    $stackPanel.Margin = '3'  # Add margin around the stack panel
    $window.Content = $stackPanel

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
    $stackPanel.Children.Add($newuserPanel)

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
    $stackPanel.Children.Add($copyuserPanel)

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
    $stackPanel.Children.Add($mobilePanel)

    # Create a ComboBox for License SKU selection
    $SkucomboBoxLabel = New-Object System.Windows.Controls.Label
    $SkucomboBoxLabel.Content = "Select License SKU:"
    $SkucomboBoxLabel.Margin = '0,0,0,4'  # Add margin below the label
    $stackPanel.Children.Add($SkucomboBoxLabel)

    $SkucomboBox = New-Object System.Windows.Controls.ComboBox
    $SkucomboBox.Margin = '0,0,0,3'
    $SkucomboBox.ItemsSource = @('Exchange Online (Plan 1)', 'Microsoft 365 Business Basic', 'Microsoft 365 E3', 'Microsoft 365 Business Premium', 'Office 365 E3')
    $stackPanel.Children.Add($SkucomboBox)

    # Create a CheckBox for adding the EntraID P2
    $AddEntraIDP2CheckBox = New-Object System.Windows.Controls.CheckBox
    $AddEntraIDP2CheckBox.Content = "Add EntraID P2"
    $AddEntraIDP2CheckBox.Margin = '0,0,0,3'  # Add margin below the checkbox
    $AddEntraIDP2CheckBox.IsChecked = $true  # Set the checkbox to checked by default
    $stackPanel.Children.Add($AddEntraIDP2CheckBox)

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

            # Validate License SKU selection
            if (-not $SkucomboBox.SelectedItem) {
                [System.Windows.MessageBox]::Show("Selecting a License SKU is mandatory. Please select an option.", "Input Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                return
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

    if ($AddEntraIDP2CheckBox.IsChecked) {
        $AddEntraIDP2 = 'Yes'
    }

    if ($result -eq $true) {
        return @{
            InputNewUser      = $newUserTextBox.Text
            InputNewMobile    = $formattedMobile  # Use the formatted or unformatted mobile number
            InputUserToCopy   = $userToCopyTextBox.Text
            InputSku          = $SkucomboBox.SelectedItem
            InputSkuEntraIDP2 = $AddEntraIDP2
        }
    } else {
        return $null
    }
}

# Call the custom input window function
$result = Show-CustomNewUserRequestWindow

# Setting variables from window function result
$NewUser = $result.InputNewUser
$Phone = $result.InputNewMobile
$UserToCopy = $result.InputUserToCopy

## Set Sku from Sku displayName
if ($result.InputSku -eq 'Exchange Online (Plan 1)') { $Sku = "EXCHANGESTANDARD" }
if ($result.InputSku -eq 'Microsoft 365 Business Basic') { $Sku = "O365_BUSINESS_ESSENTIALS" }
if ($result.InputSku -eq 'Microsoft 365 E3') { $Sku = "SPE_E3" }
if ($result.InputSku -eq 'Microsoft 365 Business Premium') { $Sku = "SPB" }
if ($result.InputSku -eq 'Office 365 E3') { $Sku = "ENTERPRISEPACK" }

$UserToCopyUPN = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -Properties Title, Fax, wWWHomePage, physicalDeliveryOfficeName, Office, Manager, Description, Department, Company

## Check for duuplicate DisplayName in AD for selected UserToCopy
if ($UserToCopyUPN.Count -gt 1) {
    Write-Host "UserToCopy has multiple values. Please check AD for accounts with duplicate DisplayName attributes. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
} elseif ($NULL -eq $UserToCopyUPN) {
    Write-Output "Could not find user $($UserToCopy) in AD to copy from. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

## Sku availability check
if ($Sku) {
    try {
        $SelectObjectPropertyList = @(
            "SkuPartNumber"
            "SkuId"
            @{
                n = "Available"
                e = { (($_.PrepaidUnits).Enabled - $_.ConsumedUnits) }
            }
        )

        $getLicCount = Get-MgSubscribedSku | Where-Object { ($_.SkuPartNumber -eq $Sku) } | Select-Object $SelectObjectPropertyList

        if ($getLicCount.Available -gt 0) {
            $getLic = $getLicCount
        } else {
            Write-Output "No available license for '$($result.InputSku)'. Please add additional licenses via the Microsoft Portal."
            $Sku = $NULL
        }
    } catch {
        Write-Output "License Sku could not be found. Or no Sku was selected."
        $Sku = $NULL
    }
}

## Check if no Sku was selected or if no Sku was found
if (!$Sku) {

    $SelectObjectPropertyList = @(
        "SkuPartNumber"
        "SkuId"
        @{
            n = "ActiveUnits"
            e = { ($_.PrepaidUnits).Enabled }
        }
        "ConsumedUnits"
    )

    $WhereObjectFilter = {
        ($_.SkuPartNumber -eq 'EXCHANGESTANDARD') -or
        ($_.SkuPartNumber -eq 'O365_BUSINESS_ESSENTIALS') -or
        ($_.SkuPartNumber -eq 'SPE_E3') -or
        ($_.SkuPartNumber -eq 'SPB') -or
        ($_.SkuPartNumber -eq 'ENTERPRISEPACK')
    }

    $selectLicense = Get-MgSubscribedSku | Select-Object $SelectObjectPropertyList | Where-Object -FilterScript $WhereObjectFilter | `
        ForEach-Object {
        [PSCustomObject]@{
            DisplayName   = switch -Regex ($_.SkuPartNumber) {
                "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
                "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
                "SPE_E3" { "Microsoft 365 E3" }
                "SPB" { "Microsoft 365 Business Premium" }
                "ENTERPRISEPACK" { "Office 365 E3" }
            }
            SkuPartNumber = $_.SkuPartNumber
            SkuId         = $_.SkuId
            Available     = ($_.ActiveUnits - $_.ConsumedUnits)
        }
    } | Sort-Object DisplayName

    $GridArguments = @{
        OutputMode = 'Single'
        Title      = 'Please select a license and click OK'
    }

    $selectLicenseTEMP = $selectLicense | ForEach-Object { $_ | Select-Object -Property 'DisplayName', 'Available' } | Out-GridView @GridArguments
    $getLic = $selectLicense | Where-Object { $_.DisplayName -in $selectLicenseTEMP.DisplayName }

}

## Building out new user variables
$Domain = $($UserToCopyUPN.UserPrincipalName -replace '.+?(?=@)')
$NewUserFirstName = $($NewUser.split(' ')[-2])
$NewUserLastName = $($NewUser -replace '.+\s')
$NewUserSamAccountName = $(($NewUser -replace '(?<=.{1}).+') + ($NewUser -replace '.+\s')).ToLower()
$NewUserEmail = $($NewUserSamAccountName + $Domain).ToLower()

$CheckNewUserUPN = $(try { Get-ADUser -Identity $NewUserSamAccountName } catch { $null })
if ($null -ne $CheckNewUserUPN) {
    Write-Host "SamAccountName exist for user $NewUser. Please check AD for accounts with duplicate SamAccountName attributes. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

$CheckNewUserEmail = $(try { Get-MgUser -Identity $NewUserEmail } catch { $null })
if ($null -ne $CheckNewUserEmail) {
    Write-Host "Account in 365 exist for user $NewUser. Please check 365 for accounts with duplicate SMTP Address. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
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
    Write-Output 'User did not enter "Y". Press any key to exit script.'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

try {
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
} catch {
    Write-Host "New User creation was not successful. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

Write-Output 'AD User has been created.'

Write-Output 'Adding AD Groups to new user.'

$CopyFromUser = Get-ADUser -Filter "DisplayName -eq '$($UserToCopy)'" -prop MemberOf
$CopyToUser = Get-ADUser -Filter "DisplayName -eq '$($NewUser)'" -prop MemberOf
$CopyFromUser.MemberOf | Where-Object { $CopyToUser.MemberOf -notcontains $_ } | Add-ADGroupMember -Members $CopyToUser

Write-Output 'Starting AD Sync'

powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta

Write-Output 'Waiting 90 seconds for AD Connect sync process.'

Start-Sleep -Seconds 90

## Check if AD User has synced to Azure loop
$Stoploop = $false
[int]$Retrycount = "0"

do {
    try {
        $NewMgUser = Get-MgUser -UserId $NewUserEmail -ErrorAction Stop
        Write-Output "User $NewUser has synced to Azure. Script will now continue."
        $Stoploop = $true
        $ADSyncCompleteYesorExit = 'yes'
    } catch {
        if ($Retrycount -gt 3) {
            Write-Host "Could not sync AD User to 365 after 3 retries."
            $Stoploop = $true
        } else {
            Write-Host "Could not sync AD User to 365 retrying in 60 seconds..."
            Start-Sleep -Seconds 60
            $Retrycount = $Retrycount + 1
        }
    }
} while ($Stoploop -eq $false)

if (!$NewMgUser) {
    $ADSyncCompleteYesorExit = Read-Host -Prompt 'AD Sync has not completed within allotted time frame. Please wait for AD sync. To resume type yes or exit'
} while ("yes", "exit" -notcontains $ADSyncCompleteYesorExit ) {
    $ADSyncCompleteYesorExit = Read-Host "Please enter your response (yes/exit)"
}

if ($ADSyncCompleteYesorExit -eq 'exit') {
    Write-Output 'You will need to set the license and add Office 365 groups via the portal. Press any key to exit script.'
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit
}

if ($ADSyncCompleteYesorExit -eq 'yes') {

    $NewMgUser = Get-MgUser -UserId $NewUserEmail -ErrorAction Stop
    if (!$NewMgUser) {
        Write-Output 'Script cannot find new user. You will need to set the license and add Office 365 groups via the portal. Press any key to exit script.'
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit
    }

    Write-Output 'Script now will resume'

    Write-Output 'Setting Usage Location for new user'

    ## Assigns US as UsageLocation
    Update-MgUser -UserId $NewUserEmail -UsageLocation US

    ## Assign primary license to new user
    if ($getLic) {
        try {
            Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $getLic.SkuId } -RemoveLicenses @() -ErrorAction stop
            Write-Output 'License added.'
        } catch {
            Write-Output 'License could not be added. You will need to set the license and add Office 365 groups via the portal. Press any key to exit script.'
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            exit
        }
    }

    ## Assign P2 license to new user
    if ($result.InputSkuEntraIDP2 -eq 'Yes') {

        $SelectObjectPropertyList = @(
            "SkuPartNumber"
            "SkuId"
            @{
                n = "ActiveUnits"
                e = { ($_.PrepaidUnits).Enabled }
            }
            "ConsumedUnits"
        )

        $WhereObjectFilter = {
            ($_.SkuPartNumber -like 'AAD_PREMIUM_P2')
        }

        $getLicenseEntraIDP2 = Get-MgSubscribedSku | Select-Object $SelectObjectPropertyList | Where-Object -FilterScript $WhereObjectFilter | `
            ForEach-Object {
            [PSCustomObject]@{
                DisplayName   = switch -Regex ($_.SkuPartNumber) {
                    "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
                }
                SkuPartNumber = $_.SkuPartNumber
                SkuId         = $_.SkuId
                Available     = ($_.ActiveUnits - $_.ConsumedUnits)
            }
        } | Sort-Object DisplayName

        if ($getLicenseEntraIDP2 -ne 0) {
            try {
                Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $getLicenseEntraIDP2.SkuId } -RemoveLicenses @() -ErrorAction stop
                Write-Output "$($_.SkuPartNumber) License added."
            } catch {
                Write-Output "$($_.SkuPartNumber) License could not be added."
            }
        } else {
            Write-Output "$($_.SkuPartNumber) no avaible license. Please add license and add manually"
        }
    }

    Write-Output 'Adding Office 365 Groups to new user.'

    ## Copy groups to new user from old user
    $All365Groups = Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id | `
        Where-Object { $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and $_.AdditionalProperties.membershipRule -eq $NULL -and $_.onPremisesSyncEnabled -ne 'False' } | `
        Select-Object Id, @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }, @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }

    Foreach ($365Group in $All365Groups) {
        try {
            New-MgGroupMember -GroupId $365Group.Id -DirectoryObjectId $(Get-MgUser -UserId $NewUserEmail).Id -ErrorAction Stop
        } catch {
            Add-DistributionGroupMember -Identity $365Group.DisplayName -Member $NewUserEmail -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction 'SilentlyContinue'
        }
    }

    $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName).Id).Count
    $NewUserGroupCount = (Get-MgUserMemberOf -UserId $(Get-MgUser -UserId $NewUserEmail).Id).Count

    Write-Output "User $($NewUser) should now be created unless any errors occurred during the process."
    Write-Output "Copy User group count: $($CopyUserGroupCount)"
    Write-Output "New User group count: $($NewUserGroupCount)"

    ## Add BookWithMeId to the extensionAttribute15 property of the new user.
    $NewUserExchGuid = (Get-Mailbox -Identity $NewUserEmail).ExchangeGuid.Guid -replace "-" -replace ""
    $extAttr15 = $NewUserExchGuid + '@compassmsp.com?anonymous&ep=plink'

    Set-ADUser -Identity $NewUserSamAccountName -Add @{extensionAttribute15 = "$extAttr15" }

    ## Sends email to SecurePath Team (soc@compassmsp.com) with the new user information.
    $MgUser = Get-MgUser -UserId $NewUserEmail

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

    ## Add additional 365 licenses

    $AddLic = Read-Host "Would you like to add additional licenses? (Y/N)"

    if ($AddLic -ne 'y') { Write-Output 'You have selected not to add any additional licenses.' }

    if ($AddLic -eq 'y') {

        $SelectObjectPropertyList = @(
            "SkuPartNumber"
            "SkuId"
            @{
                n = "ActiveUnits"
                e = { ($_.PrepaidUnits).Enabled }
            }
            "ConsumedUnits"
        )

        $WhereObjectFilter = {
            ($_.SkuPartNumber -notlike 'EXCHANGESTANDARD') -and
            ($_.SkuPartNumber -notlike 'O365_BUSINESS_ESSENTIALS') -and
            ($_.SkuPartNumber -notlike 'SPE_E3') -and
            ($_.SkuPartNumber -notlike 'SPB') -and
            ($_.SkuPartNumber -notlike 'ENTERPRISEPACK') -and
            ($_.SkuPartNumber -notlike 'PROJECT_MADEIRA_PREVIEW_IW_SKU') -and
            ($_.SkuPartNumber -notlike 'POWERAUTOMATE_ATTENDED_RPA') -and
            ($_.SkuPartNumber -notlike 'RMSBASIC') -and
            ($_.SkuPartNumber -notlike 'MCOPSTNC') -and
            ($_.SkuPartNumber -notlike 'CCIBOTS_PRIVPREV_VIRAL') -and
            ($_.SkuPartNumber -notlike 'MCOPSTN1') -and
            ($_.SkuPartNumber -notlike 'WINDOWS_STORE') -and
            ($_.SkuPartNumber -notlike 'STREAM') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_DEV') -and
            ($_.SkuPartNumber -notlike 'RIGHTSMANAGEMENT_ADHOC') -and
            ($_.SkuPartNumber -notlike 'MCOMEETADV') -and
            ($_.SkuPartNumber -notlike 'MEETING_ROOM') -and
            ($_.SkuPartNumber -notlike 'VISIO_PLAN1_DEPT') -and
            ($_.SkuPartNumber -notlike 'FLOW_FREE') -and
            ($_.SkuPartNumber -notlike 'MICROSOFT_BUSINESS_CENTER') -and
            ($_.SkuPartNumber -notlike 'PHONESYSTEM_VIRTUALUSER') -and
            ($_.SkuPartNumber -notlike 'Microsoft_Copilot_for_Finance_trial') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_VIRAL') -and
            ($_.SkuPartNumber -notlike 'Microsoft_Teams_Exploratory_Dept') -and
            ($_.SkuPartNumber -notlike 'POWERAPPS_PER_USER') -and
            ($_.SkuPartNumber -notlike 'Power BI Standard') -and
            ($_.SkuPartNumber -notlike 'AAD_PREMIUM_P2')
        }

        $selectLicense2 = Get-MgSubscribedSku | Select-Object $SelectObjectPropertyList | Where-Object -FilterScript $WhereObjectFilter | `
            ForEach-Object {
            [PSCustomObject]@{
                DisplayName   = switch -Regex ($_.SkuPartNumber) {
                    "PROJECT_P1" { "Project Plan 1" }
                    "PROJECTPROFESSIONAL" { "Project Plan 3" }
                    "VISIOCLIENT" { "Visio Plan 2" }
                    "Microsoft_Teams_Audio_Conferencing_select_dial_out" { "Microsoft Teams Audio Conferencing with dial-out to USA/CAN" }
                    "POWER_BI_PRO" { "Power BI Pro" }
                    "Microsoft_365_Copilot" { "Microsoft 365 Copilot" }
                    "Microsoft_Teams_Premium" { "Microsoft Teams Premium" }
                    "MCOEV" { "Microsoft Teams Phone Standard" }
                    "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
                    "POWER_BI_STANDARD" { "Power BI Standard" }
                    "Microsoft365_Lighthouse" { "Microsoft 365 Lighthouse" }
                }
                SkuPartNumber = $_.SkuPartNumber
                SkuId         = $_.SkuId
                Available     = ($_.ActiveUnits - $_.ConsumedUnits)
            }
        } | Sort-Object DisplayName

        $GridArguments = @{
            OutputMode = 'Multiple'
            Title      = 'Please select licenses and click OK (Hold CTRL to select multiple licenses)'
        }

        $selectLicenseTEMP2 = $selectLicense2 | ForEach-Object { $_ | Select-Object -Property 'DisplayName', 'Available' } | Out-GridView @GridArguments
        $getLic2 = $selectLicense2 | Where-Object { $_.DisplayName -in $selectLicenseTEMP2.DisplayName }
    }

    if ($GetLic2) {
        $GetLic2 | ForEach-Object {
            try {
                Set-MgUserLicense -UserId $NewMgUser.Id -AddLicenses @{SkuId = $_.SkuId } -RemoveLicenses @() -ErrorAction stop
                Write-Output "$($_.SkuPartNumber) License added."
            } catch {
                Write-Output "$($_.SkuPartNumber) License could not be added."
            }
        }
    }

    #Disconnect from Exchange and Graph
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-Graph

    ## Connect to PnP PowerShell
    <#
    $PnPAppId = "24e3c6ad-9658-4a0d-b85f-82d67d148449"
    $Org = "compassmsp.onmicrosoft.com"
    $PnPCert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { ($_.Subject -like '*CN=PnP PowerShell*') -and ($_.NotAfter -gt $([DateTime]::Now)) }
    if ($NULL -eq $PnPCert) {
        Write-Host "No valid PnP PowerShell certificates found in the LocalMachine\My store. Press any key to exit script." -ForegroundColor Red -BackgroundColor Black
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        exit
    }
    Connect-PnPOnline -Url compassmsp-admin.sharepoint.com -ClientId $PnPAppId -Tenant $Org -Thumbprint $($PNPCert.Thumbprint)
    #>
    Connect-PnPOnline -Url https://compassmsp-admin.sharepoint.com -UseWebLogin

    ## Creates OneDrive
    Request-PnPPersonalSite -UserEmails $NewUserEmail -NoWait
    Disconnect-PnPOnline

}