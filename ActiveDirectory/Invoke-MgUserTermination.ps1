#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#
.SYNOPSIS
    Handles Office 365/AD tasks during user termination.

.DESCRIPTION
    This script automates the termination process by handling both Active Directory
    and Microsoft 365 tasks including group removal, license removal, and mailbox management.

    The script will display a GUI window to collect:
    - User to terminate (email)
    - Mailbox access delegation
    - Email forwarding settings
    - OneDrive access delegation
    - OneDrive read-only setting

    IMPORTANT: This script must be run from the Primary Domain Controller with AD Connect installed.

    NOTE: Sensitive information (app IDs, certificates, etc.) is stored in a secure configuration file managed by Get-ScriptConfig.
    The config file should be placed at: C:\ProgramData\CompassScripts\config.json

.EXAMPLE
    .\Invoke-MgUserTermination.ps1

    This will launch the GUI window to collect the required information.

.NOTES
    Author: Chris Williams
    Created: 2021-12-20
    Last Modified: 2025-01-20

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    3.0.0        2025-01-20  Major Rework:
                          - Complete script reorganization and optimization
                          - Optimized UI spacing and element alignment
                          - Enhanced form layout for improved readability
                          - Added secure configuration management via Get-ScriptConfig
                          - Enhanced error handling and logging system
                          - Added progress tracking and status messaging
                          - Added Zoom phone onboarding

    2.1.0        2024-11-25  Feature Update:
                          - Reworked GUI interface
                          - Added QuickEdit and InsertMode management
                          - Removed KnowBe4 SCIM integration per SecurePath Team
                          - Added Email Forwarding functionality - KnowBe4 Notification

    2.0.0        2024-07-15  Major Feature Update:
                          - Added GUI input system
                          - Enhanced UI for variable collection
                          - Added KB4 offboarding integration
                          - Added OneDrive read-only functionality
                          - Updated KnowBe4 SCIM integration
                          - Added directory role management

    1.2.0        2023-02-12  Feature Updates:
                          - Enhanced license management
                          - Improved group handling
                          - Added KnowBe4 integration
                          - Enhanced group function cleanup
                          - Added OneDrive access management

    1.1.0        2022-06-27  Enhancement Update:
                          - Added group and license exports
                          - Improved user management functions
                          - Enhanced manager removal process
                          - Fixed group member removal
                          - Added sign-in revocation

    1.0.0        2021-12-20  Initial Release:
                          - Basic termination functionality
                          - AD user management
                          - Group removal
                          - License removal
    ------------------------------------------------------------------------------
#>

# Initialize loading animation
Clear-Host
$loadingChars = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
$i = 0
$loadingJob = Start-Job -ScriptBlock { while ($true) { Start-Sleep -Milliseconds 100 } }

try {
    Write-Host "`n  Initializing User Termination Script..." -ForegroundColor Cyan

    Write-Host "  [$($loadingChars[$i % $loadingChars.Length])] Loading core components..." -NoNewline -ForegroundColor Yellow
    $ErrorActionPreference = 'Stop'
    # Only show verbose output if -Verbose is specified
    if (-not $PSBoundParameters['Verbose']) {
        $VerbosePreference = 'SilentlyContinue'
    }
    $startTime = Get-Date
    Write-Host "`r  [✓] Core components loaded" -ForegroundColor Green

    Write-Host "  [$($loadingChars[$i % $loadingChars.Length])] Initializing progress tracking..." -NoNewline -ForegroundColor Yellow
    $progressSteps = @(
        @{ Number = 0; Name = "Initialization"; Description = "Loading configuration and connecting services" }
        @{ Number = 1; Name = "User Input"; Description = "Gathering termination details" }
        @{ Number = 2; Name = "AD Tasks"; Description = "Disabling user in Active Directory" }
        @{ Number = 3; Name = "Session Cleanup"; Description = "Removing user sessions and devices" }
        @{ Number = 4; Name = "Mailbox Setup"; Description = "Converting to shared mailbox" }
        @{ Number = 5; Name = "Directory Roles"; Description = "Removing from directory roles" }
        @{ Number = 6; Name = "Group Removal"; Description = "Removing from groups" }
        @{ Number = 7; Name = "License Removal"; Description = "Removing licenses" }
        @{ Number = 8; Name = "Notifications"; Description = "Sending notifications" }
        @{ Number = 9; Name = "Zoom Removal"; Description = "Removing from Zoom" }
        @{ Number = 10; Name = "OneDrive Setup"; Description = "Configuring OneDrive access" }
        @{ Number = 11; Name = "Final Steps"; Description = "Running AD sync and finalizing" }
    )
    Write-Host "`r  [✓] Progress tracking initialized" -ForegroundColor Green

    Write-Host "  [$($loadingChars[$i % $loadingChars.Length])] Loading functions..." -NoNewline -ForegroundColor Yellow
    $script:errorCount = 0
    $script:totalSteps = $progressSteps.Count  # Make it script-scoped and move it before the function

    function Write-ProgressStep {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$StepName,

            [Parameter(Mandatory)]
            [string]$Status
        )

        # Get the step number from the progress steps array
        $stepNumber = ($progressSteps | Where-Object { $_.Name -eq $StepName }).Number

        # Guard against division by zero or missing step number
        if ($null -eq $stepNumber -or $script:totalSteps -eq 0) {
            Write-StatusMessage -Message "Step $StepName - $Status" -Type INFO
            Write-Progress -Activity "User Termination" -Status $Status
        } else {
            Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $Status" -Type INFO
            Write-Progress -Activity "User Termination" -Status $Status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
        }
    }

    function Set-ConsoleProperties {
        [CmdletBinding()]
        param (
            [Parameter()]
            [ValidateSet('Enable', 'Disable')]
            [string]$QuickEditMode = 'Enable',

            [Parameter()]
            [ValidateSet('Enable', 'Disable')]
            [string]$InsertMode = 'Enable'
        )

        $signature = @'
        using System;
        using System.Runtime.InteropServices;

        public static class ConsoleMode {
            private const uint ENABLE_QUICK_EDIT = 0x0040;
            private const uint ENABLE_INSERT_MODE = 0x0020;
            private const int STD_INPUT_HANDLE = -10;

            [DllImport("kernel32.dll", SetLastError = true)]
            private static extern IntPtr GetStdHandle(int nStdHandle);

            [DllImport("kernel32.dll")]
            private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

            [DllImport("kernel32.dll")]
            private static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

            public static void SetMode(bool enableQuickEdit, bool enableInsert) {
                IntPtr handle = GetStdHandle(STD_INPUT_HANDLE);
                uint mode;

                if (!GetConsoleMode(handle, out mode)) {
                    throw new Exception("Failed to get console mode");
                }

                mode = enableQuickEdit ? mode | ENABLE_QUICK_EDIT : mode & ~ENABLE_QUICK_EDIT;
                mode = enableInsert ? mode | ENABLE_INSERT_MODE : mode & ~ENABLE_INSERT_MODE;

                if (!SetConsoleMode(handle, mode)) {
                    throw new Exception("Failed to set console mode");
                }
            }
        }
'@

        try {
            # Add the type if it doesn't exist
            if (-not ('ConsoleMode' -as [type])) {
                Add-Type -TypeDefinition $signature -Language CSharp
            }

            # Convert parameters to boolean values
            $quickEdit = $QuickEditMode -eq 'Enable'
            $insert = $InsertMode -eq 'Enable'

            # Set the console modes
            [ConsoleMode]::SetMode($quickEdit, $insert)

            Write-Verbose "Console properties updated successfully: QuickEdit=$QuickEditMode, Insert=$InsertMode"
        } catch {
            Write-Error "Failed to set console properties: $($_.Exception.Message)"
        }
    }

    # Disable console quick edit
    Set-ConsoleProperties -QuickEditMode Disable -InsertMode Disable

    function Write-StatusMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet('INFO', 'OK', 'SUCCESS', 'ERROR', 'WARN', 'SUMMARY')]
            [string]$Type = 'INFO'
        )

        $config = @{
            'INFO'    = @{ Status = 'INFO'; Color = 'White' }
            'OK'      = @{ Status = 'OK'; Color = 'Green' }
            'SUCCESS' = @{ Status = 'SUCCESS'; Color = 'Green' }
            'ERROR'   = @{ Status = 'ERROR'; Color = 'Red' }
            'WARN'    = @{ Status = 'WARN'; Color = 'Yellow' }
            'SUMMARY' = @{ Status = ''; Color = 'Cyan' }
        }

        if ($Type -eq 'SUMMARY') {
            Write-Host $Message -ForegroundColor $config[$Type].Color
        } else {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $statusPadded = $config[$Type].Status.PadRight(7)
            Write-Host "[$timestamp] [$statusPadded] $Message" -ForegroundColor $config[$Type].Color
        }

        if ($Type -eq 'ERROR') {
            $script:errorCount++
            if ($Message -match 'config') { $script:errorTypes.Configuration++ }
            # ... etc
        }
    }

    function Exit-Script {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Message,

            [Parameter(Mandatory = $false)]
            [ValidateSet(
                'Success',
                'Cancelled',
                'ConfigError',
                'ConnectionError',
                'UserNotFound',
                'PermissionError',
                'DuplicateUser',
                'GeneralError'
            )]
            [string]$ExitCode = 'GeneralError'
        )

        try {
            # Map exit codes to error types for tracking
            $errorTypeMap = @{
                'ConfigError'     = 'Configuration'
                'ConnectionError' = 'Connection'
                'PermissionError' = 'Permission'
                'UserNotFound'    = 'Validation'
                'GeneralError'    = 'General'
            }

            # Map exit codes to numeric values
            $exitCodes = @{
                'Success'         = 0
                'Cancelled'       = 1
                'ConfigError'     = 2
                'ConnectionError' = 3
                'UserNotFound'    = 4
                'PermissionError' = 5
                'DuplicateUser'   = 6
                'GeneralError'    = 99
            }

            # Map exit codes to message types
            $messageTypes = @{
                'Success'         = 'OK'
                'Cancelled'       = 'WARN'
                'ConfigError'     = 'ERROR'
                'ConnectionError' = 'ERROR'
                'UserNotFound'    = 'ERROR'
                'PermissionError' = 'ERROR'
                'DuplicateUser'   = 'ERROR'
                'GeneralError'    = 'ERROR'
            }

            # Track error type if it's not a success or cancellation
            if ($ExitCode -notin @('Success', 'Cancelled') -and $errorTypeMap.ContainsKey($ExitCode)) {
                Add-ErrorType -ErrorType $errorTypeMap[$ExitCode]
            }

            # Attempt to disconnect from services
            Write-StatusMessage -Message "Disconnecting from services..." -Type INFO
            try {
                Connect-ServiceEndpoints -Disconnect
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disconnect services during exit" -ErrorLevel Warning
                Add-ErrorType -ErrorType Connection
            }

            # Display final error summary if there were errors
            if ($script:errorCount -gt 0) {
                Write-StatusMessage -Message (Get-ErrorSummary) -Type SUMMARY
            }

            # Log the exit message
            Write-StatusMessage -Message $Message -Type $messageTypes[$ExitCode]
            Write-Log -Message "Script exiting with code $($exitCodes[$ExitCode]): $Message" -Level $messageTypes[$ExitCode]

            # Return the appropriate exit code
            exit $exitCodes[$ExitCode]
        } catch {
            # Catch-all for any unexpected errors during exit
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error during script exit"
            Add-ErrorType -ErrorType General
            exit 99
        }
    }

    function Get-ScriptConfig {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory = $false)]
            [string]$ConfigPath = "C:\ProgramData\CompassScripts\config.json"
        )

        # Add comment-based help here
        <#
        .SYNOPSIS
            Gets or creates configuration for Compass scripts.
        .DESCRIPTION
            Loads configuration from JSON file or creates new config with user prompts.
        .PARAMETER ConfigPath
            Path to the configuration file. Defaults to C:\ProgramData\CompassScripts\config.json
        .EXAMPLE
            $config = Get-ScriptConfig
            Loads or creates default configuration
        #>

        try {
            # Check for local config first
            Write-StatusMessage -Message "Checking for local configuration file..." -Type INFO

            if (Test-Path $ConfigPath) {
                $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
                Write-StatusMessage -Message "Loaded configuration from $ConfigPath" -Type OK
                return $config
            }

            # If no config exists, create template with prompt
            Write-StatusMessage "No configuration file found. Creating template at $ConfigPath"

            # Ensure directory exists
            $configDir = Split-Path $ConfigPath -Parent
            if (-not (Test-Path $configDir)) {
                New-Item -Path $configDir -ItemType Directory -Force | Out-Null
            }

            # Prompt for required values
            $config = @{
                ExchangeOnline = @{
                    AppId              = Read-Host "Enter Exchange Online AppId"
                    Organization       = Read-Host "Enter Organization (e.g., company.onmicrosoft.com)"
                    CertificateSubject = Read-Host "Enter Exchange Online certificate subject (e.g., CN=ExO PowerShell)"
                }
                Graph          = @{
                    AppId              = Read-Host "Enter Graph AppId"
                    TenantId           = Read-Host "Enter TenantId"
                    CertificateSubject = Read-Host "Enter Graph certificate subject (e.g., CN=Graph PowerShell)"
                }
                PnPSharePoint  = @{
                    AppId              = Read-Host "Enter PnP SharePoint AppId"
                    Url                = Read-Host "Enter SharePoint Online URL"
                    CertificateSubject = Read-Host "Enter PnP certificate subject (e.g., CN=PnP PowerShell)"
                }
                Paths          = @{
                    NewUserLogPath = "C:\Temp\NewUserCreation.log"
                    LogPath        = "C:\Temp\UserTermination.log"
                    ExportPath     = "C:\Temp\terminated_users_exports"
                }
                Email          = @{
                    NotificationFrom  = Read-Host "Enter notification from address"
                    SecurityTeamEmail = Read-Host "Enter security team email"
                }
            }

            # Save config
            $config | ConvertTo-Json | Set-Content $ConfigPath

            return $config
        } catch {
            Write-StatusMessage -Message "Critical error in configuration: $_" -Type ERROR
            throw
        }
    }

    function Write-Log {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Message,
            [Parameter()]
            [ValidateSet('INFO', 'OK', 'SUCCESS', 'ERROR', 'WARN', 'SUMMARY')]
            [string]$Level = 'INFO',
            [Parameter()]
            [string]$LogPath = $config.Paths.NewUserLogPath
        )

        try {
            # Create log directory if it doesn't exist
            $logDir = Split-Path -Path $LogPath -Parent
            if (-not (Test-Path -Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"

            # Write to log file
            Add-Content -Path $LogPath -Value $logMessage

            # Also write to status message
            Write-StatusMessage -Message $Message -Type $Level

            # Track errors
            if ($Level -eq 'ERROR') {
                $script:errorCount++
            }
        } catch {
            Write-StatusMessage -Message "Failed to write to log: $_" -Type ERROR
        }
    }

    # Function to create and show a custom WPF window for user termination
    function Show-CustomTerminationWindow {
        <#
    .SYNOPSIS
    Shows a GUI window for processing a user termination request.

    .DESCRIPTION
    Displays a WPF window that collects information needed to terminate a user,
    including delegation of access rights and mailbox settings.

    .OUTPUTS
    [PSCustomObject] Returns a custom object with the following properties:
        InputUser               : [string] Email address of the user to terminate
        InputUserFullControl    : [string] Email of user to receive full mailbox control (empty if not specified)
        InputUserFWD           : [string] Email address for mail forwarding (empty if not specified)
        InputUserOneDriveAccess: [string] Email of user to receive OneDrive access (empty if not specified)
        SetOneDriveReadOnly    : [bool] Whether to set OneDrive as read-only
    Returns $null if the user cancels the operation.
    #>

        # 1. Add required assemblies
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        # 2. UI Assembly Helper Functions
        function New-ScrollingStackPanel {
            param (
                [int]$MaxHeight = 0,
                [string]$Margin = "5"
            )
            $scrollViewer = New-FormScrollViewer -MaxHeight $MaxHeight -Margin $Margin
            $stackPanel = New-Object System.Windows.Controls.StackPanel
            $scrollViewer.Content = $stackPanel
            return @{
                ScrollViewer = $scrollViewer
                StackPanel   = $stackPanel
            }
        }

        function New-FormScrollViewer {
            param (
                [int]$MaxHeight = 0,
                [string]$Margin = "5"
            )
            $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
            $scrollViewer.VerticalScrollBarVisibility = "Auto"
            if ($MaxHeight -gt 0) {
                $scrollViewer.MaxHeight = $MaxHeight
            }
            $scrollViewer.Margin = $Margin
            return $scrollViewer
        }

        function New-FormWindow {
            param (
                [string]$Title,
                [int]$Width = 500,
                [int]$Height,
                [string]$Background = '#F0F0F0'
            )
            $window = New-Object System.Windows.Window
            $window.Title = $Title
            $window.Width = $Width
            $window.Height = $Height
            $window.WindowStartupLocation = 'CenterScreen'
            $window.Background = $Background
            return $window
        }

        function New-ButtonPanel {
            param (
                [string]$Margin = '0,10,0,0'
            )
            $buttonPanel = New-Object System.Windows.Controls.StackPanel
            $buttonPanel.Orientation = 'Horizontal'
            $buttonPanel.HorizontalAlignment = 'Right'
            $buttonPanel.Margin = $Margin
            return $buttonPanel
        }

        function New-FormDockPanel {
            param (
                [string]$Margin = '0,0,0,5'
            )
            $dockPanel = New-Object System.Windows.Controls.DockPanel
            $dockPanel.Margin = $Margin
            return $dockPanel
        }

        function New-MainPanel {
            param (
                [string]$Margin = '10'
            )
            $mainPanel = New-Object System.Windows.Controls.StackPanel
            $mainPanel.Margin = $Margin
            return $mainPanel
        }

        function New-HeaderPanel {
            param ([string]$Text)
            $headerPanel = New-Object System.Windows.Controls.Border
            $headerPanel.Background = '#E1E1E1'
            $headerPanel.Padding = '10'
            $headerPanel.Margin = '0,0,0,15'
            $headerPanel.BorderBrush = '#CCCCCC'
            $headerPanel.BorderThickness = '1'

            $headerText = New-Object System.Windows.Controls.TextBlock
            $headerText.Text = $Text
            $headerText.TextWrapping = 'Wrap'
            $headerPanel.Child = $headerText

            return $headerPanel
        }

        function New-FormButton {
            param (
                [string]$Content,
                [scriptblock]$ClickHandler,
                [string]$Margin = '0,0,0,0'
            )
            $button = New-Object System.Windows.Controls.Button
            $button.Content = $Content
            $button.Width = 100
            $button.Height = 30
            $button.Margin = $Margin
            $button.Add_Click($ClickHandler)
            return $button
        }

        function New-FormLabel {
            param ([string]$Content)
            $label = New-Object System.Windows.Controls.Label
            $label.Content = $Content
            return $label
        }

        function New-FormGroupBox {
            param (
                [string]$Header,
                [string]$Margin = '0,0,0,10'
            )
            $group = New-Object System.Windows.Controls.GroupBox
            $group.Header = $Header
            $group.Margin = $Margin

            $stack = New-Object System.Windows.Controls.StackPanel
            $stack.Margin = '5'
            $group.Content = $stack

            return @{
                Group = $group
                Stack = $stack
            }
        }

        function New-FormCheckBox {
            param (
                [string]$Content,
                [string]$ToolTip,
                [string]$Margin = "5,5,5,5",
                [bool]$IsChecked = $false
            )
            $checkbox = New-Object System.Windows.Controls.CheckBox
            $checkbox.Content = $Content
            $checkbox.ToolTip = $ToolTip
            $checkbox.Margin = $Margin
            $checkbox.IsChecked = $IsChecked
            return $checkbox
        }

        # 3. Validation Helper Functions
        function Test-EmailAddress {
            param ([string]$Email)
            return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        }

        function Show-ValidationError {
            param (
                [string]$Message,
                [string]$Title = "Input Error"
            )
            [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }

        # 4. Event Handlers
        $Script:emailGotFocusHandler = {
            if ($this.Text -eq $this.Tag) {
                $this.Text = ""
                $this.Foreground = 'Black'
            }
        }

        $Script:emailLostFocusHandler = {
            if ([string]::IsNullOrWhiteSpace($this.Text) -or $this.Text -eq $this.Tag) {
                $this.Text = $this.Tag
                $this.Foreground = 'Gray'
                $this.BorderBrush = $null
                $this.BorderThickness = 1
                return
            }
            if (-not (Test-EmailAddress -Email $this.Text)) {
                $this.BorderBrush = 'Red'
                $this.BorderThickness = 2
            } else {
                $this.BorderBrush = $null
                $this.BorderThickness = 1
            }
        }

        # 5. Input Control Initialization
        function Initialize-EmailTextBox {
            param (
                [string]$PlaceholderText,
                [string]$ToolTipText,
                [string]$Name,
                [string]$Margin = '0,0,0,10'
            )

            $textBox = New-Object System.Windows.Controls.TextBox
            $textBox.Name = $Name
            $textBox.Margin = $Margin
            $textBox.Padding = '5,3,5,3'
            $textBox.Tag = $PlaceholderText
            $textBox.Text = $PlaceholderText
            $textBox.Foreground = 'Gray'
            $textBox.ToolTip = $ToolTipText

            $textBox.Add_GotFocus($Script:emailGotFocusHandler)
            $textBox.Add_LostFocus($Script:emailLostFocusHandler)

            return $textBox
        }

        # 6. Main UI Creation and Logic
        # Create window and main containers
        $window = New-FormWindow -Title "User Termination Request" -Height 530
        $scrollViewer = New-FormScrollViewer
        $mainPanel = New-MainPanel -Margin '10'
        $scrollViewer.Content = $mainPanel
        $window.Content = $scrollViewer

        # Add header
        $mainPanel.Children.Add((New-HeaderPanel -Text "User Termination Request`nPlease fill in all required fields marked with *"))

        # Create user termination group
        $termSection = New-FormGroupBox -Header "User Information"
        $termSection.Stack.Children.Add((New-FormLabel -Content "User to Terminate (Email) *"))
        $txtUserToTerm = Initialize-EmailTextBox `
            -Name "userToTerm" `
            -PlaceholderText "Enter user's email address" `
            -ToolTipText "Enter the email address of the user to be terminated"
        $termSection.Stack.Children.Add($txtUserToTerm)
        $mainPanel.Children.Add($termSection.Group)

        # Create access delegation group
        $delegateSection = New-FormGroupBox -Header "Access Delegation"

        # OneDrive Access
        $delegateSection.Stack.Children.Add((New-FormLabel -Content "Grant OneDrive Access To (Email):"))
        $txtOneDriveAccess = Initialize-EmailTextBox `
            -Name "oneDriveAccess" `
            -PlaceholderText "Enter delegate's email address" `
            -ToolTipText "Enter the email of the person who should receive OneDrive access"
        $delegateSection.Stack.Children.Add($txtOneDriveAccess)

        # Mailbox Control
        $delegateSection.Stack.Children.Add((New-FormLabel -Content "Grant Mailbox Full Control To (Email):"))
        $txtMailboxControl = Initialize-EmailTextBox `
            -Name "mailboxControl" `
            -PlaceholderText "Enter delegate's email address" `
            -ToolTipText "Enter the email of the person who should receive mailbox access"
        $delegateSection.Stack.Children.Add($txtMailboxControl)

        # Forward Mailbox
        $delegateSection.Stack.Children.Add((New-FormLabel -Content "Forward Mailbox To (Email):"))
        $txtForwardMailbox = Initialize-EmailTextBox `
            -Name "forwardMailbox" `
            -PlaceholderText "Enter forward-to email address" `
            -ToolTipText "Enter the email address where future emails should be forwarded"
        $delegateSection.Stack.Children.Add($txtForwardMailbox)

        # OneDrive Read-Only option
        $oneDrivePanel = New-FormDockPanel -Margin "0,0,0,5"

        $chkOneDriveReadOnly = New-FormCheckBox `
            -Content "Set OneDrive as Read-Only" `
            -ToolTip "Check to make the OneDrive content read-only" `
            -Margin "10,0,0,0"

        $oneDrivePanel.Children.Add($chkOneDriveReadOnly)
        $delegateSection.Stack.Children.Add($oneDrivePanel)

        $mainPanel.Children.Add($delegateSection.Group)

        # Create buttons
        $buttonPanel = New-ButtonPanel -Margin "0,10,0,0"

        $okButton = New-FormButton -Content "OK" -Margin "0,0,10,0" -ClickHandler {
            # Validate required fields first
            if (-not $txtUserToTerm.Text -or $txtUserToTerm.Text -eq $txtUserToTerm.Tag -or -not (Test-EmailAddress -Email $txtUserToTerm.Text)) {
                Show-ValidationError -Message "Invalid or missing email for required field: User to Terminate"
                return
            }

            # Validate optional fields if they have content
            $optionalTextBoxes = @{
                'OneDrive Access' = $txtOneDriveAccess
                'Mailbox Control' = $txtMailboxControl
                'Forward Mailbox' = $txtForwardMailbox
            }

            foreach ($field in $optionalTextBoxes.GetEnumerator()) {
                if ($field.Value.Text -ne $field.Value.Tag -and -not (Test-EmailAddress -Email $field.Value.Text)) {
                    Show-ValidationError -Message "Invalid email format for: $($field.Key)"
                    return
                }
            }

            $window.DialogResult = $true
            $window.Close()
        }
        $buttonPanel.Children.Add($okButton)

        $cancelButton = New-FormButton -Content "Cancel" -ClickHandler {
            $window.DialogResult = $false
            $window.Close()
        }
        $buttonPanel.Children.Add($cancelButton)

        $mainPanel.Children.Add($buttonPanel)

        # Show the window and return results
        $result = $window.ShowDialog()

        if ($result -eq $true) {
            return @{
                InputUser               = $txtUserToTerm.Text
                InputUserFullControl    = if ($txtMailboxControl.Text -eq $txtMailboxControl.Tag) { "" } else { $txtMailboxControl.Text }
                InputUserFWD            = if ($txtForwardMailbox.Text -eq $txtForwardMailbox.Tag) { "" } else { $txtForwardMailbox.Text }
                InputUserOneDriveAccess = if ($txtOneDriveAccess.Text -eq $txtOneDriveAccess.Tag) { "" } else { $txtOneDriveAccess.Text }
                SetOneDriveReadOnly     = $chkOneDriveReadOnly.IsChecked
            }
        } else {
            return $null
        }
    }

    function Connect-ServiceEndpoints {
        <#
    .SYNOPSIS
        Manages connections to Microsoft 365 service endpoints.

    .DESCRIPTION
        Handles both connection and disconnection to Exchange Online, Microsoft Graph,
        and SharePoint Online services. Can connect/disconnect to all services or
        specific services as needed.

    .PARAMETER ExchangeOnline
        Switch to specify Exchange Online service operations.

    .PARAMETER Graph
        Switch to specify Microsoft Graph service operations.

    .PARAMETER SharePoint
        Switch to specify SharePoint Online service operations.

    .PARAMETER Disconnect
        Switch to disconnect instead of connect. If used without other switches,
        disconnects from all services.

    .EXAMPLE
        Connect-ServiceEndpoints
        Connects to all services using default configuration.

    .EXAMPLE
        Connect-ServiceEndpoints -ExchangeOnline -Graph
        Connects only to Exchange Online and Microsoft Graph services.

    .EXAMPLE
        Connect-ServiceEndpoints -Disconnect
        Disconnects from all connected services.

    .EXAMPLE
        Connect-ServiceEndpoints -Disconnect -SharePoint
        Disconnects only from SharePoint Online.

    .EXAMPLE
        Connect-ServiceEndpoints -ExchangeOnline
        Connects only to Exchange Online service.

    .NOTES
        Requires appropriate certificates and permissions configured in config.json.
        Uses global configuration variables for connection parameters.
    #>

        [CmdletBinding()]
        param(
            [Parameter()]
            [switch]$ExchangeOnline,

            [Parameter()]
            [switch]$Graph,

            [Parameter()]
            [switch]$SharePoint,

            [Parameter()]
            [switch]$Disconnect
        )

        # Validate parameters for requested services
        if ($ExchangeOnline -or (-not ($ExchangeOnline -or $Graph -or $SharePoint))) {
            $requiredExOParams = @('ExOAppId', 'Organization', 'ExOCertSubject')
            $missingExOParams = $requiredExOParams.Where({ -not (Get-Variable -Name $_ -ErrorAction SilentlyContinue) })

            if ($missingExOParams) {
                throw "Exchange Online connection requires the following parameters: $($missingExOParams -join ', ')"
            }

            Write-StatusMessage -Message "Connecting to Exchange Online..." -Type 'INFO'
            $ExOCert = Get-ChildItem Cert:\LocalMachine\My |
            Where-Object { ($_.Subject -like "*$($ExOCertSubject)*") -and ($_.NotAfter -gt $([DateTime]::Now)) }

            if ($null -eq $ExOCert) {
                Exit-Script -Message "No valid ExO PowerShell certificates found in the LocalMachine\My store" -ExitCode ConfigError
            }

            Connect-ExchangeOnline -AppId $ExOAppId -Organization $Organization -CertificateThumbprint $($ExOCert.Thumbprint) -ShowBanner:$false
            Write-StatusMessage -Message "Connected to Exchange Online" -Type 'OK'
        }

        # Validate and connect to Microsoft Graph if requested
        if ($Graph -or (-not ($ExchangeOnline -or $Graph -or $SharePoint))) {
            $requiredGraphParams = @('GraphAppId', 'TenantId', 'GraphCertSubject')
            $missingGraphParams = $requiredGraphParams.Where({ -not (Get-Variable -Name $_ -ErrorAction SilentlyContinue) })

            if ($missingGraphParams) {
                throw "Graph connection requires the following parameters: $($missingGraphParams -join ', ')"
            }

            Write-StatusMessage -Message "Connecting to Microsoft Graph..." -Type 'INFO'
            $GraphCert = Get-ChildItem Cert:\LocalMachine\My |
            Where-Object { ($_.Subject -like "*$($GraphCertSubject)*") -and ($_.NotAfter -gt $([DateTime]::Now)) }

            if ($null -eq $GraphCert) {
                Exit-Script -Message "No valid Graph PowerShell certificates found in the LocalMachine\My store" -ExitCode ConfigError
            }

            Connect-Graph -TenantId $TenantId -AppId $GraphAppId -Certificate $GraphCert -NoWelcome
            Write-StatusMessage -Message "Connected to Microsoft Graph" -Type 'OK'
        }

        # Validate and connect to SharePoint Online if requested
        if ($SharePoint) {
            $requiredPnPParams = @('PnPAppId', 'PnPUrl', 'Organization', 'PnPCertSubject')
            $missingPnPParams = $requiredPnPParams.Where({ -not (Get-Variable -Name $_ -ErrorAction SilentlyContinue) })

            if ($missingPnPParams) {
                throw "SharePoint connection requires the following parameters: $($missingPnPParams -join ', ')"
            }

            Write-StatusMessage -Message "Connecting to SharePoint Online..." -Type 'INFO'
            $PnPCert = Get-ChildItem Cert:\LocalMachine\My |
            Where-Object { ($_.Subject -like "*$($PnPCertSubject)*") -and ($_.NotAfter -gt $([DateTime]::Now)) }

            if ($null -eq $PnPCert) {
                Exit-Script -Message "No valid PnP PowerShell certificates found in the LocalMachine\My store." -ExitCode ConfigError
            }

            Connect-PnPOnline -Url $PnPUrl -ClientId $PnPAppId -Tenant $Organization -Thumbprint $($PnPCert.Thumbprint)
            Write-StatusMessage -Message "Connected to SharePoint Online" -Type 'OK'
        }
    }

    function Send-GraphMailMessage {
        <#
        .SYNOPSIS
            Sends an email message using Microsoft Graph API.

        .DESCRIPTION
            This function sends an email message using Microsoft Graph API with support for HTML content,
            CC recipients, and file attachments.

        .PARAMETER Subject
            The subject line of the email.

        .PARAMETER Content
            The body content of the email.

        .PARAMETER FromAddress
            The sender's email address. Defaults to value in $config.Email.NotificationFrom.

        .PARAMETER ToAddress
            The recipient's email address. Defaults to value in $config.Email.SecurityTeamEmail.

        .PARAMETER CcAddress
            Optional array of CC recipient email addresses.

        .PARAMETER ContentType
            The type of content in the email body. Must be either 'HTML' or 'Text'. Defaults to 'HTML'.

        .PARAMETER AttachmentPath
            Optional path to a file to attach to the email.

        .PARAMETER AttachmentName
            Optional custom name for the attached file. If not specified, uses the original filename.

        .EXAMPLE
            Send-GraphMailMessage -Subject "Test Email" -Content "Hello World"
            Sends a simple HTML email with default sender and recipient.

        .EXAMPLE
            Send-GraphMailMessage `
                -Subject "Device Setup Complete: $ENV:COMPUTERNAME" ` `
                -Content "<h1>Report Ready</h1><p>The monthly report is attached.</p>" `
                -ToAddress "user@domain.com" `
                -AttachmentPath "C:\Reports\monthly.pdf"
            Sends an HTML email with an attachment.

        .EXAMPLE
            Send-GraphMailMessage `
                -Subject "Team Update" `
                -Content "Weekly update attached" `
                -ToAddress "manager@domain.com" `
                -CcAddress @("team1@domain.com", "team2@domain.com") `
                -ContentType "Text" `
                -AttachmentPath "C:\Updates\weekly.docx"
            Sends a plain text email with CC recipients and an attachment.

        .EXAMPLE
            Send-GraphMailMessage `
                -Subject "Device Setup Complete: $ENV:COMPUTERNAME" `
                -Content "The device $ENV:COMPUTERNAME has completed configuration" `
                -ToAddress "cwooden@compassmsp.com" `
                -CcAddress "cwilliams@compassmsp.com" `
                -AttachmentPath "C:\Logs\setup.log" `
                -AttachmentName "SetupLog.txt"
            Example usage with attachment

        .NOTES
            Requires Microsoft.Graph PowerShell module and appropriate permissions.
            Uses Write-StatusMessage function for logging.
        #>
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Subject,

            [Parameter(Mandatory)]
            [string]$Content,

            [Parameter()]
            [string]$FromAddress,

            [Parameter()]
            [string]$ToAddress,

            [Parameter()]
            [string[]]$CcAddress,

            [Parameter()]
            [ValidateSet('HTML', 'Text')]
            [string]$ContentType = 'HTML',

            [Parameter()]
            [string]$AttachmentPath,

            [Parameter()]
            [string]$AttachmentName
        )

        try {
            $messageParams = @{
                subject      = $Subject
                body         = @{
                    contentType = $ContentType
                    content     = $Content
                }
                toRecipients = @(
                    @{
                        emailAddress = @{
                            address = $ToAddress
                        }
                    }
                )
            }

            # Add CC recipients if specified
            if ($CcAddress) {
                $messageParams['ccRecipients'] = @(
                    $CcAddress | ForEach-Object {
                        @{
                            emailAddress = @{
                                address = $_
                            }
                        }
                    }
                )
            }

            # Add attachment if specified
            if ($AttachmentPath) {
                $attachmentContent = Get-Content -Path $AttachmentPath -Raw -Encoding Byte
                $attachmentBase64 = [System.Convert]::ToBase64String($attachmentContent)

                $messageParams['attachments'] = @(
                    @{
                        '@odata.type' = '#microsoft.graph.fileAttachment'
                        name          = $AttachmentName ?? (Split-Path $AttachmentPath -Leaf)
                        contentType   = 'text/plain'
                        contentBytes  = $attachmentBase64
                    }
                )
            }

            $params = @{
                message         = $messageParams
                saveToSentItems = "false"
            }

            Send-MgUserMail -UserId $FromAddress -BodyParameter $params -ErrorAction Stop
            Write-StatusMessage -Message "Email notification sent successfully" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to send email notification: $_" -Type ERROR
        }
    }

    function Get-TerminationPrerequisites {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$User
        )

        try {
            # Find user in AD with more flexible search
            Write-StatusMessage -Message "Attempting to find $User in Active Directory" -Type 'INFO'

            # Try exact UPN match first
            $UserFromAD = Get-ADUser -Filter "userPrincipalName -eq '$User'" -Properties MemberOf -ErrorAction SilentlyContinue

            # If no exact UPN match, try email match
            if (-not $UserFromAD) {
                $UserFromAD = Get-ADUser -Filter "mail -eq '$User'" -Properties MemberOf -ErrorAction SilentlyContinue
            }

            # If still no match, try partial matches
            if (-not $UserFromAD) {
                $partialMatches = Get-ADUser -Filter "userPrincipalName -like '*$User*' -or mail -like '*$User*' -or displayName -like '*$User*'" `
                    -Properties MemberOf, UserPrincipalName, Mail, DisplayName

                if ($partialMatches) {
                    if ($partialMatches.Count -gt 1) {
                        Write-StatusMessage -Message "Multiple matching users found:" -Type 'WARN'
                        $index = 0
                        $partialMatches | ForEach-Object {
                            Write-Host "`n[$index] DisplayName: $($_.DisplayName)"
                            Write-Host "    UPN: $($_.UserPrincipalName)"
                            Write-Host "    Email: $($_.Mail)"
                            $index++
                        }

                        do {
                            $selection = Read-Host "`nEnter the number of the correct user or 'exit' to cancel"
                            if ($selection -eq 'exit') {
                                Exit-Script -Message "User cancelled the operation" -ExitCode Cancelled
                            }
                        } while ($selection -notmatch '^\d+$' -or [int]$selection -ge $partialMatches.Count)

                        $UserFromAD = Get-ADUser -Identity $partialMatches[$selection].DistinguishedName -Properties MemberOf
                    } else {
                        $UserFromAD = Get-ADUser -Identity $partialMatches[0].DistinguishedName -Properties MemberOf
                    }
                }
            }

            if (-not $UserFromAD) {
                Exit-Script -Message "Could not find user $User in Active Directory" -ExitCode UserNotFound
            }

            # Find Disabled Users OU
            Write-StatusMessage -Message "Attempting to find Disabled users OU" -Type 'INFO'
            $DisabledOUs = @(Get-ADOrganizationalUnit -Filter 'Name -like "*disabled*"')

            if ($DisabledOUs.count -gt 0) {
                # Set the destination OU to the first one found
                $DestinationOU = $DisabledOUs[0].DistinguishedName

                # Try to find user specific OU
                foreach ($OU in $DisabledOUs) {
                    if ($OU.DistinguishedName -like '*user*') {
                        $DestinationOU = $OU.DistinguishedName
                    }
                }
            } else {
                Exit-Script -Message "Could not find disabled OU in Active Directory" -ExitCode GeneralError
            }

            # Find user in Azure/Exchange
            Write-StatusMessage -Message "Attempting to find $($UserFromAD.UserPrincipalName) in Azure" -Type 'INFO'
            try {
                $365Mailbox = Get-Mailbox -Identity $UserFromAD.UserPrincipalName -ErrorAction Stop
                $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -ErrorAction Stop
            } catch {
                Exit-Script -Message "Could not find user in Exchange/Azure: $_" -ExitCode UserNotFound
            }

            # Get user confirmation
            $confirmMessage = @"
The user below will be disabled:
Display Name = $($UserFromAD.Name)
UserPrincipalName = $($UserFromAD.UserPrincipalName)
Mailbox name =  $($365Mailbox.DisplayName)
Azure name = $($MgUser.DisplayName)
Destination OU = $($DestinationOU)

Continue? (Y/N)
"@
            $Confirmation = Read-Host -Prompt $confirmMessage

            if ($Confirmation -ne 'y') {
                Exit-Script -Message "User termination cancelled by user. Did not enter 'Y'" -ExitCode Cancelled
            }

            # Return all the collected information
            return @{
                UserFromAD    = $UserFromAD
                DestinationOU = $DestinationOU
                Mailbox       = $365Mailbox
                MgUser        = $MgUser
            }

        } catch {
            Exit-Script -Message "Failed to validate termination prerequisites: $_" -ExitCode UserNotFound
        }
    }

    function Disable-ADUser {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [Microsoft.ActiveDirectory.Management.ADUser]
            $UserFromAD,

            [Parameter(Mandatory)]
            [string]$DestinationOU
        )

        try {
            Write-StatusMessage -Message "Performing Active Directory Steps" -Type INFO

            # Modify the AD user account
            try {
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

                Set-ADUser @SetADUserParams -ErrorAction Stop
                Write-StatusMessage -Message "User account disabled and attributes cleared" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disable user account"
                Add-ErrorType -ErrorType General
                throw
            }

            # Remove user from all AD groups
            foreach ($group in $UserFromAD.MemberOf) {
                Write-StatusMessage -Message "Removing user from group: $($group)" -Type INFO
                try {
                    Remove-ADGroupMember -Identity $group -Members $UserFromAD.SamAccountName -Confirm:$false -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully removed from group: $($group)" -Type OK
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove from AD group: $group"
                    Add-ErrorType -ErrorType Group
                }
            }
            Write-StatusMessage -Message "User removed from all AD groups" -Type OK

            # Move user to disabled OU
            Write-StatusMessage -Message "Moving user to Disabled OU" -Type INFO
            try {
                $UserFromAD | Move-ADObject -TargetPath $DestinationOU -ErrorAction Stop
                Write-StatusMessage -Message "Successfully moved user to disabled OU" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to move user to disabled OU"
                Add-ErrorType -ErrorType General
                throw
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Disable-ADUser"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Remove-UserSessions {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$UserPrincipalName
        )

        try {
            # Revoke all sessions
            Write-StatusMessage -Message "Revoking all user signed in sessions" -Type INFO
            try {
                Revoke-MgUserSignInSession -UserId $UserPrincipalName -ErrorAction Stop
                Write-StatusMessage -Message "Successfully revoked all user sessions" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to revoke user sessions"
                Add-ErrorType -ErrorType Permission
            }

            # Remove Mobile Devices
            Write-StatusMessage -Message "Removing all mobile devices" -Type INFO
            try {
                $mobileDevices = Get-MobileDevice -Mailbox $UserPrincipalName -ErrorAction Stop
                foreach ($mobileDevice in $mobileDevices) {
                    Write-StatusMessage -Message "Removing mobile device: $($mobileDevice.Id)" -Type INFO
                    try {
                        Remove-MobileDevice -DeviceID $mobileDevice.Id -Confirm:$false -ErrorAction Stop
                        Write-StatusMessage -Message "Successfully removed mobile device: $($mobileDevice.Id)" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove mobile device $($mobileDevice.Id)"
                        Add-ErrorType -ErrorType General
                    }
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get mobile devices"
                Add-ErrorType -ErrorType General
            }

            # Disable Azure AD devices
            try {
                $termUserDevices = Get-MgUserRegisteredDevice -UserId $UserPrincipalName -ErrorAction Stop
                foreach ($termUserDevice in $termUserDevices) {
                    Write-StatusMessage -Message "Disabling registered device: $($termUserDevice.Id)" -Type INFO
                    try {
                        Update-MgDevice -DeviceId $termUserDevice.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop
                        Write-StatusMessage -Message "Successfully disabled device: $($termUserDevice.Id)" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disable device $($termUserDevice.Id)"
                        Add-ErrorType -ErrorType General
                    }
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get registered devices"
                Add-ErrorType -ErrorType General
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Remove-UserSessions"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Set-TerminatedMailbox {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            $Mailbox,

            [Parameter()]
            [string]$ForwardingAddress,

            [Parameter()]
            [string]$GrantAccessTo
        )

        try {
            # Disable mailbox forwarding
            Write-StatusMessage -Message "Disabling existing mailbox forwarding" -Type INFO
            try {
                Set-Mailbox -Identity $Mailbox.Identity -ForwardingAddress $null -ForwardingSmtpAddress $null -ErrorAction Stop
                Write-StatusMessage -Message "Successfully disabled existing forwarding" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disable existing mailbox forwarding"
                Add-ErrorType -ErrorType Mailbox
            }

            # Change mailbox to shared
            Write-StatusMessage -Message "Converting to shared mailbox" -Type INFO
            try {
                Set-Mailbox -Identity $Mailbox.Identity -Type Shared -ErrorAction Stop
                Write-StatusMessage -Message "Successfully converted to shared mailbox" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to convert to shared mailbox"
                Add-ErrorType -ErrorType Mailbox
            }

            # Set forwarding if specified
            if ($ForwardingAddress) {
                try {
                    $forwardUser = Get-Mailbox $ForwardingAddress -ErrorAction Stop
                    Write-StatusMessage -Message "Setting up forwarding to $($forwardUser.PrimarySmtpAddress)" -Type INFO

                    $mailboxParams = @{
                        Identity                   = $Mailbox.Identity
                        ForwardingAddress          = $ForwardingAddress
                        DeliverToMailboxAndForward = $False
                        ErrorAction                = 'Stop'
                    }

                    Set-Mailbox @mailboxParams
                    Write-StatusMessage -Message "Successfully set up mail forwarding" -Type OK
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to set up mail forwarding"
                    Add-ErrorType -ErrorType Mailbox
                }
            }

            # Grant access if specified
            if ($GrantAccessTo) {
                try {
                    $accessUser = Get-Mailbox $GrantAccessTo -ErrorAction Stop
                    Write-StatusMessage -Message "Granting full access to $($accessUser.PrimarySmtpAddress)" -Type INFO

                    $mailboxPermissionParams = @{
                        Identity        = $Mailbox.Identity
                        User            = $GrantAccessTo
                        AccessRights    = 'FullAccess'
                        InheritanceType = 'All'
                        AutoMapping     = $true
                        ErrorAction     = 'Stop'
                    }

                    Add-MailboxPermission @mailboxPermissionParams
                    Write-StatusMessage -Message "Successfully granted full access permissions" -Type OK
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to grant mailbox permissions"
                    Add-ErrorType -ErrorType Permission
                }
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Set-TerminatedMailbox"
            Add-ErrorType -ErrorType Mailbox
            throw
        }
    }

    function Remove-UserFromGroups {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$userId,

            [Parameter()]
            [string]$ExportPath
        )

        try {
            # Define filter parameters
            $filterParams = @{
                FilterScript = {
                    $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and
                    $null -eq $_.AdditionalProperties.membershipRule -and
                    $_.onPremisesSyncEnabled -ne 'True'
                }
            }

            # Define select parameters
            $selectParams = @{
                Property = @(
                    'Id'
                    @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }
                    @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }
                    @{n = 'groupType'; e = { $_.AdditionalProperties.groupTypes } }
                    @{n = 'securityEnabled'; e = { $_.AdditionalProperties.securityEnabled } }
                )
            }

            Write-StatusMessage -Message "Finding Azure groups" -Type INFO

            try {
                $All365Groups = Get-MgUserMemberOf -UserId $userId -ErrorAction Stop |
                Where-Object @filterParams |
                Select-Object @selectParams

                Write-StatusMessage -Message "Found $($All365Groups.Count) groups to process" -Type INFO

                # Export groups if path provided
                if ($ExportPath) {
                    try {
                        $All365Groups | Export-Csv -Path $ExportPath -NoTypeInformation -ErrorAction Stop
                        Write-StatusMessage -Message "Exported user groups to: $ExportPath" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to export user groups"
                        Add-ErrorType -ErrorType General
                    }
                }

                foreach ($365Group in $All365Groups) {
                    Write-StatusMessage -Message "Processing group: $($365Group.DisplayName)" -Type INFO

                    try {
                        if ($365Group.securityEnabled -eq 'True' -or $365Group.groupType -eq 'Unified') {
                            Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $userId -ErrorAction Stop
                            Write-StatusMessage -Message "Removed from Security/Unified Group: $($365Group.DisplayName)" -Type OK
                        } else {
                            Remove-DistributionGroupMember -Identity $365Group.Id -Member $userId -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                            Write-StatusMessage -Message "Removed from Distribution Group: $($365Group.DisplayName)" -Type OK
                        }
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove from group $($365Group.DisplayName)"
                        Add-ErrorType -ErrorType Group
                    }
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get user group memberships"
                Add-ErrorType -ErrorType Group
                throw
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Remove-UserFromGroups"
            Add-ErrorType -ErrorType Group
            throw
        }
    }

    function Remove-UserFromDirectoryRoles {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]$UserId
        )

        try {
            Write-StatusMessage -Message "Checking for directory role memberships..." -Type INFO

            try {
                # Get all directory roles the user is a member of
                $directoryRoles = Get-MgUserMemberOf -UserId $UserId -ErrorAction Stop |
                Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.directoryRole' }

                if (-not $directoryRoles) {
                    Write-StatusMessage -Message "User is not a member of any directory roles" -Type INFO
                    return
                }

                Write-StatusMessage -Message "Found $($directoryRoles.Count) directory role(s)" -Type INFO

                foreach ($role in $directoryRoles) {
                    try {
                        $roleId = $role.Id
                        $roleName = $role.AdditionalProperties.displayName

                        Write-StatusMessage -Message "Removing from role: $roleName" -Type INFO
                        Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $roleId -DirectoryObjectId $UserId -ErrorAction Stop
                        Write-StatusMessage -Message "Successfully removed from role: $roleName" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove from role $roleName"
                        Add-ErrorType -ErrorType Permission
                    }
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get directory roles"
                Add-ErrorType -ErrorType Permission
                throw
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Remove-UserFromDirectoryRoles"
            Add-ErrorType -ErrorType Permission
            throw
        }
    }

    function Remove-UserLicenses {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$UserId,

            [Parameter()]
            [string]$ExportPath
        )

        try {
            Write-StatusMessage -Message "Starting license removal process" -Type INFO

            try {
                # Get and export license details if path provided
                $licenseDetails = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop |
                Select-Object SkuPartNumber, SkuId, Id

                if ($ExportPath) {
                    try {
                        $licenseDetails | Export-Csv -Path $ExportPath -NoTypeInformation -ErrorAction Stop
                        Write-StatusMessage -Message "Exported user licenses to: $ExportPath" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to export user licenses"
                        Add-ErrorType -ErrorType License
                    }
                }

                # Define primary licenses that must be removed last
                $primaryLicenses = @(
                    "O365_BUSINESS_ESSENTIALS"
                    "SPE_E3"
                    "SPB"
                    "EXCHANGESTANDARD"
                )

                # Step 1: Remove Ancillary Licenses
                foreach ($license in ($licenseDetails | Where-Object { $_.SkuPartNumber -notin $primaryLicenses })) {
                    try {
                        Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction Stop
                        Write-StatusMessage -Message "Removed Ancillary License: $($license.SkuPartNumber)" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove Ancillary License $($license.SkuPartNumber)"
                        Add-ErrorType -ErrorType License
                    }
                }

                # Step 2: Remove Primary Licenses
                foreach ($license in ($licenseDetails | Where-Object { $_.SkuPartNumber -in $primaryLicenses })) {
                    try {
                        Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction Stop
                        Write-StatusMessage -Message "Removed Primary License: $($license.SkuPartNumber)" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove Primary License $($license.SkuPartNumber)"
                        Add-ErrorType -ErrorType License
                    }
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get user licenses"
                Add-ErrorType -ErrorType License
                throw
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Remove-UserLicenses"
            Add-ErrorType -ErrorType License
            throw
        }
    }

    function Remove-UserFromZoom {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$UserId
        )

        try {
            Write-StatusMessage -Message "Checking Zoom assignments..." -Type INFO
            try {
                $ZoomSSO = Get-MgUserAppRoleAssignment -UserId $UserId -ErrorAction Stop |
                Where-Object { $_.ResourceDisplayName -eq 'Zoom Workplace Phones' }

                if ($ZoomSSO) {
                    try {
                        Remove-MgUserAppRoleAssignment -AppRoleAssignmentId $ZoomSSO.Id -UserId $UserId -ErrorAction Stop
                        Write-StatusMessage -Message "Successfully removed user from Zoom Workplace Phones" -Type OK
                    } catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to remove Zoom assignment"
                        Add-ErrorType -ErrorType Permission
                    }
                } else {
                    Write-StatusMessage -Message "User is not assigned to Zoom Workplace Phones" -Type INFO
                }
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get Zoom assignments"
                Add-ErrorType -ErrorType Permission
                throw
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Remove-UserFromZoom"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Set-TerminatedOneDrive {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$UserPrincipalName,

            [Parameter()]
            [switch]$SetReadOnly,

            [Parameter()]
            [string]$OneDriveUser,

            [Parameter()]
            [int]$MaxRetries = 3,

            [Parameter()]
            [int]$RetryDelaySeconds = 30
        )

        try {
            # Get OneDrive URL with retry logic
            Write-StatusMessage -Message "Getting OneDrive URL" -Type INFO
            $retryCount = 0
            $success = $false
            $UserOneDriveURL = $null

            do {
                try {
                    $profileProps = Get-PnPUserProfileProperty -Account $UserPrincipalName -Properties PersonalUrl -ErrorAction Stop
                    $UserOneDriveURL = $profileProps.PersonalUrl

                    if ($UserOneDriveURL) {
                        $success = $true
                        Write-StatusMessage -Message "Successfully retrieved OneDrive URL" -Type OK
                    } else {
                        throw [System.InvalidOperationException]::new("OneDrive URL is empty")
                    }
                } catch {
                    $retryCount++
                    if ($retryCount -ge $MaxRetries) {
                        Write-StatusMessage -Message "Failed to get OneDrive URL after $MaxRetries attempts" -Type WARN
                        $response = Read-Host "Would you like to try again? (Y/N)"
                        if ($response -eq 'Y') {
                            $retryCount = 0
                            Write-StatusMessage -Message "Retrying OneDrive URL retrieval..." -Type INFO
                        } else {
                            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "OneDrive URL retrieval skipped by user"
                            Add-ErrorType -ErrorType OneDrive
                            return
                        }
                    } else {
                        Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed. Waiting $RetryDelaySeconds seconds..." -Type INFO
                        Start-Sleep -Seconds $RetryDelaySeconds
                    }
                }
            } while (-not $success -and $retryCount -lt $MaxRetries)

            if (-not $UserOneDriveURL) {
                Write-StatusMessage -Message "OneDrive URL not found for user" -Type WARN
                return
            }

            # Verify OneDrive site exists
            Write-StatusMessage -Message "Verifying OneDrive site accessibility" -Type INFO
            try {
                $site = Get-PnPTenantSite -Url $UserOneDriveURL -ErrorAction Stop
                Write-StatusMessage -Message "OneDrive site verified" -Type OK
            } catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "OneDrive site not accessible"
                Add-ErrorType -ErrorType OneDrive

                $response = Read-Host "Would you like to wait for OneDrive provisioning? (Y/N)"
                if ($response -eq 'Y') {
                    $retryCount = 0
                    do {
                        try {
                            Start-Sleep -Seconds $RetryDelaySeconds
                            $site = Get-PnPTenantSite -Url $UserOneDriveURL -ErrorAction Stop
                            $success = $true
                            Write-StatusMessage -Message "OneDrive site now accessible" -Type OK
                        } catch {
                            $retryCount++
                            Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries. Waiting $RetryDelaySeconds seconds..." -Type INFO
                        }
                    } while (-not $success -and $retryCount -lt $MaxRetries)
                } else {
                    Write-StatusMessage -Message "OneDrive configuration skipped by user" -Type WARN
                    return
                }
            }

            # Set OneDrive to read-only if specified
            if ($SetReadOnly) {
                Write-StatusMessage -Message "Setting OneDrive to Read Only" -Type INFO
                try {
                    Set-PnPTenantSite -Url $UserOneDriveURL -LockState ReadOnly -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully set OneDrive to read-only" -Type OK
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to set OneDrive to read-only"
                    Add-ErrorType -ErrorType OneDrive
                }
            }

            # Grant access if OneDriveUser is provided
            if ($OneDriveUser) {
                try {
                    Write-StatusMessage -Message "Granting OneDrive access to $OneDriveUser" -Type INFO

                    $pnpParams = @{
                        Url         = $UserOneDriveURL
                        Owners      = $OneDriveUser
                        ErrorAction = 'Stop'
                    }

                    Set-PnPTenantSite @pnpParams
                    Write-StatusMessage -Message "Successfully granted OneDrive access" -Type OK
                    Write-StatusMessage -Message "OneDrive URL: $UserOneDriveURL" -Type INFO

                    do {
                        $response = Read-Host "Please copy the OneDrive URL above. Have you copied it? (y/n)"
                    } while ($response -ne 'y')
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to grant OneDrive access"
                    Add-ErrorType -ErrorType OneDrive
                }
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Set-TerminatedOneDrive"
            Add-ErrorType -ErrorType OneDrive
            throw
        }
    }

    function Start-ADSyncAndFinalize {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$User,

            [Parameter(Mandatory)]
            [string]$UserPrincipalName,

            [Parameter(Mandatory)]
            [string]$DestinationOU,

            [Parameter()]
            [string]$GrantUserFullControl,

            [Parameter()]
            [string]$SetUserMailFWD,

            [Parameter()]
            [string]$GrantUserOneDriveAccess,

            [Parameter()]
            [string]$ExportPath
        )

        try {
            # Start AD Sync
            Write-StatusMessage -Message "Starting AD sync cycle" -Type INFO
            try {
                Import-Module -Name ADSync -UseWindowsPowerShell -ErrorAction Stop -Verbose:$false
                $null = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
                Write-StatusMessage -Message "AD sync cycle initiated successfully" -Type OK
            } catch {
                try {
                    # Fallback to direct PowerShell execution if module import fails
                    $null = powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta
                    if ($LASTEXITCODE -ne 0) {
                        throw "PowerShell execution failed with exit code: $LASTEXITCODE"
                    }
                    Write-StatusMessage -Message "AD sync cycle initiated through PowerShell" -Type OK
                } catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to start AD sync cycle"
                    Add-ErrorType -ErrorType Sync
                    throw
                }
            }

            # Create summary message
            $summaryParts = @(
                "Summary of Actions:",
                "----------------------------------------",
                "User $User should now be disabled unless any errors occurred during the process.",
                "User disabled: $UserPrincipalName",
                "Moved to OU: $DestinationOU"
            )

            if ($GrantUserFullControl) {
                $summaryParts += "Mailbox access granted to: $GrantUserFullControl"
            }
            if ($SetUserMailFWD) {
                $summaryParts += "Mail forwarded to: $SetUserMailFWD"
            }
            if ($GrantUserOneDriveAccess) {
                $summaryParts += "OneDrive access granted to: $GrantUserOneDriveAccess"
            }
            if ($ExportPath) {
                $summaryParts += "Exports saved to: $ExportPath"
            }

            $summaryParts += "----------------------------------------"
            $summaryMessage = $summaryParts -join "`n"

            Write-StatusMessage -Message $summaryMessage -Type SUMMARY

            # Add error summary if there were errors
            if ($script:errorCount -gt 0) {
                Write-StatusMessage -Message (Get-ErrorSummary) -Type SUMMARY
            }

            Exit-Script -Message "$User has been successfully disabled." -ExitCode Success
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Start-ADSyncAndFinalize"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green

    Write-Host "`n  Script Ready!" -ForegroundColor Cyan
    Write-Host "  Press any key to continue..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Clear-Host

    #Region Main Execution
    try {

        # Step 0: Initialization
        Write-ProgressStep -StepName $progressSteps[0].Name -Status $progressSteps[0].Description

        # Load configuration
        $config = Get-ScriptConfig

        # Get connection parameters from config
        $Organization = $config.ExchangeOnline.Organization
        $ExOAppId = $config.ExchangeOnline.AppId
        $ExOCertSubject = $Config.ExchangeOnline.CertificateSubject
        $GraphAppId = $config.Graph.AppId
        $tenantID = $config.Graph.TenantId
        $GraphCertSubject = $Config.Graph.CertificateSubject
        $PnPAppId = $config.PnPSharePoint.AppId
        $PnPUrl = $config.PnPSharePoint.Url
        $PnPCertSubject = $Config.PnPSharePoint.CertificateSubject
        $localExportPath = $config.Paths.TermExportPath

        Connect-ServiceEndpoints

        # Call the custom input window function
        $result = Show-CustomTerminationWindow
        if (-not $result) {
            Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
        }

        # Get values from window
        $User = $result.InputUser
        $GrantUserFullControl = $result.InputUserFullControl
        $SetUserMailFWD = $result.InputUserFWD
        $GrantUserOneDriveAccess = $result.InputUserOneDriveAccess
        $SetOneDriveReadOnly = $result.SetOneDriveReadOnly

        # Set confirmation flags
        $SPOAccessConfirmation = if ($GrantUserOneDriveAccess) { 'y' } else { 'n' }

        # Validate OneDrive access user if specified
        $oneDriveUser = $null
        if ($SPOAccessConfirmation -eq 'y') {
            try {
                Write-StatusMessage -Message "Validating OneDrive access user..." -Type 'INFO'
                $oneDriveUser = Get-Mailbox $GrantUserOneDriveAccess -ErrorAction Stop
                Write-StatusMessage -Message "OneDrive access user validated" -Type 'OK'
            } catch {
                Write-StatusMessage -Message "Invalid OneDrive access user specified: $_" -Type 'ERROR'
                Write-StatusMessage -Message "OneDrive access user validation failed. Skipping OneDrive access grant." -Type 'ERROR'
                $SPOAccessConfirmation = 'n'
            }
        }

        # Step 1: User Input
        Write-ProgressStep -StepName $progressSteps[1].Name -Status $progressSteps[1].Description
        $userInfo = Get-TerminationPrerequisites -User $User

        # Extract variables for use in the rest of the script
        $UserFromAD = $userInfo.UserFromAD
        $DestinationOU = $userInfo.DestinationOU
        $365Mailbox = $userInfo.Mailbox
        $MgUser = $userInfo.MgUser

        # Step 2: AD Tasks
        Write-ProgressStep -StepName $progressSteps[2].Name -Status $progressSteps[2].Description
        Disable-ADUser -UserFromAD $UserFromAD -DestinationOU $DestinationOU

        # Step 3: Azure/Entra Tasks
        Write-ProgressStep -StepName $progressSteps[3].Name -Status $progressSteps[3].Description

        Remove-UserSessions -UserPrincipalName $UserFromAD.UserPrincipalName

        $mailboxParams = @{
            Mailbox = $365Mailbox
        }

        # Only add these parameters if they exist and have values
        if ($SetUserMailFWD) {
            $mailboxParams['ForwardingAddress'] = $SetUserMailFWD

        }

        if ($GrantUserFullControl) {
            $mailboxParams['GrantAccessTo'] = $GrantUserFullControl
        }

        Set-TerminatedMailbox @mailboxParams

        # Step 4: Remove Directory Roles
        Write-ProgressStep -StepName $progressSteps[5].Name -Status $progressSteps[5].Description
        Remove-UserFromDirectoryRoles -UserId $MgUser.Id

        # Step 5: Remove Groups
        Write-ProgressStep -StepName $progressSteps[6].Name -Status $progressSteps[6].Description
        $groupExportPath = Join-Path $localExportPath "$($User)_Groups_Id.csv"
        Remove-UserFromGroups -UserId $MgUser.Id -ExportPath $groupExportPath

        # Step 6: Remove Licenses
        Write-ProgressStep -StepName $progressSteps[7].Name -Status $progressSteps[7].Description
        $licensePath = Join-Path $localExportPath "$($User)_License_Id.csv"
        Remove-UserLicenses -UserId $UserFromAD.UserPrincipalName -ExportPath $licensePath

        # Step 7: Send Email
        Write-ProgressStep -StepName $progressSteps[8].Name -Status $progressSteps[8].Description
        $emailSubject = "KB4 – Remove User"
        $emailContent = "The following user need to be removed to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
        $MsgFrom = $config.Email.NotificationFrom
        $ToAddress = $config.Email.NotificationTo
        Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -Subject $emailSubject -Content $emailContent

        # Step 8: Remove from Zoom
        Write-ProgressStep -StepName $progressSteps[9].Name -Status $progressSteps[9].Description
        Remove-UserFromZoom -UserId $MgUser.Id

        # Step 9: Configure OneDrive
        Write-ProgressStep -StepName $progressSteps[10].Name -Status $progressSteps[10].Description

        Write-StatusMessage -Message "Disconnecting from Exchange Online and Graph. This may take a few moments..." -Type INFO

        Connect-ServiceEndpoints -Disconnect -ExchangeOnline -Graph

        Write-StatusMessage -Message "Connecting to SharePoint Online..." -Type INFO

        Connect-ServiceEndpoints -SharePoint

        $oneDriveParams = @{
            UserPrincipalName = $UserFromAD.UserPrincipalName
        }

        if ($SetOneDriveReadOnly) {
            $oneDriveParams['SetReadOnly'] = $true
        }

        if ($SPOAccessConfirmation -eq 'y') {
            $oneDriveParams['OneDriveUser'] = $oneDriveUser
        }

        Set-TerminatedOneDrive @oneDriveParams

        # Step 10: Final Sync and Summary
        Write-ProgressStep -StepName $progressSteps[11].Name -Status $progressSteps[11].Description

        $null = Disconnect-PnPOnline

        Start-ADSyncAndFinalize -User $User `
            -UserPrincipalName $UserFromAD.UserPrincipalName `
            -DestinationOU $DestinationOU `
            -GrantUserFullControl $GrantUserFullControl `
            -SetUserMailFWD $SetUserMailFWD `
            -GrantUserOneDriveAccess $GrantUserOneDriveAccess `
            -ExportPath $localExportPath

    } catch {
        Exit-Script -Message "Failed to complete termination process: $_" -ExitCode GeneralError
    }

    # Script duration tracking
    if ($script:errorCount -gt 0) {
        Write-StatusMessage "Completed with $script:errorCount errors" -Type WARN
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-StatusMessage "Script completed in $($duration.TotalMinutes.ToString('F2')) minutes" -Type INFO
} finally {
    Stop-Job $loadingJob | Out-Null
    Remove-Job $loadingJob | Out-Null
}

#Region Functions
# Add error handling helper functions
function Write-ErrorRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter()]
        [string]$CustomMessage,

        [Parameter()]
        [ValidateSet('Warning', 'Error')]
        [string]$ErrorLevel = 'Error'
    )

    # Build detailed error message
    $errorDetails = @(
        "Error Details:"
        "-------------"
        "Message: $($ErrorRecord.Exception.Message)"
        "Category: $($ErrorRecord.CategoryInfo.Category)"
        "Target: $($ErrorRecord.TargetObject)"
        "Script: $($ErrorRecord.InvocationInfo.ScriptName)"
        "Line: $($ErrorRecord.InvocationInfo.ScriptLineNumber)"
        "Command: $($ErrorRecord.InvocationInfo.MyCommand)"
    )

    if ($CustomMessage) {
        $errorDetails = @("$CustomMessage", "") + $errorDetails
    }

    # Log the error
    $errorMessage = $errorDetails -join "`n"
    Write-StatusMessage -Message $errorMessage -Type $ErrorLevel.ToUpper()
    Write-Log -Message $errorMessage -Level $ErrorLevel.ToUpper()

    # Add to error count if it's an error
    if ($ErrorLevel -eq 'Error') {
        $script:errorCount++
    }
}

function Add-ErrorType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(
            'Configuration',
            'Connection',
            'Permission',
            'Validation',
            'Sync',
            'License',
            'Group',
            'Mailbox',
            'OneDrive',
            'General'
        )]
        [string]$ErrorType
    )

    if (-not $script:errorTypes) {
        $script:errorTypes = @{
            Configuration = 0
            Connection    = 0
            Permission    = 0
            Validation    = 0
            Sync          = 0
            License       = 0
            Group         = 0
            Mailbox       = 0
            OneDrive      = 0
            General       = 0
        }
    }

    $script:errorTypes[$ErrorType]++
}

function Get-ErrorSummary {
    [CmdletBinding()]
    param()

    if ($script:errorCount -eq 0) {
        return "No errors occurred during execution"
    }

    $summary = @(
        "Error Summary:"
        "-------------"
        "Total Errors: $script:errorCount"
        ""
        "Error Types:"
    )

    foreach ($type in $script:errorTypes.Keys | Sort-Object) {
        if ($script:errorTypes[$type] -gt 0) {
            $summary += "- $type : $($script:errorTypes[$type])"
        }
    }

    return $summary -join "`n"
}