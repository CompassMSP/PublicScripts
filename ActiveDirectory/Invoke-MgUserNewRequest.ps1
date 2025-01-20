#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync -RunAsAdministrator

<#
.SYNOPSIS
    Creates a new user based on a template user in Active Directory and Microsoft 365.

.DESCRIPTION
    This script creates a new user account by copying attributes and group memberships
    from an existing template user. It handles both on-premises AD and Microsoft 365 setup.

    The script will display a GUI window to collect:
    - New user's full name
    - Template user to copy from
    - Phone number
    - Required license selection
    - Optional ancillary licenses

    IMPORTANT: This script must be run from the Primary Domain Controller with AD Connect installed.

    NOTE: Sensitive information (app IDs, certificates, etc.) is stored in a secure configuration file managed by Get-ScriptConfig.
    The config file should be placed at: C:\ProgramData\CompassScripts\config.json

.EXAMPLE
    .\Invoke-MgNewUserRequest.ps1

    This will launch the GUI window to collect the required information.

.NOTES
    Author: Chris Williams
    Created: 2022-03-02
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

    2.1.0        2024-10-15  Feature Update:
                          - Added BookWithMeId validation
                          - Enhanced AD Sync loop handling
                          - Reworked GUI interface
                          - Added QuickEdit and InsertMode functions
                          - Added SMTP duplicate checking
                          - Removed KnowBe4 SCIM integration per SecurePath Team
                          - Added Email Forwarding functionality - KnowBe4 Notification

    2.0.0        2024-05-08  Major Feature Update:
                          - Added input box system
                          - Added EntraID P2 license checkbox
                          - Enhanced UI boxes for variables
                          - Added KB4 email delivery
                          - Added MeetWithMeId and AD properties
                          - Updated KnowBe4 SCIM integration
                          - Added template user validation

    1.2.0        2024-02-12  Feature Updates:
                          - Enhanced license display output
                          - Improved group management functions
                          - Added KnowBe4 SCIM integration

    1.1.0        2022-06-27  Feature Updates:
                          - Added duplicate attribute checking
                          - Added fax attributes copying
                          - Enhanced group lookup and management
                          - Added AD sync validation

    1.0.0        2022-03-02  Initial Release:
                          - Basic user creation functionality
                          - Template user copying
                          - Group membership handling
    ------------------------------------------------------------------------------
#>

# Initialize loading animation
Clear-Host
$loadingChars = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
$i = 0
$loadingJob = Start-Job -ScriptBlock { while ($true) { Start-Sleep -Milliseconds 100 } }

try {
    Write-Host "`n  Initializing New User Creation Script..." -ForegroundColor Cyan

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
        @{ Number = 1; Name = "User Input"; Description = "Gathering new user details" }
        @{ Number = 2; Name = "Validation"; Description = "Validating inputs and prerequisites" }
        @{ Number = 3; Name = "AD Creation"; Description = "Creating user in Active Directory" }
        @{ Number = 4; Name = "AD Group Copy"; Description = "Copying AD group memberships" }
        @{ Number = 5; Name = "Azure Sync"; Description = "Syncing to Azure AD" }
        @{ Number = 6; Name = "License Setup"; Description = "Assigning licenses" }
        @{ Number = 7; Name = "Group Setup"; Description = "Adding to Microsoft 365 groups" }
        @{ Number = 8; Name = "Email Setup"; Description = "Configuring email settings" }
        @{ Number = 9; Name = "Zoom Setup"; Description = "Configuring Zoom access" }
        @{ Number = 10; Name = "OneDrive Setup"; Description = "Configuring OneDrive" }
        @{ Number = 11; Name = "Cleanup and Summary"; Description = "Running cleanup and summary" }
    )
    Write-Host "`r  [✓] Progress tracking initialized" -ForegroundColor Green

    Write-Host "  [$($loadingChars[$i % $loadingChars.Length])] Loading functions..." -NoNewline -ForegroundColor Yellow
    $script:errorCount = 0
    $script:totalSteps = $progressSteps.Count  # Make it script-scoped

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
            Write-Progress -Activity "New User Creation" -Status $Status
        } else {
            Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $Status" -Type INFO
            Write-Progress -Activity "New User Creation" -Status $Status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
        }
    }

    #Region Functions
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

        try {
            if ($Type -eq 'SUMMARY') {
                Write-Host $Message -ForegroundColor $config[$Type].Color
            }
            else {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $statusPadded = $config[$Type].Status.PadRight(7)
                Write-Host "[$timestamp] [$statusPadded] $Message" -ForegroundColor $config[$Type].Color
            }

            # Write to log file if it's not already being called from Write-Log
            if (-not ([System.Management.Automation.CallStackFrame]$MyInvocation.GetStackTrace() -match 'Write-Log')) {
                Write-Log -Message $Message -Level $Type
            }
        }
        catch {
            Write-Host "Failed to write status message: $_" -ForegroundColor Red
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
            }
            catch {
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
        }
        catch {
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
                    ClientSecret       = Read-Host "Enter PnP Client Secret"
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
            [string]$LogPath = $config.Paths.LogPath
        )

        try {
            # Create log directory if it doesn't exist
            $logDir = Split-Path -Path $LogPath -Parent
            if (-not (Test-Path -Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }

            # Format timestamp and message
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"

            # Write to log file
            Add-Content -Path $LogPath -Value $logMessage -ErrorAction Stop

            # Track errors
            if ($Level -eq 'ERROR') {
                $script:errorCount++
            }
        }
        catch {
            Write-StatusMessage -Message "Failed to write to log: $_" -Type ERROR
            Add-ErrorType -ErrorType General
        }
    }

    # Add this error handling helper function in the #Region Functions section
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

    # Add this function to track error types
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

    # Add this function to get error summary
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

    .NOTES
        Requires appropriate certificates and permissions configured in config.json.
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

        try {
            if ($Disconnect) {
                Write-StatusMessage -Message "Disconnecting from services..." -Type INFO

                # If no specific service is specified, disconnect from all
                $disconnectAll = -not ($ExchangeOnline -or $Graph -or $SharePoint)

                if ($ExchangeOnline -or $disconnectAll) {
                    try {
                        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                        Write-StatusMessage -Message "Disconnected from Exchange Online" -Type OK
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disconnect from Exchange Online"
                        Add-ErrorType -ErrorType Connection
                    }
                }

                if ($Graph -or $disconnectAll) {
                    try {
                        Disconnect-MgGraph -ErrorAction Stop
                        Write-StatusMessage -Message "Disconnected from Microsoft Graph" -Type OK
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disconnect from Microsoft Graph"
                        Add-ErrorType -ErrorType Connection
                    }
                }

                if ($SharePoint -or $disconnectAll) {
                    try {
                        Disconnect-PnPOnline -ErrorAction Stop
                        Write-StatusMessage -Message "Disconnected from SharePoint Online" -Type OK
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to disconnect from SharePoint Online"
                        Add-ErrorType -ErrorType Connection
                    }
                }
                return
            }

            # Connection logic
            if ($ExchangeOnline -or (-not ($ExchangeOnline -or $Graph -or $SharePoint))) {
                Write-StatusMessage -Message "Connecting to Exchange Online..." -Type INFO
                try {
                    $ExOCert = Get-ChildItem Cert:\LocalMachine\My |
                        Where-Object { ($_.Subject -like "*$ExOCertSubject*") -and ($_.NotAfter -gt (Get-Date)) } |
                        Select-Object -First 1

                    if (-not $ExOCert) {
                        throw "No valid Exchange Online certificate found"
                    }

                    Connect-ExchangeOnline -AppId $ExOAppId -Organization $Organization `
                        -CertificateThumbprint $ExOCert.Thumbprint -ShowBanner:$false -ErrorAction Stop
                    Write-StatusMessage -Message "Connected to Exchange Online" -Type OK
                }
                catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to connect to Exchange Online"
                    Add-ErrorType -ErrorType Connection
                    throw
                }
            }

            if ($Graph -or (-not ($ExchangeOnline -or $Graph -or $SharePoint))) {
                Write-StatusMessage -Message "Connecting to Microsoft Graph..." -Type INFO
                try {
                    $GraphCert = Get-ChildItem Cert:\LocalMachine\My |
                        Where-Object { ($_.Subject -like "*$GraphCertSubject*") -and ($_.NotAfter -gt (Get-Date)) } |
                        Select-Object -First 1

                    if (-not $GraphCert) {
                        throw "No valid Graph certificate found"
                    }

                    Connect-MgGraph -ClientId $GraphAppId -TenantId $tenantId `
                        -Certificate $GraphCert -ErrorAction Stop
                    Write-StatusMessage -Message "Connected to Microsoft Graph" -Type OK
                }
                catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to connect to Microsoft Graph"
                    Add-ErrorType -ErrorType Connection
                    throw
                }
            }

            if ($SharePoint) {
                Write-StatusMessage -Message "Connecting to SharePoint Online..." -Type INFO
                try {
                    Connect-PnPOnline -Url $PnPUrl -ClientId $PnPAppId -Tenant $Organization `
                        -CertificateThumbprint $PnPCert.Thumbprint -ErrorAction Stop
                    Write-StatusMessage -Message "Connected to SharePoint Online" -Type OK
                }
                catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to connect to SharePoint Online"
                    Add-ErrorType -ErrorType Connection
                    throw
                }
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Connect-ServiceEndpoints"
            Add-ErrorType -ErrorType Connection
            throw
        }
    }

    function Show-NewUserRequestWindow {
        # 1. Add required assemblies
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

        # 2. UI Assembly Helper Functions
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

        # 3. Validation Helper Functions
        function Test-DisplayName {
            param ([string]$DisplayName)
            return $DisplayName -match '^[A-Za-z]+ [A-Za-z]+$'
        }

        function Format-MobileNumber {
            param ([string]$MobileNumber)
            $digits = -join ($MobileNumber -replace '\D', '')
            if ($digits.Length -eq 10) {
                return "($($digits.Substring(0, 3))) $($digits.Substring(3, 3))-$($digits.Substring(6, 4))"
            }
        }

        function Show-ValidationError {
            param (
                [string]$Message,
                [string]$Title = "Input Error"
            )
            [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }

        # 4. License Processing Functions
        function Get-LicenseDisplayName {
            param ([string]$SkuPartNumber)
            $displayName = switch -Regex ($SkuPartNumber) {
                "POWERAUTOMATE_ATTENDED_RPA" { "Power Automate Premium" }
                "PROJECT_MADEIRA_PREVIEW_IW_SKU" { "Dynamics 365 Business Central for IWs" }
                "PROJECT_PLAN3_DEPT" { "Project Plan 3 (for Department)" }
                "FLOW_FREE" { "Microsoft Power Automate Free" }
                "WINDOWS_STORE" { "Windows Store for Business" }
                "RMSBASIC" { "Rights Management Service Basic Content Protection" }
                "RIGHTSMANAGEMENT_ADHOC" { "Rights Management Adhoc" }
                "POWERAPPS_VIRAL" { "Microsoft Power Apps Plan 2 Trial" }
                "POWERAPPS_PER_USER" { "Power Apps Premium" }
                "POWERAPPS_DEV" { "Microsoft PowerApps for Developer" }
                "PHONESYSTEM_VIRTUALUSER" { "Microsoft Teams Phone Resource Account" }
                "MICROSOFT_BUSINESS_CENTER" { "Microsoft Business Center" }
                "MCOPSTNC" { "Communications Credits" }
                "MCOPSTN1" { "Skype for Business PSTN Domestic Calling" }
                "MEETING_ROOM" { "Microsoft Teams Rooms Standard" }
                "MCOMEETADV" { "Microsoft 365 Audio Conferencing" }
                "CCIBOTS_PRIVPREV_VIRAL" { "Power Virtual Agents Viral Trial" }
                "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
                "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
                "SPE_E3" { "Microsoft 365 E3" }
                "SPB" { "Microsoft 365 Business Premium" }
                "ENTERPRISEPACK" { "Office 365 E3" }
                "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
                "PROJECT_P1" { "Project Plan 1" }
                "PROJECTPROFESSIONAL" { "Project Plan 3" }
                "VISIOCLIENT" { "Visio Plan 2" }
                "Microsoft_Teams_Audio_Conferencing_select_dial_out" { "Microsoft Teams Audio Conferencing with dial-out to USA/CAN" }
                "POWER_BI_PRO" { "Power BI Pro" }
                "Microsoft_365_Copilot" { "Microsoft 365 Copilot" }
                "Microsoft_Teams_Premium" { "Microsoft Teams Premium" }
                "MCOEV" { "Microsoft Teams Phone Standard" }
                "POWER_BI_STANDARD" { "Power BI Standard" }
                "Microsoft365_Lighthouse" { "Microsoft 365 Lighthouse" }
                default { $SkuPartNumber }
            }
            return $displayName
        }

        function Get-FormattedLicenseInfo {
            param ([array]$Skus)
            return $Skus | ForEach-Object {
                $available = $_.PrepaidUnits - $_.ConsumedUnits
                $SkuDisplayName = Get-LicenseDisplayName $_.SkuPartNumber
                if ([string]::IsNullOrEmpty($SkuDisplayName)) {
                    $SkuDisplayName = $_.SkuPartNumber
                }
                @{
                    DisplayName = "$($SkuDisplayName) (Available: $available)"
                    SkuId = $_.SkuId
                    SortName = $SkuDisplayName
                }
            } | Sort-Object { $_.SortName }
        }

        # 5. Event Handlers
        $Script:inputGotFocusHandler = {
            if ($this.Text -eq $this.Tag) {
                $this.Text = ""
                $this.Foreground = 'Black'
            }
        }

        $Script:inputLostFocusHandler = {
            if ([string]::IsNullOrWhiteSpace($this.Text) -or $this.Text -eq $this.Tag) {
                $this.Text = $this.Tag
                $this.Foreground = 'Gray'
                $this.BorderBrush = $null
                $this.BorderThickness = 1
                return
            }
            switch -Regex ($this.Name) {
                'newUser|userToCopy' {
                    if (-not (Test-DisplayName $this.Text)) {
                        $this.BorderBrush = 'Red'
                        $this.BorderThickness = 2
                    } else {
                        $this.BorderBrush = $null
                        $this.BorderThickness = 1
                    }
                    break
                }
                'mobile' {
                    if ($this.Text -ne $this.Tag) {
                        $formattedNumber = Format-MobileNumber $this.Text
                        if ($null -eq $formattedNumber) {
                            $this.BorderBrush = 'Red'
                            $this.BorderThickness = 2
                        } else {
                            $this.BorderBrush = $null
                            $this.BorderThickness = 1
                            if (-not $bypassFormattingCheckBox.IsChecked) {
                                $this.Text = $formattedNumber
                            }
                        }
                    }
                    break
                }
            }
        }

        # 6. Input Control Initialization
        function Initialize-InputTextBox {
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

            $textBox.Add_GotFocus($Script:inputGotFocusHandler)
            $textBox.Add_LostFocus($Script:inputLostFocusHandler)

            return $textBox
        }

        # 7. Main UI Creation and Logic
        # Get license information
        $skus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{
            Name = 'PrepaidUnits'; Expression = { $_.PrepaidUnits.Enabled }
        }
        $licenseInfo = Get-FormattedLicenseInfo -Skus $skus

        # Create window and main containers
        $window = New-Object System.Windows.Window
        $window.Title = "New User Request"
        $window.Width = 500
        $window.Height = 800
        $window.WindowStartupLocation = 'CenterScreen'
        $window.Background = '#F0F0F0'

        $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
        $scrollViewer.VerticalScrollBarVisibility = "Auto"
        $mainPanel = New-Object System.Windows.Controls.StackPanel
        $mainPanel.Margin = '10'
        $scrollViewer.Content = $mainPanel
        $window.Content = $scrollViewer

        # Add header
        $mainPanel.Children.Add((New-HeaderPanel -Text "Create New User Request`nPlease fill in all required fields marked with *"))

        # New User section
        $newUserSection = New-FormGroupBox -Header "New User Information"
        $newUserSection.Stack.Children.Add((New-FormLabel -Content "New User Name (First Last) *"))
        $newUserTextBox = Initialize-InputTextBox `
            -Name "newUser" `
            -PlaceholderText "Enter first and last name" `
            -ToolTipText "Enter the full name of the new user (e.g., John Smith)"
        $newUserSection.Stack.Children.Add($newUserTextBox)
        $mainPanel.Children.Add($newUserSection.Group)

        # Template User section
        $copyUserSection = New-FormGroupBox -Header "Template User Information"
        $copyUserSection.Stack.Children.Add((New-FormLabel -Content "User To Copy (First Last) *"))
        $userToCopyTextBox = Initialize-InputTextBox `
            -Name "userToCopy" `
            -PlaceholderText "Enter template user's name" `
            -ToolTipText "Enter the name of an existing user whose permissions should be copied"
        $copyUserSection.Stack.Children.Add($userToCopyTextBox)
        $mainPanel.Children.Add($copyUserSection.Group)

        # Mobile section
        $mobileSection = New-FormGroupBox -Header "Contact Information"
        $mobileSection.Stack.Children.Add((New-FormLabel -Content "Mobile Number"))
        $mobileTextBox = Initialize-InputTextBox `
            -Name "mobile" `
            -PlaceholderText "Enter 10-digit mobile number" `
            -ToolTipText "Enter a 10-digit mobile number (e.g., 1234567890)"
        $mobileSection.Stack.Children.Add($mobileTextBox)

        $bypassPanel = New-Object System.Windows.Controls.DockPanel
        $bypassPanel.Margin = '0,0,0,5'
        $bypassFormattingCheckBox = New-Object System.Windows.Controls.CheckBox
        $bypassFormattingCheckBox.Content = "Bypass Mobile Number Formatting"
        $bypassFormattingCheckBox.ToolTip = "Check this box to skip automatic formatting of the mobile number"
        $bypassFormattingCheckBox.Margin = '0,5,0,5'
        $bypassPanel.Children.Add($bypassFormattingCheckBox)
        $mobileSection.Stack.Children.Add($bypassPanel)
        $mainPanel.Children.Add($mobileSection.Group)

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
            "Microsoft Teams Phone Resource Account",
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
        $requiredGroup.Header = "Required License (Select One) *"
        $requiredGroup.Margin = "0,0,0,10"  # Match other group margins

        $requiredStack = New-Object System.Windows.Controls.StackPanel
        $requiredStack.Margin = "5"  # Match other stack margins

        $requiredComboBox = New-Object System.Windows.Controls.ComboBox
        $requiredComboBox.Margin = "0,0,0,10"  # Match other control margins
        $requiredComboBox.Padding = "5,3,5,3"  # Match other control padding
        $requiredComboBox.ToolTip = "Select one of the required base licenses for the user"

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
        $requiredStack.Children.Add($requiredComboBox)
        $requiredGroup.Content = $requiredStack
        $mainPanel.Children.Add($requiredGroup)

        # Ancillary Licenses Section
        $ancillaryGroup = New-Object System.Windows.Controls.GroupBox
        $ancillaryGroup.Header = "Ancillary Licenses"
        $ancillaryGroup.Margin = "0,0,0,10"  # Match other group margins

        # Create ScrollViewer for ancillary licenses
        $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
        $scrollViewer.VerticalScrollBarVisibility = "Auto"
        $scrollViewer.MaxHeight = 200
        $scrollViewer.Margin = "5"  # Match other stack margins

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
        $buttonPanel.Margin = '0,10,0,0'

        $okButton = New-FormButton -Content "OK" -Margin "0,0,10,0" -ClickHandler {
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

            # Validate Mobile Number only if entered and formatting not bypassed
            if ($mobileTextBox.Text -ne $mobileTextBox.Tag -and -not $bypassFormattingCheckBox.IsChecked) {
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
        }
        $buttonPanel.Children.Add($okButton)

        $cancelButton = New-FormButton -Content "Cancel" -ClickHandler {
            $window.DialogResult = $false
            $window.Close()
        }
        $buttonPanel.Children.Add($cancelButton)

        $mainPanel.Children.Add($buttonPanel)

        # Show the window
        $result = $window.ShowDialog()

        # Initialize formattedMobile variable
        $formattedMobile = $null
        if ($mobileTextBox.Text -ne $mobileTextBox.Tag) {
            # Only process if not placeholder text
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

    function Get-TemplateUser {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]$UserToCopy
        )

        try {
            Write-StatusMessage -Message "Getting template user details" -Type INFO

            $adUserParams = @{
                Filter     = "DisplayName -eq '$UserToCopy'"
                Properties = @(
                    'Title'
                    'Fax'
                    'wWWHomePage'
                    'physicalDeliveryOfficeName'
                    'Office'
                    'Manager'
                    'Description'
                    'Department'
                    'Company'
                )
                ErrorAction = 'Stop'
            }

            $templateUser = Get-ADUser @adUserParams
            if (-not $templateUser) {
                throw [System.InvalidOperationException]::new("Template user not found: $UserToCopy")
            }

            Write-StatusMessage -Message "Successfully retrieved template user details" -Type OK
            return $templateUser
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get template user"
            Add-ErrorType -ErrorType Validation
            throw
        }
    }

    function Test-SourceUser {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateNotNull()]
            [array]$UserToCopyUPN
        )

        try {
            Write-StatusMessage -Message "Validating source user" -Type INFO

            if ($null -eq $UserToCopyUPN) {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Source user not found"
                Add-ErrorType -ErrorType Validation
                Exit-Script -Message "Could not find user to copy from" -ExitCode UserNotFound
            }

            if ($UserToCopyUPN.Count -gt 1) {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Multiple source users found"
                Add-ErrorType -ErrorType Validation
                Exit-Script -Message "Found multiple accounts. Please check AD for duplicate DisplayName attributes" -ExitCode DuplicateUser
            }

            Write-StatusMessage -Message "Source user validated successfully" -Type OK
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Test-SourceUser"
            Add-ErrorType -ErrorType Validation
            throw
        }
    }

    function Confirm-UserCreation {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateNotNull()]
            [hashtable]$UserProperties,

            [Parameter(Mandatory)]
            [ValidateNotNull()]
            [System.Security.SecureString]$Password,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$DestinationOU,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$TemplateUser,

            [Parameter(Mandatory)]
            [string]$PlainPassword
        )

        try {
            $prompt = @"
The user below will be created:
Display Name    = $($UserProperties.FirstName) $($UserProperties.LastName)
Email Address   = $($UserProperties.Email)
Password        = $PlainPassword
First Name      = $($UserProperties.FirstName)
Last Name       = $($UserProperties.LastName)
SamAccountName  = $($UserProperties.SamAccountName)
Destination OU  = $DestinationOU
Template User   = $TemplateUser

Continue? (Y/N)
"@

            Write-StatusMessage -Message "Requesting user confirmation" -Type INFO
            $confirmation = Read-Host -Prompt $prompt

            if ($confirmation -ne 'y') {
                Write-StatusMessage -Message "User cancelled the operation" -Type WARN
                Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
            }

            Write-StatusMessage -Message "User confirmed creation" -Type OK
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Confirm-UserCreation"
            Add-ErrorType -ErrorType Validation
            throw
        }
    }

    function New-UserProperties {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$NewUser,

            [Parameter(Mandatory)]
            [string]$SourceUserUPN
        )

        try {
            # Split the new user name
            $nameParts = $NewUser -split ' '
            if ($nameParts.Count -lt 2) {
                throw [System.ArgumentException]::new("New user name must include first and last name")
            }

            $firstName = $nameParts[0]
            $lastName = $nameParts[-1]
            $displayName = $NewUser

            # Generate email and samAccountName
            try {
                $email = "$($firstName.ToLower()).$($lastName.ToLower())@domain.com"
                $samAccountName = ($firstName.Substring(0,1) + $lastName).ToLower()

                # Handle duplicate samAccountNames
                $counter = 1
                $originalSam = $samAccountName
                while (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
                    $samAccountName = $originalSam + $counter
                    $counter++
                    if ($counter > 99) {
                        throw [System.InvalidOperationException]::new("Unable to generate unique SamAccountName after 99 attempts")
                    }
                }
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to generate user properties"
                Add-ErrorType -ErrorType Validation
                throw
            }

            return @{
                FirstName     = $firstName
                LastName      = $lastName
                DisplayName   = $displayName
                Email        = $email
                SamAccountName = $samAccountName
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in New-UserProperties"
            Add-ErrorType -ErrorType Validation
            throw
        }
    }

    function Test-NewUserExists {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$SamAccountName,

            [Parameter(Mandatory)]
            [string]$Email
        )

        try {
            # Check AD for existing user
            try {
                $adUser = Get-ADUser -Filter "SamAccountName -eq '$SamAccountName' -or UserPrincipalName -eq '$Email'" -ErrorAction Stop
                if ($adUser) {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "User already exists in Active Directory"
                    Add-ErrorType -ErrorType Validation
                    Exit-Script -Message "User already exists: $($adUser.UserPrincipalName)" -ExitCode DuplicateUser
                }
            }
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                # This is expected - user should not exist
                Write-StatusMessage -Message "AD validation passed - user does not exist" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Error checking AD for existing user"
                Add-ErrorType -ErrorType Validation
                throw
            }

            # Check Exchange Online for existing mailbox
            try {
                $mailbox = Get-EXOMailbox -Identity $Email -ErrorAction Stop
                if ($mailbox) {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Mailbox already exists in Exchange Online"
                    Add-ErrorType -ErrorType Validation
                    Exit-Script -Message "Mailbox already exists: $Email" -ExitCode DuplicateUser
                }
            }
            catch [Microsoft.Exchange.Management.RestApiClient.RestApiException] {
                # This is expected - mailbox should not exist
                Write-StatusMessage -Message "Exchange validation passed - mailbox does not exist" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Error checking Exchange for existing mailbox"
                Add-ErrorType -ErrorType Validation
                throw
            }

            Write-StatusMessage -Message "All existence checks passed successfully" -Type OK
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Test-NewUserExists"
            Add-ErrorType -ErrorType Validation
            throw
        }
    }

    function New-SecureRandomPassword {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        param()

        try {
            # Define character sets
            $upperChars = 'ABCDEFGHKLMNOPRSTUVWXYZ'
            $lowerChars = 'abcdefghiklmnoprstuvwxyz'
            $numberChars = '23456789'
            $specialChars = '!@#$%^&*'

            # Generate random characters from each set
            $random = New-Object System.Random
            $password = @(
                ($upperChars.ToCharArray() | Get-Random -Count 2 | ForEach-Object { $_ })
                ($lowerChars.ToCharArray() | Get-Random -Count 3 | ForEach-Object { $_ })
                ($numberChars.ToCharArray() | Get-Random -Count 2 | ForEach-Object { $_ })
                ($specialChars.ToCharArray() | Get-Random -Count 1 | ForEach-Object { $_ })
            )

            # Shuffle the password array
            $shuffledPassword = $password | Get-Random -Count $password.Count
            $finalPassword = -join $shuffledPassword

            # Validate password meets complexity requirements
            if (-not ($finalPassword -cmatch '[A-Z]') -or
                -not ($finalPassword -cmatch '[a-z]') -or
                -not ($finalPassword -match '\d') -or
                -not ($finalPassword -match '[^\w]')) {
                throw [System.Security.SecurityException]::new("Generated password does not meet complexity requirements")
            }

            try {
                $securePassword = ConvertTo-SecureString -String $finalPassword -AsPlainText -Force
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to create secure string from password"
                Add-ErrorType -ErrorType General
                throw
            }

            return @{
                PlainPassword   = $finalPassword
                SecurePassword = $securePassword
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in New-SecureRandomPassword"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Copy-ADGroupMemberships {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$SourceUser,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$TargetUser
        )

        try {
            Write-StatusMessage -Message "Adding AD Groups to new user" -Type INFO

            try {
                $sourceGroups = Get-ADUser -Filter "DisplayName -eq '$SourceUser'" -Properties MemberOf -ErrorAction Stop
                if (-not $sourceGroups) {
                    throw [System.InvalidOperationException]::new("Source user not found: $SourceUser")
                }

                $targetGroups = Get-ADUser -Filter "DisplayName -eq '$TargetUser'" -Properties MemberOf -ErrorAction Stop
                if (-not $targetGroups) {
                    throw [System.InvalidOperationException]::new("Target user not found: $TargetUser")
                }

                $groupsToAdd = $sourceGroups.MemberOf | Where-Object { $targetGroups.MemberOf -notcontains $_ }

                foreach ($group in $groupsToAdd) {
                    try {
                        Add-ADGroupMember -Identity $group -Members $targetGroups -ErrorAction Stop
                        Write-StatusMessage -Message "Added to group: $group" -Type OK
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to add to group: $group"
                        Add-ErrorType -ErrorType Group
                    }
                }

                Write-StatusMessage -Message "AD Groups have been added to new user" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get user group memberships"
                Add-ErrorType -ErrorType Group
                throw
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Copy-ADGroupMemberships"
            Add-ErrorType -ErrorType Group
            throw
        }
    }

    function New-ADUserFromTemplate {
        [CmdletBinding(SupportsShouldProcess)]
        param (
            [Parameter(Mandatory)]
            [hashtable]$UserProperties,
            [Parameter(Mandatory)]
            [Microsoft.ActiveDirectory.Management.ADUser]$UserToCopyUPN,
            [Parameter(Mandatory)]
            [string]$Phone,
            [Parameter(Mandatory)]
            [System.Security.SecureString]$Password,
            [Parameter(Mandatory)]
            [string]$DestinationOU
        )

        $newUserParams = @{
            Name              = "$($UserProperties.FirstName) $($UserProperties.LastName)"
            SamAccountName    = $UserProperties.SamAccountName
            UserPrincipalName = $UserProperties.Email
            DisplayName       = $UserProperties.DisplayName
            GivenName         = $UserProperties.FirstName
            Surname           = $UserProperties.LastName
            MobilePhone       = $Phone
            EmailAddress      = $UserProperties.Email
            Title             = $UserToCopyUPN.Title
            Office            = $UserToCopyUPN.Office
            Manager           = $UserToCopyUPN.Manager
            Description       = $UserToCopyUPN.Description
            Department        = $UserToCopyUPN.Department
            Company           = $UserToCopyUPN.Company
            OtherAttributes   = @{
                'proxyAddresses' = "SMTP:$($UserProperties.Email)"
            }
            AccountPassword   = $Password
            Path              = $DestinationOU
            Instance          = $UserToCopyUPN
            Enabled           = $true
        }

        Write-StatusMessage -Message "Creating new AD user..." -Type INFO
        if ($PSCmdlet.ShouldProcess($UserProperties.DisplayName, "Create new AD user")) {
            New-ADUser @newUserParams
        }
        Write-StatusMessage -Message "AD user created successfully." -Type OK
    }

    # Improve Wait-ForADUserSync error handling
    function Wait-ForADUserSync {
        [CmdletBinding()]
        [OutputType([PSCustomObject])]
        param(
            [Parameter(Mandatory)]
            [string]$UserEmail,

            [Parameter()]
            [int]$MaxRetries = 5,

            [Parameter()]
            [int]$RetryIntervalSeconds = 30,

            [Parameter()]
            [int]$InitialWaitSeconds = 30,

            [Parameter()]
            [int]$SyncTimeout = 300
        )

        Write-StatusMessage -Message "Starting AD sync process for $UserEmail" -Type INFO
        $startTime = Get-Date
        [System.Collections.Generic.List[string]]$syncErrors = @()  # Initialize as a List for better performance

        try {
            # Start AD sync with retry logic
            $syncStarted = $false
            for ($i = 1; $i -le 3; $i++) {
                try {
                    $VerbosePreference = 'SilentlyContinue'
                    Import-Module -Name ADSync -UseWindowsPowerShell -ErrorAction Stop
                    $null = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
                    $syncStarted = $true
                    Write-StatusMessage -Message "AD sync started successfully (Attempt $i)" -Type SUCCESS
                    break
                } catch {
                    $syncErrors.Add("Attempt " + $i + ": " + $_.Exception.Message)
                    if ($i -eq 3) {
                        $errorMessage = "AD sync failed to start after 3 attempts. Errors: " +
                                       ($syncErrors -join " | ")
                        throw [System.TimeoutException]::new($errorMessage)
                    }
                    Write-StatusMessage -Message "Sync attempt $i failed, retrying in 5 seconds..." -Type WARN
                    Start-Sleep -Seconds 5
                }
            }

            # Monitor sync progress with improved error handling
            $retryCount = 0
            do {
                try {
                    # Check timeout
                    $elapsed = ((Get-Date) - $startTime).TotalSeconds
                    if ($elapsed -ge $SyncTimeout) {
                        throw [System.TimeoutException]::new(
                            "Sync timeout after $($elapsed.ToString('F0')) seconds"
                        )
                    }

                    # Check sync status
                    $syncStatus = Get-ADSyncScheduler
                    if ($syncStatus.SyncCycleInProgress) {
                        Write-StatusMessage -Message "Sync in progress..." -Type INFO
                        Start-Sleep -Seconds 10
                        continue
                    }

                    # Try to get user
                    $user = Get-MgUser -UserId $UserEmail -ErrorAction Stop
                    if ($user) {
                        Write-StatusMessage -Message "User $UserEmail successfully synced" -Type SUCCESS
                        return $user
                    }
                } catch [System.TimeoutException] {
                    $response = Read-Host "Sync timeout reached. Extend wait time? (Y/N)"
                    if ($response -eq 'Y') {
                        $startTime = Get-Date
                        continue
                    }
                    throw
                } catch {
                    $retryCount++
                    if ($retryCount -ge $MaxRetries) {
                        throw [System.TimeoutException]::new(
                            "Max retry attempts ($MaxRetries) reached. Last error: $($_.Exception.Message)"
                        )
                    }
                    Write-StatusMessage -Message ("Retry " + $retryCount + "/" + $MaxRetries + ": " + $_.Exception.Message) -Type INFO
                    Start-Sleep -Seconds $RetryIntervalSeconds
                }
            } while ($true)
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "AD sync failed"
            Add-ErrorType -ErrorType Sync
            throw
        }
    }

    function Set-UserLicenses {
        param(
            [Parameter(Mandatory = $true)]
            [string]$UserId,

            [Parameter(Mandatory = $true)]
            [string[]]$License
        )

        try {
            foreach ($License in $Licenses) {
                # Assign license using the required Graph API format
                Set-MgUserLicense -UserId $UserId -AddLicenses @{SkuId = $($License.SkuId) } -RemoveLicenses @() -ErrorAction Stop | Out-Null
                Write-StatusMessage -Message "Successfully assigned license $($License.DisplayName) to user: $UserId" -Type OK
            }
        } catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to assign license"
            Add-ErrorType -ErrorType License
            throw
        }
    }

    function Add-UserToZoom {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
            $MgUser
        )

        try {
            # Determine Zoom role based on department
            if ($MgUser.Department -eq 'Reactive') {
                $zoom_app_role_name = "Basic"
            }
            else {
                $zoom_app_role_name = "Licensed"
            }

            $zoom_app_name = "Zoom Workplace Phones"

            # Get Zoom app and sync details
            try {
                $zoom_ServicePrincipal = Get-MgServicePrincipal -Filter "displayName eq '$zoom_app_name'" -ErrorAction Stop
                if (-not $zoom_ServicePrincipal) {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Zoom Service Principal not found"
                    Add-ErrorType -ErrorType Permission
                    return
                }

                $zoom_synchronizationJob = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $zoom_ServicePrincipal.Id -ErrorAction Stop
                $zoom_synchronizationJobRuleId = (Get-MgServicePrincipalSynchronizationJobSchema -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_synchronizationJob.Id).SynchronizationRules.Id

                Write-StatusMessage -Message "Retrieved Zoom app details successfully" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get Zoom app details"
                Add-ErrorType -ErrorType Permission
                throw
            }

            # Assign user to Zoom app
            try {
                $params = @{
                    "PrincipalId" = $MgUser.Id
                    "ResourceId"  = $zoom_ServicePrincipal.Id
                    "AppRoleId"  = ($zoom_ServicePrincipal.AppRoles | Where-Object { $_.DisplayName -eq $zoom_app_role_name }).Id
                }

                New-MgUserAppRoleAssignment -UserId $MgUser.Id -BodyParameter $params -ErrorAction Stop
                Write-StatusMessage -Message "Successfully assigned Zoom role to user" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to assign Zoom role"
                Add-ErrorType -ErrorType Permission
                throw
            }

            # Start Zoom sync
            try {
                Start-MgServicePrincipalSynchronizationJob -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_synchronizationJob.Id -ErrorAction Stop
                Write-StatusMessage -Message "Successfully started Zoom sync" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to start Zoom sync"
                Add-ErrorType -ErrorType Sync
                throw
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Add-UserToZoom"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Send-GraphMailMessage {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Subject,

            [Parameter(Mandatory)]
            [string]$Content,

            [Parameter()]
            [string]$FromAddress = $config.Email.NotificationFrom,

            [Parameter()]
            [string]$ToAddress = $config.Email.SecurityTeamEmail,

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
            Write-StatusMessage -Message "Preparing to send email notification" -Type INFO

            # Validate parameters
            if ([string]::IsNullOrEmpty($FromAddress) -or [string]::IsNullOrEmpty($ToAddress)) {
                throw [System.ArgumentException]::new("FromAddress and ToAddress are required")
            }

            # Build message
            $messageParams = @{
                Subject      = $Subject
                ToRecipients = @(@{
                    EmailAddress = @{
                        Address = $ToAddress
                    }
                })
                Body         = @{
                    ContentType = $ContentType
                    Content    = $Content
                }
            }

            # Add CC recipients if specified
            if ($CcAddress) {
                $messageParams['CcRecipients'] = @(
                    $CcAddress | ForEach-Object {
                        @{
                            EmailAddress = @{
                                Address = $_
                            }
                        }
                    }
                )
            }

            # Add attachment if specified
            if ($AttachmentPath) {
                try {
                    if (-not (Test-Path $AttachmentPath)) {
                        throw [System.IO.FileNotFoundException]::new("Attachment file not found", $AttachmentPath)
                    }

                    $attachmentContent = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($AttachmentPath))
                    $attachmentName = if ($AttachmentName) { $AttachmentName } else { Split-Path $AttachmentPath -Leaf }

                    $messageParams['Attachments'] = @(
                        @{
                            "@odata.type" = "#microsoft.graph.fileAttachment"
                            Name         = $attachmentName
                            ContentBytes = $attachmentContent
                        }
                    )
                }
                catch {
                    Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to process attachment"
                    Add-ErrorType -ErrorType General
                    throw
                }
            }

            # Send message
            try {
                Send-MgUserMail -UserId $FromAddress -Message $messageParams -ErrorAction Stop
                Write-StatusMessage -Message "Email notification sent successfully" -Type OK
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to send email"
                Add-ErrorType -ErrorType General
                throw
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Send-GraphMailMessage"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Set-UserBookWithMeId {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$UserEmail,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$SamAccountName,

            [Parameter()]
            [ValidateRange(1, 10)]
            [int]$MaxRetries = 5,

            [Parameter()]
            [ValidateRange(30, 300)]
            [int]$RetryIntervalSeconds = 60,

            [Parameter()]
            [ValidateRange(300, 1800)]
            [int]$ProvisioningTimeout = 300  # 5 minutes default timeout
        )

        try {
            Write-StatusMessage -Message "Starting BookWithMeId configuration for $UserEmail" -Type INFO

            # Check mailbox provisioning
            try {
                $startTime = Get-Date
                $mailboxProvisioned = $false
                Write-StatusMessage -Message "Waiting for mailbox provisioning..." -Type INFO

                do {
                    try {
                        $mailbox = Get-Mailbox -Identity $UserEmail -ErrorAction Stop
                        if ($mailbox) {
                            $mailboxProvisioned = $true
                            Write-StatusMessage -Message "Mailbox found for $UserEmail" -Type OK
                            break
                        }
                    }
                    catch {
                        $elapsed = ((Get-Date) - $startTime).TotalSeconds
                        if ($elapsed -ge $ProvisioningTimeout) {
                            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Mailbox provisioning timeout after $($elapsed.ToString('F0')) seconds"
                            Add-ErrorType -ErrorType Mailbox

                            $response = Read-Host "Would you like to extend the wait time? (Y/N)"
                            if ($response -eq 'Y') {
                                $startTime = Get-Date
                                Write-StatusMessage -Message "Extending wait time for mailbox provisioning" -Type INFO
                            }
                            else {
                                Write-StatusMessage -Message "Mailbox provisioning wait cancelled by user" -Type WARN
                                return
                            }
                        }

                        Write-StatusMessage -Message "Mailbox not yet provisioned. Waiting $RetryIntervalSeconds seconds... ($($elapsed.ToString('F0')) seconds elapsed)" -Type INFO
                        Start-Sleep -Seconds $RetryIntervalSeconds
                    }
                } while (-not $mailboxProvisioned)
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to check mailbox provisioning"
                Add-ErrorType -ErrorType Mailbox
                throw
            }

            # Configure BookWithMeId
            try {
                $retryCount = 0
                $success = $false

                do {
                    try {
                        # Get Exchange GUID
                        $exchGuid = $mailbox.ExchangeGuid.Guid
                        if ([string]::IsNullOrEmpty($exchGuid)) {
                            throw [System.InvalidOperationException]::new("Exchange GUID is empty")
                        }

                        # Generate BookWithMeId
                        $extAttr15 = ($exchGuid -replace "-") + '@compassmsp.com?anonymous&ep=plink'
                        if ($extAttr15 -eq '@compassmsp.com?anonymous&ep=plink') {
                            throw [System.InvalidOperationException]::new("Generated BookWithMeId is invalid (missing ExchangeGuid)")
                        }

                        # Set AD attribute
                        Set-ADUser -Identity $SamAccountName -Add @{extensionAttribute15 = $extAttr15} -ErrorAction Stop
                        Write-StatusMessage -Message "Successfully set BookWithMeId: $extAttr15" -Type OK
                        $success = $true
                        break
                    }
                    catch {
                        $retryCount++
                        if ($retryCount -ge $MaxRetries) {
                            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to set BookWithMeId after $MaxRetries attempts"
                            Add-ErrorType -ErrorType General

                            $response = Read-Host "Would you like to retry? (Y/N)"
                            if ($response -eq 'Y') {
                                $retryCount = 0
                                Write-StatusMessage -Message "Retrying BookWithMeId configuration..." -Type INFO
                            }
                            else {
                                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "BookWithMeId configuration skipped - manual intervention required"
                                Add-ErrorType -ErrorType General
                                return
                            }
                        }
                        else {
                            Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed. Waiting $RetryIntervalSeconds seconds..." -Type INFO
                            Start-Sleep -Seconds $RetryIntervalSeconds
                        }
                    }
                } while (-not $success -and $retryCount -lt $MaxRetries)

                if (-not $success) {
                    Write-ErrorRecord `
                        -ErrorRecord ([System.InvalidOperationException]::new("Failed to set BookWithMeId")) `
                        -CustomMessage "Manual intervention required for BookWithMeId configuration" `
                        -ErrorLevel Warning
                    Add-ErrorType -ErrorType General

                    Write-StatusMessage -Message "SamAccountName: $SamAccountName" -Type INFO
                    Write-StatusMessage -Message "Expected BookWithMeId format: {ExchangeGuid}@compassmsp.com?anonymous&ep=plink" -Type INFO
                }
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to configure BookWithMeId"
                Add-ErrorType -ErrorType General
                throw
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Set-UserBookWithMeId"
            Add-ErrorType -ErrorType General
            throw
        }
    }

    function Copy-UserGroups {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]$SourceUserId,

            [Parameter(Mandatory)]
            [string]$TargetUserId,

            [Parameter()]
            [string]$ExportPath
        )

        try {
            Write-StatusMessage -Message "Getting source user group memberships" -Type INFO

            try {
                $sourceGroups = Get-MgUserMemberOf -UserId $SourceUserId -ErrorAction Stop |
                    Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' }

                if (-not $sourceGroups) {
                    Write-StatusMessage -Message "Source user is not a member of any groups" -Type WARN
                    return 0
                }

                Write-StatusMessage -Message "Found $($sourceGroups.Count) groups to copy" -Type OK

                # Export groups if path provided
                if ($ExportPath) {
                    try {
                        $sourceGroups | Export-Csv -Path $ExportPath -NoTypeInformation -ErrorAction Stop
                        Write-StatusMessage -Message "Exported group memberships to: $ExportPath" -Type OK
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to export group memberships"
                        Add-ErrorType -ErrorType General
                    }
                }

                # Copy each group membership
                $successCount = 0
                foreach ($group in $sourceGroups) {
                    try {
                        Add-MgGroupMemberByRef -GroupId $group.Id -OdataId "https://graph.microsoft.com/v1.0/users/$TargetUserId" -ErrorAction Stop
                        Write-StatusMessage -Message "Added to group: $($group.AdditionalProperties.displayName)" -Type OK
                        $successCount++
                    }
                    catch {
                        Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to add to group: $($group.AdditionalProperties.displayName)"
                        Add-ErrorType -ErrorType Group
                    }
                }

                Write-StatusMessage -Message "Successfully copied $successCount of $($sourceGroups.Count) groups" -Type OK
                return $successCount
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to get source user group memberships"
                Add-ErrorType -ErrorType Group
                throw
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Copy-UserGroups"
            Add-ErrorType -ErrorType Group
            throw
        }
    }

    function Start-NewUserFinalize {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$NewUser,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$EmailAddress,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string]$Password,

            [Parameter(Mandatory)]
            [ValidateRange(0, [int]::MaxValue)]
            [int]$TemplateGroupCount,

            [Parameter(Mandatory)]
            [ValidateRange(0, [int]::MaxValue)]
            [int]$AssignedGroupCount
        )

        try {
            Write-StatusMessage -Message "Preparing final summary" -Type INFO

            # Validate group counts
            if ($TemplateGroupCount -ne $AssignedGroupCount) {
                Write-ErrorRecord `
                    -ErrorRecord ([System.InvalidOperationException]::new(
                        "Group count mismatch: Template=$TemplateGroupCount, Assigned=$AssignedGroupCount"
                    )) `
                    -CustomMessage "Group assignment verification failed" `
                    -ErrorLevel Warning
                Add-ErrorType -ErrorType Group
            }

            # Build summary parts
            $summaryParts = @(
                "Summary of Actions:",
                "----------------------------------------",
                "$NewUser should now be created unless any errors occurred during the process.",
                "",
                "User Creation Status:",
                "- Display Name: $NewUser",
                "- Email Address: $EmailAddress",
                "- Password: $Password",
                "- Template Groups: $TemplateGroupCount",
                "- Assigned Groups: $AssignedGroupCount",
                "----------------------------------------",
                "",
                "IMPORTANT: Please record this password now - it will be needed for the user's first login."
            )

            # Add warning for group mismatch
            if ($TemplateGroupCount -ne $AssignedGroupCount) {
                $summaryParts += @(
                    "",
                    "WARNING: Group count mismatch detected",
                    "- Template Groups: $TemplateGroupCount",
                    "- Assigned Groups: $AssignedGroupCount",
                    "Please verify group assignments manually"
                )
            }

            # Display summary
            try {
                $summaryMessage = $summaryParts -join "`n"
                Write-StatusMessage -Message $summaryMessage -Type SUMMARY
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to display summary message"
                Add-ErrorType -ErrorType General
                throw
            }

            # Log completion
            try {
                Write-Log -Message "New user creation completed for $NewUser ($EmailAddress)" -Level SUCCESS
                Write-Log -Message "Groups: $AssignedGroupCount of $TemplateGroupCount assigned" -Level INFO
            }
            catch {
                Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Failed to log completion"
                Add-ErrorType -ErrorType General
            }

            # Exit with appropriate status
            if ($TemplateGroupCount -eq $AssignedGroupCount) {
                Exit-Script -Message "$NewUser has been successfully created" -ExitCode Success
            }
            else {
                Exit-Script -Message "$NewUser created with warnings - please verify group assignments" -ExitCode Success
            }
        }
        catch {
            Write-ErrorRecord -ErrorRecord $_ -CustomMessage "Critical error in Start-NewUserFinalize"
            Add-ErrorType -ErrorType General
            Exit-Script -Message "Failed to complete user creation finalization" -ExitCode GeneralError
        }
    }
    #EndRegion Functions

    Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green

    Write-Host "`n  Script Ready!" -ForegroundColor Cyan
    Write-Host "  Press any key to continue..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Clear-Host

    #Region Main Execution

    try {
        # Step 0: Initialization
        Write-ProgressStep -StepName $progressSteps[0].Name -Status $progressSteps[0].Description
        $config = Get-ScriptConfig

        # These variables should be set before Connect-ServiceEndpoints
        $Organization = $config.ExchangeOnline.Organization
        $ExOAppId = $config.ExchangeOnline.AppId
        $ExOCertSubject = $Config.ExchangeOnline.CertificateSubject
        $GraphAppId = $config.Graph.AppId
        $tenantID = $config.Graph.TenantId
        $GraphCertSubject = $Config.Graph.CertificateSubject
        $PnPAppId = $config.PnPSharePoint.AppId
        $PnPUrl = $config.PnPSharePoint.Url
        $PnPClientSecret = $config.PnPSharePoint.ClientSecret

        # After loading config
        if (-not $config.ExchangeOnline -or -not $config.Graph) {
            Exit-Script -Message "Missing required configuration sections" -ExitCode ConfigError
        }

        Write-StatusMessage -Message "Connecting to required services..." -Type INFO
        Connect-ServiceEndpoints

        # Step 1: User Input
        Write-ProgressStep -StepName $progressSteps[1].Name -Status $progressSteps[1].Description
        Write-StatusMessage -Message "Opening new user request window..." -Type INFO
        $result = Show-NewUserRequestWindow
        if (-not $result) {
            Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
        }

        # Set variables from window function result
        $NewUser = $result.InputNewUser
        $Phone = $result.InputNewMobile
        $UserToCopy = $result.InputUserToCopy
        $RequiredLicense = $result.InputRequiredLicense
        $AncillaryLicenses = $result.InputAncillaryLicenses

        # Step 2: Validation
        Write-ProgressStep -StepName $progressSteps[2].Name -Status $progressSteps[2].Description
        $UserToCopyUPN = Get-TemplateUser -UserToCopy $UserToCopy
        Test-SourceUser -UserToCopyUPN $UserToCopyUPN
        $userProperties = New-UserProperties -NewUser $NewUser -SourceUserUPN $UserToCopyUPN.UserPrincipalName
        Test-NewUserExists -SamAccountName $userProperties.SamAccountName -Email $userProperties.Email

        # Step 3: AD Creation
        Write-ProgressStep -StepName $progressSteps[3].Name -Status $progressSteps[3].Description
        $passwordResult = New-SecureRandomPassword
        $destinationOU = $UserToCopyUPN.DistinguishedName.split(",", 2)[1]
        Confirm-UserCreation -UserProperties $userProperties -Password $passwordResult.SecurePassword -DestinationOU $destinationOU -TemplateUser $UserToCopy -PlainPassword $passwordResult.PlainPassword
        New-ADUserFromTemplate -UserProperties $userProperties -UserToCopyUPN $UserToCopyUPN -Phone $Phone -Password $passwordResult.SecurePassword -DestinationOU $destinationOU
        $finalPassword = $passwordResult.PlainPassword

        # Step 4: AD Group Copy
        Write-ProgressStep -StepName $progressSteps[4].Name -Status $progressSteps[4].Description
        Copy-ADGroupMemberships -SourceUser $UserToCopy -TargetUser "$($userProperties.FirstName) $($userProperties.LastName)"

        # Step 5: Azure Sync
        Write-ProgressStep -StepName $progressSteps[5].Name -Status $progressSteps[5].Description
        $MgUser = Wait-ForADUserSync -UserEmail $userProperties.Email

        if (-not $MgUser) {
            Write-StatusMessage -Message "Failed to get user from Azure AD after sync" -Type 'ERROR'
            Exit-Script -Message "Azure AD sync completed but user was not found" -ExitCode GeneralError
        }

        try {
            # Step 6: License Setup
            Write-ProgressStep -StepName $progressSteps[6].Name -Status $progressSteps[6].Description
            Write-StatusMessage -Message "Setting Usage Location for new user" -Type INFO
            Update-MgUser -UserId $MgUser.Id -UsageLocation US
            Write-StatusMessage -Message "Usage Location has been set for new user" -Type OK

            Set-UserLicenses -UserId $MgUser.Id -License $RequiredLicense

            Start-Sleep -Seconds 60

            if ($null -ne $AncillaryLicenses) {
                Set-UserLicenses -UserId $MgUser.Id -License $AncillaryLicenses
            }

            # Step 7: Group Setup
            Write-ProgressStep -StepName $progressSteps[7].Name -Status $progressSteps[7].Description
            $MgUserCopy = Get-MgUser -UserId $UserToCopyUPN.UserPrincipalName
            Copy-UserGroups -SourceUserId $MgUserCopy.Id -TargetUserId $MgUser.Id
            $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $MgUserCopy.Id).Count
            $NewUserGroupCount = (Get-MgUserMemberOf -UserId $MgUser.Id).Count

            # Step 8: Email Setup
            Write-ProgressStep -StepName $progressSteps[8].Name -Status $progressSteps[8].Description
            $emailSubject = "KB4 – New User"
            $emailContent = "The following user need to be added to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
            $MsgFrom = $config.Email.NotificationFrom
            $ToAddress = $config.Email.NotificationTo
            Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -Subject $emailSubject -Content $emailContent
            Set-UserBookWithMeId -UserEmail $userProperties.Email -SamAccountName $userProperties.SamAccountName

            # Step 9: Zoom Setup
            Write-ProgressStep -StepName $progressSteps[9].Name -Status $progressSteps[9].Description
            Add-UserToZoom -MgUser $MgUser

            # Step 10: OneDrive Setup
            Write-ProgressStep -StepName $progressSteps[10].Name -Status $progressSteps[10].Description
            Write-StatusMessage -Message "Provisioning OneDrive for new user." -Type INFO

            <# This is not working as expected.
            try {
                Get-MgUserDefaultDrive -UserId $MgUser.Id -ErrorAction Stop
                Write-StatusMessage -Message "OneDrive has been provisioned for new user." -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to provision OneDrive: $_" -Type ERROR
            }
            #>

            # Step 11: Cleanup and Summary
            Write-ProgressStep -StepName $progressSteps[11].Name -Status $progressSteps[11].Description
            Write-StatusMessage -Message "Disconnecting from Exchange Online and Graph." -Type INFO

            Connect-ServiceEndpoints -Disconnect

            Write-StatusMessage -Message "Building final summary..." -Type INFO

            Start-NewUserFinalize -NewUser $NewUser `
                -EmailAddress $userProperties.Email `
                -Password $finalPassword `
                -TemplateGroupCount $CopyUserGroupCount `
                -AssignedGroupCount $NewUserGroupCount

        } catch {
            Write-StatusMessage -Message "Failed to update user properties: $($_.Exception.Message)" -Type 'ERROR'
            Exit-Script -Message "Failed to configure user in Azure AD" -ExitCode GeneralError
        }

        # Clear the progress bar when done
        Write-Progress -Activity "New User Creation" -Completed

    } catch {
        Write-Progress -Activity "New User Creation" -Completed
        Exit-Script -Message "Failed to complete user creation process: $_" -ExitCode GeneralError
    }

    # At end of script
    if ($script:errorCount -gt 0) {
        Write-StatusMessage "Completed with $script:errorCount errors" -Type WARN
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-StatusMessage "Script completed in $($duration.TotalMinutes.ToString('F2')) minutes" -Type INFO

    # Before Exit-Script, add error summary if there were errors
    if ($script:errorCount -gt 0) {
        Write-StatusMessage -Message (Get-ErrorSummary) -Type SUMMARY
    }

} finally {
    Stop-Job $loadingJob | Out-Null
    Remove-Job $loadingJob | Out-Null
}