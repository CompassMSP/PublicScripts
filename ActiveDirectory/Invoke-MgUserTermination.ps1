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

$script:TestMode = $false  # Default to false

# Disable console quick edit
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

Set-ConsoleProperties -QuickEditMode Disable -InsertMode Disable

# Initialize loading animation
Clear-Host
$loadingChars = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
$i = 0
$loadingJob = Start-Job -ScriptBlock { while ($true) { Start-Sleep -Milliseconds 100 } }

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
    @{ Number = 1; Name = "Initialization"; Description = "Loading configuration and connecting services" }
    @{ Number = 2; Name = "User Input"; Description = "Gathering termination details" }
    @{ Number = 3; Name = "AD Tasks"; Description = "Disabling user in Active Directory" }
    @{ Number = 4; Name = "Session Cleanup"; Description = "Removing user sessions and devices" }
    @{ Number = 5; Name = "Exchange Tasks"; Description = "Convert to SharedMailbox and setting forwarding/grant acces" }
    @{ Number = 6; Name = "Directory Roles"; Description = "Removing from directory roles" }
    @{ Number = 7; Name = "Group Removal"; Description = "Removing and exporting Entra/Exchange groups" }
    @{ Number = 8; Name = "License Removal"; Description = "Removing and exporting Entra licenses" }
    @{ Number = 9; Name = "Remove Zoom SSO"; Description = "Removing user from Zoom SSO" }
    @{ Number = 10; Name = "Notifications"; Description = "Sending SecurePath Offboarding notifications" }
    @{ Number = 11; Name = "Disconnecting from Exchange and Graph"; Description = "Disconnecting from Exchange and Graph" }
    @{ Number = 12; Name = "OneDrive Setup"; Description = "Configuring OneDrive access" }
    @{ Number = 13; Name = "Final Steps"; Description = "Running AD sync and finalizing" }
)
Write-Host "`r  [✓] Progress tracking initialized" -ForegroundColor Green

Write-Host "  [$($loadingChars[$i % $loadingChars.Length])] Loading functions..." -NoNewline -ForegroundColor Yellow
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
        Write-Progress -Activity "User Termination" -Status $Status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $Status" -Type INFO
        Write-Progress -Activity "User Termination" -Status $Status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
    }
}

#Region Standard Functions

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
        } else {
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $statusPadded = $config[$Type].Status.PadRight(7)
            Write-Host "[$timestamp] [$statusPadded] $Message" -ForegroundColor $config[$Type].Color
        }
    } catch {
        Write-Host "Failed to write status message: $_" -ForegroundColor Red
    }
}

function Exit-Script {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
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

        # Only disconnect if not in test mode
        if (-not $script:TestMode) {
            Write-StatusMessage -Message "Disconnecting from services..." -Type INFO
            try {
                Connect-ServiceEndpoints -Disconnect
            } catch {
                Write-StatusMessage -Message "Failed to disconnect services during exit" -Type ERROR
            }
        }

        # Log the exit message
        Write-StatusMessage -Message $Message -Type $messageTypes[$ExitCode]

        # In test mode, don't actually exit
        if ($script:TestMode) {
            Write-StatusMessage -Message "Test Mode: Script would exit here with code $($exitCodes[$ExitCode])" -Type WARN
            return
        }

        # Return the appropriate exit code
        exit $exitCodes[$ExitCode]
    } catch {
        # Catch-all for any unexpected errors during exit
        Write-StatusMessage -Message "Critical error during script exit" -Type ERROR
        if (-not $script:TestMode) {
            exit 99
        }
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
            TestMode       = @{
                Email = Read-Host "Enter test email address for development"
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

    # If Disconnect is specified, handle disconnections
    if ($Disconnect) {
        Write-StatusMessage -Message "Disconnecting from services..." -Type INFO
        # If no specific services selected, set all to true for disconnecting everything
        $disconnectAll = -not ($ExchangeOnline -or $Graph -or $SharePoint)

        # Disconnect from Exchange Online
        if (($ExchangeOnline -or $disconnectAll) -and (Get-ConnectionInformation)) {
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                Write-StatusMessage -Message "Disconnected from Exchange Online" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to disconnect from Exchange Online: $_" -Type WARN
            }
        }

        # Disconnect from Microsoft Graph
        if (($Graph -or $disconnectAll) -and (Get-MgContext)) {
            try {
                $null = Disconnect-MgGraph -ErrorAction Stop
                Write-StatusMessage -Message "Disconnected from Microsoft Graph" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to disconnect from Microsoft Graph: $_" -Type WARN
            }
        }

        # Disconnect from SharePoint
        if (($SharePoint -or $disconnectAll)) {
            try {
                # Try to disconnect only if there's an active connection
                try {
                    $pnpContext = Get-PnPContext -ErrorAction Stop
                    if ($pnpContext) {
                        Disconnect-PnPOnline -ErrorAction Stop
                        Write-StatusMessage -Message "Disconnected from SharePoint Online" -Type OK
                    }
                } catch {
                }
            } catch {
                Write-StatusMessage -Message "Failed to disconnect from SharePoint Online: $_" -Type WARN
            }
        }

        return
    }

    # If not disconnecting, handle connections
    if (-not $Disconnect) {
        # If no specific services selected, connect to all
        $connectAll = -not ($ExchangeOnline -or $Graph -or $SharePoint)
        Write-StatusMessage -Message "Connecting to services..." -Type INFO

        # Connect to Exchange Online
        if ($ExchangeOnline -or $connectAll) {
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

        # Connect to Microsoft Graph
        if ($Graph -or $connectAll) {
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

        # Connect to SharePoint Online
        if ($SharePoint -or $connectAll) {
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
        # If in test mode, override the ToAddress
        if ($script:TestMode) {
            Write-StatusMessage -Message "TEST MODE: Redirecting email to $script:TestEmailAddress" -Type WARN
            $ToAddress = $script:TestEmailAddress
            $CcAddress = @() # Clear CC in test mode
        }

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

#Region Custom Functions

function Get-UserTerminationInput {
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
        TestModeEnabled       : [bool] Whether test mode is enabled
    Returns $null if the user cancels the operation.
    #>

    # 1. Add required assemblies
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    # Initialize test mode if not already set
    if ($null -eq $script:TestMode) {
        $script:TestMode = $false
    }

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

    # Add test mode checkbox to button panel
    $testModeButton = New-FormCheckBox `
        -Content "Test Mode" `
        -ToolTip "Enable to redirect emails to: $($config.TestMode.Email)" `
        -IsChecked ($script:TestMode -eq $true) `
        -Margin "0,5,10,0"

    $buttonPanel.Children.Add($testModeButton)

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
            TestModeEnabled         = ($testModeButton.IsChecked -eq $true)
        }
    } else {
        return $null
    }
}

function Get-TerminationPrerequisites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$User,

        [Parameter()]
        [string]$UserPropertiesPath,

        [Parameter()]
        [string]$ADGroupsPath
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
            $MgUser = Get-MgUser -UserId $UserFromAD.UserPrincipalName -Property Id, Mail, DisplayName, Department | Select-Object Id, Mail, DisplayName, Department -ErrorAction Stop
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

        # After confirmation, export properties and groups if paths provided
        if ($UserPropertiesPath -or $ADGroupsPath) {
            try {
                # Export user properties
                if ($UserPropertiesPath) {
                    Write-StatusMessage -Message "Exporting user properties" -Type INFO
                    $propertyList = @(
                        'displayName',
                        'SamAccountName',
                        'UserPrincipalName',
                        'mail',
                        'proxyAddresses',
                        'company',
                        'Title',
                        'Manager',
                        'physicalDeliveryOfficeName',
                        'Department',
                        'l',
                        'c',
                        'facsimileTelephoneNumber',
                        'mobile',
                        'telephoneNumber',
                        'DistinguishedName',
                        'extensionAttribute1',
                        'extensionAttribute2',
                        'extensionAttribute3',
                        'extensionAttribute4',
                        'extensionAttribute5',
                        'extensionAttribute6',
                        'extensionAttribute7',
                        'extensionAttribute8',
                        'extensionAttribute9',
                        'extensionAttribute10',
                        'extensionAttribute11',
                        'extensionAttribute12',
                        'extensionAttribute13',
                        'extensionAttribute14',
                        'extensionAttribute15'
                    )
                    $userProperties = Get-ADUser -Identity $UserFromAD.SamAccountName -Properties $propertyList
                    $userProperties | Select-Object $propertyList | Export-Csv -Path $UserPropertiesPath -NoTypeInformation
                    Write-StatusMessage -Message "Exported user properties to: $UserPropertiesPath" -Type OK
                }

                # Export group memberships
                if ($ADGroupsPath) {
                    Write-StatusMessage -Message "Exporting group memberships" -Type INFO
                    $groups = $UserFromAD.MemberOf | ForEach-Object {
                        Get-ADGroup -Identity $_ -Properties Name, Description, GroupCategory, GroupScope |
                            Select-Object Name, Description, GroupCategory, GroupScope, DistinguishedName
                    }
                    $groups | Export-Csv -Path $ADGroupsPath -NoTypeInformation
                    Write-StatusMessage -Message "Exported group memberships to: $ADGroupsPath" -Type OK
                }
            }
            catch {
                Write-StatusMessage -Message "Failed to export user data: $_" -Type ERROR
            }
        }

        # Return all the collected information
        return @{
            UserFromAD    = $UserFromAD
            DestinationOU = $DestinationOU
            Mailbox       = $365Mailbox
            MgUser        = $MgUser
            UserPropertiesPath = $UserPropertiesPath
            ADGroupsPath = $ADGroupsPath
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
                    'Manager',
                    'physicalDeliveryOfficeName',
                    'Department',
                    'facsimileTelephoneNumber',
                    'l', # l is for Location because Microsoft AD attributes are stupid
                    'c', # c is for Country because Microsoft AD attributes are stupid
                    'wWWHomePage'
                    'mobile',
                    'telephoneNumber',
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
            Write-StatusMessage -Message "Failed to disable user account" -Type ERROR
            throw
        }

        # Remove user from all AD groups
        foreach ($group in $UserFromAD.MemberOf) {
            Write-StatusMessage -Message "Removing user from group: $($group)" -Type INFO
            try {
                Remove-ADGroupMember -Identity $group -Members $UserFromAD.SamAccountName -Confirm:$false -ErrorAction Stop
                Write-StatusMessage -Message "Successfully removed from group: $($group)" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to remove from AD group: $group" -Type ERROR
            }
        }
        Write-StatusMessage -Message "User removed from all AD groups" -Type OK

        # Move user to disabled OU
        Write-StatusMessage -Message "Moving user to Disabled OU" -Type INFO
        try {
            $UserFromAD | Move-ADObject -TargetPath $DestinationOU -ErrorAction Stop
            Write-StatusMessage -Message "Successfully moved user to disabled OU" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to move user to disabled OU" -Type ERROR
            throw
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Disable-ADUser" -Type ERROR
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
        }
        catch {
            Write-StatusMessage -Message "Failed to revoke user sessions" -Type ERROR
        }

        # Remove authentication methods
        Write-StatusMessage -Message "Removing user authentication methods" -Type INFO
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $UserPrincipalName -ErrorAction Stop

            foreach ($authMethod in $authMethods) {
                $authType = $authMethod.AdditionalProperties.'@odata.type'

                try {
                    switch ($authType) {
                        "#microsoft.graph.passwordAuthenticationMethod" {
                            continue
                        }
                        "#microsoft.graph.phoneAuthenticationMethod" {
                            Remove-MgUserAuthenticationPhoneMethod -UserId $UserPrincipalName -PhoneAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Phone Authentication Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                            Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $UserPrincipalName -WindowsHelloForBusinessAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Windows Hello for Business Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                            Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $UserPrincipalName -MicrosoftAuthenticatorAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Microsoft Authenticator Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.fido2AuthenticationMethod" {
                            Remove-MgUserAuthenticationFido2Method -UserId $UserPrincipalName -Fido2AuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed FIDO2 Authenticator Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.softwareOathAuthenticationMethod" {
                            Remove-MgUserAuthenticationSoftwareOathMethod -UserId $UserPrincipalName -SoftwareOathAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Software Oath Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                            Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $UserPrincipalName -TemporaryAccessPassAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Temporary Access Pass Method: $($authMethod.Id)" -Type OK
                        }
                        default {
                            Write-StatusMessage -Message "Skipping unknown authentication method: $authType" -Type ERROR
                        }
                    }
                }
                catch {
                    Write-StatusMessage -Message "Failed to remove authentication method $($authMethod.Id) of type $authType" -Type ERROR
                }
            }
        }
        catch {
            Write-StatusMessage -Message "Failed to get user authentication methods" -Type ERROR
        }

        # Remove Mobile Devices
        Write-StatusMessage -Message "Removing all mobile devices" -Type INFO
        try {
            $mobileDevices = Get-MobileDevice -Mailbox $UserPrincipalName -ErrorAction Stop
            foreach ($mobileDevice in $mobileDevices) {
                Write-StatusMessage -Message "Removing mobile device: $($mobileDevice.Id)" -Type INFO
                try {
                    Remove-MobileDevice -Identity $mobileDevice.Id -Confirm:$false -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully removed mobile device: $($mobileDevice.Id)" -Type OK
                }
                catch {
                    Write-StatusMessage -Message "Failed to remove mobile device $($mobileDevice.Id)" -Type ERROR
                }
            }
        }
        catch {
            Write-StatusMessage -Message "Failed to get mobile devices" -Type ERROR
        }

        # Disable Azure AD devices
        try {
            $termUserDevices = Get-MgUserRegisteredDevice -UserId $UserPrincipalName -ErrorAction Stop
            foreach ($termUserDevice in $termUserDevices) {
                Write-StatusMessage -Message "Disabling registered device: $($termUserDevice.Id)" -Type INFO
                try {
                    Update-MgDevice -DeviceId $termUserDevice.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully disabled device: $($termUserDevice.Id)" -Type OK
                }
                catch {
                    Write-StatusMessage -Message "Failed to disable device $($termUserDevice.Id)" -Type ERROR
                }
            }
        }
        catch {
            Write-StatusMessage -Message "Failed to get registered devices" -Type ERROR
        }

    } catch {
        Write-StatusMessage -Message "Critical error in Remove-UserSessions" -Type ERROR
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
            Write-StatusMessage -Message "Failed to disable existing mailbox forwarding" -Type ERROR
        }

        # Change mailbox to shared
        Write-StatusMessage -Message "Converting to shared mailbox" -Type INFO
        try {
            Set-Mailbox -Identity $Mailbox.Identity -Type Shared -ErrorAction Stop
            Write-StatusMessage -Message "Successfully converted to shared mailbox" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to convert to shared mailbox" -Type ERROR
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
                Write-StatusMessage -Message "Failed to set up mail forwarding" -Type ERROR
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

                $null = Add-MailboxPermission @mailboxPermissionParams
                Write-StatusMessage -Message "Successfully granted full access permissions" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to grant mailbox permissions" -Type ERROR
            }
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Set-TerminatedMailbox" -Type ERROR
        throw
    }
}

function Remove-UserFromEntraDirectoryRoles {
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
                    Write-StatusMessage -Message "Failed to remove from role $roleName" -Type ERROR
                }
            }
        } catch {
            Write-StatusMessage -Message "Failed to get directory roles" -Type ERROR
            throw
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Remove-UserFromDirectoryRoles" -Type ERROR
        throw
    }
}

function Remove-UserFromEntraGroups {
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
                # Not a directory role
                $_.AdditionalProperties['@odata.type'] -ne '#microsoft.graph.directoryRole' -and
                # Not a dynamic group
                $null -eq $_.AdditionalProperties.membershipRule -and
                # Only sync-enabled groups (not false)
                $null -eq $_.AdditionalProperties.onPremisesSyncEnabled
            }
        }

        # Define select parameters with a custom groupType classification
        $selectParams = @{
            Property = @(
                'Id'
                @{n = 'DisplayName'; e = { $_.AdditionalProperties.displayName } }
                @{n = 'Mail'; e = { $_.AdditionalProperties.mail } }
                @{n = 'groupType'; e = {
                    if ($_.AdditionalProperties.securityEnabled -eq $true) {
                        return "Security"
                    } elseif ($_.AdditionalProperties.groupTypes -contains "Unified") {
                        return "Unified"
                    } else {
                        return "Distribution"
                    }
                }}
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
                    Write-StatusMessage -Message "Failed to export user groups" -Type ERROR
                }
            }

            foreach ($365Group in $All365Groups) {
                Write-StatusMessage -Message "Processing group: $($365Group.DisplayName)" -Type INFO

                try {
                    if ($365Group.securityEnabled -eq $true -or $365Group.groupType -eq 'Unified') {
                        Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $userId -ErrorAction Stop
                        Write-StatusMessage -Message "Removed from Security/Unified Group: $($365Group.DisplayName)" -Type OK
                    } else {
                        Remove-DistributionGroupMember -Identity $365Group.Id -Member $userId -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                        Write-StatusMessage -Message "Removed from Distribution Group: $($365Group.DisplayName)" -Type OK
                    }
                } catch {
                    Write-StatusMessage -Message "Failed to remove from group $($365Group.DisplayName)" -Type ERROR
                }
            }
        } catch {
            Write-StatusMessage -Message "Failed to get user group memberships" -Type ERROR
            throw
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Remove-UserFromGroups" -Type ERROR
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
                    Write-StatusMessage -Message "Failed to export user licenses" -Type ERROR
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
                    $null = Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction Stop
                    Write-StatusMessage -Message "Removed Ancillary License: $($license.SkuPartNumber)" -Type OK
                } catch {
                    Write-StatusMessage -Message "Failed to remove Ancillary License $($license.SkuPartNumber)" -Type ERROR
                }
            }

            # Step 2: Remove Primary Licenses
            foreach ($license in ($licenseDetails | Where-Object { $_.SkuPartNumber -in $primaryLicenses })) {
                try {
                    $null = Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction Stop
                    Write-StatusMessage -Message "Removed Primary License: $($license.SkuPartNumber)" -Type OK
                } catch {
                    Write-StatusMessage -Message "Failed to remove Primary License $($license.SkuPartNumber)" -Type ERROR
                }
            }
        } catch {
            Write-StatusMessage -Message "Failed to get user licenses" -Type ERROR
            throw
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Remove-UserLicenses" -Type ERROR
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
                    Write-StatusMessage -Message "Failed to remove Zoom assignment" -Type ERROR
                }
            } else {
                Write-StatusMessage -Message "User is not assigned to Zoom Workplace Phones" -Type INFO
            }
        } catch {
            Write-StatusMessage -Message "Failed to get Zoom assignments" -Type ERROR
            throw
        }
    } catch {
        Write-StatusMessage -Message "Critical error in Remove-UserFromZoom" -Type ERROR
        throw
    }
}

function Set-TerminatedOneDrive {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TermUser,

        [Parameter()]
        [switch]$SetReadOnly,

        [Parameter()]
        [string]$OneDriveUser
    )

    # Skip if no actions are requested
    if (-not $SetReadOnly -and -not $OneDriveUser) {
        Write-StatusMessage -Message "No OneDrive actions requested. Skipping OneDrive configuration." -Type INFO
        return
    }

    try {
        # Ensure SharePoint connection
        Write-StatusMessage -Message "Connecting to SharePoint Online..." -Type INFO
        Connect-ServiceEndpoints -SharePoint

        # Get OneDrive URL
        Write-StatusMessage -Message "Getting OneDrive URL for $TermUser" -Type INFO
        $onedriveProps = Get-PnPUserProfileProperty -Account $TermUser -Properties PersonalUrl -ErrorAction Stop

        if (-not $onedriveProps.PersonalUrl) {
            Write-StatusMessage -Message "No OneDrive URL found for $TermUser. Skipping OneDrive configuration." -Type WARN
            return
        }

        $UserOneDriveURL = $onedriveProps.PersonalUrl
        Write-StatusMessage -Message "Successfully retrieved OneDrive URL" -Type OK

        # Handle OneDrive access grant if specified
        if ($OneDriveUser) {
            try {
                Write-StatusMessage -Message "Granting OneDrive access to $OneDriveUser" -Type INFO
                Set-PnPTenantSite -Url $UserOneDriveURL -Owners $OneDriveUser -ErrorAction Stop
                Write-StatusMessage -Message "Successfully granted OneDrive access" -Type OK
                Write-StatusMessage -Message "OneDrive URL: $UserOneDriveURL" -Type INFO

                do {
                    $response = Read-Host "Please copy the OneDrive URL above. Have you copied it? (y/n)"
                } while ($response -notmatch '^[yY]$')

            } catch {
                Write-StatusMessage -Message "Failed to grant OneDrive access: $_" -Type ERROR
            }
        }

        # Handle read-only setting if specified
        if ($SetReadOnly) {
            try {
                Write-StatusMessage -Message "Setting OneDrive to read-only" -Type INFO
                Set-PnPTenantSite -Url $UserOneDriveURL -LockState ReadOnly -ErrorAction Stop
                Write-StatusMessage -Message "Successfully set OneDrive to read-only" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to set OneDrive to read-only: $_" -Type ERROR
            }
        }

    } catch {
        Write-StatusMessage -Message "Error processing OneDrive configuration: $_" -Type ERROR
    } finally {
        # Disconnect from SharePoint if we connected
        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
            Write-StatusMessage -Message "Disconnecting from SharePoint Online..." -Type INFO
            Connect-ServiceEndpoints -SharePoint -Disconnect
        }
    }
}

function Start-ADSyncAndFinalize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

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

    # Start AD Sync
    Write-StatusMessage -Message "Starting AD sync cycle" -Type INFO
    try {
        Import-Module -Name ADSync -UseWindowsPowerShell -WarningAction:SilentlyContinue -ErrorAction Stop
        $null = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        Write-StatusMessage -Message "AD sync cycle initiated successfully" -Type OK
    } catch {
        try {
            # Fallback to direct PowerShell execution if module import fails
            $null = powershell.exe -command Start-ADSyncSyncCycle -PolicyType Delta
            Write-StatusMessage -Message "AD sync cycle initiated through PowerShell" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to start AD sync cycle" -Type ERROR
            throw
        }
    }

    # Build summary parts
    $summaryParts = @(
        "Summary of Actions:",
        "----------------------------------------",
        "$($User.displayName) should now be offboarded unless any errors occurred during the process.",
        "If any info below is blank then something went wrong in the script. ",
        "User Termination Status:",
        "- EntraID: $($User.Id)",
        "- Display Name: $($User.displayName)",
        "- Email Address: $($User.mail)",
        "- Moved to OU: $DestinationOU"
    )

    if ($GrantUserFullControl) {
        $summaryParts += "- Mailbox access granted to: $GrantUserFullControl"
    }
    if ($SetUserMailFWD) {
        $summaryParts += "- Mail forwarded to: $SetUserMailFWD"
    }
    if ($GrantUserOneDriveAccess) {
        $summaryParts += "- OneDrive access granted to: $GrantUserOneDriveAccess"
    }
    if ($ExportPath) {
        $summaryParts += "- Groups and licenses exported to: $ExportPath"
    }

    $summaryParts += "----------------------------------------"
    $summaryMessage = $summaryParts -join "`n"

    Write-StatusMessage -Message $summaryMessage -Type SUMMARY

    # Show duration
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-StatusMessage "Script completed in $($duration.TotalMinutes.ToString('F2')) minutes" -Type INFO

    # Give user time to read/copy the summary
    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
    Read-Host | Out-Null

    # Clear the progress bar
    Write-Progress -Activity "User Termination" -Completed

    # Clean up and exit
    Stop-Job $loadingJob | Out-Null
    Remove-Job $loadingJob | Out-Null


    Exit-Script -Message "$User has been successfully disabled." -ExitCode Success
}

Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green
Write-Host "`n  Ready to create new user account..." -ForegroundColor Cyan

#Region Main Execution


# Step 1: Initialization
Write-ProgressStep -StepName $progressSteps[1].Name -Status $progressSteps[1].Description

# Load configuration
$config = Get-ScriptConfig

$script:TestEmailAddress = $config.TestMode.Email

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

Connect-ServiceEndpoints -ExchangeOnline -Graph

# Call the custom input window function

$result = Get-UserTerminationInput
if (-not $result) {
    Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
}

if ($result.TestModeEnabled -eq 'True') { $script:TestMode = $true }

# Set variables with validation
$User = $result.InputUser.Trim()

# Optional fields - use $null for unset values
$GrantUserFullControl = if ($result.InputUserFullControl) { $result.InputUserFullControl.Trim() } else { $null }
$SetUserMailFWD = if ($result.InputUserFWD) { $result.InputUserFWD.Trim() } else { $null }
$GrantUserOneDriveAccess = if ($result.InputUserOneDriveAccess) { $result.InputUserOneDriveAccess.Trim() } else { $null }
$SetOneDriveReadOnly = $result.SetOneDriveReadOnly

Write-StatusMessage -Message "Termination input collected successfully" -Type OK

# Validate OneDrive access user if specified
$oneDriveUser = $null
if ($GrantUserOneDriveAccess) {
    # Will be $null if not provided
    try {
        Write-StatusMessage -Message "Validating OneDrive access user..." -Type 'INFO'
        $oneDriveUser = Get-Mailbox $GrantUserOneDriveAccess -ErrorAction Stop
        Write-StatusMessage -Message "OneDrive access user validated" -Type 'OK'
    } catch {
        Write-StatusMessage -Message "Invalid OneDrive access user specified: $_" -Type 'ERROR'
        Write-StatusMessage -Message "OneDrive access user validation failed. Skipping OneDrive access grant." -Type 'ERROR'
        $GrantUserOneDriveAccess = $null  # Reset to null if validation fails
    }
}

# Step 2: User Input
Write-ProgressStep -StepName $progressSteps[2].Name -Status $progressSteps[2].Description
# Should this be a function in the New User Script?
$userPropertiesPath = Join-Path $localExportPath "$($User)_Properties.csv"
$adGroupsPath = Join-Path $localExportPath "$($User)_ADGroups.csv"
$UserInfo = Get-TerminationPrerequisites `
    -User $User `
    -UserPropertiesPath $userPropertiesPath `
    -ADGroupsPath $adGroupsPath

# Extract variables for use in the rest of the script
$UserFromAD = $userInfo.UserFromAD
$DestinationOU = $userInfo.DestinationOU
$365Mailbox = $userInfo.Mailbox
$MgUser = $userInfo.MgUser

# Step 3: AD Tasks
Write-ProgressStep -StepName $progressSteps[3].Name -Status $progressSteps[3].Description
Disable-ADUser -UserFromAD $UserFromAD -DestinationOU $DestinationOU

# Step 4: Azure/Entra Tasks
Write-ProgressStep -StepName $progressSteps[4].Name -Status $progressSteps[4].Description
Remove-UserSessions -UserPrincipalName $UserFromAD.UserPrincipalName

# Step 5: Convert to SharedMailbox, Set forwarding/grant access
Write-ProgressStep -StepName $progressSteps[5].Name -Status $progressSteps[5].Description

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

# Step 6: Remove Directory Roles
Write-ProgressStep -StepName $progressSteps[6].Name -Status $progressSteps[6].Description
Remove-UserFromEntraDirectoryRoles -UserId $MgUser.Id

# Step 7: Remove Groups
Write-ProgressStep -StepName $progressSteps[7].Name -Status $progressSteps[7].Description
$GroupExportPath = Join-Path $localExportPath "$($User)_EntraGroups.csv"
Remove-UserFromEntraGroups -UserId $MgUser.Id -ExportPath $groupExportPath

# Step 8: Remove Licenses
Write-ProgressStep -StepName $progressSteps[8].Name -Status $progressSteps[8].Description
$licensePath = Join-Path $localExportPath "$($User)_EntraLicense.csv"
Remove-UserLicenses -UserId $UserFromAD.UserPrincipalName -ExportPath $licensePath

# Step 9: Remove from Zoom
Write-ProgressStep -StepName $progressSteps[9].Name -Status $progressSteps[9].Description
Remove-UserFromZoom -UserId $MgUser.Id

#Step 10: Send Email Notification - SecurePath
Write-ProgressStep -StepName $progressSteps[10].Name -Status $progressSteps[10].Description
$emailSubject = "KB4 – Remove User"
$emailContent = "The following user need to be removed to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
$MsgFrom = $config.Email.NotificationFrom
$ToAddress = $config.Email.NotificationTo
Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -Subject $emailSubject -Content $emailContent

# Step 11: Disconnect from Exchange Online and Graph
Write-ProgressStep -StepName $progressSteps[11].Name -Status $progressSteps[11].Description
Connect-ServiceEndpoints -Disconnect -ExchangeOnline -Graph

# Step 12: Configure OneDrive
Write-ProgressStep -StepName $progressSteps[12].Name -Status $progressSteps[12].Description

# Only create and run OneDrive params if needed
if ($SetOneDriveReadOnly -or $GrantUserOneDriveAccess) {
    $oneDriveParams = @{
        TermUser = $UserFromAD.UserPrincipalName
    }

    if ($SetOneDriveReadOnly) {
        $oneDriveParams['SetReadOnly'] = $true
    }

    if ($GrantUserOneDriveAccess) {
        $oneDriveParams['OneDriveUser'] = $oneDriveUser
    }

    Set-TerminatedOneDrive @oneDriveParams
}

# Step 13: Final Sync and Summary
Write-ProgressStep -StepName $progressSteps[13].Name -Status $progressSteps[13].Description

Start-ADSyncAndFinalize -User $MgUser `
    -DestinationOU $DestinationOU `
    -GrantUserFullControl $GrantUserFullControl `
    -SetUserMailFWD $SetUserMailFWD `
    -GrantUserOneDriveAccess $GrantUserOneDriveAccess `
    -ExportPath $localExportPath
