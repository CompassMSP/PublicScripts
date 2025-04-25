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
    Last Modified: 2025-07-04

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    4.0.0      2025-07-04   UI Refactor:
                                - Refactored the UI to use winui 3 styles for easier management and better visuals

    3.3.0      2025-04-01   Zoom Phone Removal:
                                - Removed provisioning steps for Zoom Phone and Contact Center as we are moving to 8x8

    3.2.0      2025-02-03   Zoom Phone Offboarding:
                                - Added removeal steps for Zoom Phone and Contact Cente

    3.0.0      2025-01-20   Major Rework:
                                - Complete script reorganization and optimization
                                - Optimized UI spacing and element alignment
                                - Enhanced form layout for improved readability
                                - Added secure configuration management via Get-ScriptConfig
                                - Enhanced error handling and logging system
                                - Added progress tracking and status messaging
                                - Added Zoom phone offboarding

    2.1.0      2024-11-25   Feature Update:
                                - Reworked GUI interface
                                - Added QuickEdit and InsertMode management
                                - Removed KnowBe4 SCIM integration per SecurePath Team
                                - Added Email Forwarding functionality - KnowBe4 Notification

    2.0.0      2024-07-15   Major Feature Update:
                                - Added GUI input system
                                - Enhanced UI for variable collection
                                - Added KB4 offboarding integration
                                - Added OneDrive read-only functionality
                                - Updated KnowBe4 SCIM integration
                                - Added directory role management

    1.2.0      2023-02-12   Feature Updates:
                                - Enhanced license management
                                - Improved group handling
                                - Added KnowBe4 integration
                                - Enhanced group function cleanup
                                - Added OneDrive access management

    1.1.0      2022-06-27   Enhancement Update:
                                - Added group and license exports
                                - Improved user management functions
                                - Enhanced manager removal process
                                - Fixed group member removal
                                - Added sign-in revocation

    1.0.0      2021-12-20   Initial Release:
                                - Basic termination functionality
                                - AD user management
                                - Group removal
                                - License removal
    ------------------------------------------------------------------------------
#>

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

$script:TestMode = $false  # Default to false

$ErrorActionPreference = 'Stop'
# Only show verbose output if -Verbose is specified
if (-not $PSBoundParameters['Verbose']) {
    $VerbosePreference = 'SilentlyContinue'
}

# Initialize loading animation
Clear-Host

Write-Host "`n  Initializing User Termination Script v4.0.0..." -ForegroundColor Cyan
$startTime = Get-Date
Write-Host "  Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray

Write-Host "  [ ] Loading functions..." -NoNewline -ForegroundColor Yellow

function Write-ProgressStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$StepName
    )

    # Find the step object by name
    $step = $progressSteps | Where-Object { $_.Name -eq $StepName }

    if (-not $step) {
        Write-Warning "Progress step '$StepName' not found."
        return
    }

    $stepNumber = $script:currentStep
    $status = $step.Description

    # Guard against division by zero or missing values
    if ($null -eq $stepNumber -or $script:totalSteps -eq 0) {
        Write-StatusMessage -Message "Step $StepName - $status" -Type INFO
        Write-Progress -Activity "User Termination" -Status $status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $status" -Type INFO
        Write-Progress -Activity "User Termination" -Status $status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
    }
    $script:currentStep += 1
}
function Write-ProgressStep {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$StepName
    )

    # Find the step object by name
    $step = $progressSteps | Where-Object { $_.Name -eq $StepName }

    if (-not $step) {
        Write-Warning "Progress step '$StepName' not found."
        return
    }

    $stepNumber = $step.Number
    $status = $step.Description

    # Guard against division by zero or missing values
    if ($null -eq $stepNumber -or $script:totalSteps -eq 0) {
        Write-StatusMessage -Message "Step $StepName - $status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
    }
}

#Region Standard Functions

function Write-StatusMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'OK', 'ERROR', 'WARN', 'SUMMARY')]
        [string]$Type = 'INFO'
    )

    $config = @{
        'INFO'    = @{ Status = 'INFO'; Color = 'White' }
        'OK'      = @{ Status = 'OK'; Color = 'Green' }
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
            message         = @{
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
            saveToSentItems = $false
        }

        # Add CC recipients if specified
        if ($CcAddress) {
            $messageParams.message['ccRecipients'] = @(
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

            $messageParams.message['attachments'] = @(
                @{
                    '@odata.type' = '#microsoft.graph.fileAttachment'
                    name          = $AttachmentName ?? (Split-Path $AttachmentPath -Leaf)
                    contentType   = 'text/plain'
                    contentBytes  = $attachmentBase64
                }
            )
        }

        # Use Graph API directly
        $graphUri = "https://graph.microsoft.com/v1.0/users/$FromAddress/sendMail"
        Invoke-MgGraphRequest -Method POST -Uri $graphUri -Body $messageParams -ContentType "application/json"
        Write-StatusMessage -Message "Email notification sent successfully" -Type OK
    } catch {
        Write-StatusMessage -Message "Failed to send email notification: $_" -Type ERROR
    }
}

#Region Custom Functions

function Show-CustomAlert {
    param (
        [string]$Message,
        [ValidateSet("Error", "Warning", "Info", "Success")]
        [string]$AlertType = "Error",
        [string]$Title
    )

    if (-not $Title) {
        $Title = $AlertType
    }

    switch ($AlertType) {
        "Error" { $color = "#E81123"; $sound = [System.Media.SystemSounds]::Hand }
        "Warning" { $color = "#FFB900"; $sound = [System.Media.SystemSounds]::Exclamation }
        "Info" { $color = "#0078D7"; $sound = [System.Media.SystemSounds]::Asterisk }
        "Success" { $color = "#107C10"; $sound = [System.Media.SystemSounds]::Beep }
    }

    $XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="110" Width="400"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style TargetType="TextBlock">
                <Setter Property="Foreground" Value="White"/>
                <Setter Property="FontSize" Value="14"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>

    <Border CornerRadius="12" Background="#2D2D30" Padding="15" BorderBrush="$color" BorderThickness="2">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Top Bar for Dragging -->
            <Border Name="Top_Bar" Background="Transparent" Height="5" Grid.Row="0" />

            <!-- Content -->
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Icon -->
                <Viewbox Grid.Row="0" Grid.RowSpan="2" Margin="0,0,15,0" Width="40" Height="40" VerticalAlignment="Top">
                    <Canvas Width="48" Height="48">
                        <Ellipse Width="48" Height="48" Fill="$color"/>
                        <Rectangle Width="6" Height="20" Fill="White" Canvas.Left="21" Canvas.Top="10"/>
                        <Rectangle Width="6" Height="6" Fill="White" Canvas.Left="21" Canvas.Top="34"/>
                    </Canvas>
                </Viewbox>

                <!-- Message -->
                <TextBlock Grid.Column="1" Grid.Row="0" TextWrapping="Wrap" Text="$Message" Margin="0,0,0,10"/>

                <!-- Button -->
                <Button Name="OkButton" Grid.Column="1" Grid.Row="1" Content="OK" Width="80" HorizontalAlignment="Right" IsDefault="True"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

    Add-Type -AssemblyName PresentationFramework

    $stringReader = New-Object System.IO.StringReader $xaml
    $xmlReader = [System.Xml.XmlReader]::Create($stringReader)
    $alertWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

    $okButton = $alertWindow.FindName("OkButton")
    $okButton.Add_Click({ $alertWindow.Close() })

    # Find the top bar and add the MouseDown event
    $topBar = $alertWindow.FindName("Top_Bar")
    $topBar.Add_MouseDown({
            param($s, $e)
            if ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
                $alertWindow.DragMove()
            }
        })

    # Play the appropriate sound
    $sound.Play()

    $alertWindow.ShowDialog() | Out-Null
}

function Get-UserTermination {
    <#
    .SYNOPSIS
    Shows a modern GUI window for processing a user termination request.

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
        TestModeEnabled        : [bool] Whether test mode is enabled
    Returns $null if the user cancels the operation.
    #>

    # Add required assemblies
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    # XAML Design
    [xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Termination Request" Height="780" Width="600"
    WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml" />
            </ResourceDictionary.MergedDictionaries>

            <!-- TabControl Style -->
            <Style TargetType="TabControl">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Padding" Value="0"/>
            </Style>

            <!-- TabItem Style -->
            <Style TargetType="TabItem">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="Padding" Value="20,10"/>
                <Setter Property="Margin" Value="0,0,4,0"/>
                <Setter Property="FontSize" Value="14"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="TabItem">
                            <Border x:Name="Border"
                                    Background="{TemplateBinding Background}"
                                    BorderBrush="{DynamicResource TextControlBorderBrush}"
                                    BorderThickness="0,0,0,2"
                                    CornerRadius="4,4,0,0">
                                <ContentPresenter x:Name="ContentSite"
                                                ContentSource="Header"
                                                HorizontalAlignment="Center"
                                                VerticalAlignment="Center"/>
                            </Border>
                            <ControlTemplate.Triggers>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="{DynamicResource TextControlBackgroundPointerOver}"/>
                                    <Setter Property="BorderThickness" Value="0,0,0,2"/>
                                    <Setter Property="BorderBrush" Value="{DynamicResource SystemAccentColor}"/>
                                    <Setter Property="TextElement.Foreground" Value="{DynamicResource SystemAccentColor}"/>
                                </Trigger>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="{DynamicResource TextControlBackgroundPointerOver}"/>
                                    <Setter TargetName="Border" Property="BorderThickness" Value="0,0,0,2"/>
                                    <Setter TargetName="Border" Property="BorderBrush" Value="{DynamicResource SystemAccentColorLight1}"/>
                                </Trigger>
                            </ControlTemplate.Triggers>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>
        </ResourceDictionary>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="20,20,20,0">
            <TextBlock Text="User Termination Request" FontSize="24" FontWeight="SemiBold" Margin="0,0,0,20"/>
            <Border Background="{DynamicResource TextControlBackgroundPointerOver}"
                    BorderBrush="{DynamicResource TextControlBorderBrush}"
                    BorderThickness="1"
                    Padding="10"
                    Margin="0,0,0,20">
                <TextBlock TextWrapping="Wrap">
                    Please fill in the required information for user termination. Fields marked with * are required.
                </TextBlock>
            </Border>
        </StackPanel>

        <!-- Main Content -->
        <ScrollViewer Grid.Row="1" Margin="20,10,20,20" VerticalScrollBarVisibility="Auto">
            <StackPanel>
                <!-- User Information Section -->
                <GroupBox Header="User Information"
                         Margin="0,0,0,15"
                         BorderBrush="{DynamicResource TextControlBorderBrush}"
                         Background="{DynamicResource TextControlBackgroundPointerOver}">
                    <StackPanel Margin="15">
                        <Label>
                            <TextBlock>
                                <Run Text="User to Terminate (Email)"/>
                                <Run Text=" *" Foreground="Red"/>
                            </TextBlock>
                        </Label>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtUserToTerminate"
                                   Height="32"
                                   Padding="8,5,8,5"
                                   VerticalContentAlignment="Center"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter user's email address"
                                     VerticalAlignment="Center"
                                     Margin="8,0,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtUserToTerminate}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>
                    </StackPanel>
                </GroupBox>

                <!-- Access Delegation Section -->
                <GroupBox Header="Access Delegation"
                         Margin="0,0,0,15"
                         BorderBrush="{DynamicResource TextControlBorderBrush}"
                         Background="{DynamicResource TextControlBackgroundPointerOver}">
                    <Grid Margin="15">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <!-- Left Column -->
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <Label Content="Grant Full Mailbox Control To (Email)"/>
                            <Grid Margin="0,0,0,15">
                                <TextBox x:Name="txtMailboxControl"
                                       Height="32"
                                       Padding="8,5,8,5"
                                       VerticalContentAlignment="Center"/>
                                <TextBlock IsHitTestVisible="False"
                                         Text="Enter delegate's email address"
                                         VerticalAlignment="Center"
                                         Margin="8,0,0,0"
                                         Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                    <TextBlock.Style>
                                        <Style TargetType="{x:Type TextBlock}">
                                            <Setter Property="Visibility" Value="Collapsed"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding Text, ElementName=txtMailboxControl}" Value="">
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </TextBlock.Style>
                                </TextBlock>
                            </Grid>

                            <Label Content="Forward Mailbox To (Email)"/>
                            <Grid Margin="0,0,0,15">
                                <TextBox x:Name="txtForwardMailbox"
                                       Height="32"
                                       Padding="8,5,8,5"
                                       VerticalContentAlignment="Center"/>
                                <TextBlock IsHitTestVisible="False"
                                         Text="Enter forward-to email address"
                                         VerticalAlignment="Center"
                                         Margin="8,0,0,0"
                                         Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                    <TextBlock.Style>
                                        <Style TargetType="{x:Type TextBlock}">
                                            <Setter Property="Visibility" Value="Collapsed"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding Text, ElementName=txtForwardMailbox}" Value="">
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </TextBlock.Style>
                                </TextBlock>
                            </Grid>
                        </StackPanel>

                        <!-- Right Column -->
                        <StackPanel Grid.Column="1" Margin="10,0,0,0">
                            <Label Content="Grant OneDrive Access To (Email)"/>
                            <Grid Margin="0,0,0,15">
                                <TextBox x:Name="txtOneDriveAccess"
                                       Height="32"
                                       Padding="8,5,8,5"
                                       VerticalContentAlignment="Center"/>
                                <TextBlock IsHitTestVisible="False"
                                         Text="Enter delegate's email address"
                                         VerticalAlignment="Center"
                                         Margin="8,0,0,0"
                                         Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                    <TextBlock.Style>
                                        <Style TargetType="{x:Type TextBlock}">
                                            <Setter Property="Visibility" Value="Collapsed"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding Text, ElementName=txtOneDriveAccess}" Value="">
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </TextBlock.Style>
                                </TextBlock>
                            </Grid>

                            <CheckBox x:Name="chkOneDriveReadOnly"
                                    Content="Set OneDrive as Read-Only"
                                    Margin="0,0,0,5"/>
                        </StackPanel>
                    </Grid>
                </GroupBox>

                <!-- Out of Office Section -->
                <GroupBox Header="Out of Office Settings"
                         Margin="0,0,0,15"
                         BorderBrush="{DynamicResource TextControlBorderBrush}"
                         Background="{DynamicResource TextControlBackgroundPointerOver}">
                    <StackPanel Margin="15">
                        <Label Content="Out of Office Message"/>
                        <TextBox x:Name="txtOutOfOffice"
                               Height="80"
                               TextWrapping="Wrap"
                               AcceptsReturn="True"
                               VerticalScrollBarVisibility="Auto"
                               Padding="8,5,8,5"/>
                    </StackPanel>
                </GroupBox>
            </StackPanel>
        </ScrollViewer>

        <!-- Footer -->
        <Border Grid.Row="2"
                BorderBrush="{DynamicResource TextControlBorderBrush}"
                BorderThickness="0,1,0,0"
                Background="{DynamicResource TextControlBackgroundPointerOver}"
                Padding="20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <CheckBox x:Name="chkTestMode" Content="Test Mode" Grid.Column="0"/>
                <TextBlock x:Name="tbStatus" Grid.Column="1" VerticalAlignment="Center" Margin="20,0"/>
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="btnSubmit" Content="Submit" Style="{DynamicResource AccentButtonStyle}" Padding="20,5" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnReset" Content="Reset Form" Padding="20,5" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnCancel" Content="Cancel" Padding="20,5" Height="32"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Parse XAML
    $XAMLReader = [System.Xml.XmlNodeReader]::new($XAML)
    $Window = [Windows.Markup.XamlReader]::Load($XAMLReader)

    # Create namespace manager for XPath queries
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XAML.NameTable)
    $nsManager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")

    # Get all form controls by name and create variables
    $XAML.SelectNodes("//*[@x:Name]", $nsManager) | ForEach-Object {
        $Name = $_.Name
        Set-Variable -Name $Name -Value $Window.FindName($Name) -Force
    }

    # Validation function
    function Test-EmailAddress {
        param ([string]$Email)
        return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    }

    function Show-ValidationError {
        param (
            [string]$Message,
            [string]$Title = "Validation Error"
        )
        Show-CustomAlert -Message $Message -AlertType "Error" -Title $Title
    }

    function Show-StatusMessage {
        param (
            [string]$Message,
            [ValidateSet("Info", "Success", "Warning", "Error")]
            [string]$Type = "Info"
        )
        $tbStatus.Text = $Message
        switch ($Type) {
            "Info" { $tbStatus.Foreground = "#0078D7" }
            "Success" { $tbStatus.Foreground = "#107C10" }
            "Warning" { $tbStatus.Foreground = "#FFB900" }
            "Error" { $tbStatus.Foreground = "#E81123" }
        }
    }

    # Function to reset the form
    function Reset-Form {
        $txtUserToTerminate.Text = ""
        $txtMailboxControl.Text = ""
        $txtForwardMailbox.Text = ""
        $txtOneDriveAccess.Text = ""
        $chkOneDriveReadOnly.IsChecked = $false
        $chkTestMode.IsChecked = $false
        Show-StatusMessage -Message "Form has been reset" -Type "Info"
    }

    # Function to get form data
    function Get-FormData {
        $outOfOfficeMessage = $txtOutOfOffice.Text.Trim()
        return [PSCustomObject]@{
            InputUser               = $txtUserToTerminate.Text
            InputUserFullControl    = $txtMailboxControl.Text
            InputUserFWD            = $txtForwardMailbox.Text
            InputUserOneDriveAccess = $txtOneDriveAccess.Text
            SetOneDriveReadOnly     = $chkOneDriveReadOnly.IsChecked
            TestModeEnabled         = $chkTestMode.IsChecked
            OutOfOfficeMessage      = $txtOutOfOffice.Text.Trim()
            SetOOO                  = [bool]$outOfOfficeMessage
        }
    }

    # Add email validation to text boxes
    $emailTextBoxes = @($txtUserToTerminate, $txtMailboxControl, $txtForwardMailbox, $txtOneDriveAccess)
    foreach ($textBox in $emailTextBoxes) {
        $textBox.Add_TextChanged({
                if ($this.Text -and -not (Test-EmailAddress -Email $this.Text)) {
                    $this.BorderBrush = 'Red'
                    $this.BorderThickness = 2
                } else {
                    $this.BorderBrush = $null
                    $this.BorderThickness = 1
                }
            })
    }

    # Add button click handlers
    $btnSubmit.Add_Click({
            # Validate required user email
            if (-not $txtUserToTerminate.Text -or -not (Test-EmailAddress -Email $txtUserToTerminate.Text)) {
                Show-ValidationError -Message "Please enter a valid email address for the user to terminate."
                return
            }

            # Validate optional email fields if they're not empty
            $optionalEmails = @{
                'Mailbox Control' = $txtMailboxControl
                'Forward Mailbox' = $txtForwardMailbox
                'OneDrive Access' = $txtOneDriveAccess
            }

            foreach ($field in $optionalEmails.GetEnumerator()) {
                if ($field.Value.Text -and -not (Test-EmailAddress -Email $field.Value.Text)) {
                    Show-ValidationError -Message "Please enter a valid email address for $($field.Key)."
                    return
                }
            }

            $Window.DialogResult = $true
            $Window.Close()
        })

    $btnCancel.Add_Click({
            $Window.DialogResult = $false
            $Window.Close()
        })

    $btnReset.Add_Click({ Reset-Form })

    # Show the window
    $result = $Window.ShowDialog()

    # Return results if the form was submitted
    if ($result -eq $true) {
        return Get-FormData
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
            } catch {
                Write-StatusMessage -Message "Failed to export user data: $_" -Type ERROR
            }
        }

        # Return all the collected information
        return @{
            selectUserFromAD    = $UserFromAD
            selectDestinationOU = $DestinationOU
            selectMailbox       = $365Mailbox
            selectMgUser        = $MgUser
            UserPropertiesPath  = $UserPropertiesPath
            ADGroupsPath        = $ADGroupsPath
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
            }

            Set-ADUser @SetADUserParams -ErrorAction Stop
            Write-StatusMessage -Message "User account disabled and attributes cleared" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to disable user account" -Type ERROR
            throw
        }

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

function Remove-ADUserGroups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADUser]
        $ADUser,

        [Parameter()]
        [string]$ExportPath
    )

    try {
        Write-StatusMessage -Message "Starting AD group removal process" -Type INFO

        # Get and export group details if path provided
        $groupDetails = $ADUser.MemberOf | ForEach-Object {
            Get-ADGroup $_ -Properties Name, DistinguishedName
        } | Select-Object Name, DistinguishedName

        if ($ExportPath) {
            try {
                $groupDetails | Export-Csv -Path $ExportPath -NoTypeInformation -ErrorAction Stop
                Write-StatusMessage -Message "Exported user groups to: $ExportPath" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to export user groups" -Type ERROR
            }
        }

        # Remove from all groups
        foreach ($group in $groupDetails) {
            try {
                Remove-ADGroupMember -Identity $group.DistinguishedName -Members $ADUser.SamAccountName -Confirm:$false -ErrorAction Stop
                Write-StatusMessage -Message "Removed from group: $($group.Name)" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to remove from group $($group.Name)" -Type ERROR
            }
        }

    } catch {
        Write-StatusMessage -Message "Error in Remove-UserGroups: $($_.Exception.Message)" -Type ERROR
        throw
    }
}

function Remove-UserSessions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User
    )

    try {
        # Revoke all sessions
        Write-StatusMessage -Message "Revoking all user signed in sessions" -Type INFO
        try {
            Revoke-MgUserSignInSession -UserId $User.Id -ErrorAction Stop
            Write-StatusMessage -Message "Successfully revoked all user sessions" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to revoke user sessions" -Type ERROR
        }

        # Remove authentication methods
        Write-StatusMessage -Message "Removing user authentication methods" -Type INFO
        try {
            $authMethods = Get-MgUserAuthenticationMethod -UserId $User.Id -ErrorAction Stop

            foreach ($authMethod in $authMethods) {
                $authType = $authMethod.AdditionalProperties.'@odata.type'

                try {
                    switch ($authType) {
                        "#microsoft.graph.passwordAuthenticationMethod" {
                            continue
                        }
                        "#microsoft.graph.phoneAuthenticationMethod" {
                            Remove-MgUserAuthenticationPhoneMethod -UserId $User.Id -PhoneAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Phone Authentication Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                            Remove-MgUserAuthenticationWindowsHelloForBusinessMethod -UserId $User.Id -WindowsHelloForBusinessAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Windows Hello for Business Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                            Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId $User.Id -MicrosoftAuthenticatorAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Microsoft Authenticator Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.fido2AuthenticationMethod" {
                            Remove-MgUserAuthenticationFido2Method -UserId $User.Id -Fido2AuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed FIDO2 Authenticator Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.softwareOathAuthenticationMethod" {
                            Remove-MgUserAuthenticationSoftwareOathMethod -UserId $User.Id -SoftwareOathAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Software Oath Method: $($authMethod.Id)" -Type OK
                        }
                        "#microsoft.graph.temporaryAccessPassAuthenticationMethod" {
                            Remove-MgUserAuthenticationTemporaryAccessPassMethod -UserId $User.Id -TemporaryAccessPassAuthenticationMethodId $authMethod.Id -ErrorAction Stop
                            Write-StatusMessage -Message "Removed Temporary Access Pass Method: $($authMethod.Id)" -Type OK
                        }
                        default {
                            Write-StatusMessage -Message "Skipping unknown authentication method: $authType" -Type ERROR
                        }
                    }
                } catch {
                    Write-StatusMessage -Message "Failed to remove authentication method $($authMethod.Id) of type $authType" -Type ERROR
                }
            }
        } catch {
            Write-StatusMessage -Message "Failed to get user authentication methods" -Type ERROR
        }

        # Remove Mobile Devices
        Write-StatusMessage -Message "Removing all mobile devices" -Type INFO
        try {
            $mobileDevices = Get-MobileDevice -Mailbox $User.mail -ErrorAction Stop
            foreach ($mobileDevice in $mobileDevices) {
                Write-StatusMessage -Message "Removing mobile device: $($mobileDevice.Id)" -Type INFO
                try {
                    Remove-MobileDevice -Identity $mobileDevice.Id -Confirm:$false -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully removed mobile device: $($mobileDevice.Id)" -Type OK
                } catch {
                    Write-StatusMessage -Message "Failed to remove mobile device $($mobileDevice.Id)" -Type ERROR
                }
            }
        } catch {
            Write-StatusMessage -Message "Failed to get mobile devices" -Type ERROR
        }

        # Disable Azure AD devices
        try {
            $termUserDevices = Get-MgUserRegisteredDevice -UserId $User.Id -ErrorAction Stop
            foreach ($termUserDevice in $termUserDevices) {
                Write-StatusMessage -Message "Disabling registered device: $($termUserDevice.Id)" -Type INFO
                try {
                    Update-MgDevice -DeviceId $termUserDevice.Id -BodyParameter @{ AccountEnabled = $false } -ErrorAction Stop
                    Write-StatusMessage -Message "Successfully disabled device: $($termUserDevice.Id)" -Type OK
                } catch {
                    Write-StatusMessage -Message "Failed to disable device $($termUserDevice.Id)" -Type ERROR
                }
            }
        } catch {
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
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User
    )

    try {
        Write-StatusMessage -Message "Checking for directory role memberships..." -Type INFO

        try {
            # Get all directory roles the user is a member of
            $directoryRoles = Get-MgUserMemberOf -UserId $User.Id -ErrorAction Stop |
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
                    Remove-MgDirectoryRoleMemberByRef -DirectoryRoleId $roleId -DirectoryObjectId $User.Id -ErrorAction Stop
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
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

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
                $_.AdditionalProperties.groupTypes -notcontains "DynamicMembership" -and
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
                    }
                }
                @{n = 'securityEnabled'; e = { $_.AdditionalProperties.securityEnabled } }
            )
        }

        Write-StatusMessage -Message "Finding Azure groups" -Type INFO

        try {
            $All365Groups = Get-MgUserMemberOf -UserId $User.Id -ErrorAction Stop |
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
                        Remove-MgGroupMemberByRef -GroupId $365Group.Id -DirectoryObjectId $User.Id -ErrorAction Stop
                        Write-StatusMessage -Message "Removed from Security/Unified Group: $($365Group.DisplayName)" -Type OK
                    } else {
                        Remove-DistributionGroupMember -Identity $365Group.Id -Member $User.Id -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
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
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

        [Parameter()]
        [string]$ExportPath,

        [Parameter()]
        [int]$MaxRetries = 3,

        [Parameter()]
        [int]$RetryDelaySeconds = 5
    )

    try {
        Write-StatusMessage -Message "Starting license removal process" -Type INFO

        try {
            # Get and export license details if path provided
            $licenseDetails = Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop |
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

            # Track failed removals for retry
            $failedLicenses = [System.Collections.ArrayList]::new()

            # Step 1: Remove Ancillary Licenses
            foreach ($license in ($licenseDetails | Where-Object { $_.SkuPartNumber -notin $primaryLicenses })) {
                try {

                    # Assign the license
                    $licenseBody = @{
                        addLicenses    = @()
                        removeLicenses = @($license.SkuId)
                    } | ConvertTo-Json -Depth 3

                    $uri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/assignLicense"
                    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $licenseBody -ContentType "application/json" -ErrorAction Stop

                    Write-StatusMessage -Message "Removed Ancillary License: $($license.SkuPartNumber)" -Type OK
                    Start-Sleep -Seconds 2  # Brief pause after successful removal
                } catch {
                    Write-StatusMessage -Message "Failed to remove Ancillary License $($license.SkuPartNumber) - will retry later" -Type WARN
                    $null = $failedLicenses.Add($license)
                }
            }

            # Step 2: Remove Primary Licenses
            foreach ($license in ($licenseDetails | Where-Object { $_.SkuPartNumber -in $primaryLicenses })) {
                try {

                    # Assign the license
                    $licenseBody = @{
                        addLicenses    = @()
                        removeLicenses = @($license.SkuId)
                    } | ConvertTo-Json -Depth 3

                    $uri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/assignLicense"
                    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $licenseBody -ContentType "application/json" -ErrorAction Stop

                    Write-StatusMessage -Message "Removed Primary License: $($license.SkuPartNumber)" -Type OK
                    Start-Sleep -Seconds 2  # Brief pause after successful removal
                } catch {
                    Write-StatusMessage -Message "Failed to remove Primary License $($license.SkuPartNumber) - will retry later" -Type WARN
                    $null = $failedLicenses.Add($license)
                }
            }

            # Step 3: Retry failed removals
            if ($failedLicenses.Count -gt 0) {
                Write-StatusMessage -Message "Waiting 10 seconds before retrying failed removals..." -Type INFO
                Start-Sleep -Seconds 10

                $remainingRetries = $MaxRetries - 1  # Already tried once above
                $stillFailed = [System.Collections.ArrayList]::new()

                while ($failedLicenses.Count -gt 0 -and $remainingRetries -gt 0) {
                    Write-StatusMessage -Message "Retrying failed removals (attempt $($MaxRetries - $remainingRetries) of $MaxRetries)" -Type INFO

                    foreach ($license in $failedLicenses) {
                        try {

                            # Assign the license
                            $licenseBody = @{
                                addLicenses    = @()
                                removeLicenses = @(@{ skuId = $license.SkuId })
                            } | ConvertTo-Json -Depth 3

                            $uri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/assignLicense"
                            $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $licenseBody -ContentType "application/json" -ErrorAction Stop

                            Write-StatusMessage -Message "Successfully removed license on retry: $($license.SkuPartNumber)" -Type OK
                            Start-Sleep -Seconds 2
                        } catch {
                            $null = $stillFailed.Add($license)
                            Write-StatusMessage -Message "Failed to remove license $($license.SkuPartNumber) on retry" -Type WARN
                        }
                    }

                    $failedLicenses.Clear()
                    if ($stillFailed.Count -gt 0) {
                        $null = $failedLicenses.AddRange($stillFailed)
                        $stillFailed.Clear()
                        $remainingRetries--

                        if ($remainingRetries -gt 0) {
                            Write-StatusMessage -Message "Waiting $RetryDelaySeconds seconds before next retry..." -Type INFO
                            Start-Sleep -Seconds $RetryDelaySeconds
                        }
                    } else {
                        break  # All retries successful
                    }
                }

                # Report any licenses that couldn't be removed
                if ($failedLicenses.Count -gt 0) {
                    Write-StatusMessage -Message "The following licenses could not be removed after $MaxRetries attempts:" -Type ERROR
                    foreach ($license in $failedLicenses) {
                        Write-StatusMessage -Message "- $($license.SkuPartNumber)" -Type ERROR
                    }
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
        $summaryParts += "- Attributes, Groups and licenses exported to: $ExportPath"
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
}

Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green

#Region Main Execution

Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green

Write-Host "  [ ] Initializing progress tracking..." -NoNewline -ForegroundColor Yellow
$progressSteps = @(
    @{ Name = "Initialization"; Description = "Loading configuration and connecting services" }
    @{ Name = "User Input"; Description = "Gathering termination details" }
    @{ Name = "AD Tasks"; Description = "Disabling user in Active Directory" }
    @{ Name = "Session Cleanup"; Description = "Removing user sessions and devices" }
    @{ Name = "Exchange Tasks"; Description = "Convert to SharedMailbox and setting forwarding/grant acces" }
    @{ Name = "Directory Roles"; Description = "Removing from directory roles" }
    @{ Name = "Group Removal"; Description = "Removing and exporting Entra/Exchange groups" }
    @{ Name = "License Removal"; Description = "Removing and exporting Entra licenses" }
    @{ Name = "Notifications"; Description = "Sending email notifications" }
    @{ Name = "Disconnecting from Exchange and Graph"; Description = "Disconnecting from Exchange and Graph" }
    @{ Name = "OneDrive Setup"; Description = "Configuring OneDrive access" }
    @{ Name = "Summary"; Description = "Running AD sync and Summary" }
)
$script:totalSteps = $progressSteps.Count
$script:currentStep = 0
Write-Host "`r  [✓] Progress tracking initialized" -ForegroundColor Green

try {

    Write-Host "`n  Beginning User Termination..." -ForegroundColor Cyan

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

    Connect-ServiceEndpoints -ExchangeOnline -Graph

    # Call the custom input window function

    $result = Get-UserTermination
    if (-not $result) {
        Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
    }

    if ($result.TestModeEnabled -eq 'True') { $script:TestMode = $true }
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

    # Step: User Input
    Write-ProgressStep -StepName 'User Input'
    # Should this be a function in the New User Script?
    $userPropertiesPath = Join-Path $config.Paths.TermExportPath "$($result.InputUser)_Properties.csv"
    $adGroupsPath = Join-Path $config.Paths.TermExportPath "$($result.InputUser)_ADGroups.csv"
    $UserInfo = Get-TerminationPrerequisites `
        -User $result.InputUser `
        -UserPropertiesPath $userPropertiesPath `
        -ADGroupsPath $adGroupsPath

    # Extract variables for use in the rest of the script

    # Step: AD Tasks
    Write-ProgressStep -StepName 'AD Tasks'
    Disable-ADUser -UserFromAD $userInfo.selectUserFromAD -DestinationOU $userInfo.selectDestinationOU
    Remove-ADUserGroups -ADUser $userInfo.selectUserFromAD

    # Step: Azure/Entra Tasks
    Write-ProgressStep -StepName 'Session Cleanup'
    Remove-UserSessions -User $UserInfo.selectMgUser

    # Set Out of Office Message
    if ($result.setOOO) {
        Set-MailboxAutoReplyConfiguration `
            -Identity $result.InputUser `
            -AutoReplyState Enabled `
            -ExternalMessage $result.OutOfOfficeMessage `
            -InternalMessage $null `
            -ExternalAudience All
    }

    # Step: Convert to SharedMailbox, Set forwarding/grant access
    Write-ProgressStep -StepName 'Exchange Tasks'

    $mailboxParams = @{
        Mailbox = $userInfo.selectMailbox
    }

    # Only add these parameters if they exist and have values
    if ($SetUserMailFWD) {
        $mailboxParams['ForwardingAddress'] = $SetUserMailFWD
    }

    if ($GrantUserFullControl) {
        $mailboxParams['GrantAccessTo'] = $GrantUserFullControl
    }

    Set-TerminatedMailbox @mailboxParams

    # Step: Remove Directory Roles
    Write-ProgressStep -StepName 'Directory Roles'
    Remove-UserFromEntraDirectoryRoles -User $userInfo.selectMgUser

    # Step: Remove Groups
    Write-ProgressStep -StepName 'Group Removal'
    $groupExportPath = Join-Path $config.Paths.TermExportPath "$($result.InputUser)_Groups_Id.csv"
    Remove-UserFromEntraGroups -User $userInfo.selectMgUser -ExportPath $groupExportPath

    # Step: Remove Licenses
    Write-ProgressStep -StepName 'License Removal'
    $licensePath = Join-Path $config.Paths.TermExportPath "$($result.InputUser)_License_Id.csv"
    Remove-UserLicenses -User $userInfo.selectMgUser -ExportPath $licensePath

    # Step: Send notifications
    Write-ProgressStep -StepName 'Notifications'
    $MsgFrom = $config.Email.NotificationFrom
    $CcAddress  = $config.Email.NotificationCcAddress

    # Email to SOC for KnowBe4
    try {
        $ToAddress = $config.Email.NotificationToKnowBe4
        $emailSubject = "KB4 – Remove User"
        $emailContent = @"
The following user need to be removed to the CompassMSP KnowBe4 account. <p> $($userInfo.selectMgUser.DisplayName) <br> $($userInfo.selectMgUser.Mail)"
"@

        Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -CcAddress $CcAddress -Subject $emailSubject -Content $emailContent
    } catch {
        Write-StatusMessage -Message "Failed to send KnowBe4 notification email: $($_.Exception.Message)" -Type ERROR
    }

    # Email Compass West for 8x8
    try {
        $ToAddress = $config.Email.NotificationTo8x8
        $emailSubject = "8x8 – Remove User"
        $emailContent = @"
The following user need to be removed from 8x8. <p> $($userInfo.selectMgUser.DisplayName) <br> $($userInfo.selectMgUser.Mail)"
"@

        Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -CcAddress $CcAddress -Subject $emailSubject -Content $emailContent
    } catch {
        Write-StatusMessage -Message "Failed to send 8x8 notification email: $($_.Exception.Message)" -Type ERROR
    }

    # Step : Disconnect from Exchange Online and Graph
    Write-ProgressStep -StepName 'Disconnecting from Exchange and Graph'
    Connect-ServiceEndpoints -Disconnect -ExchangeOnline -Graph

    # Step: Configure OneDrive
    Write-ProgressStep -StepName 'OneDrive Setup'

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

    # Step: Final Sync and Summary
    Write-ProgressStep -StepName 'Summary'

    Start-ADSyncAndFinalize -User $userInfo.selectMgUser `
        -GrantUserFullControl $GrantUserFullControl `
        -SetUserMailFWD $SetUserMailFWD `
        -GrantUserOneDriveAccess $GrantUserOneDriveAccess `
        -ExportPath $config.Paths.TermExportPath

    # Clear the progress bar
    Write-Progress -Activity "User Termination" -Completed

    Exit-Script -Message "$User has been successfully disabled." -ExitCode Success
} catch {

    Write-StatusMessage -Message "Script failed: $($_.Exception.Message)" -Type ERROR
    Write-StatusMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -Type ERROR

    # Clear the progress bar
    Write-Progress -Activity "User Termination" -Status "Failed" -PercentComplete 100

    Exit-Script -Message "Script failed during execution" -ExitCode GeneralError
}
