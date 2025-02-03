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
    Last Modified: 2025-02-03

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    3.2.0        2025-02-03  Zoom Phone Onboarding:
                          - Added provisioning steps for Zoom Phone and Contact Center

    3.1.0        2025-01-25  Password System Update:
                          - Replaced New-SecureRandomPassword with New-ReadablePassword
                          - Added human-readable password generation using word list
                          - Added interactive password acceptance/rejection
                          - Added GitHub wordlist integration
                          - Added support for custom word lists
                          - Added configurable word count (2-20 words)
                          - Added spaces/no-spaces password formatting options

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
    @{ Number = 1; Name = "Initialization"; Description = "Loading configuration and connecting services" }
    @{ Number = 2; Name = "User Input"; Description = "Gathering new user details" }
    @{ Number = 3; Name = "Validation"; Description = "Validating inputs and building user creation prerequisites" }
    @{ Number = 4; Name = "New User AD Creation"; Description = "Creating user in Active Directory" }
    @{ Number = 5; Name = "AD Group Copy"; Description = "Copying AD group memberships" }
    @{ Number = 6; Name = "Azure Sync"; Description = "Syncing to Azure AD" }
    @{ Number = 7; Name = "License Setup"; Description = "Assigning licenses" }
    @{ Number = 8; Name = "Entra Group Copy"; Description = "Copying Entra group memberships" }
    @{ Number = 9; Name = "Email to SOC for KnowBe4"; Description = "Sending SOC notification email for KnowBe4 setup" }
    @{ Number = 10; Name = "OneDrive Provisioning"; Description = "Provisioning new users OneDrive" }
    @{ Number = 11; Name = "Configuring BookWithMeId"; Description = "Configuring BookWithMeId" }
    @{ Number = 12; Name = "Zoom Phone Setup"; Description = "Adding User to Zoom Phone" }
    @{ Number = 13; Name = "Cleanup and Summary"; Description = "Running cleanup and summary" }
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
        Write-Progress -Activity "New User Creation" -Status $Status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $Status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $Status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
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

function Get-NewUserRequestInput {
    <#
        .SYNOPSIS
        Shows a GUI window for creating a new user request.

        .DESCRIPTION
        Displays a WPF window that collects information needed to create a new user,
        including user details, mobile number, and license selections.

        .OUTPUTS
        [PSCustomObject] Returns a custom object with the following properties:
            InputNewUser           : [string] The new user's display name (First Last format)
            InputNewMobile        : [string] Formatted mobile number or null if not provided
            InputUserToCopy       : [string] Template user's display name to copy permissions from
            InputRequiredLicense  : [hashtable] Selected required license with properties:
                - SkuId          : [string] The license SKU ID
                - DisplayName    : [string] The friendly name of the license
            InputAncillaryLicenses: [array] Array of selected additional licenses, each containing:
                - SkuId          : [string] The license SKU ID
                - DisplayName    : [string] The friendly name of the license
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

    function New-FormComboBox {
        param (
            [string]$ToolTip,
            [string]$Margin = "0,0,0,10",
            [string]$Padding = "5,3,5,3",
            [string]$DisplayMemberPath
        )
        $comboBox = New-Object System.Windows.Controls.ComboBox
        $comboBox.Margin = $Margin
        $comboBox.Padding = $Padding
        $comboBox.ToolTip = $ToolTip
        if ($DisplayMemberPath) {
            $comboBox.DisplayMemberPath = $DisplayMemberPath
        }
        return $comboBox
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
            "SHAREPOINTSTORAGE" { "SharePoint Storage" }
            "Teams_Premium_(for_Departments)" { "Teams Premium (for Departments)" }
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
                SkuId       = $_.SkuId
                SortName    = $SkuDisplayName
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

    # Define required and ignored licenses
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
        "Dynamics 365 Business Central for IWs",
        "SharePoint Storage",
        "Teams Premium (for Departments)"
    )

    # Create window and main containers
    $window = New-FormWindow -Title "New User Request" -Height 950
    $scrollViewer = New-FormScrollViewer
    $mainPanel = New-MainPanel -Margin '10'
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

    $bypassPanel = New-FormDockPanel -Margin "0,0,0,5"
    $bypassFormattingCheckBox = New-FormCheckBox `
        -Content "Bypass Mobile Number Formatting" `
        -ToolTip "Check this box to skip automatic formatting of the mobile number" `
        -Margin "0,5,0,5"
    $bypassFormattingCheckBox.VerticalAlignment = 'Center'
    $bypassPanel.Children.Add($bypassFormattingCheckBox)
    $mobileSection.Stack.Children.Add($bypassPanel)
    $mainPanel.Children.Add($mobileSection.Group)

    # Zoom License Section (only visible if Zoom User is checked)
    $zoomSection = New-FormGroupBox -Header "Zoom License Type"
    $zoomComboBox = New-FormComboBox `
        -ToolTip "Select the type of Zoom license for the user" `
        -DisplayMemberPath "DisplayName"

    # Add Zoom license options
    $zoomLicenses = @(
        [PSCustomObject]@{
            DisplayName = "Zoom Basic"
            Value       = "false"
        },
        [PSCustomObject]@{
            DisplayName = "Zoom Workplace"
            Value       = "true"
        }
    )

    foreach ($license in $zoomLicenses) {
        $zoomComboBox.Items.Add($license)
    }

    # Create a panel for Zoom options
    $zoomPanel = New-FormDockPanel

    # Add zoom user checkbox - now enabled by default
    $zoomUserCheckBox = New-FormCheckBox -Content "Enable Zoom" -ToolTip "Enable to create Zoom account for user" -IsChecked $true -Margin "0,5,10,5"

    $zoomPanel.Children.Add($zoomUserCheckBox)
    $zoomSection.Stack.Children.Add($zoomPanel)
    $zoomSection.Stack.Children.Add($zoomComboBox)

    # Enable the combo box since checkbox is checked by default
    $zoomComboBox.IsEnabled = $true

    # Pre-select Basic license
    $zoomComboBox.SelectedIndex = 0  # Select first item (Zoom Basic)

    # Add event handler for checkbox
    $zoomUserCheckBox.Add_Checked({
            $zoomComboBox.IsEnabled = $true
            $zoomComboBox.SelectedIndex = 0  # Select Basic when re-enabled
        })

    $zoomUserCheckBox.Add_Unchecked({
            $zoomComboBox.IsEnabled = $false
            $zoomComboBox.SelectedIndex = -1
        })

    $mainPanel.Children.Add($zoomSection.Group)

    # Required License Section
    $requiredSection = New-FormGroupBox -Header "365:Required License (Select One) *"
    $requiredComboBox = New-FormComboBox `
        -ToolTip "Select one of the required base licenses for the user" `
        -DisplayMemberPath "DisplayName"

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

    $requiredSection.Stack.Children.Add($requiredComboBox)
    $mainPanel.Children.Add($requiredSection.Group)

    # Ancillary Licenses Section
    $ancillarySection = New-FormGroupBox -Header "365: Ancillary Licenses"
    $scrollingPanel = New-ScrollingStackPanel -MaxHeight 200
    $licensesStack = $scrollingPanel.StackPanel

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
            $skucb = New-FormCheckBox -Content $license.DisplayName -ToolTip $license.SkuId
            $skucb.Tag = $license.SkuId
            if ($license.DisplayName -like "*Microsoft Entra ID P2*") {
                $skucb.IsChecked = $true
            }
            $SkuCheckBoxes += $skucb
            $licensesStack.Children.Add($skucb)
        }
    }
    $ancillarySection.Group.Content = $scrollingPanel.ScrollViewer
    $mainPanel.Children.Add($ancillarySection.Group)

    # Create and add OK and Cancel buttons
    $buttonPanel = New-ButtonPanel -Margin "0,10,0,0"

    # Add test mode checkbox to button panel
    $testModeButton = New-FormCheckBox `
        -Content "Test Mode" `
        -ToolTip "Enable to redirect emails to: $($config.TestMode.Email)" `
        -IsChecked ($script:TestMode -eq $true) `
        -Margin "0,5,10,0"

    $buttonPanel.Children.Add($testModeButton)

    $okButton = New-FormButton -Content "OK" -Margin "0,0,10,0" -ClickHandler {
        # Validate New User input
        if (-not $newUserTextBox.Text) {
            Show-ValidationError -Message "New User is a mandatory field. Please enter a valid Display Name."
            return
        }
        if (-not (Test-DisplayName $newUserTextBox.Text)) {
            Show-ValidationError -Message "Invalid format for New User. Please use 'First Last' name format."
            return
        }

        # Validate User To Copy input
        if (-not $userToCopyTextBox.Text) {
            Show-ValidationError -Message "User To Copy is a mandatory field. Please enter a Display Name."
            return
        }
        if (-not (Test-DisplayName $userToCopyTextBox.Text)) {
            Show-ValidationError -Message "Invalid format for User To Copy. Please use 'First Last' name format."
            return
        }

        # Validate Mobile Number
        if ($mobileTextBox.Text -ne $mobileTextBox.Tag -and -not $bypassFormattingCheckBox.IsChecked) {
            $formattedNumber = Format-MobileNumber $mobileTextBox.Text
            if ($null -eq $formattedNumber) {
                Show-ValidationError -Message "Invalid mobile number format. Please enter a valid 10-digit mobile number."
                return
            }
        }

        # Validate required license selection
        if ($null -eq $requiredComboBox.SelectedItem) {
            Show-ValidationError -Message "Please select a required license." -Title "Required License Missing"
            return
        }

        # Check license availability
        if ($requiredComboBox.SelectedItem.DisplayName -match "Available: (\d+)") {
            $availableCount = [int]$Matches[1]
            if ($availableCount -eq 0) {
                $licenseName = $requiredComboBox.SelectedItem.DisplayName -replace ' \(Available: \d+\)$', ''
                Show-ValidationError -Message "$licenseName has no licenses available." -Title "No Available Licenses"
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
                    Show-ValidationError -Message "$licenseName has no licenses available." -Title "No Available Licenses"
                    return
                }
            }
        }

        # Get selected licenses
        $script:selectedLicenses = @()
        $script:selectedLicenses += $requiredComboBox.SelectedItem.SkuId
        $script:selectedLicenses += ($SkuCheckBoxes | Where-Object { $_.IsChecked } | ForEach-Object { $_.Tag })

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
            TestModeEnabled        = ($testModeButton.IsChecked -eq $true)
            ZoomUserEnabled        = ($zoomUserCheckBox.IsChecked -eq $true)
            IsWorkPlaceLicense     = ($zoomUserCheckBox.IsChecked -and
                $zoomComboBox.SelectedItem -and
                $zoomComboBox.SelectedItem.DisplayName -eq 'Zoom Workplace')
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
        Write-StatusMessage -Message "Getting template user details for: $UserToCopy" -Type INFO

        $adUserParams = @{
            Filter      = "DisplayName -eq '$UserToCopy'"
            Properties  = @(
                'Company',
                'Title',
                'Manager',
                'physicalDeliveryOfficeName',
                'Department',
                'Description'.
                'facsimileTelephoneNumber',
                'l', # l is for Location because Microsoft AD attributes are stupid
                'c', # c is for Country because Microsoft AD attributes are stupid
                'wWWHomePage'
            )
            ErrorAction = 'Stop'
        }

        $templateUser = Get-ADUser @adUserParams

        # Check for null or multiple users
        if ($null -eq $templateUser) {
            Write-StatusMessage -Message "Could not find user $UserToCopy in AD to copy from" -Type ERROR
            Exit-Script -Message "Template user not found: $UserToCopy" -ExitCode UserNotFound
        }

        if ($templateUser.Count -gt 1) {
            Write-StatusMessage -Message "Found multiple users with DisplayName: $UserToCopy" -Type ERROR
            Exit-Script -Message "Multiple template users found - please check AD for duplicate DisplayName attributes" -ExitCode DuplicateUser
        }

        Write-StatusMessage -Message "Successfully retrieved template user details" -Type OK
        return $templateUser

    } catch {
        Write-StatusMessage -Message "Failed to get template user: $_" -Type ERROR
        Exit-Script -Message "Critical error getting template user" -ExitCode GeneralError
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

    function New-DuplicatePromptForm {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Title,

            [Parameter(Mandatory)]
            [string]$ExistingValue,

            [Parameter(Mandatory)]
            [string]$PromptText
        )

        try {
            Add-Type -AssemblyName System.Windows.Forms

            $form = New-Object System.Windows.Forms.Form
            $form.Text = $Title
            $form.Size = New-Object System.Drawing.Size(450, 210)
            $form.StartPosition = "CenterScreen"

            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10, 20)
            $label.Size = New-Object System.Drawing.Size(430, 50)
            $label.Text = $PromptText
            $form.Controls.Add($label)

            $textBox = New-Object System.Windows.Forms.TextBox
            $textBox.Location = New-Object System.Drawing.Point(10, 70)
            $textBox.Size = New-Object System.Drawing.Size(410, 20)
            $textBox.Text = $ExistingValue
            $form.Controls.Add($textBox)

            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Location = New-Object System.Drawing.Point(100, 120)
            $okButton.Size = New-Object System.Drawing.Size(75, 23)
            $okButton.Text = "OK"
            $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Controls.Add($okButton)

            $cancelButton = New-Object System.Windows.Forms.Button
            $cancelButton.Location = New-Object System.Drawing.Point(200, 120)
            $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
            $cancelButton.Text = "Cancel"
            $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Controls.Add($cancelButton)

            $form.AcceptButton = $okButton
            $form.CancelButton = $cancelButton

            return $form
        } catch {
            Write-StatusMessage -Message "Failed to create form: $_" -Type ERROR
            throw
        }
    }

    try {
        Write-StatusMessage -Message "Generating properties for new user: $NewUser" -Type INFO

        # Split the new user name
        $nameParts = $NewUser -split ' '
        if ($nameParts.Count -lt 2) {
            Write-StatusMessage -Message "Invalid user name format: Must include first and last name" -Type ERROR
            Exit-Script -Message "New user name must include first and last name" -ExitCode GeneralError
        }

        try {
            # Get domain from source user
            $domain = $SourceUserUPN -replace '.+?(?=@)'
            if ([string]::IsNullOrEmpty($domain)) {
                Write-StatusMessage -Message "Failed to extract domain from source user" -Type ERROR
                Exit-Script -Message "Invalid source user domain" -ExitCode GeneralError
            }

            # Parse name parts using original logic
            $firstName = $nameParts[-2]  # Second to last part
            $lastName = $nameParts[-1]   # Last part
            $displayName = $NewUser

            # Generate initial samAccountName and email
            $samAccountName = (($NewUser -replace '(?<=.{1}).+') + ($lastName)).ToLower()
            $email = ($samAccountName + $domain).ToLower()

            # Check for duplicate samAccountName in AD
            if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
                Write-StatusMessage -Message "SamAccountName $samAccountName already exists" -Type WARN

                # For SAM account name prompt:
                $formDuplicateSam = New-DuplicatePromptForm `
                    -Title "Duplicate SAM Account Name" `
                    -ExistingValue $samAccountName `
                    -PromptText "SAM account name '$samAccountName' already exists.`nPlease enter a different SAM account name:"

                $resultSam = $formDuplicateSam.ShowDialog()

                if ($resultSam -eq [System.Windows.Forms.DialogResult]::OK) {
                    $samAccountName = $textBox.Text
                    # Verify the new samAccountName is unique
                    if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
                        Write-StatusMessage -Message "New SamAccountName $samAccountName is also in use" -Type ERROR
                        Exit-Script -Message "Unable to generate unique SamAccountName" -ExitCode DuplicateUser
                    }
                    Write-StatusMessage -Message "Using custom SamAccountName: $samAccountName" -Type OK
                    # Update email with new samAccountName
                    $email = ($samAccountName + $domain).ToLower()
                } else {
                    Write-StatusMessage -Message "User cancelled SAM account name selection" -Type WARN
                    Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
                }
            }

            # Check for existing mailbox
            try {
                $mailbox = Get-Mailbox -Filter "EmailAddresses -like '*$email*'" -ErrorAction Stop
                if ($mailbox) {
                    Write-StatusMessage -Message "Email address $email (or similar) already exists for mailbox: $($mailbox.UserPrincipalName)" -Type WARN

                    # For email prompt:
                    $formDuplicateEmail = New-DuplicatePromptForm `
                        -Title "Duplicate Email Address" `
                        -ExistingValue $email `
                        -PromptText "Email address '$email' already exists.`nPlease enter a different email address:"

                    $resultEmail = $formDuplicateEmail.ShowDialog()

                    if ($resultEmail -eq [System.Windows.Forms.DialogResult]::OK) {
                        $email = $textBox.Text
                        # Verify the new email is unique
                        try {
                            $checkMailbox = Get-Mailbox -Filter "EmailAddresses -like '*$email*'" -ErrorAction Stop
                            if ($checkMailbox) {
                                Write-StatusMessage -Message "New email address $email is also in use by: $($checkMailbox.displayName)" -Type ERROR
                                Exit-Script -Message "Unable to generate unique email address" -ExitCode DuplicateUser
                            }
                        } catch [Microsoft.Exchange.Management.RestApiClient.RestApiException] {
                            Write-StatusMessage -Message "Using custom email address: $email" -Type OK
                        }
                    } else {
                        Write-StatusMessage -Message "User cancelled email address selection" -Type WARN
                        Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
                    }
                }
            } catch [Microsoft.Exchange.Management.RestApiClient.RestApiException] {
                # This is expected - mailbox should not exist
                Write-StatusMessage -Message "Exchange validation passed - mailbox does not exist" -Type OK
            }

            Write-StatusMessage -Message "Successfully generated user properties" -Type OK
            return @{
                FirstName      = $firstName
                LastName       = $lastName
                DisplayName    = $displayName
                Email          = $email
                SamAccountName = $samAccountName
            }
        } catch {
            Write-StatusMessage -Message "Failed to generate user properties: $_" -Type ERROR
            Exit-Script -Message "Failed to generate user properties" -ExitCode GeneralError
        }
    } catch {
        Write-StatusMessage -Message "Critical error in New-UserProperties: $_" -Type ERROR
        Exit-Script -Message "Critical error generating user properties" -ExitCode GeneralError
    }
}

function New-ReadablePassword {
    <#
    .SYNOPSIS
        Generates a human-readable password using random words and special characters.

    .DESCRIPTION
        Creates a memorable password by combining random words from a curated wordlist with special characters
        and numbers. Allows user to accept or reject generated passwords. Returns both plain text and SecureString versions.

    .PARAMETER WordCount
        Number of words to use in the password (2-20). Default is 3.

    .PARAMETER AddSpaces
        Adds spaces between words in the final password.

    .PARAMETER WordListPath
        Optional path to a custom wordlist file. If not provided, uses default GitHub wordlist.

    .PARAMETER GitHubToken
        GitHub Personal Access Token for accessing private word list repository.

    .EXAMPLE
        $password = New-ReadablePassword -GitHubToken "your-github-pat"
        # Prompts user with generated password like: "Mountain7$ Forest#2 Lake"

    .NOTES
        Name: New-ReadablePassword
        Author: Chris Williams
        Version: 1.0.0
        DateCreated: 2025-Jan-25
    #>

    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [ValidateRange(2, 20)]
        [int]$WordCount = 3,
        [switch]$AddSpaces,
        [string]$WordListPath,
        [Parameter(Mandatory)]
        [string]$GitHubToken
    )

    try {
        Write-StatusMessage -Message "Generating secure word-based password" -Type INFO

        do {
            # Get word list
            $FullList = if ($WordListPath -and (Test-Path $WordListPath)) {
                Get-Content $WordListPath
            } else {
                $headers = @{
                    "Authorization" = "token $GitHubToken"
                    "Accept"        = "application/json"
                }
                (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/ryanchrisw/CompassDeploy/refs/heads/main/Wordlist/wordlist" -Headers $headers).Content.Trim().split("`n")
            }

            # Group words by length
            $WordsByLength = $FullList | Group-Object Length -AsHashTable

            # Select appropriate word lengths based on count
            $WordList = switch ($WordCount) {
                { $_ -le 3 } { $WordsByLength[7] + $WordsByLength[8] + $WordsByLength[9] }
                4 { $WordsByLength[4..7] | ForEach-Object { $_ } }
                5 { $WordsByLength[4..6] | ForEach-Object { $_ } }
                default { $WordsByLength[3..5] | ForEach-Object { $_ } }
            }

            # Generate password
            $SpecialChars = [char[]]((33, 35) + (36..38) + (40..46) + (60..62) + (64))
            $Numbers = [char[]](48..57)

            $Password = 1..$WordCount | ForEach-Object {
                if ($_ -eq $WordCount) {
                    $WordList | Get-Random
                } else {
                    "$($WordList | Get-Random)$([char[]]($SpecialChars + $Numbers) | Get-Random)"
                }
            }

            $plainPassword = if ($AddSpaces) { $Password -join ' ' } else { $Password -join '' }

            # Display password and get confirmation
            Write-Host "`nGenerated Password: $plainPassword" -ForegroundColor Cyan
            $response = Read-Host "Accept this password? (y/n)"

        } while ($response -ne 'y')

        Write-StatusMessage -Message "Password accepted" -Type OK
        return @{
            PlainPassword  = $plainPassword
            SecurePassword = ConvertTo-SecureString -String $plainPassword -AsPlainText -Force
        }
    } catch {
        Write-StatusMessage -Message "Critical error in password generation: $_" -Type ERROR
        Exit-Script -Message "Critical password generation failure" -ExitCode GeneralError
    }
}

function Confirm-UserCreation {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [hashtable]$NewUserProperties,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$DestinationOU,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TemplateUser,

        [Parameter(Mandatory)]
        [string]$Password
    )

    try {
        Write-StatusMessage -Message "Preparing user creation summary" -Type INFO

        # Build summary with consistent formatting
        $summary = @"
New User Details:
----------------
- Display Name    = $($NewUserProperties.DisplayName)
- Email Address   = $($NewUserProperties.Email)
- Password        = $Password
- First Name      = $($NewUserProperties.FirstName)
- Last Name       = $($NewUserProperties.LastName)
- SamAccountName  = $($NewUserProperties.SamAccountName)
- Destination OU  = $DestinationOU
- Template User   = $TemplateUser
"@

        # Display summary and get confirmation
        Write-StatusMessage -Message $summary -Type SUMMARY
        Write-StatusMessage -Message "Please review the details above carefully" -Type WARN
        $confirmation = Read-Host "Do you want to proceed with user creation? (Y/N)"

        if ($confirmation.ToUpper() -ne 'Y') {
            Write-StatusMessage -Message "User creation cancelled." -Type WARN
            Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
        }

        Write-StatusMessage -Message "User creation confirmed." -Type OK

    } catch {
        Write-StatusMessage -Message "Error during user creation confirmation: $_" -Type ERROR
        Exit-Script -Message "Failed to confirm user creation" -ExitCode GeneralError
    }
}

function New-ADUserFromTemplate {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)]
        [hashtable]$NewUser,

        [Parameter(Mandatory)]
        [Microsoft.ActiveDirectory.Management.ADUser]$SourceUser,

        [Parameter(Mandatory)]
        [string]$Phone,

        [Parameter(Mandatory)]
        [System.Security.SecureString]$Password,

        [Parameter(Mandatory)]
        [string]$DestinationOU
    )

    try {
        Write-StatusMessage -Message "Preparing to create new AD user: $($NewUser.DisplayName)" -Type INFO

        $newUserParams = @{
            Name              = "$($NewUser.FirstName) $($NewUser.LastName)"
            SamAccountName    = $NewUser.SamAccountName
            UserPrincipalName = $NewUser.Email
            EmailAddress      = $NewUser.Email
            DisplayName       = $NewUser.DisplayName
            GivenName         = $NewUser.FirstName
            Surname           = $NewUser.LastName
            MobilePhone       = $Phone
            OtherAttributes   = @{
                'proxyAddresses' = "SMTP:$($NewUser.Email)"
            }
            AccountPassword   = $Password
            Path              = $DestinationOU
            Instance          = $SourceUser
            Enabled           = $true
            ErrorAction       = 'Stop'
        }

        if ($PSCmdlet.ShouldProcess($UserProperties.DisplayName, "Create new AD user")) {
            try {
                Write-StatusMessage -Message "Creating new AD user..." -Type INFO
                New-ADUser @newUserParams
                Write-StatusMessage -Message "AD user created successfully" -Type OK
            } catch {
                Write-StatusMessage -Message "Failed to create AD user: $_" -Type ERROR
                Exit-Script -Message "AD user creation failed" -ExitCode GeneralError
            }
        }
    } catch {
        Write-StatusMessage -Message "Critical error in New-ADUserFromTemplate: $_" -Type ERROR
        Exit-Script -Message "Critical error during AD user creation" -ExitCode GeneralError
    }
}

function Copy-UserADGroups {
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
        Write-StatusMessage -Message "Copying AD group memberships from $SourceUser to $TargetUser" -Type INFO

        # Get source user and their groups
        try {
            $sourceGroups = Get-ADUser -Filter "DisplayName -eq '$SourceUser'" -Properties MemberOf -ErrorAction Stop
            if (-not $sourceGroups) {
                Write-StatusMessage -Message "Source user not found: $SourceUser" -Type ERROR
                return
            }
            Write-StatusMessage -Message "Found source user with $($sourceGroups.MemberOf.Count) groups" -Type INFO
        } catch {
            Write-StatusMessage -Message "Failed to get source user groups: $_" -Type ERROR
            return
        }

        # Get target user
        try {
            $getTargetUser = Get-ADUser -Filter "DisplayName -eq '$TargetUser'" -Properties MemberOf -ErrorAction Stop
            if (-not $getTargetUser) {
                Write-StatusMessage -Message "Target user not found: $TargetUser" -Type ERROR
                return
            }
            Write-StatusMessage -Message "Found target user" -Type INFO
        } catch {
            Write-StatusMessage -Message "Failed to get target user: $_" -Type ERROR
            return
        }

        # Calculate groups to add (groups source has that target doesn't)
        $groupsToAdd = $sourceGroups.MemberOf | Where-Object { $getTargetUser.MemberOf -notcontains $_ }

        if (-not $groupsToAdd) {
            Write-StatusMessage -Message "No new groups to add - target user already has all source groups" -Type OK
            return
        }

        Write-StatusMessage -Message "Adding $($groupsToAdd.Count) groups to  $TargetUser" -Type INFO

        # Add groups with individual error handling
        $successCount = 0
        foreach ($group in $groupsToAdd) {
            try {
                Add-ADGroupMember -Identity $group -Members $getTargetUser -ErrorAction Stop
                Write-StatusMessage -Message "Added to group: $((Get-ADGroup $group).Name)" -Type OK
                $successCount++
            } catch {
                Write-StatusMessage -Message "Failed to add to group $((Get-ADGroup $group).Name): $_" -Type WARN
            }
        }

        # Final status check
        if ($successCount -eq $groupsToAdd.Count) {
            Write-StatusMessage -Message "Successfully added all $successCount groups" -Type OK
        } else {
            Write-StatusMessage -Message "Added $successCount of $($groupsToAdd.Count) groups - some assignments failed" -Type WARN
        }

    } catch {
        Write-StatusMessage -Message "Error in Copy-UserADGroups: $_" -Type ERROR
    }
}

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
        [int]$SyncTimeout = 300  # 5 minutes default
    )

    Write-StatusMessage -Message "Starting AD sync process for $UserEmail" -Type INFO
    $syncStartTime = Get-Date

    try {
        # Initial wait before starting sync
        Write-StatusMessage -Message "Waiting $InitialWaitSeconds seconds before starting sync..." -Type INFO
        Start-Sleep -Seconds $InitialWaitSeconds

        # Start AD sync with retry logic
        $syncStarted = $false
        for ($i = 1; $i -le 3; $i++) {
            try {
                Write-StatusMessage -Message "Attempting to start AD sync (Attempt $i of 3)" -Type INFO
                Import-Module -Name ADSync -UseWindowsPowerShell -WarningAction:SilentlyContinue -ErrorAction Stop
                $null = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
                $syncStarted = $true
                Write-StatusMessage -Message "AD sync started successfully" -Type OK
                break
            } catch {
                Write-StatusMessage -Message "Sync attempt $i failed: $_" -Type WARN
                if ($i -eq 3) {
                    Write-StatusMessage -Message "Failed to start AD sync after 3 attempts" -Type ERROR
                    Exit-Script -Message "AD sync failed to start" -ExitCode GeneralError
                }
                Start-Sleep -Seconds 5
            }
        }

        # Monitor sync progress
        $retryCount = 0
        do {
            try {
                # Check timeout
                $elapsed = ((Get-Date) - $syncStartTime).TotalSeconds
                if ($elapsed -ge $SyncTimeout) {
                    Write-StatusMessage -Message "Sync timeout after $($elapsed.ToString('F0')) seconds" -Type ERROR
                    Exit-Script -Message "AD sync timeout" -ExitCode GeneralError
                }

                # Check sync status
                $syncStatus = Get-ADSyncScheduler
                if ($syncStatus.SyncCycleInProgress) {
                    Write-StatusMessage -Message "Sync in progress... ($($elapsed.ToString('F0')) seconds elapsed)" -Type INFO
                    Start-Sleep -Seconds 10
                    continue
                }

                # Try to get user
                Write-StatusMessage -Message "Checking for user in Azure AD..." -Type INFO
                $user = Get-MgUser -UserId $UserEmail -Property Id, Mail, DisplayName, GivenName, Surname, Department, officeLocation | Select-Object Id, Mail, DisplayName, GivenName, Surname, Department, officeLocation -ErrorAction Stop
                if ($user) {
                    Write-StatusMessage -Message "User $UserEmail successfully synced to Azure AD" -Type OK
                    return $user
                }

            } catch {
                $retryCount++
                if ($retryCount -ge $MaxRetries) {
                    Write-StatusMessage -Message "Max retry attempts ($MaxRetries) reached" -Type ERROR
                    Exit-Script -Message "Failed to verify user sync after maximum retries" -ExitCode UserNotFound
                }
                Write-StatusMessage -Message "Retry $($retryCount) of $($MaxRetries): User not found in Azure AD yet" -Type WARN
                Start-Sleep -Seconds $RetryIntervalSeconds
            }
        } while ($true)

    } catch {
        Write-StatusMessage -Message "Critical error during AD sync process: $_" -Type ERROR
        Exit-Script -Message "AD sync process failed" -ExitCode GeneralError
    }
}

function Set-UserLicenses {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$License,

        [Parameter()]
        [switch]$Required
    )

    try {
        $licenseType = if ($Required) { "required" } else { "ancillary" }
        Write-StatusMessage -Message "Starting $licenseType license assignment for user: $($User.displayName)" -Type INFO

        foreach ($lic in $License) {
            try {
                Write-StatusMessage -Message "Assigning license: $($lic.DisplayName)" -Type INFO

                $null = Set-MgUserLicense -UserId $User.Id `
                    -AddLicenses @{SkuId = $($lic.SkuId) } `
                    -RemoveLicenses @() `
                    -ErrorAction Stop

                Write-StatusMessage -Message "Successfully assigned license: $($lic.DisplayName)" -Type OK
            } catch {
                if ($Required) {
                    Write-StatusMessage -Message "Failed to assign license $($lic.DisplayName): $_" -Type ERROR
                    Exit-Script -Message "Required license assignment failed" -ExitCode GeneralError
                } else {
                    Write-StatusMessage -Message "Failed to assign license $($lic.DisplayName): $_" -Type WARN
                }
            }
        }
    } catch {
        if ($Required) {
            Write-StatusMessage -Message "Error in Set-UserLicenses: $_" -Type ERROR
            Exit-Script -Message "Critical error during license assignment" -ExitCode GeneralError
        } else {
            Write-StatusMessage -Message "Error in Set-UserLicenses: $_" -Type WARN
        }
    }
}

function Copy-UserEntraGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $SourceUser,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $TargetUser
    )

    try {
        Write-StatusMessage -Message "Starting group membership copy from $($SourceUser.DisplayName) to $($TargetUser.DisplayName)" -Type INFO

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

        try {
            $groups = Get-MgUserMemberOf -UserId $SourceUser.Id -ErrorAction Stop |
            Where-Object @filterParams |
            Select-Object @selectParams

            if (-not $groups) {
                Write-StatusMessage -Message "Source user has no groups to copy" -Type INFO
                return
            }

            Write-StatusMessage -Message "Found $($groups.Count) groups to copy" -Type INFO

            foreach ($group in $groups) {
                try {
                    if ($group.securityEnabled -eq 'True' -or $group.groupType -eq 'Unified') {
                        $null = New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $TargetUser.Id -ErrorAction Stop
                        Write-StatusMessage -Message "Added to Security/Unified Group: $($group.DisplayName)" -Type OK
                    } else {
                        $null = Add-DistributionGroupMember -Identity $group.Id -Member $TargetUser.Id -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                        Write-StatusMessage -Message "Added to Distribution Group: $($group.DisplayName)" -Type OK
                    }
                } catch {
                    Write-StatusMessage -Message "Failed to add to group $($group.DisplayName): $_" -Type WARN
                }
            }

            Write-StatusMessage -Message "Group membership copy completed" -Type OK

        } catch {
            Write-StatusMessage -Message "Failed to get source user groups: $_" -Type WARN
        }

    } catch {
        Write-StatusMessage -Message "Error in Copy-UserEntraGroups: $_" -Type WARN
    }
}

function Add-UserToZoom {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [string]$ClientSecret,

        [Parameter(Mandatory)]
        [string]$AccountId,

        [Parameter()]
        [string]$UserToCopy,

        [Parameter()]
        [string]$Password,

        [Parameter()]
        [string]$IsWorkPlaceLicense
    )

    try {
        # First create the Zoom user via API
        Write-StatusMessage -Message "Creating Zoom user via API" -Type INFO

        # Get Bearer Token
        try {
            $pair = "$ClientId`:$ClientSecret"
            $bytes = [System.Text.Encoding]::UTF8.GetBytes($pair)
            $base64 = [Convert]::ToBase64String($bytes)

            $tokenParams = @{
                Method      = "POST"
                Uri         = "https://zoom.us/oauth/token"
                ContentType = "application/x-www-form-urlencoded"
                Body        = @{
                    grant_type = 'account_credentials'
                    account_id = $AccountId
                }
                Headers     = @{
                    Host          = 'zoom.us'
                    Authorization = "Basic $base64"
                }
            }

            $tokenResponse = Invoke-WebRequest @tokenParams
            $BearerToken = ($tokenResponse.Content | ConvertFrom-Json).access_token
            Write-StatusMessage -Message "Successfully obtained Zoom API token" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to get Zoom API token: $_" -Type ERROR
            return
        }

        # Create Zoom User
        try {
            $headers = @{
                "Content-Type"  = "application/json"
                "Authorization" = "Bearer $BearerToken"
            }

            # Determine if user should have phone features based on department
            if ($IsWorkPlaceLicense -eq 'true') {
                $body = @{
                    action    = "autoCreate"
                    user_info = @{
                        email        = $User.Mail
                        first_name   = $User.GivenName
                        last_name    = $User.Surname
                        display_name = $User.DisplayName
                        password     = $Password
                        type         = 2
                        feature      = @{
                            zoom_phone    = $true
                            zoom_one_type = 16
                        }
                    }
                }
            } elseif ($User.Department -eq 'Reactive') {
                $body = @{
                    action    = "autoCreate"
                    user_info = @{
                        email        = $User.Mail
                        first_name   = $User.GivenName
                        last_name    = $User.Surname
                        display_name = $User.DisplayName
                        password     = $Password
                        type         = 1
                    }
                }
            } else {
                $body = @{
                    action    = "autoCreate"
                    user_info = @{
                        email        = $User.Mail
                        first_name   = $User.GivenName
                        last_name    = $User.Surname
                        display_name = $User.DisplayName
                        password     = $Password
                        type         = 1
                        feature      = @{
                            zoom_phone = $true
                        }
                    }
                }
            }


            $response = Invoke-WebRequest -Uri 'https://api.zoom.us/v2/users' -Method POST -Headers $headers -Body ($body | ConvertTo-Json -Depth 3)

            if ($response.StatusCode -eq 201) {
                Write-StatusMessage -Message "Successfully created Zoom user via API" -Type OK

                # Wait for user to be fully provisioned
                Write-StatusMessage -Message "Waiting for Zoom user provisioning to complete" -Type INFO
                $maxAttempts = 6  # Maximum number of attempts
                $attempt = 0
                $userProvisioned = $false

                while (-not $userProvisioned -and $attempt -lt $maxAttempts) {
                    try {
                        $checkResponse = Invoke-WebRequest `
                            -Uri "https://api.zoom.us/v2/users/$($User.Mail)" `
                            -Method GET `
                            -Headers $headers

                        if ($checkResponse.StatusCode -eq 200) {
                            $userProvisioned = $true
                            Write-StatusMessage -Message "Zoom user provisioning completed" -Type OK
                        }
                    } catch {
                        $attempt++
                        if ($attempt -lt $maxAttempts) {
                            Write-StatusMessage -Message "Waiting for user provisioning (attempt $attempt of $maxAttempts)" -Type INFO
                            Start-Sleep -Seconds 5
                        } else {
                            Write-StatusMessage -Message "Timeout waiting for user provisioning" -Type ERROR
                            return
                        }
                    }
                }
            } else {
                Write-StatusMessage -Message "Failed to create Zoom user. Status code: $($response.StatusCode)" -Type ERROR
                return
            }
        } catch {
            Write-StatusMessage -Message "Failed to create Zoom user via API: $_" -Type ERROR
            return
        }

        # After creating the user, assign calling plan if not Reactive department
        if ($IsWorkPlaceLicense -eq 'true') {
            Write-StatusMessage -Message "Provisioning Zoom Phone Plan for $($User.DisplayName)" -Type INFO
            Write-StatusMessage -Message "Phone plan included with Workplace license" -Type OK
        } elseif ($User.Department -eq 'Reactive') {
            Write-StatusMessage -Message "Provisioning Zoom Contact Center for $($User.DisplayName)" -Type INFO

            # Define Zoom Conact Center template ID mapping
            $templateMap = @{
                'South Florida' = "unIR_2-xQemLQ9pfvDKk3w"
                'Northeast'     = "-Pil2tjcRaOTaIl6LUV_6g"
                'North Florida' = "nQdKAwkhRh-ds0Yvp7QKsw"
                'Mid-West'      = "ySmugPJ0Qc6O3IaWHi9jHA"
                'Mid-Atlantic'  = "K-6ajeNsS4KJRf-6xfRnaA"
            }

            $body = @{
                user_email  = $User.Mail
                template_id = $null  # Will be set based on location
            }

            if ($templateMap.ContainsKey($User.officeLocation)) {
                $body.template_id = $templateMap[$User.officeLocation]
                Write-StatusMessage -Message "Using template ID for location: $($User.officeLocation)" -Type INFO

                $headers["Content-Type"] = "application/json"

                try {
                    $response = Invoke-WebRequest -Uri 'https://api.zoom.us/v2/contact_center/users' `
                        -Method POST `
                        -Headers $headers `
                        -Body ($body | ConvertTo-Json)

                    if ($response.StatusCode -eq 201) {
                        Write-StatusMessage -Message "Successfully provisioned Contact Center" -Type OK

                        # Get skills from user to copy
                        try {
                            Write-StatusMessage -Message "Getting skills from template user" -Type INFO
                            $skillsResponse = Invoke-WebRequest -Uri "https://api.zoom.us/v2/contact_center/users/$UserToCopy/skills" `
                                -Method GET `
                                -Headers $headers

                            if ($skillsResponse.StatusCode -eq 200) {
                                $skillsContent = $skillsResponse.Content | ConvertFrom-Json
                                $user_proficiency_level = $skillsContent.skills[0].user_proficiency_level

                                # Set skills for new user
                                $skillBody = @{
                                    skills = @(
                                        @{
                                            skill_id              = "sS21RNM2PSQOy8djVEbxang" # Tech Proficiency skill Id
                                            max_proficiency_level = $user_proficiency_level
                                        }
                                    )
                                }

                                $setSkillResponse = Invoke-WebRequest `
                                    -Uri "https://api.zoom.us/v2/contact_center/users/$($User.Mail)/skills" `
                                    -Method POST `
                                    -Headers $headers `
                                    -ContentType 'application/json' `
                                    -Body ($skillBody | ConvertTo-Json)

                                if ($setSkillResponse.StatusCode -eq 201) {
                                    Write-StatusMessage -Message "Successfully copied skill proficiency level" -Type OK
                                } else {
                                    Write-StatusMessage -Message "Failed to set skill proficiency level. Status code: $($setSkillResponse.StatusCode)" -Type ERROR
                                }
                            } else {
                                Write-StatusMessage -Message "Failed to get template user skills. Status code: $($skillsResponse.StatusCode)" -Type ERROR
                            }
                        } catch {
                            Write-StatusMessage -Message "Failed to copy skill proficiency: $_" -Type ERROR
                        }
                    }
                } catch {
                    Write-StatusMessage -Message "Failed to provision Contact Center: $_" -Type ERROR
                    return
                }
            } else {
                Write-StatusMessage -Message "No template ID found for location: $($User.officeLocation)" -Type ERROR
                return
            }
        } else {
            Write-StatusMessage -Message "Provisioning Zoom Phone Plan for $($User.DisplayName)" -Type INFO
            try {
                # Get available calling plans
                $response = Invoke-WebRequest -Uri 'https://api.zoom.us/v2/phone/calling_plans' -Method GET -Headers $headers
                $responseContent = $response.Content | ConvertFrom-Json
                $callingPlan = $responseContent.calling_plans | Where-Object { $_.name -eq 'US/CA Unlimited Calling Plan' }

                if ($callingPlan -and $callingPlan.available -gt 0) {
                    Write-StatusMessage -Message "Zoom licenses are available for assignment" -Type INFO

                    # Assign calling plan to user
                    $body = @{
                        calling_plans = @(
                            @{
                                type = 200  # Type for the "US/CA Unlimited Calling Plan"
                            }
                        )
                    } | ConvertTo-Json

                    $assignPlanResponse = Invoke-WebRequest `
                        -Uri "https://api.zoom.us/v2/phone/users/$($User.Mail)/calling_plans" `
                        -Method POST `
                        -Headers $headers `
                        -ContentType 'application/json' `
                        -Body $body

                    if ($assignPlanResponse.StatusCode -eq 201) {
                        Write-StatusMessage -Message "Calling plan successfully assigned to the user" -Type OK
                    } else {
                        Write-StatusMessage -Message "Failed to assign calling plan" -Type ERROR
                    }
                } else {
                    Write-StatusMessage -Message "No Zoom licenses available" -Type ERROR
                    return
                }
            } catch {
                Write-StatusMessage -Message "Failed to assign calling plan: $_" -Type ERROR
                return
            }
        }

        # Now handle the Enterprise App assignment
        if ($IsWorkPlaceLicense -eq 'true') {
            $zoom_app_role_name = "Licensed"
        } else {
            $zoom_app_role_name = "Basic"
        }

        $zoom_app_name = "Zoom Workplace Phones"

        # Get Zoom app and sync details
        try {
            $zoom_ServicePrincipal = Get-MgServicePrincipal -Filter "displayName eq '$zoom_app_name'" -ErrorAction Stop
            if (-not $zoom_ServicePrincipal) {
                Write-StatusMessage -Message "Zoom Service Principal not found" -Type WARN
                return
            }

            $zoom_syncJob = Get-MgServicePrincipalSynchronizationJob -ServicePrincipalId $zoom_ServicePrincipal.Id -ErrorAction Stop
            $zoom_syncJobRuleId = (Get-MgServicePrincipalSynchronizationJobSchema -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_syncJob.Id).SynchronizationRules.Id

            Write-StatusMessage -Message "Retrieved Zoom app details successfully" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to get Zoom app details: $_" -Type WARN
            return
        }

        # Assign user to Zoom app
        try {
            $params = @{
                "PrincipalId" = $User.Id
                "ResourceId"  = $zoom_ServicePrincipal.Id
                "AppRoleId"   = ($zoom_ServicePrincipal.AppRoles | Where-Object { $_.DisplayName -eq $zoom_app_role_name }).Id
            }

            $null = New-MgUserAppRoleAssignment -UserId $User.Id -BodyParameter $params -ErrorAction Stop
            Write-StatusMessage -Message "Successfully assigned Zoom role to user" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to assign Zoom role: $_" -Type WARN
            return
        }

        # Start Zoom sync
        try {
            $params = @{
                parameters = @(
                    @{
                        subjects = @(
                            @{
                                objectId       = "$($User.Id)"
                                objectTypeName = "User"
                            }
                        )
                        ruleId   = $zoom_syncJobRuleId
                    }
                )
            }

            $null = New-MgServicePrincipalSynchronizationJobOnDemand -ServicePrincipalId $zoom_ServicePrincipal.Id -SynchronizationJobId $zoom_syncJob.Id -BodyParameter $params
            Write-StatusMessage -Message "Successfully started Zoom sync" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to start Zoom sync: $_" -Type WARN
        }
    } catch {
        Write-StatusMessage -Message "Error in Add-UserToZoom: $_" -Type WARN
    }
}

function Set-UserBookWithMeId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SamAccountName,

        [Parameter()]
        [int]$MaxRetries = 6,

        [Parameter()]
        [int]$RetryIntervalSeconds = 30
    )

    try {
        Write-StatusMessage -Message "Configuring BookWithMeId for $($User.displayName)" -Type INFO

        # Get Exchange GUID with retry logic
        $retryCount = 0
        $mailbox = $null
        $success = $false

        do {
            try {
                $mailbox = Get-Mailbox -Identity $User.Mail -ErrorAction Stop
                if ($mailbox) {
                    $success = $true
                    Write-StatusMessage -Message "Retrieved mailbox successfully" -Type OK
                    break
                }
            } catch {
                $retryCount++
                if ($retryCount -ge $MaxRetries) {
                    Write-StatusMessage -Message "Failed to get mailbox after $MaxRetries attempts" -Type WARN
                    return
                }
                Write-StatusMessage -Message "Mailbox not ready, attempt $retryCount of $MaxRetries. Waiting $RetryIntervalSeconds seconds..." -Type INFO
                Start-Sleep -Seconds $RetryIntervalSeconds
            }
        } while ($retryCount -lt $MaxRetries)

        if (-not $success) {
            Write-StatusMessage -Message "Failed to get mailbox after all retries" -Type WARN
            return
        }

        # Get Exchange GUID
        $exchangeGuid = $mailbox.ExchangeGuid.Guid
        if ([string]::IsNullOrEmpty($exchangeGuid)) {
            Write-StatusMessage -Message "Exchange GUID not found for $($User.displayName)" -Type WARN
            return
        }

        # Generate BookWithMeId
        $formattedGuid = $exchangeGuid -replace "-"
        $bookWithMeId = "${formattedGuid}@compassmsp.com?anonymous&ep=plink"

        if ($bookWithMeId -eq '@compassmsp.com?anonymous&ep=plink') {
            Write-StatusMessage -Message "Generated BookWithMeId is invalid (missing ExchangeGuid)" -Type WARN
            Write-StatusMessage -Message "Please add BookWithMeId to extensionAttribute15 manually for $SamAccountName" -Type WARN
            return
        }

        # Set AD attribute
        try {
            Set-ADUser -Identity $SamAccountName -Add @{extensionAttribute15 = $bookWithMeId } -ErrorAction Stop
            Write-StatusMessage -Message "Successfully set BookWithMeId for $($User.displayName)" -Type OK
        } catch {
            Write-StatusMessage -Message "Failed to set extensionAttribute15: $_" -Type WARN
            Write-StatusMessage -Message "Please set BookWithMeId ($bookWithMeId) manually for $SamAccountName" -Type WARN
        }

    } catch {
        Write-StatusMessage -Message "Error in Set-UserBookWithMeId: $_" -Type WARN
    }
}

function Start-NewUserFinalize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser]
        $User,

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

    Write-StatusMessage -Message "Preparing final summary" -Type INFO

    # Validate group counts
    if ($TemplateGroupCount -ne $AssignedGroupCount) {
        Write-StatusMessage -Message "Group count mismatch: Template=$TemplateGroupCount, Assigned=$AssignedGroupCount" -Type WARN
    }

    # Build summary parts
    $summaryParts = @(
        "Summary of Actions:",
        "----------------------------------------",
        "$($User.displayName) should now be created unless any errors occurred during the process.",
        "If any info below is blank then something went wrong in the script. ",
        "User Creation Status:",
        "- EntraID: $($User.Id)",
        "- Display Name: $($User.displayName)",
        "- Email Address: $($User.mail)",
        "- Password: $Password",
        "- Template User   = $TemplateUser",
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
    } catch {
        Write-StatusMessage -Message "Failed to display summary message: $_" -Type WARN
    }

}


#EndRegion Functions

Write-Host "`r  [✓] Functions loaded" -ForegroundColor Green
Write-Host "`n  Ready to process user request..." -ForegroundColor Cyan
Write-Host "`n  Beginning New User Request..." -ForegroundColor Cyan
#Region Main Execution

# Step 1: Initialization
Write-ProgressStep -StepName $progressSteps[1].Name -Status $progressSteps[1].Description

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

Connect-ServiceEndpoints -ExchangeOnline -Graph

# Step 2: User Input
Write-ProgressStep -StepName $progressSteps[2].Name -Status $progressSteps[2].Description
$userInput = Get-NewUserRequestInput

# Set variables after input
if ($userInput.TestModeEnabled -eq 'True') { $script:TestMode = $true }
$script:TestEmailAddress = $config.TestMode.Email

# Step 3: Validation and Preparation (AD)
Write-ProgressStep -StepName $progressSteps[3].Name -Status $progressSteps[3].Description
$UserToCopyAD = Get-TemplateUser -UserToCopy $userInput.InputUserToCopy
$destinationOU = $UserToCopyAD.DistinguishedName.split(",", 2)[1]                 # Validates template user
$newUserProperties = New-UserProperties -NewUser $userInput.InputNewUser -SourceUserUPN $UserToCopyAD.UserPrincipalName

# Step 4: AD User Creation
Write-ProgressStep -StepName $progressSteps[4].Name -Status $progressSteps[4].Description
$passwordResult = New-ReadablePassword -GitHubToken $config.GitHub.Token

# Show summary and get confirmation before creating

Confirm-UserCreation -NewUserProperties $newUserProperties `
    -DestinationOU $destinationOU `
    -TemplateUser $userInput.InputUserToCopy `
    -Password $passwordResult.PlainPassword

# Only proceeds if user confirms
$Phone = if ($userInput.InputNewMobile) { $userInput.InputNewMobile } else { $null }

New-ADUserFromTemplate -NewUser $newUserProperties `
    -SourceUser $UserToCopyAD `
    -Phone $Phone `
    -Password $passwordResult.SecurePassword `
    -DestinationOU $destinationOU

# Step 5: AD Group Copy
Write-ProgressStep -StepName $progressSteps[5].Name -Status $progressSteps[5].Description
Copy-UserADGroups -SourceUser $userInput.InputUserToCopy -TargetUser $userInput.InputUserToCopy

# Step 6: Azure Sync
Write-ProgressStep -StepName $progressSteps[6].Name -Status $progressSteps[6].Description
$MgUser = Wait-ForADUserSync -UserEmail $newUserProperties.Email

if ($MgUser) {

    # Step 7: License Setup
    Write-ProgressStep -StepName $progressSteps[7].Name -Status $progressSteps[7].Description
    Write-StatusMessage -Message "Setting Usage Location for new user" -Type INFO
    Update-MgUser -UserId $MgUser.Id -UsageLocation US
    Write-StatusMessage -Message "Usage Location has been set for new user" -Type OK

    # Required license - will exit on failure
    Set-UserLicenses -User $MgUser -License $userInput.InputRequiredLicense -Required

    Start-Sleep -Seconds 60  # Wait for license to apply

    # Ancillary licenses - will continue on failure
    if ($null -ne $userInput.InputAncillaryLicenses) {
        Set-UserLicenses -User $MgUser -License $userInput.InputAncillaryLicenses
    }

    # Step 8: Entra Group Copy
    Write-ProgressStep -StepName $progressSteps[8].Name -Status $progressSteps[8].Description
    $MgUserCopyAD = Get-MgUser -UserId $UserToCopyAD.UserPrincipalName
    Copy-UserEntraGroups -SourceUser $MgUserCopyAD -TargetUser $MgUser
    $CopyUserGroupCount = (Get-MgUserMemberOf -UserId $MgUserCopyAD.Id).Count
    $NewUserGroupCount = (Get-MgUserMemberOf -UserId $MgUser.Id).Count

    # Step 9: Email to SOC for KnowBe4
    Write-ProgressStep -StepName $progressSteps[9].Name -Status $progressSteps[9].Description
    $emailSubject = "KB4 – New User"
    $emailContent = "The following user need to be added to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
    $MsgFrom = $config.Email.NotificationFrom
    $ToAddress = $config.Email.NotificationTo
    Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -Subject $emailSubject -Content $emailContent

    # Step 10: OneDrive Provisioning
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

    # Step 11: BookWithMeId Setup
    Write-ProgressStep -StepName $progressSteps[11].Name -Status $progressSteps[11].Description
    Set-UserBookWithMeId -User $MgUser -SamAccountName $newUserProperties.SamAccountName

    # Step 12: Zoom Phone Creation
    Write-ProgressStep -StepName $progressSteps[12].Name -Status $progressSteps[12].Description
    if ($userInput.ZoomUserEnabled) {

        $zoomParams = @{
            User         = $MgUser
            ClientId     = $config.Zoom.ClientId
            ClientSecret = $config.Zoom.ClientSecret
            AccountId    = $config.Zoom.AccountId
            UserToCopy   = $MgUserCopyAD.mail
            Password     = $passwordResult.PlainPassword
        }

        # Add IsWorkPlaceLicense parameter only if true
        if ($userInput.IsWorkPlaceLicense) {
            $zoomParams['IsWorkPlaceLicense'] = 'true'
        }

        # Call function with splatted parameters
        Add-UserToZoom @zoomParams
    }

    # Step 13: Cleanup and Summary
    Write-ProgressStep -StepName $progressSteps[13].Name -Status $progressSteps[13].Description
    Write-StatusMessage -Message "Disconnecting from Exchange Online and Graph." -Type INFO

    Connect-ServiceEndpoints -Disconnect

    Write-StatusMessage -Message "Building final summary..." -Type INFO

    Start-NewUserFinalize -User $MgUser `
        -Password $passwordResult.PlainPassword `
        -TemplateGroupCount $CopyUserGroupCount `
        -AssignedGroupCount $NewUserGroupCount

    # Show duration
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-StatusMessage "Script completed in $($duration.TotalMinutes.ToString('F2')) minutes" -Type INFO

    # Give user time to read/copy the summary
    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
    Read-Host | Out-Null

    # Clear the progress bar
    Write-Progress -Activity "New User Creation" -Completed

    # Clean up and exit
    Stop-Job $loadingJob | Out-Null
    Remove-Job $loadingJob | Out-Null

    Exit-Script -Message "$($MgUser.displayName) has been successfully created." -ExitCode Success

} else {
    Write-StatusMessage -Message "Failed to get user from Azure AD after sync" -Type 'ERROR'
    Exit-Script -Message "Azure AD sync completed but user was not found" -ExitCode GeneralError
}