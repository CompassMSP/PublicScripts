#requires -Version 7.0
#requires -RunAsAdministrator
#requires -Modules activeDirectory,ExchangeOnlineManagement,Microsoft.Graph.Users,Microsoft.Graph.Groups,ADSync
<#
TODO: Add Department Group Mapping on line 3102 at $setDepartmentMappings
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
    Last Modified: 2025-04-07

    Version History:
    ------------------------------------------------------------------------------
    Version    Date         Changes
    -------    ----------  ---------------------------------------------------
    4.0.0      2025-04-07   Major UI and functions refactor:
                                - Changed UI to use winui 3 styles for easier management and better visuals
                                - Refactor user creation functions in preparation for Forms/Power Automate Flow execution
                                - Added JSON input for user creation data
                                - Changed license assignment and groups to use Graph API due to PowerShell SDK bugs
                                - Added additional error checks and logging
                                - Enhanced group operations and better error handling
                                - Added mailbox provisioning check before group operations
                                - Added comprehensive operation summary with detailed status tracking
                                - Enhanced error messages and user feedback throughout the process

    3.3.0      2025-04-01   Zoom Phone Removal:
                                - Removed provisioning steps for Zoom Phone and Contact Center as we are moving to 8x8

    3.2.0      2025-02-03   Zoom Phone Onboarding:
                                - Added provisioning steps for Zoom Phone and Contact Center

    3.1.0      2025-01-25   Password System Update:
                                - Replaced New-SecureRandomPassword with New-ReadablePassword
                                - Added human-readable password generation using word list
                                - Added interactive password acceptance/rejection
                                - Added GitHub wordlist integration
                                - Added support for custom word lists
                                - Added configurable word count (2-20 words)
                                - Added spaces/no-spaces password formatting options

    3.0.0      2025-01-20   Major Rework:
                                - Complete script reorganization and optimization
                                - Optimized UI spacing and element alignment
                                - Enhanced form layout for improved readability
                                - Added secure configuration management via Get-ScriptConfig
                                - Enhanced error handling and logging system
                                - Added progress tracking and status messaging

    2.1.0      2024-10-15   Feature Update:
                                - Added BookWithMeId validation
                                - Enhanced AD Sync loop handling
                                - Reworked GUI interface
                                - Added QuickEdit and InsertMode functions
                                - Added SMTP duplicate checking
                                - Removed KnowBe4 SCIM integration per SecurePath Team
                                - Added Email Forwarding functionality - KnowBe4 Notification

    2.0.0      2024-05-08   Major Feature Update:
                                - Added input box system
                                - Added EntraID P2 license checkbox
                                - Enhanced UI boxes for variables
                                - Added KB4 email delivery
                                - Added MeetWithMeId and AD properties
                                - Updated KnowBe4 SCIM integration
                                - Added template user validation

    1.2.0       2024-02-12  Feature Updates:
                                - Enhanced license display output
                                - Improved group management functions
                                - Added KnowBe4 SCIM integration

    1.1.0       2022-06-27  Feature Updates:
                                - Added duplicate attribute checking
                                - Added fax attributes copying
                                - Enhanced group lookup and management
                                - Added AD sync validation

    1.0.0       2022-03-02  Initial Release:
                                - Basic user creation functionality
                                - Template user copying
                                - Group membership handling
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

Clear-Host

Write-Host "`n  Initializing New User Creation Script v4.0.0..." -ForegroundColor Cyan
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
        Write-Progress -Activity "New User Creation" -Status $status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
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

    $stepNumber = $script:currentStep
    $status = $step.Description

    # Guard against division by zero or missing values
    if ($null -eq $stepNumber -or $script:totalSteps -eq 0) {
        Write-StatusMessage -Message "Step $StepName - $status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $status
    } else {
        Write-StatusMessage -Message "Step $stepNumber of $script:totalSteps : $StepName - $status" -Type INFO
        Write-Progress -Activity "New User Creation" -Status $status -PercentComplete (($stepNumber / $script:totalSteps) * 100)
    }
    $script:currentStep += 1
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

# License Helper Functions
function Get-LicenseDisplayName {
    param ([string]$SkuPartNumber)
    $displayName = switch -Regex ($SkuPartNumber) {
        "MCOPSTNC" { "Communications Credits" }
        "PROJECT_MADEIRA_PREVIEW_IW_SKU" { "Dynamics 365 Business Central for IWs" }
        "EXCHANGESTANDARD" { "Exchange Online (Plan 1)" }
        "FLOW_FREE" { "Microsoft Power Automate Free" }
        "MICROSOFT_BUSINESS_CENTER" { "Microsoft Business Center" }
        "Microsoft_Copilot_for_Finance_trial" { "Microsoft Copilot for Finance trial" }
        "Microsoft365_Lighthouse" { "Microsoft 365 Lighthouse" }
        "MCOMEETADV" { "Microsoft 365 Audio Conferencing" }
        "Microsoft_365_Copilot" { "Microsoft 365 Copilot" }
        "O365_BUSINESS_ESSENTIALS" { "Microsoft 365 Business Basic" }
        "SPB" { "Microsoft 365 Business Premium" }
        "SPE_E3" { "Microsoft 365 E3" }
        "AAD_PREMIUM_P2" { "Microsoft Entra ID P2" }
        "POWERAPPS_DEV" { "Microsoft PowerApps for Developer" }
        "POWERAPPS_VIRAL" { "Microsoft Power Apps Plan 2 Trial" }
        "Microsoft_Teams_Audio_Conferencing_select_dial_out" { "Microsoft Teams Audio Conferencing with dial-out to USA/CAN" }
        "Microsoft_Teams_Premium" { "Microsoft Teams Premium" }
        "MCOEV" { "Microsoft Teams Phone Standard" }
        "PHONESYSTEM_VIRTUALUSER" { "Microsoft Teams Phone Resource Account" }
        "MEETING_ROOM" { "Microsoft Teams Rooms Standard" }
        "ENTERPRISEPACK" { "Office 365 E3" }
        "POWERAPPS_PER_USER" { "Power Apps Premium" }
        "POWERAUTOMATE_ATTENDED_RPA" { "Power Automate Premium" }
        "POWER_BI_PRO" { "Power BI Pro" }
        "POWER_BI_STANDARD" { "Power BI Standard" }
        "CCIBOTS_PRIVPREV_VIRAL" { "Power Virtual Agents Viral Trial" }
        "PROJECT_P1" { "Project Plan 1" }
        "PROJECTPROFESSIONAL" { "Project Plan 3" }
        "PROJECT_PLAN3_DEPT" { "Project Plan 3 (for Department)" }
        "RIGHTSMANAGEMENT_ADHOC" { "Rights Management Adhoc" }
        "RMSBASIC" { "Rights Management Service Basic Content Protection" }
        "SHAREPOINTSTORAGE" { "SharePoint Storage" }
        "MCOPSTN1" { "Skype for Business PSTN Domestic Calling" }
        "Teams_Premium_(for_Departments)" { "Teams Premium (for Departments)" }
        "VISIOCLIENT" { "Visio Plan 2" }
        "WINDOWS_STORE" { "Windows Store for Business" }
        default { $SkuPartNumber }
    }
    return $displayName
}

function Get-FormattedLicenseInfo {
    param (
        [array]$Skus,
        [array]$IgnoredLicenses = @(
            "Communications Credits",
            "Dynamics 365 Business Central for IWs",
            "Microsoft 365 Copilot",
            "Microsoft 365 Lighthouse",
            "Microsoft Business Center",
            "Microsoft Copilot for Finance trial",
            "Microsoft Power Apps Plan 2 Trial",
            "Microsoft Power Automate Free",
            "Microsoft PowerApps for Developer",
            "Microsoft Teams Phone Resource Account",
            "Microsoft Teams Phone Standard",
            "Microsoft Teams Rooms Standard",
            "Power Apps Premium",
            "Power Automate Premium",
            "Power BI Pro",
            "Power BI Standard",
            "Power Virtual Agents Viral Trial",
            "Project Plan 3 (for Department)",
            "Rights Management Adhoc",
            "Rights Management Service Basic Content Protection",
            "STREAM",
            "SharePoint Storage",
            "Skype for Business PSTN Domestic Calling",
            "Teams_Premium_(for_Departments)",
            "Windows Store for Business"
        )
    )
    return $Skus | ForEach-Object {
        $available = $_.PrepaidUnits - $_.ConsumedUnits
        $SkuDisplayName = Get-LicenseDisplayName $_.SkuPartNumber
        if ([string]::IsNullOrEmpty($SkuDisplayName)) {
            $SkuDisplayName = $_.SkuPartNumber
        }

        # Skip if license is in ignored list
        if ($IgnoredLicenses -notcontains $SkuDisplayName) {
            @{
                DisplayName = "$($SkuDisplayName) (Available: $available)"
                SkuId       = $_.SkuId
                SortName    = $SkuDisplayName
            }
        }
    } | Where-Object { $_ -ne $null } | Sort-Object { $_.SortName }
}

# UI Functions
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

function Get-NewUserRequest {
    #region Assembly and Namespace loading
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    #region XAML Design
    [xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Creation Form" Height="800" Width="820"
    WindowStartupLocation="CenterScreen">

    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/PresentationFramework.Fluent;component/Themes/Fluent.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <StackPanel Grid.Row="0">
            <TextBlock Text="User Creation Form" FontSize="24" FontWeight="SemiBold" Margin="0,0,0,20"/>
            <Border Background="{DynamicResource TextControlBackgroundPointerOver}"
                    BorderBrush="{DynamicResource TextControlBorderBrush}"
                    BorderThickness="1"
                    Padding="10"
                    Margin="0,0,0,20">
                <TextBlock TextWrapping="Wrap">
                    Please fill in the required information for the new user. Fields marked with * are required.
                </TextBlock>
            </Border>
            <WrapPanel Margin="0,0,0,20">
                <Button x:Name="btnLoadJson" Content="Load JSON" Style="{DynamicResource AccentButtonStyle}" Width="120" Height="32" Margin="0,0,10,0"/>
                <Button x:Name="btnSaveJson" Content="Save JSON" Width="120" Height="32" Margin="0,0,10,0"/>
                <Button x:Name="btnRefreshLicenses" Content="Refresh Licenses" Width="120" Height="32" Margin="0,0,10,0"/>
            </WrapPanel>
        </StackPanel>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="0,10">
            <!-- User Information Tab -->
            <TabItem Header="User Information">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10" Padding="0,0,20,0">
                    <StackPanel>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <!-- Left Column -->
                            <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                <Label>
                                    <TextBlock>
                                        <Run Text="Required License"/>
                                        <Run Text=" *" Foreground="Red"/>
                                    </TextBlock>
                                </Label>
                                <ComboBox x:Name="cboRequiredLicense" Height="32" Margin="0,0,0,15"/>

                                <Label>
                                    <TextBlock>
                                        <Run Text="Display Name"/>
                                        <Run Text=" *" Foreground="Red"/>
                                    </TextBlock>
                                </Label>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtDisplayName" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter full name (First Last)"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtDisplayName}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Email Address"/>
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,15">
                                    <Grid Width="150">
                                        <TextBox x:Name="txtSamAccountName" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                        <TextBlock IsHitTestVisible="False"
                                                 Text="Username"
                                                 VerticalAlignment="Center"
                                                 Margin="8,0,0,0"
                                                 Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                            <TextBlock.Style>
                                                <Style TargetType="{x:Type TextBlock}">
                                                    <Setter Property="Visibility" Value="Collapsed"/>
                                                    <Style.Triggers>
                                                        <DataTrigger Binding="{Binding Text, ElementName=txtSamAccountName}" Value="">
                                                            <Setter Property="Visibility" Value="Visible"/>
                                                        </DataTrigger>
                                                    </Style.Triggers>
                                                </Style>
                                            </TextBlock.Style>
                                    </TextBlock>
                                    </Grid>
                                    <TextBlock Text="@" VerticalAlignment="Center" Margin="5,0" Padding="0,5,0,0"/>
                                    <ComboBox x:Name="cboDomain" Width="150" Height="32" VerticalAlignment="Center"/>
                                    <Button x:Name="btnRefreshDomains" Content="⟳" Width="32" Height="32" VerticalAlignment="Center" Margin="5,0,0,0"/>
                                </StackPanel>

                                <Label Content="Mobile Phone"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtMobilePhone" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter mobile phone number"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtMobilePhone}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Time Zone"/>
                                <ComboBox x:Name="cboTimeZone" Height="32" Margin="0,0,0,15"/>

                                <Label Content="365 Usage Location"/>
                                <ComboBox x:Name="cboUsageLocation" Height="32" Margin="0,0,0,15"/>
                            </StackPanel>

                            <!-- Right Column -->
                            <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                <Label Content="Copy User Operations"/>
                                <ComboBox x:Name="cboCopyUserOperations" Height="32" Margin="0,0,0,15"/>

                                <Label Content="User To Copy"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtUserToCopy" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter email of user to copy"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtUserToCopy}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Ancillary Licenses (Multiple Selection)"/>
                                <ListBox x:Name="lstAncillaryLicenses"
                                    SelectionMode="Multiple"
                                        MinHeight="100"
                                        MaxHeight="270"
                                        ScrollViewer.VerticalScrollBarVisibility="Auto"
                                        Margin="0,0,0,15"/>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Employee Information Tab -->
            <TabItem Header="Employee Information">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10" Padding="0,0,20,0">
                    <StackPanel>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <!-- Left Column -->
                            <StackPanel Grid.Column="0" Margin="0,0,10,0">
                                <Label Content="Given Name"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtGivenName" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter first name"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtGivenName}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Job Title"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtJobTitle" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter job title"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtJobTitle}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Company Name"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtCompanyName" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter company name"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtCompanyName}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Employee Hire Date"/>
                                <DatePicker x:Name="dateEmployeeHireDate" Height="32" Margin="0,0,0,15"/>

                                <Label Content="Business Phone"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtBusinessPhone" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter business phone number"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtBusinessPhone}" Value="">
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
                                <Label Content="Surname"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtSurname" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter last name"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtSurname}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Department"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtDepartment" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter department name"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtDepartment}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Office Location"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtOfficeLocation" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter office location"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtOfficeLocation}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Manager"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtManager" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter manager's email"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtManager}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>

                                <Label Content="Fax Number"/>
                                <Grid Margin="0,0,0,15">
                                    <TextBox x:Name="txtFaxNumber" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                                    <TextBlock IsHitTestVisible="False"
                                             Text="Enter fax number"
                                             VerticalAlignment="Center"
                                             Margin="8,0,0,0"
                                             Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                        <TextBlock.Style>
                                            <Style TargetType="{x:Type TextBlock}">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Text, ElementName=txtFaxNumber}" Value="">
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- Location Information Tab -->
            <TabItem Header="Location Information">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,10" Padding="0,0,20,0">
                    <StackPanel>
                        <Label Content="Street Address"/>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtStreetAddress" Height="64" Padding="8,5,8,5" TextWrapping="Wrap" AcceptsReturn="True" VerticalContentAlignment="Top"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter street address"
                                     VerticalAlignment="Top"
                                     Margin="8,5,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtStreetAddress}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>

                        <Label Content="City"/>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtCity" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter city"
                                     VerticalAlignment="Center"
                                     Margin="8,0,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtCity}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>

                        <Label Content="State"/>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtState" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter state"
                                     VerticalAlignment="Center"
                                     Margin="8,0,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtState}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>

                        <Label Content="Postal Code"/>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtPostalCode" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter postal code"
                                     VerticalAlignment="Center"
                                     Margin="8,0,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtPostalCode}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>

                        <Label Content="Country"/>
                        <Grid Margin="0,0,0,15">
                            <TextBox x:Name="txtCountry" Height="32" Padding="8,5,8,5" VerticalContentAlignment="Center"/>
                            <TextBlock IsHitTestVisible="False"
                                     Text="Enter country"
                                     VerticalAlignment="Center"
                                     Margin="8,0,0,0"
                                     Foreground="{DynamicResource TextControlPlaceholderForeground}">
                                <TextBlock.Style>
                                    <Style TargetType="{x:Type TextBlock}">
                                        <Setter Property="Visibility" Value="Collapsed"/>
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Text, ElementName=txtCountry}" Value="">
                                                <Setter Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </TextBlock.Style>
                            </TextBlock>
                        </Grid>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

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

                <CheckBox x:Name="cbTestMode" Content="Test Mode" Grid.Column="0"/>
                <TextBlock x:Name="tbStatus" Grid.Column="1" VerticalAlignment="Center" Margin="20,0"/>
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="btnSubmit" Content="Submit" Style="{DynamicResource AccentButtonStyle}" Padding="20,5" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnReset" Content="Reset" Padding="20,5" Height="32" Margin="0,0,10,0"/>
                    <Button x:Name="btnCancel" Content="Cancel" Padding="20,5" Height="32"/>
        </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Parse the XAML
    $XAMLReader = [System.Xml.XmlNodeReader]::new($XAML)
    $Window = [Windows.Markup.XamlReader]::Load($XAMLReader)

    # Create namespace manager for XPath queries
    $nsManager = New-Object System.Xml.XmlNamespaceManager($XAML.NameTable)
    $nsManager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml")

    # Get all form controls by name and create variables
    $XAML.SelectNodes("//*[@x:Name]", $nsManager) | ForEach-Object {
        $Name = $_.Name
        $Variable = New-Variable -Name $Name -Value $Window.FindName($Name) -Force
    }

    # Function to load JSON data
    function Invoke-LoadJsonFile {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $openFileDialog.Title = "Select a JSON file"

        if ($openFileDialog.ShowDialog() -eq "OK") {
            try {
                $jsonContent = Get-Content -Path $openFileDialog.FileName -Raw | ConvertFrom-Json

                # Set the required license
                if ($jsonContent.requiredLicense) {
                    foreach ($item in $cboRequiredLicense.Items) {
                        $itemText = $item.Content.ToString() -replace '\s+\(Available:.*?\)', ''
                        if ($itemText -eq $jsonContent.requiredLicense) {
                            $cboRequiredLicense.SelectedItem = $item
                            break
                        }
                    }
                }

                # Set ancillary licenses (multi-select)
                $ancillaryLicenseData = if ($jsonContent.ancillaryLicense) { $jsonContent.ancillaryLicense }

                if ($ancillaryLicenseData) {
                    # Convert input to array regardless of type
                    $licenses = @()
                    if ($ancillaryLicenseData -is [string]) {
                        # If it's a single string, split by comma if it contains commas, otherwise use as is
                        if ($ancillaryLicenseData -match ',') {
                            $licenses = $ancillaryLicenseData -split ',' | ForEach-Object { $_.Trim() }
                        } else {
                            $licenses = @($ancillaryLicenseData.Trim())
                        }
                    } elseif ($ancillaryLicenseData -is [array]) {
                        # If it's already an array, use it directly
                        $licenses = $ancillaryLicenseData
                    }

                    # Clear any existing selections
                    $lstAncillaryLicenses.SelectedItems.Clear()

                    # Process each license
                    foreach ($license in $licenses) {
                        $trimmedLicense = $license.Trim()
                        for ($i = 0; $i -lt $lstAncillaryLicenses.Items.Count; $i++) {
                            # Only strip the availability count from the ListBox items
                            $itemText = $lstAncillaryLicenses.Items[$i].Content -replace '\s*\(Available:.*\)', ''

                            # Compare the stripped ListBox item text with the JSON license name
                            if ($itemText -eq $trimmedLicense) {
                                $lstAncillaryLicenses.SelectedItems.Add($lstAncillaryLicenses.Items[$i])
                                break
                            }
                        }
                    }
                }

                # Set the employee hire date
                if ($jsonContent.employeeHireDate -and $jsonContent.employeeHireDate -ne "") {
                    try {
                        $dateEmployeeHireDate.SelectedDate = [DateTime]::Parse($jsonContent.employeeHireDate)
                    } catch {
                        Write-StatusMessage "Failed to parse date: $($jsonContent.employeeHireDate)" -Type ERROR
                    }
                }

                # Set copy user operations
                if ($jsonContent.copyUserOperations) {
                    foreach ($item in $cboCopyUserOperations.Items) {
                        if ($item -eq $jsonContent.copyUserOperations) {
                            $cboCopyUserOperations.SelectedItem = $item
                            break
                        }
                    }
                }

                # Set the domain and username from userPrincipalName
                if ($jsonContent.userPrincipalName -and $jsonContent.userPrincipalName -match '@') {
                    $upnParts = $jsonContent.userPrincipalName -split '@'
                    $txtSamAccountName.Text = $upnParts[0]

                    # Try to set domain from either the domain field or from the UPN
                    $domainToSet = if ($jsonContent.domain) { $jsonContent.domain } else { $upnParts[1] }
                    foreach ($item in $cboDomain.Items) {
                        if ($item.ToString() -eq $domainToSet) {
                            $cboDomain.SelectedItem = $item
                            break
                        }
                    }
                } elseif ($jsonContent.domain) {
                    # If no UPN but domain exists, try to set just the domain
                    foreach ($item in $cboDomain.Items) {
                        if ($item.ToString() -eq $jsonContent.domain) {
                            $cboDomain.SelectedItem = $item
                            break
                        }
                    }
                }

                # Populate the form fields with JSON data
                $txtDisplayName.Text = $jsonContent.displayName
                $txtSamAccountName.Text = $jsonContent.userPrincipalName.Split('@')[0]
                $txtMobilePhone.Text = $jsonContent.mobilePhone
                $cboTimeZone.SelectedItem = $jsonContent.timeZone
                $txtUserToCopy.Text = $jsonContent.userToCopy
                $txtGivenName.Text = $jsonContent.givenName
                $txtSurname.Text = $jsonContent.surname
                $txtJobTitle.Text = $jsonContent.jobTitle
                $txtDepartment.Text = $jsonContent.department
                $txtCompanyName.Text = $jsonContent.companyName
                $txtOfficeLocation.Text = $jsonContent.officeLocation
                $txtManager.Text = $jsonContent.manager
                $txtBusinessPhone.Text = $jsonContent.businessPhone
                $txtFaxNumber.Text = $jsonContent.faxNumber
                $txtStreetAddress.Text = $jsonContent.streetAddress
                $txtCity.Text = $jsonContent.city
                $txtState.Text = $jsonContent.state
                $txtPostalCode.Text = $jsonContent.postalCode
                $txtCountry.Text = $jsonContent.country

                # Set department groups (multi-select) - DISABLED
                if ($jsonContent.departmentGroupsDISABLED) {
                    $groups = $jsonContent.departmentGroups -split ','
                    foreach ($group in $groups) {
                        $trimmedGroup = $group.Trim()
                        for ($i = 0; $i -lt $lstDepartmentGroups.Items.Count; $i++) {
                            if ($lstDepartmentGroups.Items[$i] -eq $trimmedGroup) {
                                $lstDepartmentGroups.SelectedItems.Add($lstDepartmentGroups.Items[$i])
                            }
                        }
                    }
                }

                Show-CustomAlert -Message "JSON file loaded successfully" -AlertType "Success" -Title "Success"
            } catch {
                Show-CustomAlert -Message "Error loading JSON file: $_" -AlertType "Error" -Title "Error"
            }
        }
    }

    # Function to save JSON data
    function Save-JsonData {
        # Get form data using the existing Get-FormData function
        $formDataJSON = Get-FormData

        if ($cboRequiredLicense.SelectedItem) {
            $formDataJSON.requiredLicense = $cboRequiredLicense.SelectedItem.Content -replace '\s*\(Available:.*\)', ''
        } else { "" }

        # Get the ancillary licenses as an array of display names
        $selectedAncillaryLicenses = @()
        foreach ($item in $lstAncillaryLicenses.SelectedItems) {
            # Strip out the (Available: X) part and add to array
            $licenseName = $item.Content -replace '\s*\(Available:.*\)', ''
            $selectedAncillaryLicenses += $licenseName
        }

        if ($selectedAncillaryLicenses -ne 0) {
            $formDataJSON.ancillaryLicense = $selectedAncillaryLicenses
        }

        # Get the department groups as an array
        if ($lstDepartmentGroups) {
            $selectedDepartmentGroups = @()
            foreach ($item in $lstDepartmentGroups.SelectedItems) {
                $selectedDepartmentGroups += $item.Content
            }
        }

        # Open save file dialog
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $saveFileDialog.Title = "Save JSON File"
        $saveFileDialog.DefaultExt = "json"

        if ($saveFileDialog.ShowDialog() -eq "OK") {
            try {
                $formDataJSON | ConvertTo-Json -Depth 5 | Set-Content -Path $saveFileDialog.FileName
                Show-CustomAlert -Message "JSON file saved successfully" -AlertType "Success" -Title "Success"
            } catch {
                Show-CustomAlert -Message "Error saving JSON file: $_" -AlertType "Error" -Title "Error"
            }
        }
    }

    # Function to reset the form
    function Reset-Form {
        $cboRequiredLicense.SelectedIndex = -1
        $txtDisplayName.Text = ""
        $txtSamAccountName.Text = ""
        $txtMobilePhone.Text = ""
        $cboTimeZone.SelectedIndex = -1
        $cboCopyUserOperations.SelectedIndex = -1
        $txtUserToCopy.Text = ""
        $lstAncillaryLicenses.SelectedItems.Clear()
        $txtGivenName.Text = ""
        $txtSurname.Text = ""
        $txtJobTitle.Text = ""
        $txtDepartment.Text = ""
        $txtOfficeLocation.Text = ""
        $txtManager.Text = ""
        $dateEmployeeHireDate.SelectedDate = $null
        $txtCompanyName.Text = ""
        $txtBusinessPhone.Text = ""
        $txtFaxNumber.Text = ""
        $txtStreetAddress.Text = ""
        $txtCity.Text = ""
        $txtState.Text = ""
        $txtPostalCode.Text = ""
        $txtCountry.Text = ""
        #$lstDepartmentGroups.SelectedItems.Clear()
        Show-StatusMessage -Message "Form has been reset" -Type "Info"
    }

    # Function to get form data and store in variables
    function Get-FormData {
        # Helper function to return $null for empty strings
        function Get-ValueOrNull($value) {
            if ([string]::IsNullOrWhiteSpace($value)) {
                return $null
            }
            return $value
        }

        # Store form data in a custom object
        $formData = [PSCustomObject]@{
            requiredLicense        = @()
            displayName            = Get-ValueOrNull $txtDisplayName.Text
            samAccountName         = Get-ValueOrNull $txtSamAccountName.Text
            domain                 = if ($cboDomain.SelectedItem) { $cboDomain.SelectedItem.ToString() } else { "" }
            userPrincipalName      = if ($txtSamAccountName.Text -and $cboDomain.SelectedItem) { "$($txtSamAccountName.Text)@$($cboDomain.SelectedItem)" } else { "" }
            mobilePhone            = Get-ValueOrNull $txtMobilePhone.Text
            timeZone               = if ($cboTimeZone.SelectedItem) { $cboTimeZone.SelectedItem } else { $null }
            usageLocation          = if ($cboUsageLocation.SelectedItem) { $cboUsageLocation.SelectedItem.ToString() } else { "" }
            copyUserOperations     = if ($cboCopyUserOperations.SelectedItem -eq 'None') { $null } elseif ($cboCopyUserOperations.SelectedItem) { $cboCopyUserOperations.SelectedItem } else { $null }
            userToCopy             = Get-ValueOrNull $txtUserToCopy.Text
            ancillaryLicense       = @()
            givenName              = Get-ValueOrNull $txtGivenName.Text
            surname                = Get-ValueOrNull $txtSurname.Text
            jobTitle               = Get-ValueOrNull $txtJobTitle.Text
            department             = Get-ValueOrNull $txtDepartment.Text
            companyName            = Get-ValueOrNull $txtCompanyName.Text
            officeLocation         = Get-ValueOrNull $txtOfficeLocation.Text
            employeeHireDate       = if ($dateEmployeeHireDate.SelectedDate) { $dateEmployeeHireDate.SelectedDate.ToString("yyyy-MM-dd") } else { $null }
            manager                = Get-ValueOrNull $txtManager.Text
            businessPhone          = Get-ValueOrNull $txtBusinessPhone.Text
            faxNumber              = Get-ValueOrNull $txtFaxNumber.Text
            streetAddress          = Get-ValueOrNull $txtStreetAddress.Text
            city                   = Get-ValueOrNull $txtCity.Text
            state                  = Get-ValueOrNull $txtState.Text
            postalCode             = Get-ValueOrNull $txtPostalCode.Text
            country                = Get-ValueOrNull $txtCountry.Text
            departmentGroupOptions = @()
            testModeEnabled        = $false
        }

        # Store required licenses in an array of objects with DisplayName and SkuId
        foreach ($item in $cboRequiredLicense.SelectedItem) {
            $formData.requiredLicense += [PSCustomObject]@{
                DisplayName = $item.Content
                SkuId       = $item.Tag
            }
        }

        # Store ancillary licenses in an array of objects with DisplayName and SkuId
        foreach ($item in $lstAncillaryLicenses.SelectedItems) {
            $formData.ancillaryLicense += [PSCustomObject]@{
                DisplayName = $item.Content
                SkuId       = $item.Tag
            }
        }

        # Store department groups in an array
        if ($lstDepartmentGroups) {
            $selectedDepartmentGroups = @()
            foreach ($item in $lstDepartmentGroups.SelectedItems) {
                $selectedDepartmentGroups += $item.Content
            }
        }

        if ($cbTestMode.IsChecked -eq $true) {
            $formData.testModeEnabled = $true
        } else {
            $formData.testModeEnabled = $false
        }

        return $formData
    }

    # Function to initialize and refresh licenses
    function Initialize-Licenses {
        try {
            Write-StatusMessage "Retrieving licenses..." -Type INFO
            $btnRefreshLicenses.IsEnabled = $false

            # Clear current items
            $cboRequiredLicense.Items.Clear()
            $lstAncillaryLicenses.Items.Clear()

            # Get license info from Microsoft Graph
            $skus = Get-MgSubscribedSku | Select-Object SkuId, SkuPartNumber, ConsumedUnits, @{
                Name = 'PrepaidUnits'; Expression = { $_.PrepaidUnits.Enabled }
            }

            # Format license info if needed (or assume licenseInfo = $skus)
            $licenseInfo = Get-FormattedLicenseInfo -Skus $skus

            # Define the licenses you care about
            $requiredLicenses = @(
                "Exchange Online (Plan 1)",
                "Office 365 E3",
                "Microsoft 365 Business Basic",
                "Microsoft 365 E3",
                "Microsoft 365 Business Premium"
            )

            # Populate the combo box with matching licenses
            foreach ($license in $licenseInfo) {
                foreach ($reqLicense in $requiredLicenses) {
                    if ($license.DisplayName -like "*$reqLicense*") {
                        $comboItem = New-Object System.Windows.Controls.ComboBoxItem
                        $comboItem.Content = $license.DisplayName
                        $comboItem.Tag = $license.SkuId  # Store the SkuId for use later
                        $cboRequiredLicense.Items.Add($comboItem)
                    }
                }
            }

            # Loop through licenseInfo and only add licenses not in the required list
            foreach ($license in $licenseInfo) {
                $isRequired = $false
                foreach ($reqLicense in $requiredLicenses) {
                    if ($license.DisplayName -like "*$reqLicense*") {
                        $isRequired = $true
                        break
                    }
                }

                if (-not $isRequired) {
                    $listItem = New-Object System.Windows.Controls.ListBoxItem
                    $listItem.Content = $license.DisplayName
                    $listItem.Tag = $license.SkuId
                    $lstAncillaryLicenses.Items.Add($listItem)
                }
            }

            Write-StatusMessage "Licenses refreshed successfully" -Type OK
        } catch {
            Write-StatusMessage "Error retrieving licenses: $($_.Exception.Message)" -Type ERROR
            Show-CustomAlert -Message "Error retrieving licenses: $($_.Exception.Message)" -AlertType "Error" -Title "Error"
        } finally {
            $btnRefreshLicenses.IsEnabled = $true
        }
    }

    # Function to validate required data
    function Invoke-ValidateForm {
        param (
            [Parameter()]
            $DisplayName, # TextBox for Display Name

            [Parameter()]
            $RequiredLicense  # ComboBox for Required License
        )

        # Validate DisplayName for "First Last" format using a regex pattern
        $namePattern = "^[A-Za-z]+\s[A-Za-z]+$"  # Matches First Last format with only letters

        if (-not $DisplayName -or $DisplayName -notmatch $namePattern) {
            # Invalid format or empty
            Show-CustomAlert -Message "Please enter a valid full name (First Last)"
            return $false
        }

        # Check if an item is selected in the ComboBox
        if (-not $RequiredLicense -or -not $cboRequiredLicense.SelectedItem) {
            # No item selected
            Show-CustomAlert -Message "Please select a required license."
            return $false
        }

        # Validate required license availability
        $requiredLicenseText = $cboRequiredLicense.SelectedItem.Content
        if ($requiredLicenseText -match "\(Available:\s*(\d+)\)") {
            $availableLicenses = [int]$Matches[1]
            if ($availableLicenses -le 0) {
                $licenseName = $requiredLicenseText -replace '\s*\(Available:.*\)', ''
                Show-CustomAlert -Message "The selected required license '$licenseName' has no available licenses."
                return $false
            }
        }

        # Validate ancillary licenses availability
        foreach ($selectedItem in $lstAncillaryLicenses.SelectedItems) {
            $licenseText = $selectedItem.Content
            if ($licenseText -match "\(Available:\s*(\d+)\)") {
                $availableLicenses = [int]$Matches[1]
                if ($availableLicenses -le 0) {
                    $licenseName = $licenseText -replace '\s*\(Available:.*\)', ''
                    Show-CustomAlert -Message "The selected ancillary license '$licenseName' has no available licenses."
                    return $false
                }
            }
        }

        # If both validations pass, return true
        return $true
    }

    # Function to initialize domains
    function Initialize-Domains {
        try {
            Write-StatusMessage "Retrieving domains..." -Type INFO
            $btnRefreshDomains.IsEnabled = $false

            # Get domains from Graph API
            $domains = Get-MgDomain -All | Where-Object { $_.IsVerified -eq $true } | Sort-Object Id

            if ($null -eq $domains -or $domains.Count -eq 0) {
                Write-StatusMessage "No verified domains found" -Type WARN
                return
            }

            # Clear and reload domains
            $cboDomain.Items.Clear()

            # Add verified domains and find default domain
            $defaultDomain = $null
            foreach ($domain in $domains) {
                $cboDomain.Items.Add($domain.Id)
                if ($domain.IsDefault) {
                    $defaultDomain = $domain.Id
                }
            }

            # Select default domain
            if ($defaultDomain) {
                $cboDomain.SelectedItem = $defaultDomain
                Write-StatusMessage "Selected default domain: $defaultDomain" INFO
            } elseif ($cboDomain.Items.Count -gt 0) {
                $cboDomain.SelectedIndex = 0
            }

            Write-StatusMessage "Retrieved $($domains.Count) domains" -Type OK
        } catch {
            Write-StatusMessage "Error retrieving domains: $($_.Exception.Message)" -Type ERROR
            Show-CustomAlert -Message "Error retrieving domains: $($_.Exception.Message)" -AlertType "Error" -Title "Error"
        } finally {
            $btnRefreshDomains.IsEnabled = $true
        }
    }

    function Initialize-UsageLocation {
        param (
            [System.Windows.Controls.ComboBox]$ComboBox
        )

        # Define country codes array
        $countryCodes = @(
            "AD", "AE", "AF", "AG", "AI", "AL", "AM", "AO", "AQ", "AR", "AS", "AT", "AU", "AW", "AX", "AZ", "BA", "BB",
            "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BL", "BM", "BN", "BO", "BQ", "BR", "BS", "BT", "BV", "BW", "BY",
            "BZ", "CA", "CC", "CD", "CF", "CG", "CH", "CI", "CK", "CL", "CM", "CN", "CO", "CR", "CU", "CV", "CW", "CX",
            "CY", "CZ", "DE", "DJ", "DK", "DM", "DO", "DZ", "EC", "EE", "EG", "EH", "ER", "ES", "ET", "FI", "FJ", "FK",
            "FM", "FO", "FR", "GA", "GB", "GD", "GE", "GF", "GG", "GH", "GI", "GL", "GM", "GN", "GP", "GQ", "GR", "GS",
            "GT", "GU", "GW", "GY", "HK", "HM", "HN", "HR", "HT", "HU", "ID", "IE", "IL", "IM", "IN", "IO", "IQ", "IR",
            "IS", "IT", "JE", "JM", "JO", "JP", "KE", "KG", "KH", "KI", "KM", "KN", "KP", "KR", "KW", "KY", "KZ", "LA",
            "LB", "LC", "LI", "LK", "LR", "LS", "LT", "LU", "LV", "LY", "MA", "MC", "MD", "ME", "MF", "MG", "MH", "MK",
            "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ", "NA", "NC", "NE",
            "NF", "NG", "NI", "NL", "NO", "NP", "NR", "NU", "NZ", "OM", "PA", "PE", "PF", "PG", "PH", "PK", "PL", "PM",
            "PN", "PR", "PS", "PT", "PW", "PY", "QA", "RE", "RO", "RS", "RU", "RW", "SA", "SB", "SC", "SD", "SE", "SG",
            "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SR", "SS", "ST", "SV", "SX", "SY", "SZ", "TC", "TD", "TF",
            "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TO", "TR", "TT", "TV", "TW", "TZ", "UA", "UG", "UM", "US", "UY",
            "UZ", "VA", "VC", "VE", "VG", "VI", "VN", "VU", "WF", "WS", "YE", "YT", "ZA", "ZM", "ZW"
        )

        # Clear existing items
        $ComboBox.Items.Clear()

        # Add all country codes to the ComboBox
        foreach ($code in $countryCodes) {
            [void]$ComboBox.Items.Add($code)
        }

        # Set default value to US
        $ComboBox.SelectedItem = "US"

        Write-StatusMessage "Usage location initialized with default value: US" -Type INFO
    }

    # Function to initialize department groups
    function Initialize-DepartmentGroups {
        try {
            Write-StatusMessage "Initializing department groups..." -Type INFO

            # Clear existing items
            $lstDepartmentGroups.Items.Clear()

            # Define department group options
            $departmentGroupOptions = @(
                "Marketing Group",
                "Sales Group",
                "Engineering Group",
                "Finance Group",
                "HR Group",
                "Executive Group",
                "IT Support Group",
                "Customer Service Group"
            )

            # Add items to the listbox
            foreach ($option in $departmentGroupOptions) {
                $lstDepartmentGroups.Items.Add($option)
            }

            Write-StatusMessage "Department groups initialized successfully" -Type OK
        } catch {
            Write-StatusMessage "Error initializing department groups: $($_.Exception.Message)" -Type ERROR
            Show-CustomAlert -Message "Error initializing department groups: $($_.Exception.Message)" -AlertType "Error" -Title "Error"
        }
    }

    # Function to get selected items
    function Get-SelectedDepartments {
        $selectedDepartments = @()
        if ($chkNFLROC.IsChecked) { $selectedDepartments += "NFL ROC" }
        if ($chkSFLOC.IsChecked) { $selectedDepartments += "SFL OC" }
        if ($chkNEROC.IsChecked) { $selectedDepartments += "NE ROC" }
        if ($chkBilling.IsChecked) { $selectedDepartments += "Billing" }
        if ($chkPSALL.IsChecked) { $selectedDepartments += "PS ALL" }
        if ($chkPST1.IsChecked) { $selectedDepartments += "PS T1" }
        if ($chkPST2.IsChecked) { $selectedDepartments += "PS T2" }
        return $selectedDepartments
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


    # Add event handlers
    $btnLoadJson.Add_Click({
            Reset-Form
            Invoke-LoadJsonFile
        })
    $btnSaveJson.Add_Click({ Save-JsonData })
    $btnReset.Add_Click({ Reset-Form })
    $btnRefreshLicenses.Add_Click({ Initialize-Licenses })
    #$btnRefreshDepartments.Add_Click({ Initialize-DepartmentGroups })
    $btnSubmit.Add_Click({
            # Run the validation first
            $isValid = Invoke-ValidateForm -DisplayName $txtDisplayName.Text -RequiredLicense $cboRequiredLicense

            if ($isValid) {
                # If validation passes, collect the form data
                $formData = Get-FormData
                $Window.Close()  # Close the window after submission
                return $formData  # Return the form data
            } else {
                # If validation fails, do not close the window and optionally show a message
                Write-StatusMessage "Validation failed. Please fix the errors and try again." -Type ERROR
            }
        })

    $btnCancel.Add_Click({
            $Window.DialogResult = $false
            $Window.Close()
        })

    # Add event handler for the refresh button
    $btnRefreshDomains.Add_Click({ Initialize-Domains })

    # Initialize licenses
    Initialize-Licenses

    # Initialize domains
    Initialize-Domains

    # Set default location to US
    Initialize-UsageLocation -ComboBox $cboUsageLocation

    # Initialize department groups
      # Initialize-DepartmentGroups

    # Define copy user operations options
    $copyUserOperationsOptions = @(
        "None",
        "Copy Attributes",
        "Copy Groups",
        "Copy Attributes and Groups"
    )

    # Populate the copy user operations dropdown
    foreach ($option in $copyUserOperationsOptions) {
        $cboCopyUserOperations.Items.Add($option)
    }

    # Define timezone options
    $timeZoneOptions = @(
        'Eastern Standard Time',
        'Central Standard Time',
        'Mountain Standard Time',
        'US Mountain Standard Time (Arizona)',
        'Pacific Standard Time'
    )

    # Populate the timezone dropdown
    foreach ($timeZone in $timeZoneOptions) {
        $cboTimeZone.Items.Add($timeZone)
    }

    # Add validation for DisplayName
    $txtDisplayName.Add_TextChanged({
            $namePattern = "^[A-Za-z]+\s[A-Za-z]+$"
            if ($this.Text -and -not ($this.Text -match $namePattern)) {
                $this.BorderBrush = 'Red'
                $this.BorderThickness = 2
                $this.ToolTip = "Name must be in 'First Last' format"
            } else {
                $this.BorderBrush = $null
                $this.BorderThickness = 1
                $this.ToolTip = $null
            }
        })

    # Show the form
    $Window.ShowDialog() | Out-Null

    # Return the form data
    return Get-FormData
}

function New-DuplicatePromptForm {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$ExistingValue,

        [Parameter(Mandatory)]
        [string]$PromptText,

        [ValidateSet("Error", "Warning", "Info", "Success")]
        [string]$AlertType = "Info"  # Default to "Info" if not specified
    )

    Add-Type -AssemblyName PresentationFramework

    # Determine the color and sound based on the AlertType
    switch ($AlertType) {
        "Error" { $color = "#E81123"; $sound = [System.Media.SystemSounds]::Hand }
        "Warning" { $color = "#FFB900"; $sound = [System.Media.SystemSounds]::Exclamation }
        "Info" { $color = "#0078D7"; $sound = [System.Media.SystemSounds]::Asterisk }
        "Success" { $color = "#107C10"; $sound = [System.Media.SystemSounds]::Beep }
    }

    [xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DuplicatePromptForm" Height="190" Width="530"
        WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent">

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

    <Border CornerRadius="12" Background="#2D2D30" Padding="10" BorderBrush="$color" BorderThickness="2">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Top Bar for Dragging -->
            <Border Name="Top_Bar" Background="Transparent" Height="5" Grid.Row="0" />

            <!-- Content -->
            <Grid Grid.Row="1" Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Icon -->
                <Viewbox Grid.Row="0" Grid.RowSpan="3" Margin="0,0,15,0" Width="40" Height="40" VerticalAlignment="Top">
                    <Canvas Width="48" Height="48">
                        <Ellipse Width="48" Height="48" Fill="$color"/>
                        <Rectangle Width="6" Height="20" Fill="White" Canvas.Left="21" Canvas.Top="10"/>
                        <Rectangle Width="6" Height="6" Fill="White" Canvas.Left="21" Canvas.Top="34"/>
                    </Canvas>
                </Viewbox>

                <!-- Message -->
                <TextBlock x:Name="PromptLabel" Grid.Column="1" Grid.Row="0" Grid.ColumnSpan="2" TextWrapping="Wrap" Margin="0,0,0,5"/>

                <!-- Input Box -->
                <TextBox x:Name="InputBox" Grid.Column="1" Grid.Row="1" Grid.ColumnSpan="2" Margin="0,5,0,10"/>

                <!-- Buttons -->
                <Button x:Name="OkButton" Grid.Column="1" Grid.Row="2" Width="75" Margin="5" IsDefault="True" Content="OK" HorizontalAlignment="Right"/>
                <Button x:Name="CancelButton" Grid.Column="2" Grid.Row="2" Width="75" Margin="5" IsCancel="True" Content="Cancel" HorizontalAlignment="Left"/>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

    $reader = (New-Object System.Xml.XmlNodeReader $XAML)
    $form = [Windows.Markup.XamlReader]::Load($reader)

    $PromptLabel = $form.FindName("PromptLabel")
    $InputBox = $form.FindName("InputBox")
    $OkButton = $form.FindName("OkButton")
    $CancelButton = $form.FindName("CancelButton")

    # Set values
    $form.Title = $Title
    $PromptLabel.Text = $PromptText
    $InputBox.Text = $ExistingValue

    # Add event handlers
    $OkButton.Add_Click({
            if ($InputBox.Text -eq $ExistingValue) {
                Show-CustomAlert -Message "You must change duplicate value: '$($ExistingValue)'"
                return
            }
            $form.DialogResult = $true
            $form.Close()
        })
    $CancelButton.Add_Click({
            $form.DialogResult = $false
            $form.Close()
        })

    # Find the top bar and add the MouseDown event
    $topBar = $form.FindName("Top_Bar")
    $topBar.Add_MouseDown({
            param($s, $e)
            if ($e.ChangedButton -eq [System.Windows.Input.MouseButton]::Left) {
                $form.DragMove()
            }
        })

    # Play the appropriate sound
    $sound.Play()

    $result = $form.ShowDialog()

    if ($result -eq $true) {
        return $InputBox.Text
    } else {
        return $null
    }
}

# Main Exection Functions

function Get-TemplateUser {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]$UserToCopy
    )

    try {
        Write-StatusMessage -Message "Getting template user details for: $UserToCopy" -Type INFO

        $adUserParams = @{
            Filter      = "DisplayName -eq '$UserToCopy' -or UserPrincipalName -eq '$UserToCopy'"
            Properties  = @(
                'Company',
                'physicalDeliveryOfficeName', # ldapDisplayName for Office
                'Title',
                'Department',
                'facsimileTelephoneNumber',
                'streetAddress'
                'l', # ldapDisplayName for City
                'st', # ldapDisplayName for State
                'postalCode',
                'c' # ldapDisplayName for Country
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

function Get-ADUserCopiedAttributes {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)][Microsoft.ActiveDirectory.Management.ADUser]$TemplateUser
    )

    try {
        Write-StatusMessage -Message "Copying attributes from template user: $($TemplateUser.SamAccountName)" -Type INFO
        return [pscustomobject]@{
            companyName    = $templateUser.Company
            officeLocation = $templateUser.physicalDeliveryOfficeName
            jobTitle       = $templateUser.Title
            department     = $templateUser.Department
            faxNumber      = $templateUser.facsimileTelephoneNumber
            streetAddress  = $templateUser.streetAddress
            city           = $templateUser.l
            state          = $templateUser.st
            postalCode     = $templateUser.postalCode
            country        = $templateUser.c
        }
    } catch {
        Write-StatusMessage -Message "Failed to copy attributes from template user: $_" -Type ERROR
    }
}

function New-UserProperties {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName,

        [Parameter(Mandatory)]
        [string]$Domain,

        [string]$FirstName,
        [string]$LastName,
        [string]$userPrincipalName
    )

    try {
        Write-StatusMessage -Message "Generating properties for new user: $DisplayName" -Type INFO

        # Determine first and last name
        if (-not $FirstName -or -not $LastName) {
            $nameParts = $DisplayName -split ' '
            if ($nameParts.Count -lt 2) {
                Write-StatusMessage -Message "Invalid display name format: Must include first and last name" -Type ERROR
                Exit-Script -Message "Display name must include first and last name" -ExitCode GeneralError
            }
            $FirstName = $nameParts[-2]  # Second to last part
            $LastName = $nameParts[-1]   # Last part
        }

        # Use provided email or generate one
        if (-not $userPrincipalName) {
            $samAccountName = (($FirstName.Substring(0, 1) + $LastName).ToLower())
            $userPrincipalName = ($samAccountName + $Domain).ToLower()
        } else {
            $samAccountName = ($userPrincipalName -split '@')[0]
        }

        # Check for duplicate samAccountName in AD
        if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
            Write-StatusMessage -Message "SamAccountName $samAccountName already exists" -Type WARN

            # For SAM account name prompt:
            $formDuplicateSam = New-DuplicatePromptForm `
                -Title "Duplicate SAM Account Name" `
                -ExistingValue $samAccountName `
                -PromptText "Please enter a different samAccountName: '$samAccountName' already exists."

            if ($formDuplicateSam -ne $samAccountName) {
                $samAccountName = $formDuplicateSam
                # Verify the new samAccountName is unique
                if (Get-ADUser -Filter "SamAccountName -eq '$samAccountName'" -ErrorAction SilentlyContinue) {
                    Write-StatusMessage -Message "New SamAccountName $samAccountName is also in use" -Type ERROR
                    Exit-Script -Message "Unable to generate unique SamAccountName" -ExitCode DuplicateUser
                }
                Write-StatusMessage -Message "Using custom SamAccountName: $samAccountName" -Type OK
                # Update email with new samAccountName
                $userPrincipalName = ($samAccountName + $Domain).ToLower()
            } else {
                Write-StatusMessage -Message "User cancelled SAM account name selection" -Type WARN
                Exit-Script -Message "Operation cancelled by user" -ExitCode Cancelled
            }
        }

        # Check for existing mailbox
        try {
            $mailbox = Get-Mailbox -Filter "EmailAddresses -like '*$userPrincipalName*'" -ErrorAction Stop
            if ($mailbox) {
                Write-StatusMessage -Message "Email address $userPrincipalName (or similar) already exists for mailbox: $($mailbox.UserPrincipalName)" -Type WARN

                # For email prompt:
                $formDuplicateEmail = New-DuplicatePromptForm `
                    -Title "Duplicate Email Address" `
                    -ExistingValue $samAccountName `
                    -PromptText "Please enter a different emailAddress: '$userPrincipalName' already exists."

                if ($formDuplicateEmail -ne $userPrincipalName) {
                    $samAccountName = $formDuplicateEmail
                    # Verify the new email is unique
                    try {
                        $userPrincipalName = ($samAccountName + $Domain).ToLower()
                        $checkMailbox = Get-Mailbox -Filter "EmailAddresses -like '*$userPrincipalName*'" -ErrorAction Stop
                        if ($checkMailbox) {
                            Write-StatusMessage -Message "New email address $userPrincipalName is also in use by: $($checkMailbox.displayName)" -Type ERROR
                            Exit-Script -Message "Unable to generate unique email address" -ExitCode DuplicateUser
                        }
                    } catch [Microsoft.Exchange.Management.RestApiClient.RestApiException] {
                        Write-StatusMessage -Message "Using custom email address: $userPrincipalName" -Type OK
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
            FirstName         = $FirstName
            LastName          = $LastName
            DisplayName       = $DisplayName
            userPrincipalName = $userPrincipalName
            Email             = $userPrincipalName
            SamAccountName    = $samAccountName
        }

    } catch {
        Write-StatusMessage -Message "Critical error in New-UserProperties: $_" -Type ERROR
        Exit-Script -Message "Critical error generating user properties" -ExitCode GeneralError
    }
}

function New-ADUserStandard {
    [CmdletBinding(SupportsShouldProcess)]
    param (
        [Parameter(Mandatory)][hashtable]$NewUser,
        [Parameter(Mandatory)][System.Security.SecureString]$Password,
        [Parameter(Mandatory)][string]$DestinationOU,
        [Parameter()][Microsoft.ActiveDirectory.Management.ADUser]$TemplateUser
    )

    try {
        Write-StatusMessage -Message "Creating new AD user: $($NewUser.DisplayName)" -Type INFO
        $newUserParams = @{
            Name              = "$($NewUser.FirstName) $($NewUser.LastName)"
            SamAccountName    = $NewUser.SamAccountName
            UserPrincipalName = $NewUser.Email
            EmailAddress      = $NewUser.Email
            DisplayName       = $NewUser.DisplayName
            GivenName         = $NewUser.FirstName
            Surname           = $NewUser.LastName
            AccountPassword   = $Password
            Path              = $DestinationOU
            OtherAttributes   = @{ proxyAddresses = "SMTP:$($NewUser.Email)" }
            Enabled           = $true
            ErrorAction       = 'Stop'
        }

        if ($PSCmdlet.ShouldProcess($NewUser.DisplayName, "Create AD user")) {
            New-ADUser @newUserParams
            Write-StatusMessage -Message "Successfully created AD user: $($NewUser.DisplayName)" -Type OK
        }
    } catch {
        Write-StatusMessage -Message "Failed to create AD user: $_" -Type ERROR
        Exit-Script -Message "Failed to create AD user" -ExitCode GeneralError
    }
}

function Set-ADUserOptionalFields {
    param (
        [Parameter(Mandatory)][string]$SamAccountName,
        [Parameter(Mandatory)][pscustomobject]$UserInput,
        [Parameter()][pscustomobject]$TemplateAttributes
    )

    function Format-PhoneNumber {
        param (
            [string]$PhoneNumber
        )

        # Return as-is if null or whitespace
        if ([string]::IsNullOrWhiteSpace($PhoneNumber)) {
            return $PhoneNumber
        }

        # If it already looks like a formatted number, don't touch it
        if ($PhoneNumber -match '^(1-\d{3}-\d{3}-\d{4})$' -or
            $PhoneNumber -match '^\(\d{3}\) \d{3}-\d{4}$' -or
            $PhoneNumber -match '^\+\d{1,3} \d{3}-\d{3}-\d{4}$') {
            return $PhoneNumber
        }

        # Clean the number: keep only digits, unless it starts with a +
        $cleanedNumber = $PhoneNumber -replace '[^\d]', ''
        if ($PhoneNumber.StartsWith('+')) {
            $cleanedNumber = '+' + $cleanedNumber
        }

        # Format based on different cases
        switch -regex ($cleanedNumber) {
            '^1(\d{3})(\d{3})(\d{4})$' {
                return "1-$($matches[1])-$($matches[2])-$($matches[3])"
            }
            '^\+(\d{1,3})(\d{3})(\d{3})(\d{4})$' {
                return "+$($matches[1]) $($matches[2])-$($matches[3])-$($matches[4])"
            }
            '^(\d{3})(\d{3})(\d{4})$' {
                return "($($matches[1])) $($matches[2])-$($matches[3])"
            }
            default {
                Write-StatusMessage -Message "Could not format phone number: $PhoneNumber" -Type WARN
                return $PhoneNumber
            }
        }
    }


    try {
        Write-StatusMessage -Message "Setting optional fields for user: $SamAccountName" -Type INFO
        # Merge TemplateAttributes and UserInput, with UserInput taking precedence
        $mergedInput = @{}

        # List of allowed properties
        $allowedUserProps = @(
            'companyName',
            'officeLocation',
            'department',
            'jobTitle',
            'mobilePhone',
            'businessPhone',
            'faxNumber',
            'streetAddress',
            'city',
            'state',
            'postalCode',
            'country'
        )

        # Add allowed UserInput values first
        foreach ($prop in $allowedUserProps) {
            $value = $UserInput.$prop
            if ($value) {
                $mergedInput[$prop] = $value
            }
        }

        # Fill in missing allowed properties from TemplateAttributes
        foreach ($prop in $allowedUserProps) {
            if (-not $mergedInput.ContainsKey($prop)) {
                $value = $TemplateAttributes.$prop
                if ($value) {
                    $mergedInput[$prop] = $value
                }
            }
        }

        if ($mergedInput.mobilePhone) {
            $mergedInput.mobilePhone = Format-PhoneNumber -PhoneNumber $mergedInput.mobilePhone
        }

        if ($mergedInput.businessPhone) {
            $mergedInput.businessPhone = Format-PhoneNumber -PhoneNumber $mergedInput.businessPhone
        }

        if ($mergedInput.facsimileTelephoneNumber) {
            $mergedInput.facsimileTelephoneNumber = Format-PhoneNumber -PhoneNumber $mergedInput.facsimileTelephoneNumber
        }

        # Map to AD attribute names
        $optionalFields = @{
            company                    = $mergedInput.companyName
            physicalDeliveryOfficeName = $mergedInput.officeLocation
            department                 = $mergedInput.department
            title                      = $mergedInput.jobTitle
            description                = $mergedInput.jobTitle
            mobile                     = $mergedInput.mobilePhone
            telephoneNumber            = $mergedInput.businessPhone
            facsimileTelephoneNumber   = $mergedInput.faxNumber
            streetAddress              = $mergedInput.streetAddress
            l                          = $mergedInput.city
            st                         = $mergedInput.state
            postalCode                 = $mergedInput.postalCode
            c                          = $mergedInput.country
        }

        $filtered = $optionalFields.GetEnumerator() | Where-Object { $_.Value }
        if ($filtered.Count -gt 0) {
            $update = @{}
            foreach ($item in $filtered) { $update[$item.Key] = $item.Value }

            Set-ADUser -Identity $SamAccountName -Replace $update
            Write-StatusMessage -Message "Successfully set optional fields for user: $SamAccountName" -Type OK
        }
    } catch {
        Write-StatusMessage -Message "Failed to set optional fields for user: $SamAccountName - $_" -Type ERROR
    }
}

function Set-ADUserManager {
    param (
        [Parameter(Mandatory)][string]$SamAccountName,
        [Parameter(Mandatory)][string]$ManagerInput
    )

    try {
        Write-StatusMessage -Message "Setting manager for user: $SamAccountName" -Type INFO
        $manager = Get-ADUser -Filter "DisplayName -eq '$ManagerInput' -or UserPrincipalName -eq '$ManagerInput' -or DistinguishedName -eq '$ManagerInput'" -Properties DistinguishedName
        if ($manager) {
            Set-ADUser -Identity $SamAccountName -Manager $manager.DistinguishedName
            Write-StatusMessage -Message "Successfully set manager for user: $SamAccountName" -Type OK
        } else {
            Write-StatusMessage -Message "Manager '$ManagerInput' not found in AD." -Type WARN
        }
    } catch {
        Write-StatusMessage -Message "Failed to set manager for user: $SamAccountName - $_" -Type ERROR
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

        [Parameter()]
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
- Template User   = $(if ($TemplateUser) {$TemplateUser} else {"No template user selected"})
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
            $sourceGroups = Get-ADUser -Filter "DisplayName -eq '$SourceUser' -or UserPrincipalName -eq '$SourceUser'" -Properties MemberOf -ErrorAction Stop
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
            $getTargetUser = Get-ADUser -Filter "DisplayName -eq '$TargetUser' -or UserPrincipalName -eq '$TargetUser'" -Properties MemberOf -ErrorAction Stop
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
                $properties = @(
                    'Id',
                    'Mail',
                    'DisplayName',
                    'GivenName',
                    'Surname',
                    'Department',
                    'officeLocation',
                    'City'
                )
                try {
                    $user = Get-MgUser -UserId $UserEmail -Property $properties -ErrorAction Stop | Select-Object $properties
                    if ($user) {
                        Write-StatusMessage -Message "User $UserEmail successfully synced to Azure AD" -Type OK
                        return $user
                    }
                } catch [Microsoft.Graph.PowerShell.Runtime.RestException] {
                    if ($_.Exception.Response.StatusCode -eq 404) {
                        # User not found yet, this is expected during sync
                        $retryCount++
                        if ($retryCount -ge $MaxRetries) {
                            Write-StatusMessage -Message "Max retry attempts ($MaxRetries) reached" -Type ERROR
                            Exit-Script -Message "Failed to verify user sync after maximum retries" -ExitCode UserNotFound
                        }
                        Write-StatusMessage -Message "Retry $($retryCount) of $($MaxRetries): User not found in Azure AD yet" -Type WARN
                        Start-Sleep -Seconds $RetryIntervalSeconds
                    } else {
                        # Unexpected error, rethrow
                        throw
                    }
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
        [switch]$Required,

        [Parameter()]
        [int]$MaxRetries = 3,

        [Parameter()]
        [int]$RetryDelaySeconds = 5
    )

    try {
        $licenseType = if ($Required) { "required" } else { "ancillary" }
        Write-StatusMessage -Message "Starting $licenseType license assignment for user: $($User.displayName)" -Type INFO

        $totalLicenses = $License.Count
        $currentLicense = 0

        foreach ($lic in $License) {
            $currentLicense++
            Write-StatusMessage -Message "Processing license $currentLicense of $($totalLicenses): $($lic.DisplayName)" -Type INFO

            # Validate license object
            if (-not $lic.SkuId) {
                $errorMsg = "Invalid license object: Missing SkuId"
                if ($Required) {
                    Write-StatusMessage -Message $errorMsg -Type ERROR
                    Exit-Script -Message $errorMsg -ExitCode GeneralError
                } else {
                    Write-StatusMessage -Message $errorMsg -Type WARN
                    continue
                }
            }

            $retryCount = 0
            $licenseAssigned = $false

            do {
                try {
                    # Assign the license
                    $licenseBody = @{
                        addLicenses    = @(@{ skuId = $lic.SkuId })
                        removeLicenses = @()
                    } | ConvertTo-Json -Depth 3

                    $uri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/assignLicense"
                    $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $licenseBody -ContentType "application/json" -ErrorAction Stop

                    # Wait a moment for the license to be processed
                    Start-Sleep -Seconds 2

                    # Verify license assignment
                    $getSku = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($User.Id)/licenseDetails" -ErrorAction Stop

                    if ($getSku.value.skuId -contains $lic.SkuId) {
                        $licenseAssigned = $true
                        Write-StatusMessage -Message "Successfully assigned license: $($lic.DisplayName)" -Type OK
                        break
                    } else {
                        $retryCount++
                        if ($retryCount -lt $MaxRetries) {
                            Write-StatusMessage -Message "License verification failed, retrying ($retryCount of $MaxRetries)..." -Type WARN
                            Start-Sleep -Seconds $RetryDelaySeconds
                        }
                    }
                } catch {
                    $retryCount++
                    if ($retryCount -lt $MaxRetries) {
                        Write-StatusMessage -Message "Error assigning license (attempt $retryCount of $MaxRetries): $($_.Exception.Message)" -Type WARN
                        Start-Sleep -Seconds $RetryDelaySeconds
                    } else {
                        throw
                    }
                }
            } while ($retryCount -lt $MaxRetries)

            if (-not $licenseAssigned) {
                $errorMsg = "Failed to assign license $($lic.DisplayName) after $MaxRetries attempts"
                if ($Required) {
                    Write-StatusMessage -Message $errorMsg -Type ERROR
                    Exit-Script -Message $errorMsg -ExitCode GeneralError
                } else {
                    Write-StatusMessage -Message $errorMsg -Type WARN
                }
            }
        }
    } catch {
        $errorMsg = "Error in Set-UserLicenses: $($_.Exception.Message)"
        if ($Required) {
            Write-StatusMessage -Message $errorMsg -Type ERROR
            Exit-Script -Message $errorMsg -ExitCode GeneralError
        } else {
            Write-StatusMessage -Message $errorMsg -Type WARN
        }
    }
}

function Wait-ForMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Email,

        [int]$MaxWaitTime = 180, # Max wait time in seconds (3 minutes)
        [int]$Interval = 30 # Interval between checks in seconds (30 seconds)
    )

    Write-StatusMessage -Message "Waiting for mailbox to be created for $Email..." -Type INFO
    $elapsedTime = 0

    while ($elapsedTime -lt $MaxWaitTime) {
        try {
            $getMailbox = Get-Mailbox -Identity $Email -ErrorAction Stop
            if ($getMailbox) {
                Write-StatusMessage -Message "Mailbox found for $Email" -Type OK
                return $true
            }
        } catch {
            Write-StatusMessage -Message "$Email not found. Retrying in $Interval seconds..." -Type INFO
        }

        Start-Sleep -Seconds $Interval
        $elapsedTime += $Interval
    }

    Write-StatusMessage -Message "Timeout reached. Mailbox for $Email was not found within $MaxWaitTime seconds." -Type WARN
    return $false
}

function Get-PaginatedGraphResponse {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE', 'PATCH')]
        [string]$Method = "GET",

        [Parameter(Mandatory = $false)]
        [object]$Body = $null,

        [Parameter(Mandatory = $false)]
        [string]$ContentType = "application/json",

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 5
    )

    try {
        $allResults = [System.Collections.Generic.List[object]]::new()
        $nextLink = $Uri
        $pageCount = 0

        while ($nextLink) {
            $retryCount = 0
            $success = $false

            do {
                try {
                    $pageCount++
                    Write-StatusMessage -Message "Fetching page $pageCount..." -Type INFO

                    $response = Invoke-MgGraphRequest -Uri $nextLink -Method $Method -Body $Body -ContentType $ContentType -ErrorAction Stop
                    $success = $true

                    # Handle the response based on its structure
                    if ($response.value) {
                        $allResults.AddRange($response.value)
                        Write-StatusMessage -Message "Retrieved $($response.value.Count) items" -Type INFO
                    } else {
                        $allResults.Add($response)
                    }

                    # Check for next page
                    $nextLink = $response.'@odata.nextLink'

                    # Add a small delay between requests to prevent rate limiting
                    if ($nextLink) {
                        Start-Sleep -Milliseconds 100
                    }
                } catch {
                    $retryCount++
                    if ($retryCount -lt $MaxRetries) {
                        Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed: $($_.Exception.Message)" -Type WARN
                        Start-Sleep -Seconds $RetryDelaySeconds
                    } else {
                        throw
                    }
                }
            } while (-not $success -and $retryCount -lt $MaxRetries)
        }

        # Remove duplicates based on ID
        $uniqueResults = $allResults | Group-Object -Property id | ForEach-Object { $_.Group | Select-Object -First 1 }
        Write-StatusMessage -Message "Retrieved $($uniqueResults.Count) unique items" -Type INFO

        return $uniqueResults
    } catch {
        Write-StatusMessage -Message "Failed to get paginated response: $($_.Exception.Message)" -Type ERROR
        throw
    }
}

function Get-CopyUserGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$userToCopy,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 5
    )

    Write-StatusMessage -Message "Starting group collection from copy user" -Type INFO

    try {
        $encodedFilter = [uri]::EscapeDataString("mail eq '$userToCopy'")
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$encodedFilter"

        $retryCount = 0
        $success = $false

        do {
            try {
                $userResponse = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                $success = $true
            } catch {
                $retryCount++
                if ($retryCount -lt $MaxRetries) {
                    Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed: $($_.Exception.Message)" -Type WARN
                    Start-Sleep -Seconds $RetryDelaySeconds
                } else {
                    throw
                }
            }
        } while (-not $success -and $retryCount -lt $MaxRetries)

        if (-not $userResponse.value -or $userResponse.value.Count -eq 0) {
            Write-StatusMessage -Message "User '$userToCopy' not found" -Type WARN
            return @()
        }

        if ($userResponse.value.Count -gt 1) {
            Write-StatusMessage -Message "Multiple users found for '$userToCopy', skipping" -Type WARN
            return @()
        }

        $copyUser = $userResponse.value[0]
        $uri = "https://graph.microsoft.com/v1.0/users/$($copyUser.id)/memberOf"
        $response = Get-PaginatedGraphResponse -Uri $uri -Method GET

        if (-not $response -or $response.Count -eq 0) {
            Write-StatusMessage -Message "No groups found for user '$userToCopy'" -Type WARN
            return @()
        }

        $filteredGroups = $response | Where-Object {
            $_.'@odata.type' -eq "#microsoft.graph.group" -and
            -not ($_.groupTypes -contains 'DynamicMembership') -and
            -not $_.onPremisesSyncEnabled
        } | ForEach-Object {
            [PSCustomObject]@{
                id          = $_.id
                displayName = $_.displayName
                mail        = $_.mail
                groupType   = if ($_.groupTypes -contains "Unified") {
                    "Unified"
                } elseif ($_.mailEnabled -and $_.securityEnabled) {
                    "Mail-Enabled Security"
                } elseif ($_.securityEnabled) {
                    "Security"
                } elseif ($_.mailEnabled) {
                    "Distribution"
                } else {
                    "Unknown"
                }
            }
        }

        Write-StatusMessage -Message "Found $($filteredGroups.Count) groups for user '$userToCopy'" -Type INFO
        return $filteredGroups
    } catch {
        Write-StatusMessage -Message "Failed to get groups from copy user '$userToCopy': $($_.Exception.Message)" -Type ERROR
        return @()
    }
}

function Get-DepartmentGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$departments,

        [Parameter(Mandatory = $true)]
        [hashtable]$DepartmentMappings
    )

    try {
        $departments += 'All'
        $groupsToAdd = @()
        $invalidDepartments = @()

        foreach ($department in $departments) {
            if ($DepartmentMappings.ContainsKey($department)) {
                $groupsToAdd += $DepartmentMappings[$department]
                Write-StatusMessage -Message "Found groups for department '$department'" -Type INFO
            } else {
                $invalidDepartments += $department
                Write-StatusMessage -Message "No predefined groups for department '$department'" -Type WARN
            }
        }

        if ($groupsToAdd.Count -eq 0) {
            Write-StatusMessage -Message "No groups found for specified departments" -Type WARN
            return @()
        }

        $groupFilters = $groupsToAdd | ForEach-Object { "displayName eq '$_'" }
        $filterQuery = "($($groupFilters -join ' or '))"
        $encodedFilter = [uri]::EscapeDataString($filterQuery)
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=$encodedFilter"

        $response = Get-PaginatedGraphResponse -Uri $uri -Method GET

        $filteredGroups = $response | Where-Object {
            -not $_.onPremisesSyncEnabled
        } | ForEach-Object {
            [PSCustomObject]@{
                id          = $_.id
                displayName = $_.displayName
                mail        = $_.mail
                groupType   = if ($_.groupTypes -contains "Unified") {
                    "Unified"
                } elseif ($_.mailEnabled -and $_.securityEnabled) {
                    "Mail-Enabled Security"
                } elseif ($_.securityEnabled) {
                    "Security"
                } elseif ($_.mailEnabled) {
                    "Distribution"
                } else {
                    "Unknown"
                }
            }
        }

        Write-StatusMessage -Message "Found $($filteredGroups.Count) department groups" -Type INFO
        return $filteredGroups
    } catch {
        Write-StatusMessage -Message "Failed to fetch department groups: $($_.Exception.Message)" -Type ERROR
        return @()
    }
}

function Add-UserToGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [array]$Groups,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UserId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Source,

        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory = $false)]
        [int]$RetryDelaySeconds = 5
    )

    # Initialize counters and tracking
    $successCount = 0
    $failureCount = 0
    $failedGroups = @()

    # Separate groups by type
    $graphGroups = $Groups | Where-Object { $_.GroupType -eq 'Unified' -or $_.GroupType -eq 'Security' }
    $exchangeGroups = $Groups | Where-Object { $_.GroupType -eq 'Distribution' -or $_.GroupType -eq 'Mail-Enabled Security' }
    $unknownGroups = $Groups | Where-Object { $_.GroupType -notin @('Unified', 'Security', 'Distribution', 'Mail-Enabled Security') }

    Write-StatusMessage -Message "Starting group assignments: $($graphGroups.Count) Graph groups, $($exchangeGroups.Count) Exchange groups" -Type INFO

    # Process Graph groups
    if ($graphGroups.Count -gt 0) {
        Write-StatusMessage -Message "Processing Graph groups..." -Type INFO

        foreach ($group in $graphGroups) {
            $retryCount = 0
            $success = $false

            do {
                try {
                    Write-StatusMessage -Message "Processing Graph group: $($group.DisplayName)" -Type INFO

                    $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/members/`$ref"
                    $body = @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$userId"
                    } | ConvertTo-Json

                    Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json" -ErrorAction Stop
                    $success = $true
                    $successCount++
                    Write-StatusMessage -Message "Successfully added user to Graph group: $($group.DisplayName)" -Type OK
                } catch {
                    $retryCount++
                    if ($retryCount -lt $MaxRetries) {
                        # Check for specific Graph API errors
                        $errorMessage = $_.Exception.Message
                        if ($errorMessage -like "*Request_ResourceNotFound*") {
                            Write-StatusMessage -Message "Group '$($group.DisplayName)' not found in Graph API" -Type WARN
                        } elseif ($errorMessage -like "*Request_BadRequest*") {
                            Write-StatusMessage -Message "Invalid request for group '$($group.DisplayName)': $errorMessage" -Type WARN
                        } else {
                            Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed for Graph group '$($group.DisplayName)': $errorMessage" -Type WARN
                        }
                        Start-Sleep -Seconds $RetryDelaySeconds
                    } else {
                        $failureCount++
                        $failedGroups += $group.DisplayName
                        Write-StatusMessage -Message "Failed to add user to Graph group '$($group.DisplayName)' after $MaxRetries attempts: $($_.Exception.Message)" -Type ERROR
                        $success = $true
                    }
                }
            } while (-not $success -and $retryCount -lt $MaxRetries)

            # Add a small delay between groups to prevent overwhelming Graph API
            Start-Sleep -Milliseconds 100
        }
    }

    # Process Exchange groups
    if ($exchangeGroups.Count -gt 0) {
        Write-StatusMessage -Message "Processing Exchange groups..." -Type INFO

        foreach ($group in $exchangeGroups) {
            $retryCount = 0
            $success = $false

            do {
                try {
                    Write-StatusMessage -Message "Processing Exchange group: $($group.DisplayName)" -Type INFO

                    if ($group.mail) {
                        Add-DistributionGroupMember -Identity $group.mail -Member $userId -BypassSecurityGroupManagerCheck -ErrorAction Stop
                        $success = $true
                        $successCount++
                        Write-StatusMessage -Message "Successfully added user to Exchange group: $($group.DisplayName)" -Type OK
                    } else {
                        throw "Group does not have a mail address"
                    }
                } catch {
                    $retryCount++
                    if ($retryCount -lt $MaxRetries) {
                        Write-StatusMessage -Message "Attempt $retryCount of $MaxRetries failed for Exchange group '$($group.DisplayName)': $($_.Exception.Message)" -Type WARN
                        Start-Sleep -Seconds $RetryDelaySeconds
                    } else {
                        $failureCount++
                        $failedGroups += $group.DisplayName
                        Write-StatusMessage -Message "Failed to add user to Exchange group '$($group.DisplayName)' after $MaxRetries attempts: $($_.Exception.Message)" -Type ERROR
                        $success = $true
                    }
                }
            } while (-not $success -and $retryCount -lt $MaxRetries)

            # Add a small delay between groups to prevent overwhelming Exchange
            Start-Sleep -Seconds 1
        }
    }

    # Handle unknown group types
    if ($unknownGroups.Count -gt 0) {
        Write-StatusMessage -Message "Skipping $($unknownGroups.Count) unknown group types" -Type WARN
        foreach ($group in $unknownGroups) {
            Write-StatusMessage -Message "Unknown group type: $($group.DisplayName) ($($group.GroupType))" -Type WARN
        }
    }

    # Output summary
    Write-StatusMessage -Message "Group assignment summary: $successCount successful, $failureCount failed" -Type INFO
    if ($failedGroups.Count -gt 0) {
        Write-StatusMessage -Message "Failed groups: $($failedGroups -join ', ')" -Type WARN
    }

    # Return both success and failure information
    return @{
        SuccessCount = $successCount
        FailureCount = $failureCount
        FailedGroups = $failedGroups
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
        [int]$AssignedGroupCount,

        [Parameter()]
        [hashtable]$GroupOperationSummary = @{
            CopyUserGroups   = @{
                Count  = 0
                Groups = @()
                Failed = @()
            }
            DepartmentGroups = @{
                Count  = 0
                Groups = @()
                Failed = @()
            }
            TotalFailed      = 0
        }
    )

    Write-StatusMessage -Message "Preparing final summary" -Type INFO

    # Validate user object
    if (-not $User.Id -or -not $User.displayName -or -not $User.mail) {
        Write-StatusMessage -Message "Warning: User object is missing required properties" -Type WARN
    }

    # Calculate successful assignments
    $totalSuccessful = $AssignedGroupCount - $GroupOperationSummary.TotalFailed

    # Build summary parts
    $summaryParts = @(
        "SUMMARY OF ACTIONS",
        "========================================",
        "User Creation Status:",
        "----------------------------------------",
        "- EntraID: $($User.Id)",
        "- Display Name: $($User.displayName)",
        "- Email Address: $($User.mail)",
        "- Password: $Password",
        "- Template User: $(if ($TemplateUser) {$userInput.userToCopy} else {'No template user selected.'})",
        "",
        "Group Assignment Status:",
        "----------------------------------------",
        "- Total Groups Attempted: $AssignedGroupCount",
        "- Successfully Added: $totalSuccessful",
        "- Failed Additions: $($GroupOperationSummary.TotalFailed)"
    )

    # Add group operation details
    if ($GroupOperationSummary.CopyUserGroups.Count -gt 0) {
        $summaryParts += @(
            "",
            "Template User Groups:",
            "- Total Groups: $($GroupOperationSummary.CopyUserGroups.Count)",
            "- Successfully Added: $($GroupOperationSummary.CopyUserGroups.Count - $GroupOperationSummary.CopyUserGroups.Failed.Count)"
        )
        if ($GroupOperationSummary.CopyUserGroups.Groups.Count -gt 0) {
            $summaryParts += "- Groups: $($GroupOperationSummary.CopyUserGroups.Groups -join ', ')"
        }
        if ($GroupOperationSummary.CopyUserGroups.Failed.Count -gt 0) {
            $summaryParts += "- Failed Groups: $($GroupOperationSummary.CopyUserGroups.Failed -join ', ')"
        }
    }

    if ($GroupOperationSummary.DepartmentGroups.Count -gt 0) {
        $summaryParts += @(
            "",
            "Department Groups:",
            "- Total Groups: $($GroupOperationSummary.DepartmentGroups.Count)",
            "- Successfully Added: $($GroupOperationSummary.DepartmentGroups.Count - $GroupOperationSummary.DepartmentGroups.Failed.Count)"
        )
        if ($GroupOperationSummary.DepartmentGroups.Groups.Count -gt 0) {
            $summaryParts += "- Groups: $($GroupOperationSummary.DepartmentGroups.Groups -join ', ')"
        }
        if ($GroupOperationSummary.DepartmentGroups.Failed.Count -gt 0) {
            $summaryParts += "- Failed Groups: $($GroupOperationSummary.DepartmentGroups.Failed -join ', ')"
        }
    }

    # Add warnings and important notes
    $warnings = @()

    if ($skippedTemplateUserGroups) {
        $warnings += "Template group copy was skipped"
    }

    # Check for group count mismatches
    $totalGroupsFound = $GroupOperationSummary.CopyUserGroups.Count + $GroupOperationSummary.DepartmentGroups.Count
    if ($totalGroupsFound -ne $AssignedGroupCount) {
        $warnings += "Group count mismatch detected: Found=$totalGroupsFound, Attempted=$AssignedGroupCount"
    }

    if ($GroupOperationSummary.TotalFailed -gt 0) {
        $warnings += "Group assignments failed: $($GroupOperationSummary.TotalFailed) total failures"

        # Add specific failure details
        if ($GroupOperationSummary.CopyUserGroups.Failed.Count -gt 0) {
            $warnings += "- Template user group failures: $($GroupOperationSummary.CopyUserGroups.Failed.Count)"
        }
        if ($GroupOperationSummary.DepartmentGroups.Failed.Count -gt 0) {
            $warnings += "- Department group failures: $($GroupOperationSummary.DepartmentGroups.Failed.Count)"
        }
    }

    if ($warnings.Count -gt 0) {
        $summaryParts += @(
            "",
            "WARNINGS:",
            "----------------------------------------"
        )
        $warnings | ForEach-Object {
            $summaryParts += "- $_"
        }
    }

    # Add important notes
    $summaryParts += @(
        "",
        "IMPORTANT NOTES:",
        "----------------------------------------",
        "1. Please record this password now - it should be needed for the user's first login.",
        "2. Verify all group assignments in the EntraID portal.",
        "3. Check the user's mailbox status in Exchange Online.",
        "4. Review any warnings above and take appropriate action.",
        "5. If any group assignments failed, manual remediation may be required."
    )

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

#Region Main Execution

Write-Host "  [ ] Initializing progress tracking..." -NoNewline -ForegroundColor Yellow
$progressSteps = @(
    @{ Number = 0; Name = "Initialization"; Description = "Loading configuration and connecting services" }
    @{ Number = 1; Name = "User Input"; Description = "Gathering new user details" }
    @{ Number = 2; Name = "Validation"; Description = "Validating inputs and building user creation prerequisites" }
    @{ Number = 3; Name = "New User Creation"; Description = "Creating user in Entra" }
    @{ Number = 4; Name = "License Setup"; Description = "Assigning licenses" }
    @{ Number = 5; Name = "Set Timezone"; Description = "Setting Timezone for new user" }
    @{ Number = 6; Name = "Mailbox Provisioning"; Description = "Waiting for Exchange to provision mailbox" }
    @{ Number = 7; Name = "Entra Group Assignment"; Description = "Assigning Entra Groups" }
    @{ Number = 8; Name = "Email to SOC for KnowBe4"; Description = "Sending SOC notification email for KnowBe4 setup" }
    @{ Number = 9; Name = "OneDrive Provisioning"; Description = "Provisioning new users OneDrive" }
    @{ Number = 10; Name = "Configuring BookWithMeId"; Description = "Configuring BookWithMeId" }
    @{ Number = 11; Name = "Cleanup and Summary"; Description = "Running cleanup and summary" }
)
$script:totalSteps = $progressSteps.Count
$script:currentStep = 0
Write-Host "`r  [✓] Progress tracking initialized" -ForegroundColor Green

Write-Host "`n  Beginning New User Request..." -ForegroundColor Cyan

try {

    # Initialize tracking variables
    $script:startTime = Get-Date
    $script:errorCount = 0
    $script:warningCount = 0
    $script:successCount = 0

    # Step: Initialization
    Write-ProgressStep -StepName 'Initialization'

    # Load configuration
    $config = Get-ScriptConfig
    if (-not $config) {
        throw "Failed to load configuration"
    }

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

    # Step: User Input
    Write-ProgressStep -StepName 'User Input'
    $userInput = Get-NewUserRequest
    if (-not $userInput) {
        throw "Failed to get user input"
    }

    # Set variables after input
    if ($userInput.TestModeEnabled -eq 'True') {
        $script:TestMode = $true
        Write-StatusMessage -Message "Test mode enabled - using test email: $($config.TestMode.Email)" -Type INFO
    }
    $script:TestEmailAddress = $config.TestMode.Email

    # Process copy operations
    switch ($userInput.copyUserOperations) {
        'Copy Attributes' { $copyUserAttribues = $true }
        'Copy Groups' { $copyUserGroups = $true }
        'Copy Attributes and Groups' {
            $copyUserAttribues = $true
            $copyUserGroups = $true
        }
    }

    # Step: Validation and Preparation (AD)
    Write-ProgressStep -StepName 'Validation'

    $passwordResult = New-ReadablePassword -GitHubToken $config.GitHub.Token
    if (-not $passwordResult) {
        throw "Failed to generate password"
    }

    # Get template user (if copying)
    if ($userInput.userToCopy) {
        $templateUser = Get-TemplateUser -UserToCopy $userInput.userToCopy
        if (-not $templateUser) {
            throw "Failed to get template user: $($userInput.userToCopy)"
        }
        $templateAttributes = Get-ADUserCopiedAttributes -TemplateUser $templateUser
        $templateUserManager = $templateAttributes.Manager

        if ($userInput.domain) {
            $domain = '@' + $($userInput.domain)
        } else {
            $domain = $templateUser.UserPrincipalName -replace '.+?(?=@)'
        }

        $destinationOU = $templateUser.DistinguishedName.split(",", 2)[1]
    } else {
        if ($userInput.domain) {
            $domain = $userInput.domain
        } else {
            $domain = '@compassmsp.com'
        }

        $destinationOU = 'OU=Offices,OU=CompassMSP,DC=COMPASSMSP,DC=com'
        Write-StatusMessage -Message "Default destination OU selected. Move to correct OU after account creation." -Type WARN
    }

    # Initialize the parameters for New-UserProperties and check for duplicate SamAccountName/Mail
    $newUserParams = @{
        DisplayName = $userInput.displayName
        Domain      = $domain
    }

    # Conditionally add FirstName and LastName if they exist
    if ($userInput.givenName -and $userInput.surname) {
        $newUserParams.FirstName = $userInput.givenName
        $newUserParams.LastName = $userInput.surname
    }

    # Conditionally add Email if it exists
    if ($userInput.userPrincipalName) {
        $newUserParams.userPrincipalName = $userInput.userPrincipalName
    }

    # Call the New-UserProperties function with the constructed parameters
    $newUserProperties = New-UserProperties @newUserParams
    if (-not $newUserProperties) {
        throw "Failed to create user properties"
    }

    # Show summary and get confirmation before creating
    $confirmUserParams = @{
        NewUserProperties = $newUserProperties
        DestinationOU     = $destinationOU
        Password          = $passwordResult.PlainPassword
    }

    if ($templateUser) {
        $confirmUserParams.TemplateUser = $userInput.userToCopy
    }

    Confirm-UserCreation @confirmUserParams

    # Step: AD User Creation / Set Attributes
    Write-ProgressStep -StepName 'New User AD Creation'

    New-ADUserStandard -NewUser $newUserProperties -Password $passwordResult.SecurePassword -DestinationOU $destinationOU

    # Set optional fields (from template + form)
    $setUserParams = @{
        SamAccountName = $newUserProperties.SamAccountName
        UserInput      = $userInput
    }

    if ($templateAttributes) {
        $setUserParams.TemplateAttributes = $templateAttributes
    }

    Set-ADUserOptionalFields @setUserParams

    # Set manager
    if ($userInput.manager -or $templateUserManager) {
        $setUserManagerParams = @{
            SamAccountName = $newUserProperties.SamAccountName
        }

        if ($userInput.manager) {
            $setUserManagerParams.ManagerInput = $userInput.manager
        } elseif ($copyUserAttribues -eq $true) {
            $setUserManagerParams.ManagerInput = $templateUserManager
        }

        Set-ADUserManager @setUserManagerParams
    } else {
        Write-StatusMessage -Message 'No manager user object selected. Skipping...' -Type WARN
    }

    # Step: AD Group Copy
    Write-ProgressStep -StepName 'AD Group Copy'
    if ($copyUserGroups -eq $true) {
        Copy-UserADGroups -SourceUser $userInput.userToCopy -TargetUser $newUserProperties.displayName
    } else {
        Write-StatusMessage -Message 'No group copy operation selected. Skipping...' -Type INFO
    }

    # Azure Sync
    Write-ProgressStep -StepName 'Azure Sync'
    $MgUser = Wait-ForADUserSync -UserEmail $newUserProperties.Email
    if (-not $MgUser) {
        throw "Failed to sync user to Azure AD"
    }

    # License Assignment
    Write-ProgressStep -StepName 'License Assignment'
    Write-StatusMessage -Message "Setting Usage Location for new user..." -Type INFO
    if ($userInput.usageLocation) {
        $setUsageLocation = $userInput.usageLocation
    } else {
        $setUsageLocation = 'US'
    }
    Update-MgUser -UserId $MgUser.Id -UsageLocation $setUsageLocation

    # Required license - will exit on failure
    Set-UserLicenses -User $MgUser -License $userInput.requiredLicense -Required

    # Ancillary licenses - will continue on failure
    if ($userInput.ancillaryLicense) {
        Set-UserLicenses -User $MgUser -License $userInput.ancillaryLicense
    }

    # Step: Wait for Mailbox
    Write-ProgressStep -StepName 'Mailbox Provisioning'
    Write-StatusMessage -Message "Waiting for the mailbox to provision in 365..." -Type INFO
    # Wait for mailbox to be created
    if (-not (Wait-ForMailbox -Email $MgUser.Mail)) {
        Write-StatusMessage -Message "Mailbox not yet provisioned. Some group operations may fail. Please verify group assignments script completes." -Type WARN
    }

    # Step: TimeZone Assignment
    # Set Timezone after license
    Write-StatusMessage -Message "Setting Timezone for new user" -Type INFO
    if ($userInput.timeZone) {
        if ($userInput.timeZone -eq 'US Mountain Standard Time (Arizona)') {
            $userinput.timeZone = 'US Mountain Standard Time'
        }
        Set-MailboxRegionalConfiguration -Identity $($MgUser.Mail) -TimeZone $userinput.timeZone
    } else {
        Write-StatusMessage -Message "Timezone for new user not selected. Skipping" -Type ERROR
    }

    # Step: Entra Group Assignment
    Write-ProgressStep -StepName 'Entra Group Assignment'

    $allFilteredGroups = [System.Collections.Generic.List[object]]::new()
    $groupOperationSummary = @{
        CopyUserGroups   = @{
            Count  = 0
            Groups = @()
            Failed = @()
        }
        DepartmentGroups = @{
            Count  = 0
            Groups = @()
            Failed = @()
        }
        TotalFailed      = 0
    }

    # Start Group Add Operations
    try {
        # Process copy user groups if selected
        if ($userInput.userToCopy -and $copyUserGroups) {
            Write-StatusMessage -Message 'Copy template user groups selected...' -Type INFO
            try {
                $copyGroups = Get-CopyUserGroups -userToCopy $userInput.userToCopy
                if ($copyGroups.Count -gt 0) {
                    $allFilteredGroups.AddRange($copyGroups)
                    $groupOperationSummary.CopyUserGroups.Count = $copyGroups.Count
                    $groupOperationSummary.CopyUserGroups.Groups = $copyGroups.DisplayName
                    Write-StatusMessage -Message "Found $($copyGroups.Count) groups from template user" -Type INFO
                } else {
                    Write-StatusMessage -Message 'No groups found for template user' -Type WARN
                    $groupOperationSummary.CopyUserGroups.Failed += "No groups found for template user"
                }
            } catch {
                Write-StatusMessage -Message "Failed to get groups from template user: $($_.Exception.Message)" -Type ERROR
                $groupOperationSummary.CopyUserGroups.Failed += "Failed to get template user groups: $($_.Exception.Message)"
                $groupOperationSummary.TotalFailed += $copyGroups?.Count ?? 0
            }
        } else {
            Write-StatusMessage -Message 'No copy operations selected. Trying group mapping via selected department group options.' -Type WARN
            $skippedTemplateUserGroups = $true
        }

        # Process department groups if selected
        if ($userInput.DepartmentGroups) {
            Write-StatusMessage -Message "Processing department groups for: $($userInput.DepartmentGroups -join ', ')" -Type INFO
            try {
                $setDepartmentMappings = @{
                    'All'       = @('All Company')
                    'NFL - ROC' = @('CMSP - NFL REGION', '[NFL] OnCall', '[NFL] Alerts')
                    'SFL - ROC' = @('CMSP - SFL REGION', '[SFL] OnCall', '[SFL] Alerts')
                }

                $deptGroups = Get-DepartmentGroups -departments $userInput.DepartmentGroups -DepartmentMappings $setDepartmentMappings
                if ($deptGroups.Count -gt 0) {
                    $allFilteredGroups.AddRange($deptGroups)
                    $groupOperationSummary.DepartmentGroups.Count = $deptGroups.Count
                    $groupOperationSummary.DepartmentGroups.Groups = $deptGroups.DisplayName
                    Write-StatusMessage -Message "Found $($deptGroups.Count) department groups" -Type INFO
                } else {
                    Write-StatusMessage -Message 'No department groups found' -Type WARN
                    $groupOperationSummary.DepartmentGroups.Failed += "No department groups found"
                }
            } catch {
                Write-StatusMessage -Message "Failed to get department groups: $($_.Exception.Message)" -Type ERROR
                $groupOperationSummary.DepartmentGroups.Failed += "Failed to get department groups: $($_.Exception.Message)"
                $groupOperationSummary.TotalFailed += $deptGroups?.Count ?? 0
            }
        }

        # Process combined groups if any were found
        if ($allFilteredGroups.Count -gt 0) {
            Write-StatusMessage -Message "Processing $($allFilteredGroups.Count) total groups..." -Type INFO

            # Get unique groups and validate
            $uniqueGroups = $allFilteredGroups | Select-Object -Unique -Property id, DisplayName, Mail, GroupType
            Write-StatusMessage -Message "Found $($uniqueGroups.Count) unique groups after deduplication" -Type INFO

            try {
                # Add groups to user
                $groupAddResults = Add-UserToGroups -UserId $mgUser.id -Groups $uniqueGroups -Source "combined groups"

                Write-StatusMessage -Message "User now belongs to $($groupAddResults.SuccessCount) groups" -Type INFO
                if ($groupAddResults.FailureCount -gt 0) {
                    Write-StatusMessage -Message "Failed to add $($groupAddResults.FailureCount) groups" -Type WARN
                    $groupOperationSummary.TotalFailed += $groupAddResults.FailureCount

                    # Add failed groups to respective categories
                    foreach ($failedGroup in $groupAddResults.FailedGroups) {
                        if ($copyGroups.DisplayName -contains $failedGroup) {
                            $groupOperationSummary.CopyUserGroups.Failed += $failedGroup
                        }
                        if ($deptGroups.DisplayName -contains $failedGroup) {
                            $groupOperationSummary.DepartmentGroups.Failed += $failedGroup
                        }
                    }
                }
            } catch {
                Write-StatusMessage -Message "Error during group addition: $($_.Exception.Message)" -Type ERROR
                Write-StatusMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -Type ERROR

                # Set failure counts but continue execution
                $groupOperationSummary.TotalFailed += $uniqueGroups.Count
                Write-StatusMessage -Message "Marking all groups as failed but continuing execution" -Type WARN

                # Create empty results object for downstream processing
                $groupAddResults = @{
                    SuccessCount = 0
                    FailureCount = $uniqueGroups.Count
                    FailedGroups = $uniqueGroups.DisplayName
                }
            }
        } else {
            Write-StatusMessage -Message 'No groups found to add. Please assign groups manually' -Type WARN
            # Initialize empty results for downstream processing
            $groupAddResults = @{
                SuccessCount = 0
                FailureCount = 0
                FailedGroups = @()
            }
        }
    } catch {
        Write-StatusMessage -Message "Unexpected error during group operations: $($_.Exception.Message)" -Type ERROR
        Write-StatusMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -Type ERROR

        # Initialize results for downstream processing
        $groupAddResults = @{
            SuccessCount = 0
            FailureCount = $allFilteredGroups.Count
            FailedGroups = $allFilteredGroups.DisplayName
        }

        # Ensure TotalFailed reflects any uncaught failures
        if ($groupOperationSummary.TotalFailed -eq 0) {
            $groupOperationSummary.TotalFailed = $allFilteredGroups.Count
        }
    }

    # Step: Email to SOC for KnowBe4
    Write-ProgressStep -StepName 'Email to SOC for KnowBe4'
    try {
        $emailSubject = "KB4 – New User"
        $emailContent = "The following user need to be added to the CompassMSP KnowBe4 account. <p> $($MgUser.DisplayName) <br> $($MgUser.Mail)"
        $MsgFrom = $config.Email.NotificationFrom
        $ToAddress = $config.Email.NotificationTo
        Send-GraphMailMessage -FromAddress $MsgFrom -ToAddress $ToAddress -Subject $emailSubject -Content $emailContent
    } catch {
        Write-StatusMessage -Message "Failed to send KnowBe4 notification email: $($_.Exception.Message)" -Type ERROR
    }

    # Step: OneDrive Provisioning
    Write-ProgressStep -StepName 'OneDrive Provisioning'
    Write-StatusMessage -Message "OneDrive provisioning is currently disabled" -Type INFO

    # Step: BookWithMeId Setup
    Write-ProgressStep -StepName 'Configuring BookWithMeId'
    Set-UserBookWithMeId -User $MgUser -SamAccountName $newUserProperties.SamAccountName

    # Step 11: Cleanup and Summary
    Write-ProgressStep -StepName 'Cleanup and Summary'
    Write-StatusMessage -Message "Disconnecting from Exchange Online and Graph." -Type INFO

    Connect-ServiceEndpoints -Disconnect

    Write-StatusMessage -Message "Building final summary..." -Type INFO

    Start-NewUserFinalize -User $MgUser `
        -Password $passwordResult.PlainPassword `
        -AssignedGroupCount $groupAddResults.SuccessCount `
        -GroupOperationSummary $groupOperationSummary

    # Show duration
    $endTime = Get-Date
    $duration = $endTime - $startTime
    Write-StatusMessage "Script completed in $($duration.TotalMinutes.ToString('F2')) minutes" -Type INFO

    # Give user time to read/copy the summary
    Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
    Read-Host | Out-Null

    # Clear the progress bar
    Write-Progress -Activity "New User Creation" -Completed

    Exit-Script -Message "$($MgUser.displayName) has been successfully created." -ExitCode Success

} catch {
    Write-StatusMessage -Message "Script failed: $($_.Exception.Message)" -Type ERROR
    Write-StatusMessage -Message "Stack Trace: $($_.ScriptStackTrace)" -Type ERROR

    # Clear the progress bar
    Write-Progress -Activity "New User Creation" -Status "Failed" -PercentComplete 100

    Exit-Script -Message "Script failed during execution" -ExitCode GeneralError
}