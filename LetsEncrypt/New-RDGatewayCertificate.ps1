#Requires -Modules Posh-ACME, RemoteDesktop

<#
.SYNOPSIS
    Sets up initial Let's Encrypt certificate for RD Gateway using DNS or HTTP validation.

.DESCRIPTION
    This script handles the initial setup of Let's Encrypt certificates for Remote Desktop Gateway servers.
    It supports both DNS and HTTP validation methods, with two ways to configure DNS validation:
    1. Interactive setup (-DNS switch)
    2. Configuration file (-LoadConfig)

    Supported DNS providers:
    - DNSMadeEasy
    - AWS Route53
    - Cloudflare (API Token or Global API Key)

    For certificate renewals, use Renew-RDGatewayCertificate.ps1.

.PARAMETER DNS
    Use DNS validation with interactive provider setup.

.PARAMETER HTTP
    Use HTTP validation (default).

.PARAMETER LoadConfig
    Path to a JSON configuration file for DNS validation settings.

.PARAMETER Email
    Contact email address for Let's Encrypt registration.

.PARAMETER Domain
    Domain name for the certificate.

.PARAMETER Stage
    Use Let's Encrypt staging environment for testing.

.PARAMETER BaseDirectory
    Directory to store certificates and ACME account information.
    Default: "C:\ProgramData\RDGatewayCerts"

.EXAMPLE
    # Basic HTTP validation
    .\New-RDGatewayCertificate.ps1 -Domain "gateway.contoso.com" -Email "admin@contoso.com"

.EXAMPLE
    # Interactive DNS validation setup
    .\New-RDGatewayCertificate.ps1 -DNS -Domain "gateway.contoso.com" -Email "admin@contoso.com"

.EXAMPLE
    # Use config file for DNS validation
    .\New-RDGatewayCertificate.ps1 -LoadConfig "config.json"

    # Sample config.json:
    {
        "Email": "admin@contoso.com",
        "Domain": "gateway.contoso.com",
        "DNSProvider": "DMEasy",
        "DMEKey": "your-api-key",
        "DMESecret": "your-secret-key"
    }

.EXAMPLE
    # Initial setup with DNS and scheduled task creation
    .\New-RDGatewayCertificate.ps1 -DNS -Domain "gateway.contoso.com" -Email "admin@contoso.com"

    # Create renewal task
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Renew-RDGatewayCertificate.ps1" -Domain "gateway.contoso.com"'
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 3AM
    Register-ScheduledTask -TaskName "RD Gateway Certificate Renewal" `
        -Action $action -Trigger $trigger -User "SYSTEM" -RunLevel Highest

.NOTES
    - Script must run with administrative privileges
    - For DNS validation, credentials are saved by Posh-ACME for renewals
    - HTTP validation requires port 80 access
    - Use staging environment (-Stage) for testing
    - Create scheduled task for automatic renewals

.AUTHORS
    - Chris Williams
    - Last Updated: 01/22/2025
#>

[CmdletBinding(DefaultParameterSetName = 'HTTP')]
param(
    [Parameter(ParameterSetName = 'DNS')]
    [switch]$DNS,

    [Parameter(ParameterSetName = 'HTTP')]
    [switch]$HTTP,

    [Parameter(ParameterSetName = 'Config')]
    [string]$LoadConfig,

    [Parameter(Mandatory = $true, ParameterSetName = 'DNS')]
    [Parameter(Mandatory = $true, ParameterSetName = 'HTTP')]
    [ValidatePattern('^[^@]+@[^@]+\.[^@]+$')]
    [string]$Email,

    [Parameter(Mandatory = $true, ParameterSetName = 'DNS')]
    [Parameter(Mandatory = $true, ParameterSetName = 'HTTP')]
    [ValidatePattern('^[a-zA-Z0-9][a-zA-Z0-9-\.]*[a-zA-Z0-9]$')]
    [string]$Domain,

    [switch]$Stage,

    [string]$BaseDirectory = "C:\ProgramData\RDGatewayCerts"
)

#region Core Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "$timestamp [$Level] $Message"

        # Write to console
        Write-Host $logMessage

        # Set up log file path
        $logDir = Join-Path $BaseDirectory "Logs"
        $logFile = Join-Path $logDir "RDGateway-Certificate.log"

        # Ensure log directory exists
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        # Append to log file
        $logMessage | Out-File -FilePath $logFile -Append -Encoding UTF8

        # Rotate logs if file gets too large (>10MB)
        if ((Get-Item $logFile).Length -gt 10MB) {
            $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
            $archiveLog = Join-Path $logDir "RDGateway-Certificate-$timestamp.log"
            Move-Item -Path $logFile -Destination $archiveLog
            # Keep only last 5 archived logs
            Get-ChildItem -Path $logDir -Filter "RDGateway-Certificate-*.log" |
                Sort-Object LastWriteTime -Descending |
                Select-Object -Skip 5 |
                Remove-Item -Force
        }
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
        Write-Host $logMessage
    }
}

function Install-RDGatewayCertificate {
    param(
        [Parameter(Mandatory)]
        [ValidateScript({ Test-Path $_ })]
        [string]$CertPath,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Thumbprint,

        [Parameter(Mandatory)]
        [string]$CertPassword
    )

    try {
        # Validate certificate
        $pfxCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $pfxCert.Import($CertPath, $CertPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)

        if ($pfxCert.Thumbprint -ne $Thumbprint) {
            throw "Certificate thumbprint mismatch"
        }

        Write-Log "Certificate validated successfully. Thumbprint: $Thumbprint"

        # Get broker FQDN
        $broker = "$env:COMPUTERNAME.$((Get-WmiObject Win32_ComputerSystem).Domain)"
        Write-Log "Using connection broker: $broker"

        $securePassword = ConvertTo-SecureString -String $CertPassword -AsPlainText -Force

        # Install for RD Gateway roles
        foreach ($role in @('RDGateway', 'RDWebAccess', 'RDRedirector', 'RDPublishing')) {
            Write-Log "Installing certificate for $role..."
            Set-RDCertificate -Role $role -ImportPath $CertPath -Password $securePassword -ConnectionBroker $broker -Force
            Write-Log "Successfully installed certificate for $role"
        }

        # Handle RDWebClient if available
        try {
            Write-Log "Checking for RDWebClient module..."
            Import-Module RDWebClientManagement -ErrorAction Stop

            Write-Log "Converting PFX to CER for RDWebClient..."
            $cerPath = $CertPath -replace '\.pfx$', '.cer'
            $webClientCert = $null

            try {
                $webClientCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                $webClientCert.Import($CertPath, $CertPassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
                [System.IO.File]::WriteAllBytes($cerPath, $webClientCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

                Write-Log "Installing certificate for RDWebClient..."
                Import-RDWebClientBrokerCert $cerPath -ErrorAction Stop
                Install-RDWebClientPackage -ErrorAction Stop
                Publish-RDWebClientPackage -Type Production -Latest -ErrorAction Stop
                Write-Log "RDWebClient configuration completed successfully"
            } finally {
                if ($webClientCert) { $webClientCert.Dispose() }
                if (Test-Path $cerPath) {
                    Remove-Item $cerPath -Force
                    Write-Log "Cleaned up temporary CER file"
                }
            }
        } catch {
            if ($_.Exception.Message -match "Could not load file or assembly 'RDWebClientManagement'") {
                Write-Log "RDWebClient module not found - skipping configuration" -Level Warning
            } else {
                Write-Log "Failed to configure RDWebClient: $_" -Level Error
                throw
            }
        }

        Write-Log "Certificate installation completed successfully"
    }
    catch {
        Write-Log "Failed to install certificate: $_" -Level Error
        throw
    }
    finally {
        if ($pfxCert) { $pfxCert.Dispose() }
    }
}
#endregion

#region DNS Configuration Functions
function Import-CertificateConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    try {
        Write-Log "Loading configuration from $ConfigPath"

        if (-not (Test-Path $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }

        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

        # Convert to hashtable for easier use
        $configHash = @{
            Email = $config.Email
            Domain = $config.Domain
            DNSProvider = $config.DNSProvider
            DNSCredentials = @{}
        }

        # Validate required fields
        if (-not $configHash.Email -or -not $configHash.Domain) {
            throw "Configuration must include Email and Domain"
        }

        # Handle DNS credentials based on provider
        if ($config.DNSProvider) {
            switch ($config.DNSProvider) {
                'DMEasy' {
                    $configHash.DNSCredentials = @{
                        DMEKey = $config.DMEKey
                        DMESecret = $config.DMESecret
                    }
                }
                'Route53' {
                    $configHash.DNSCredentials = @{
                        R53AccessKey = $config.R53AccessKey
                        R53SecretKey = $config.R53SecretKey
                    }
                }
                'Cloudflare' {
                    if ($config.CFUseToken) {
                        $configHash.DNSCredentials = @{
                            CFUseToken = $true
                            CFToken = $config.CFToken
                        }
                    } else {
                        $configHash.DNSCredentials = @{
                            CFUseToken = $false
                            CFAuthEmail = $config.CFAuthEmail
                            CFAuthKey = $config.CFAuthKey
                        }
                    }
                }
                default {
                    throw "Unsupported DNS provider in configuration: $($config.DNSProvider)"
                }
            }
        }

        Write-Log "Configuration loaded successfully"
        return $configHash
    } catch {
        Write-Log "Failed to load configuration: $_" -Level Error
        throw
    }
}

function Get-DnsPluginParameters {
    param(
        [Parameter(Mandatory)]
        [string]$Domain,
        [Parameter(Mandatory)]
        [string]$Email,
        [string]$DNSProvider,
        [hashtable]$DNSCredentials
    )

    try {
        # Build base certificate parameters
        $certParams = @{
            Domain = $Domain
            Contact = $Email
        }

        if ($DNSProvider -and -not $DNSCredentials) {
            throw "DNS credentials are required when specifying a DNS provider"
        }

        if ($DNSProvider -and $DNSCredentials) {
            Write-Log "Using DNS validation with $DNSProvider"
            # Use provided DNS credentials
            switch ($DNSProvider) {
                'DMEasy' {
                    $certParams += @{
                        Plugin = "DMEasy"
                        PluginArgs = @{
                            DMEKey = $DNSCredentials.DMEKey
                            DMESecret = (ConvertTo-SecureString $DNSCredentials.DMESecret -AsPlainText -Force)
                        }
                    }
                }
                'Route53' {
                    $certParams += @{
                        Plugin = "Route53"
                        PluginArgs = @{
                            R53AccessKey = $DNSCredentials.R53AccessKey
                            R53SecretKey = $DNSCredentials.R53SecretKey
                        }
                    }
                }
                'Cloudflare' {
                    if ($DNSCredentials.CFUseToken) {
                        $certParams += @{
                            Plugin = "Cloudflare"
                            PluginArgs = @{
                                CFToken = $DNSCredentials.CFToken
                            }
                        }
                    } else {
                        $certParams += @{
                            Plugin = "Cloudflare"
                            PluginArgs = @{
                                CFAuthEmail = $DNSCredentials.CFAuthEmail
                                CFAuthKey = $DNSCredentials.CFAuthKey
                            }
                        }
                    }
                }
                default {
                    throw "Unsupported DNS provider: $DNSProvider"
                }
            }
        } else {
            # Interactive mode
            Write-Log "Starting interactive DNS provider setup..."
            Write-Host "`nSelect DNS Provider:"
            Write-Host "1. DNSMadeEasy"
            Write-Host "2. AWS Route53"
            Write-Host "3. Cloudflare (API Token)"
            Write-Host "4. Cloudflare (Global API Key)"

            $choice = Read-Host "`nEnter choice (1-4)"

            switch ($choice) {
                "1" {
                    $dmeKey = Read-Host "Enter DNSMadeEasy API Key"
                    $dmeSecret = Read-Host "Enter DNSMadeEasy Secret Key"
                    $secureSecret = ConvertTo-SecureString $dmeSecret -AsPlainText -Force
                    $certParams += @{
                        Plugin = "DMEasy"
                        PluginArgs = @{
                            DMEKey = $dmeKey
                            DMESecret = $secureSecret
                        }
                    }
                }
                "2" {
                    $certParams += @{
                        Plugin = "Route53"
                        PluginArgs = @{
                            R53AccessKey = Read-Host "Enter AWS Access Key"
                            R53SecretKey = Read-Host "Enter AWS Secret Key"
                        }
                    }
                }
                "3" {
                    $certParams += @{
                        Plugin = "Cloudflare"
                        PluginArgs = @{
                            CFToken = Read-Host "Enter Cloudflare API Token"
                        }
                    }
                }
                "4" {
                    $certParams += @{
                        Plugin = "Cloudflare"
                        PluginArgs = @{
                            CFAuthEmail = Read-Host "Enter Cloudflare Account Email"
                            CFAuthKey = Read-Host "Enter Cloudflare Global API Key"
                        }
                    }
                }
                default { throw "Invalid choice" }
            }
        }

        return $certParams
    } catch {
        Write-Log "Failed to configure DNS plugin parameters: $_" -Level Error
        throw
    }
}
#endregion

#region Main Execution
try {
    # At start of script
    $MinPoshACMEVersion = "4.0.0"
    $poshACME = Get-Module -ListAvailable -Name Posh-ACME | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $poshACME -or $poshACME.Version -lt [Version]$MinPoshACMEVersion) {
        throw "Posh-ACME module version $MinPoshACMEVersion or higher is required"
    }

    # 1. Initial Setup
    Write-Log "Starting certificate setup for $Domain"
    $env:POSHACME_HOME = Join-Path $BaseDirectory "Posh-ACME"

    # 2. Load Config (if specified)
    if ($LoadConfig) {
        $config = Import-CertificateConfig -ConfigPath $LoadConfig
        if (-not $config.DNSProvider) {
            throw "Config file must specify DNSProvider for DNS validation"
        }
        $Email = $config.Email
        $Domain = $config.Domain
        $DNS = $true
        Write-Log "Loaded configuration for $Domain using $($config.DNSProvider)"
    }

    # 3. Initialize ACME Environment
    try {
        if ($Stage) {
            Write-Log "Using Let's Encrypt Staging environment"
            Set-PAServer LE_STAGE -ErrorAction Stop
        } else {
            Write-Log "Using Let's Encrypt Production environment"
            Set-PAServer LE_PROD -ErrorAction Stop
        }
    } catch {
        Write-Log "Failed to set ACME server: $_" -Level Error
        throw
    }

    # 4. Check/Create ACME Account
    $account = Get-PAAccount -ErrorAction SilentlyContinue
    if (-not $account) {
        Write-Log "Creating new ACME account for $Email..."
        $account = New-PAAccount -Contact $Email -AcceptTOS
        Write-Log "ACME account created successfully: $($account.ID)"
    } else {
        Write-Log "Using existing ACME account: $($account.ID)"
        Set-PAAccount -ID $account.ID
    }

    # 5. Check for existing certificate
    $existingCert = Get-PACertificate $Domain -ErrorAction SilentlyContinue
    if ($existingCert) {
        Write-Log "Certificate already exists for $Domain. Use Renew-RDGatewayCertificate.ps1 for renewals." -Level Warning
        return $existingCert
    }

    # 6. Set up validation parameters
    Write-Log "Setting up certificate validation..."
    $certParams = if ($DNS) {
        if ($LoadConfig) {
            # Use config file DNS settings
            Get-DnsPluginParameters -Domain $Domain -Email $Email -DNSProvider $config.DNSProvider -DNSCredentials $config.DNSCredentials
        } else {
            # Interactive DNS setup
            Get-DnsPluginParameters -Domain $Domain -Email $Email
        }
    } else {
        Write-Log "Using HTTP validation"
        @{ Domain = $Domain; Contact = $Email }
    }

    $DefaultCertPassword = 'poshacme' # Consider making this a parameter

    $certParams += @{
        AcceptTOS = $true
        Install = $true
        PfxPass = $DefaultCertPassword
    }

    # 7. Request certificate
    Write-Log "Requesting new certificate for $Domain..."
    $cert = New-PACertificate @certParams
    if (-not $cert -or -not (Test-Path $cert.PfxFile)) {
        throw "Failed to obtain certificate"
    }

    Write-Log "Certificate obtained successfully"
    Write-Log "Certificate Details:"
    Write-Log "  Thumbprint: $($cert.Thumbprint)"
    Write-Log "  Subject: $($cert.Subject)"
    Write-Log "  Valid From: $($cert.NotBefore.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Log "  Valid To: $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss'))"

    # 8. Install certificate
    Write-Log "Installing certificate for RD Gateway roles..."
    Install-RDGatewayCertificate -CertPath $cert.PfxFile -Thumbprint $cert.Thumbprint -CertPassword $certParams.PfxPass

    # After certificate installation
    if ($cert.ChainPemFile -and (Test-Path $cert.ChainPemFile)) {
        Remove-Item $cert.ChainPemFile -Force
        Write-Log "Cleaned up chain file"
    }

    Write-Log "Certificate setup completed successfully"

    # Download renewal script and create scheduled task
    Write-Log "Setting up automatic renewal..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $RenewScript = Join-Path $BaseDirectory "Renew-RDGatewayCertificate.ps1"

    try {
        Write-Log "Downloading renewal script to $RenewScript"
        (New-Object System.Net.WebClient).DownloadFile(
            "https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/LetsEncrypt/Renew-RDGatewayCertificate.ps1",
            $RenewScript
        )

        Write-Log "Creating scheduled task for automatic renewal"
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$RenewScript`" -Domain `"$Domain`""

        $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 3AM

        Register-ScheduledTask -TaskName "RD Gateway Certificate Renewal" `
            -Action $action `
            -Trigger $trigger `
            -User "SYSTEM" `
            -RunLevel Highest `
            -Description "Automatically renews Let's Encrypt certificate for RD Gateway" `
            -ErrorAction Stop

        Write-Log "Automatic renewal setup completed successfully"
    }
    catch {
        Write-Log "Warning: Failed to setup automatic renewal: $_" -Level Warning
        Write-Log "Please setup the renewal task manually" -Level Warning
    }

    return $cert
} catch {
    Write-Log "Certificate setup failed: $_" -Level Error
    throw
}
#endregion

