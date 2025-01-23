#Requires -Modules Posh-ACME, RemoteDesktop

<#
.SYNOPSIS
    Renews Let's Encrypt certificates for RD Gateway using existing configuration.

.DESCRIPTION
    This script handles the renewal of Let's Encrypt certificates for Remote Desktop Gateway servers.
    It uses the existing Posh-ACME configuration and certificate settings.

    Designed to be run as a scheduled task, it will:
    1. Check if certificate needs renewal (within 30 days of expiry)
    2. Renew certificate if needed using existing validation method
    3. Install renewed certificate for all RD Gateway roles
    4. Log all actions to file

.PARAMETER Domain
    The domain name of the certificate to renew.

.PARAMETER BaseDirectory
    Directory containing Posh-ACME configuration and certificates.
    Default: "C:\ProgramData\RDGatewayCerts"

.EXAMPLE
    # Basic renewal check and process
    .\Renew-RDGatewayCertificate.ps1 -Domain "gateway.contoso.com"

.EXAMPLE
    # Create a scheduled task for automatic renewal
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' `
        -Argument '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\Scripts\Renew-RDGatewayCertificate.ps1" -Domain "gateway.contoso.com"'

    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 3AM

    Register-ScheduledTask -TaskName "RD Gateway Certificate Renewal" `
        -Action $action `
        -Trigger $trigger `
        -User "SYSTEM" `
        -RunLevel Highest

.NOTES
    - Script must run with administrative privileges
    - Logs are stored in $BaseDirectory\Logs
    - Uses existing Posh-ACME configuration
    - Will not request new certificates, only renew existing ones
    - Recommended to run weekly via scheduled task

.AUTHORS
    - Chris Williams
    - Last Updated: 01/22/2025
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Domain,
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

#region Certificate Management Functions
function Test-CertificateRenewal {
    param(
        [string]$Domain,
        [int]$RenewalDays = 30
    )

    try {
        Write-Log "Checking certificate status for $Domain..."
        $existingCert = Get-PACertificate $Domain -ErrorAction SilentlyContinue

        if ($existingCert) {
            $expiry = $existingCert.NotAfter
            $daysUntilExpiry = ($expiry - (Get-Date)).Days
            Write-Log "Found existing certificate, expires in $daysUntilExpiry days"

            if ($daysUntilExpiry -le $RenewalDays) {
                Write-Log "Certificate needs renewal"
                return @{
                    NeedsRenewal = $true
                    IsNew = $false
                    Order = Get-PAOrder $Domain
                }
            }
            return @{ NeedsRenewal = $false }
        }
        Write-Log "No existing certificate found"
        return @{
            NeedsRenewal = $true
            IsNew = $true
            Order = $null
        }
    } catch {
        Write-Log "Error checking certificate status: $_" -Level Error
        throw
    }
}
#endregion

#region Main Execution
try {
    $MinPoshACMEVersion = "4.0.0"
    $poshACME = Get-Module -ListAvailable -Name Posh-ACME | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $poshACME -or $poshACME.Version -lt [Version]$MinPoshACMEVersion) {
        throw "Posh-ACME module version $MinPoshACMEVersion or higher is required"
    }

    Write-Log "Starting certificate renewal check for $Domain"
    $env:POSHACME_HOME = Join-Path $BaseDirectory "Posh-ACME"

    # 1. Verify ACME account exists
    $account = Get-PAAccount -ErrorAction SilentlyContinue
    if (-not $account) {
        throw "No ACME account found. Please run New-RDGatewayCertificate.ps1 first to set up the initial certificate."
    }
    Write-Log "Found ACME account"

    # 2. Check certificate status
    $certStatus = Test-CertificateRenewal -Domain $Domain

    # 3. Handle renewal if needed
    if ($certStatus.NeedsRenewal) {
        if ($certStatus.IsNew) {
            throw "No existing certificate found. Please run New-RDGatewayCertificate.ps1 first."
        }

        Write-Log "Renewing certificate..."
        $cert = Submit-Renewal -Order $certStatus.Order -AllowInsecureRedirect

        if (-not $cert -or -not (Test-Path $cert.PfxFile)) {
            throw "Certificate renewal failed"
        }

        if ($cert) {
            Write-Log "Certificate renewed successfully"
            Write-Log "Certificate Details:"
            Write-Log "  Thumbprint: $($cert.Thumbprint)"
            Write-Log "  Subject: $($cert.Subject)"
            Write-Log "  Valid From: $($cert.NotBefore.ToString('yyyy-MM-dd HH:mm:ss'))"
            Write-Log "  Valid To: $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss'))"

            Write-Log "Installing renewed certificate..."
            Install-RDGatewayCertificate -CertPath $cert.PfxFile -Thumbprint $cert.Thumbprint -CertPassword 'poshacme'

            # Cleanup chain file after successful installation
            if ($cert.ChainPemFile -and (Test-Path $cert.ChainPemFile)) {
                Remove-Item $cert.ChainPemFile -Force
                Write-Log "Cleaned up chain file"
            }
        }
    } else {
        Write-Log "Certificate is still valid, no renewal needed"
    }

    # At the end of successful renewal
    if ($certStatus.NeedsRenewal) {
        Write-Log "Certificate renewal and installation completed successfully"
        return $cert  # Return the renewed certificate object
    } else {
        Write-Log "Certificate is still valid, no renewal needed"
        return $null  # Or return the existing cert if you prefer
    }
} catch {
    Write-Log "Renewal process failed: $_" -Level Error
    throw
}
#endregion