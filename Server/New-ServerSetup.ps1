<#Author       : Chris Williams
# Creation Date: 3-10-2025
# Usage        : New Server Setup

#********************************************************************************
# Date                         Version      Changes
#------------------------------------------------------------------------
# 3-10-2025                     1.0        Initial Version
#
#*********************************************************************************
#
#>

# Install VMware Tools
Set-Location -Path "d:\"
.\setup.exe /S /v'/qn REBOOT="ReallySuppress"'

function Initialize-NewServer {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$IPAddress,

        [Parameter(Mandatory = $true)]
        [string]$Gateway,

        [Parameter(Mandatory = $true)]
        [string]$DnsServer,

        [Parameter(Mandatory = $false)]
        [string]$DenyRDP = $true,

        [Parameter(Mandatory = $true)]
        [string]$TimeZone,

        [Parameter(Mandatory = $true)]
        [string]$NewName,

        [Parameter(Mandatory = $false)]
        [bool]$JoinDomain = $false,

        [Parameter(Mandatory = $false)]
        [string]$DomainName,

        [Parameter(Mandatory = $false)]
        [string]$DomainUser,

        [Parameter(Mandatory = $false)]
        [string]$DomainPassword,

        [Parameter(Mandatory = $false)]
        [string]$OUPath
    )

    # Get all operational Ethernet adapters
    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }

    if ($adapters.Count -eq 1) {
        $interfaceIndex = $adapters[0].InterfaceIndex
    } else {
        # Look for adapter with an APIPA address
        $apipaAdapter = Get-NetIPAddress -IPAddress "169.254.*" -ErrorAction SilentlyContinue | Select-Object -First 1

        if (-not $apipaAdapter) {
            Write-Warning "Multiple adapters found but none have an APIPA address."
            return
        }

        $interfaceIndex = $apipaAdapter.InterfaceIndex
    }


    # Configure IP settings
    New-NetIPAddress -InterfaceIndex $interfaceIndex -IPAddress $IPAddress -AddressFamily IPv4 -PrefixLength 24 -DefaultGateway $Gateway -ErrorAction SilentlyContinue
    Set-DnsClientServerAddress -InterfaceIndex $interfaceIndex -ServerAddresses $DnsServer -ErrorAction SilentlyContinue

    # Disable firewall
    Get-NetFirewallProfile | Set-NetFirewallProfile -Enabled False

    # Disable UAC
    New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\policies\system" -Name EnableLUA -PropertyType DWord -Value 0 -Force

    if ($DenyRDP -eq $true) {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\" -Name fDenyTSConnections -Value 1
    } else {
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\" -Name fDenyTSConnections -Value 0
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\" -Name UserAuthentication -Value 1
    }

    # Disabled Enhanced Security
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}" -Name IsInstalled -Value 0
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}" -Name IsInstalled -Value 0
    Rundll32 iesetup.dll, IEHardenUser
    Rundll32 iesetup.dll, IEHardenLMSettings
    Rundll32 iesetup.dll, IEHardenAdmin

    # Set timezone
    Set-TimeZone -Name $TimeZone

    # Add log off button to public desktop
    $SourceFilePath = "C:\Windows\System32\logoff.exe"
    $ShortcutPath = "C:\Users\Public\Desktop\Logoff.lnk"
    $WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $SourceFilePath
    $shortcut.IconLocation = "shell32.dll,44"
    $shortcut.Save()

    ## Install Cato Cert
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $CatoCert = "CatoNetworksTrustedRootCA.cer"
    $CatoCertURI = "https://clientdownload.catonetworks.com/public/certificates/CatoNetworksTrustedRootCA.cer"
    (New-Object System.Net.WebClient).DownloadFile("$CatoCertURI", "C:\windows\temp\$CatoCert")
    Import-Certificate -FilePath "C:\windows\temp\$CatoCert" -CertStoreLocation Cert:\LocalMachine\Root


    if ($JoinDomain) {
        try {
            $SecurePassword = ConvertTo-SecureString -String $DomainPassword -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential ($DomainUser, $SecurePassword)

            Add-Computer -DomainName $DomainName -NewName $ComputerName -Credential $Credential -OUPath $OUPath -Restart -Force
        } catch {
            Write-Error "Domain join failed: $_"
        }
    } else {
        Rename-Computer -NewName $ComputerName -Force -Restart
    }
}

$Params = @{
    ComputerName   = "COMPASS-DC2"
    IPAddress      = "10.21.17.36"
    Gateway        = "10.21.17.1"
    DnsServer      = "10.21.17.19"
    TimeZone       = "Eastern Standard Time"
    JoinDomain     = $true
    DomainName     = "domain.local"
    DomainUser     = "domain\Administrator"
    DomainPassword = "PasswordHERE"
    OUPath         = "OU=RDS_Servers,OU=Servers,OU=NOG,DC=NorwichOrthopedics,DC=local"
}

Initialize-NetworkAndJoinDomain @Params

function Set-WindowsLicense {

    <#
    https://admin.microsoft.com/#/subscriptions/vlnew/downloadsandkeys

            slmgr.vbs /dlv

            irm https://get.activated.win | iex
    #>

    param(
        [switch]$WhatIf
    )

    # Get Windows version info
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $caption = $os.Caption

    # Define license keys
    $licenses = @{
        'Microsoft Windows 11 Pro'       = '3N2XV-VD43Y-GJ4D8-6TGKX-6CQGY'
        'Windows Server 2019 Standard'   = '2CPJV-9NTMM-CV68D-WPJCR-8464D'
        'Windows Server 2019 Datacenter' = 'CHT93-QKNG4-6GWM7-V6MCV-8HXRH'
        'Windows Server 2022 Standard'   = 'WX7F2-3NJ7P-7PWWM-GBMPX-PBB3W'
        'Windows Server 2022 Datacenter' = '8XWD3-XN78D-JXTY2-VXQMD-H24Q4'
        'Windows Server 2025 Datacenter' = 'QN2CQ-3FHBJ-W6GH8-WTRFT-8T4G6'
    }

    # Match OS to license
    $key = switch -Wildcard ($caption) {
        "*Windows 11*" { $licenses['Microsoft Windows 11 Pro'] }
        "*Server 2019*Standard*" { $licenses['Windows Server 2019 Standard'] }
        "*Server 2019*Datacenter*" { $licenses['Windows Server 2019 Datacenter'] }
        "*Server 2022*Standard*" { $licenses['Windows Server 2022 Standard'] }
        "*Server 2022*Datacenter*" { $licenses['Windows Server 2022 Datacenter'] }
        "*Server 2025*Datacenter*" { $licenses['Windows Server 2025 Datacenter'] }
        default { throw "Unsupported OS: $caption" }
    }

    Write-Host "Detected OS: $caption"
    Write-Host "License Key: $key"

    if ($WhatIf) {
        Write-Host "WhatIf: Would run command: slmgr.vbs /ipk $key"
    } else {
        Write-Host "Applying license key..."
        $result = Start-Process "slmgr.vbs" -ArgumentList "/ipk $key" -Wait -PassThru
        if ($result.ExitCode -eq 0) {
            Write-Host "License key applied successfully"
        } else {
            throw "Failed to apply license key"
        }
    }
}

## Set Windows License
Set-WindowsLicense


## Install Compass Automate
function Install-CompassAutomate {
    <#
    .SYNOPSIS
        Installs ConnectWise Automate agent with location-specific configuration.

    .DESCRIPTION
        This function installs the ConnectWise Automate agent with specific location settings.
        It supports installation using either a location ID or a predefined location name.
        If no parameters are specified, it defaults to the "NewPc (Catch All)" location.

    .PARAMETER Id
        The location ID for the Automate agent. Defaults to '1' (NewPc Catch All).
        Cannot be used with -Name parameter.

    .PARAMETER Name
        The predefined location name. Valid options are:
        - North Florida
        - North East
        - North East (TT)
        - Mid-Atlantic
        - Equinix
        - CoreSite
        - VMEarth
        - OneCompass
        Cannot be used with -Id parameter.

    .PARAMETER Server
        The Automate server address. Defaults to 'rmm.compassmsp.com'.

    .PARAMETER Token
        The authentication token for installation. Has a default value.

    .PARAMETER WhatIf
        Shows what would happen if the function runs.
        The function is not run.

    .EXAMPLE
        Install-CompassAutomate
        Installs agent with default settings (ID: 1, NewPc Catch All).

    .EXAMPLE
        Install-CompassAutomate -Id "2141"
        Installs agent for North Florida location using ID.

    .EXAMPLE
        Install-CompassAutomate -Name "North Florida"
        Installs agent for North Florida location using name.

    .EXAMPLE
        Install-CompassAutomate -Id "2141" -Server "custom.server.com"
        Installs agent for North Florida with custom server.

    .NOTES
        Requires internet connectivity to download and install the agent.
        The function will throw an error if the installation fails.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Id', SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Token = 'e980c8bb7f40ec60e85020433a5f0f5f',

        [Parameter(Mandatory = $false, ParameterSetName = 'Id')]
        [string]$Id = '1',

        [Parameter(Mandatory = $false, ParameterSetName = 'Name')]
        [ValidateSet("North Florida", "North East", "North East (TT)",
            "Mid-Atlantic", "Equinix", "CoreSite", "VMEarth", "OneCompass")]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Server = 'rmm.mycompass.cloud'
    )

    # Define named locations
    $locations = @(
        [pscustomobject]@{Name = "North Florida"; Id = "2141" },
        [pscustomobject]@{Name = "North East"; Id = "2333" },
        [pscustomobject]@{Name = "North East (TT)"; Id = "2323" },
        [pscustomobject]@{Name = "Mid-Atlantic"; Id = "3" },
        [pscustomobject]@{Name = "Equinix"; Id = "518" },
        [pscustomobject]@{Name = "CoreSite"; Id = "1737" },
        [pscustomobject]@{Name = "VMEarth"; Id = "1723" },
        [pscustomobject]@{Name = "OneCompass"; Id = "265" }
    )

    # Get location ID based on parameter set
    $LocationId = if ($PSCmdlet.ParameterSetName -eq 'Name') {
            ($locations | Where-Object Name -EQ $Name).Id
    } else {
        $Id
    }

    # Get location name for display
    $locationName = if ($LocationId -eq '1') {
        "NewPc (Catch All)"
    } else {
            ($locations | Where-Object Id -EQ $LocationId).Name
    }

    # Set TLS 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Install Automate
    if ($PSCmdlet.ShouldProcess(
            "Installing Automate agent for $locationName (ID: $LocationId) on server $Server",
            "Install Automate agent?",
            "Install-CompassAutomate"
        )) {
        Write-Host "Installing Automate for location: $locationName (ID: $LocationId)"
        Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/CWAutomate/CMSP_Automate-Module.psm1'); `
            Install-Automate -Server $Server -Transcript -Token $Token -LocationID $LocationId
    }
}

<#
"North Florida"
"North East"
"North East (TT)"
"Mid-Atlantic"
"Equinix"
"CoreSite"
"VMEarth"
"OneCompass"
#>

Install-CompassAutomate -Name VMEarth

## Install Duo
function Install-DuoClient {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Compass', 'BUI', 'VMEarth', 'Johnstone')]
        [string]$ClientName
    )

    # Function to build argument list
    function Get-DuoArguments {
        param([string]$ClientName)

        # Common Duo parameters
        $commonParams = @(
            'AUTOPUSH="#1"'
            'FAILOPEN="#1"'
            'OfflineAvailable="#0"'
            'SMARTCARD="#1"'
            'RDPONLY="#0"'
            '/qn'
        )

        # Client-specific configurations
        $clientConfigs = @{
            Compass   = @{
                IKEY = "DITDY5ZR5YAF1WUO440U"
                SKEY = "nEh7xIqYtdFZtRSahNRawMAB7Vd1U3e6h83ACDwV"
                HOST = "api-3e208a92.duosecurity.com"
            }
            BUI       = @{
                IKEY = "DI4ZRKW46M4EOY666OZW"
                SKEY = "gDaGTvPFQXyt0rgLhNuGGQ6lTtRlvFQNjAh58064"
                HOST = "api-3e208a92.duosecurity.com"
            }
            VMEarth   = @{
                IKEY = "DIS58U9OSCBIPEKOR93Q"
                SKEY = "NiqX6CmBkNyio7GMXhE793KipIhEJqjegazHgvHT"
                HOST = "api-3e208a92.duosecurity.com"
            }
            Johnstone = @{
                IKEY = "DIACQQXZUQQRV6VSAGA8"
                SKEY = "40BqPiWKhSomoFHKYtK11vf4lWAqAY5zhQ3QxfgC"
                HOST = "api-518a4989.duosecurity.com"
            }
        }

        $config = $clientConfigs[$ClientName]
        $clientArgs = @(
            "IKEY=`"$($config.IKEY)`""
            "SKEY=`"$($config.SKEY)`""
            "HOST=`"$($config.HOST)`""
        )

        return $clientArgs + $commonParams
    }

    try {
        Write-Host "Installing Duo for $ClientName..."
        $tempPath = "C:\Windows\temp"

        # Download and extract Duo
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $duoZip = Join-Path $tempPath "DuoWinLogon_MSIs-latest.zip"
        $duoPath = Join-Path $tempPath "DuoWinLogon_MSIs-latest"

        (New-Object System.Net.WebClient).DownloadFile(
            "https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip",
            $duoZip
        )

        Expand-Archive -LiteralPath $duoZip -DestinationPath $duoPath -Force

        # Get client-specific arguments
        $argumentList = Get-DuoArguments -ClientName $ClientName

        # Install
        Start-Process -FilePath (Join-Path $duoPath "DuoWindowsLogon64.msi") -ArgumentList $argumentList -Wait

        # Duo Gateway Properties
        try {
            if (Get-Command Get-RDServer -ErrorAction SilentlyContinue) {
                $gateway = Get-RDServer -Role RDS-GATEWAY -ErrorAction SilentlyContinue
                if ($gateway) {
                    New-ItemProperty -Path "HKLM:\SOFTWARE\Duo Security\DuoTsg" -Name AuthorizedSession_MaxDuration -PropertyType DWord -Value 600 -Force | Out-Null
                    New-ItemProperty -Path "HKLM:\SOFTWARE\Duo Security\DuoTsg" -Name AuthorizedSession_IdleTimeout -PropertyType DWord -Value 180 -Force | Out-Null
                }
            }
        } catch {
            # Do nothing
        }


        Write-Host "Duo installation completed for $ClientName"
    } catch {
        Write-Error ("Failed to install Duo for " + $ClientName + ": " + $_.Exception.Message)
    }
}

# Install Duo separately
Install-DuoClient -ClientName "VMEarth"  # or "BUI", "VMEarth", "Johnstone"

winget upgrade --all --accept-source-agreements --accept-package-agreements

$apps = @(
    #"Microsoft.Office",
    #"Microsoft.Teams",
    #"ConnectWise.ConnectWiseManageClient64-bit",
    "Microsoft.PowerShell",
    "Microsoft.WindowsTerminal",
    "Microsoft.WindowsApp",
    "Google.Chrome",
    "Mozilla.Firefox",
    "Notepad++.Notepad++",
    "7zip.7zip",
    "PuTTY.PuTTY",
    "Foxit.FoxitReader"
)

Foreach ($app in $apps) {
    winget install --id $app -e --source winget --accept-source-agreements --accept-package-agreements
}


DISM /Online /NoRestart /Enable-Feature /all /FeatureName:NetFx3 /Source:D:\sources\sxs

function Install-Application {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('SeaMonkey', 'WinSCP', 'BoxTools', 'TreeSize', 'VCRedist', 'All')]
        [string]$Name,
        [string]$Uri,
        [string]$Filename,
        [string]$Arguments,
        [switch]$Wait
    )

    # Application configurations
    $apps = @{
        SeaMonkey = @{
            Uri       = "https://portableapps.com/redir2/?a=SeaMonkeyPortable&s=s&d=pa&f=SeaMonkeyPortable_2.53.16_English.paf.exe"
            Filename  = "SeaMonkeyPortable_2.53.16_English.paf.exe"
            Arguments = "/q"
            Wait      = $true
        }
        WinSCP    = @{
            Uri       = "https://cdn.winscp.net/files/WinSCP-6.1.1-Setup.exe"
            Filename  = "WinSCP-6.1.1-Setup.exe"
            Arguments = "/VERYSILENT /ALLUSERS"
        }
        BoxTools  = @{
            Uri       = "https://e3.boxcdn.net/box-installers/boxedit/win/currentrelease/BoxToolsInstaller-AdminInstall.msi"
            Filename  = "BoxToolsInstaller-AdminInstall.msi"
            Arguments = "/quiet"
        }
        TreeSize  = @{
            Uri       = "https://downloads.jam-software.de/treesize_free/TreeSizeFreeSetup.exe"
            Filename  = "TreeSizeFreeSetup.exe"
            Arguments = "/quiet"
        }
        VCRedist  = @{
            Uri       = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
            Filename  = "vc_redist.x64.exe"
            Arguments = "/quiet"
        }
    }

    if ($Name -eq 'All') {
        foreach ($app in $apps.Keys) {
            Install-Application -Name $app @($apps[$app])
        }
        return
    }

    try {
        Write-Host "Installing $Name..."
        $tempPath = "C:\Windows\temp\$Filename"

        # Download file
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object System.Net.WebClient).DownloadFile($Uri, $tempPath)

        # Install
        $processParams = @{
            FilePath     = $tempPath
            ArgumentList = $Arguments
        }
        if ($Wait) { $processParams['Wait'] = $true }

        Start-Process @processParams
        Write-Host "$Name installation started"
    } catch {
        Write-Error ("Failed to install " + $Name + ": " + $_.Exception.Message)
    }
}

# Install single app
Install-Application -Name WinSCP

# Install all apps
Install-Application -Name All

## Install IISCrypto
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$IISCryptoURI = "https://www.nartac.com/Downloads/IISCrypto/IISCryptoCli.exe"
$IISCrypto = "IISCryptoCli.exe"
(New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/Security/Compass_IISCryptTemplate.ictpl", "C:\windows\temp\Compass_IISCryptTemplate.ictpl")
(New-Object System.Net.WebClient).DownloadFile("$IISCryptoURI", "C:\windows\temp\$IISCrypto")
Start-Process -FilePath "C:\windows\temp\IISCryptoCli.exe" -Wait -ArgumentList "/template C:\windows\temp\Compass_IISCryptTemplate.ictpl /reboot"

## Install ScreenConnect ad-hoc
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ScreenConnectClientURI = "https://remote.mycompass.cloud/Bin/ScreenConnect.ClientSetup.msi?e=Access&y=Guest&c=CompassMSP&c=Cloud%20-%20Azure&c=&c=&c=&c=&c=&c="
$ScreenConnectClientSetup = "ScreenConnect.ClientSetup.msi"
(New-Object System.Net.WebClient).DownloadFile("$ScreenConnectClientURI", "C:\windows\temp\$ScreenConnectClientSetup")
Start-Process -FilePath "C:\windows\temp\ScreenConnect.ClientSetup.msi" -Wait -ArgumentList "/qn"

netdom ComputerName ELC-File /ADD file
IPConfig /RegisterDNS

netsh advfirewall reset
