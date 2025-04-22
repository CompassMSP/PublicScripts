
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

    if ($adapters.ifIndex.Count -eq 1) {
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

            $domainJoinParams = @{
                DomainName = $DomainName
                NewName    = $ComputerName
                Credential = $Credential
            }

            if ($OUPath) {
                $domainJoinParams['OUPath'] = $OUPath
            }

            $domainJoinParams['Restart'] = $true
            $domainJoinParams['Force'] = $true

            Add-Computer @domainJoinParams
        } catch {
            Write-Error "Domain join failed: $_"
        }
    } else {
        Rename-Computer -NewName $ComputerName -Force -Restart
    }
}

$Params = @{
    ComputerName   = "vmearth-vprx2"
    IPAddress      = "10.21.19.139"
    Gateway        = "10.21.19.1"
    DnsServer      = "10.21.19.26"
    TimeZone       = "Eastern Standard Time"
    JoinDomain     = $true
    DomainName     = "vmearth.com"
    DomainUser     = "vmearth\chris.admin"
    DomainPassword = "OrangeMonkey21!"
    OUPath         = "OU=Veeam Servers,OU=Computers,OU=CompassMSP,DC=vmearth,DC=com"
}

Initialize-NewServer @Params

function Set-WindowsLicense {

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

Install-CompassAutomate -Name VMEarth