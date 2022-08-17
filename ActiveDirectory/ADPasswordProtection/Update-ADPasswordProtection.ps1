<#
.DESCRIPTION
This script updates the Lithnet AD Password Protection database with latest HIBP password list. Must be run on each DC.

.EXAMPLE
Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Update-ADPasswordProtection.ps1' | Invoke-Expression

.LINK
https://haveibeenpwned.com/Passwords
https://github.com/lithnet/ad-password-protection
https://github.com/CompassMSP/PublicScripts/blob/master/ActiveDirectory/Install-ADPasswordProtection.ps1

Chris Williams
#>
#Requires -RunAsAdministrator

function Write-Log {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory = $true)]
        [Alias('LogPath')]
        [string]$Path,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info")]
        [string]$Level = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoClobber,

        [Parameter(Mandatory = $false)]
        [switch]$DailyMode
    )

    Begin {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'Continue'
        if ($DailyMode) {
            $Path = $Path.Replace('.', "-$(Get-Date -UFormat "%Y%m%d").")
        }
    }
    Process {
        # If the file already exists and NoClobber was specified, do not write to the log.
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
        }

        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            New-Item $Path -Force -ItemType File
        }

        else {
            # Nothing to see here yet.
        }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR:'
            }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING:'
            }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO:'
            }
        }

        # Write log entry to $Path
        #try to write to the log file. Rety if it is locked
        $StopWriteLogloop = $false
        [int]$WriteLogRetrycount = "0"
        do {
            try {
                "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append -ErrorAction Stop
                $StopWriteLogloop = $true
            }
            catch {
                if ($WriteLogRetrycount -gt 5) {
                    $StopWriteLogloop = $true
                }
                else {
                    Start-Sleep -Milliseconds 500
                    $WriteLogRetrycount++
                }
            }
        }While ($StopWriteLogloop -eq $false)
    }
    End {
    }
}

$LogDirectory = 'C:\Windows\Temp\UpdatePasswordProtection.log'

#Check if computer is a DC
if ((Get-WmiObject Win32_ComputerSystem).domainRole -lt 4) {
    Write-Log -Level Error -Path $LogDirectory -Message 'Computer is not a DC. Script will exit'
    Start-Process $LogDirectory
    exit
}

#Check if DC has enough free space
if ((Get-PSDrive C).free -lt 20GB) {
    Write-Log -Level Error -Path $LogDirectory -Message 'DC has less than 20 GB free. Script will exit'
    Start-Process $LogDirectory
    exit 
}

#Sets TLS 1.2 in POSH
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Checks if 7Zip POSH Module is installed
$7Zip4PoshInstalled = Get-InstalledModule 7Zip4PowerShell

#Install 7Zip POSH Module
if ($7Zip4PoshInstalled -eq $NULL) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue
    Set-PSRepository -Name 'PSGallery' -SourceLocation "https://www.powershellgallery.com/api/v2" -InstallationPolicy Trusted
    Install-Module -Name 7Zip4PowerShell -Force
} else { 
    Write-Log -Level Info -Path $LogDirectory -Message '7Zip4Powershell Modules is already install.'
}

#Downloads latest version of the HIBP Database
$LatestVersionUrl = (Invoke-WebRequest https://haveibeenpwned.com/Passwords -MaximumRedirection 0).Links | Where-Object {$_.href -like "*pwned-passwords-ntlm-ordered-by-hash-v*.7z"} | Select-Object -expand href

#Variables built out for script 
$LatestVersionZip = $($LatestVersionUrl -replace '[a-zA-Z]+://[a-zA-Z]+\.[a-zA-Z]+\.[a-zA-Z]+/[a-zA-Z]+/')
$LatestVersionTXT = $($LatestVersionZip -replace '.7z') + '.txt'
$LatestVersionLog = $($LatestVersionZip -replace 'pwned-passwords-ntlm-ordered-by-hash-') 
$LatestVersionLog = $($LatestVersionLog -replace '.7z')

$PassProtectionPath = 'C:\Program Files\Lithnet\Active Directory Password Protection\'

#Checks for older database version
$OldVersionLog = Get-ChildItem -Path $PassProtectionPath | Where-Object {$_.Name -like 'v*'}

#Checks if latest version is installed 
if ((Get-ChildItem -Path $PassProtectionPath).Name -notcontains $LatestVersionLog) {
    $HIBPDBUpdate = $true
    Write-Log -Level Info -Path $LogDirectory -Message "HIBP Database needs updating. Will now download the latest DB."
} else { 
    Write-Log -Level Error -Path $LogDirectory -Message "HIBP Database is on latest version. Script will now exit."
    Start-Process $LogDirectory
    exit
}

if ($HIBPDBUpdate -eq $true) {
    Write-Log -Level Info -Path $LogDirectory -Message 'Downloading HIBP hashes.'

    #(New-Object System.Net.WebClient).DownloadFile("$LatestVersionUrl","C:\temp\$LatestVersionZip")
    Start-BitsTransfer -Source $LatestVersionUrl -Destination "C:\temp\$($LatestVersionZip)"

    #Extract HIBP Hashes
    Write-Log -Level Info -Path $LogDirectory -Message 'Extracting HIBP hashes'

    try {
        Expand-7Zip -ArchiveFileName C:\Temp\$LatestVersionZip -TargetPath 'C:\Temp\'
    }
    catch {
        Write-Log -Level Error -Path $LogDirectory -Message "Ran into an issue extracting the file $LatestVersionZip"
        Start-Process $LogDirectory
        exit
    }

    Remove-Item C:\Temp\$LatestVersionZip -Force -ErrorAction SilentlyContinue
}

if ((Get-ChildItem -Path C:\Temp).Name -contains $LatestVersionTXT) {
    Write-Log -Level Info -Path $LogDirectory -Message 'Loading LithnetPasswordProtection module and Store'
    try {
        Import-Module LithnetPasswordProtection
        Open-Store 'C:\Program Files\Lithnet\Active Directory Password Protection\Store'

        Write-Log -Level Info -Path $LogDirectory -Message 'Import compromised password hashes'

        Import-CompromisedPasswordHashes -Filename C:\Temp\$($LatestVersionTXT)

        Write-Log -Level Info -Path $LogDirectory -Message 'Adding new version file'

        New-Item $($PassProtectionPath + '\' + $LatestVersionLog) -Type File

        $PDC = (Get-ADForest | Select-Object -ExpandProperty RootDomain | Get-ADDomain).PDCEmulator

        $LocalDC = [System.Net.Dns]::GetHostByName($env:computerName).HostName

        if ($PDC -eq $LocalDC) {
            Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Invoke-ADPasswordAudit.ps1'); Invoke-ADPasswordAudit -NotificationEmail $NotificationEmail -SMTPRelay $SMTPRelay -FromEmail $FromEmail
        }

        if ($OldVersionLog -ne $NULL) {
            Write-Log -Level Info -Path $LogDirectory -Message 'Removing old version file'
            Remove-Item $OldVersionLog
        }
    }
    catch {
        Write-Log -Level Error -Path $LogDirectory -Message "Ran into an issue importing $LatestVersionZip"
    }
}
