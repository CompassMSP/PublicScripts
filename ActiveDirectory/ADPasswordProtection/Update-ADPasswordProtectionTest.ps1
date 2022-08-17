<#
.DESCRIPTION
This script updates the Lithnet AD Password Protection database with latest HIBP password list. Must be run on each DC.

.EXAMPLE
Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Update-ADPasswordProtection.ps1'); Update-ADPasswordProtection -NotificationEmail 'cwilliams@compassmsp.com' -SMTPRelay 'compassmsp-com.mail.protection.outlook.com' -FromEmail 'cwilliams@compassmsp.com'
.LINK
https://haveibeenpwned.com/Passwords
https://github.com/lithnet/ad-password-protection
https://github.com/CompassMSP/PublicScripts/blob/master/ActiveDirectory/Install-ADPasswordProtection.ps1

Chris Williams
#>
#Requires -Version 5 -RunAsAdministrator

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true,
        HelpMessage = 'example.mail.protection.outlook.com')]
    [string]$SMTPRelay,

    [Parameter(Mandatory = $true)]
    [string]$NotificationEmail,

    [Parameter(Mandatory = $true)]
    [string]$FromEmail
)

$StoreFilesInDBFormatLink = 'https://rmm.compassmsp.com/softwarepackages/ADPasswordProtectionStore.zip'
$StoreFilesInDBFormatFile = 'C:\Temp\ADPasswordAuditStore.zip'
$LogDirectory = 'C:\Windows\Temp\PasswordProtection.log'
$PassProtectionPath = 'C:\Program Files\Lithnet\Active Directory Password Protection'

$Errors = @()

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

function Expand-ZIP {
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]
        [String]$ZipFile,

        [parameter(Mandatory = $true)]
        [String]$OutPath
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    if (Test-Path -Path $OutPath) {
        Remove-Item $OutPath -Recurse -Force
    }

    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $OutPath)
}

#Check if computer is a DC
if ((Get-WmiObject Win32_ComputerSystem).domainRole -lt 4) {
    Write-Log -Level Warn -Path $LogDirectory -Message 'Computer is not a DC. Script will exit'
    Start-Process $LogDirectory
    exit
}

#Check if DC has enough free space
if ((Get-PSDrive C).free -lt 20GB) {
    Write-Log -Level Warn -Path $LogDirectory -Message 'DC has less than 20 GB free. Script will exit'
    Start-Process $LogDirectory
    exit 
}

#Downloads latest version of the HIBP Database
$LatestVersionUrl = (Invoke-WebRequest https://haveibeenpwned.com/Passwords -MaximumRedirection 0).Links | Where-Object {$_.href -like "*pwned-passwords-ntlm-ordered-by-hash-v*.7z"} | Select-Object -expand href
$compassLatestVersion = (Invoke-WebRequest https://rmm.compassmsp.com/softwarepackages/hibp-latest.txt).Content 

#Variables built out for script 
$LatestVersionZip = $($LatestVersionUrl -replace '[a-zA-Z]+://[a-zA-Z]+\.[a-zA-Z]+\.[a-zA-Z]+/[a-zA-Z]+/')
$LatestVersionLog = $($LatestVersionZip -replace 'pwned-passwords-ntlm-ordered-by-hash-') 
$LatestVersionLog = $($LatestVersionLog -replace '.7z')

if ($compassLatestVersion -ne $LatestVersionLog ) {
    Write-Log -Level Warn -Path $LogDirectory -Message 'The Compass database is out of date. Please open a ticket with internal support. Script will now exit.'
    Start-Process $LogDirectory
    exit
}
#Checks for older database version
$OldVersionLog = Get-ChildItem -Path $PassProtectionPath | Where-Object {$_.Name -like 'v*'}

#Checks if latest version is installed 
if ((Get-ChildItem -Path $PassProtectionPath).Name -notcontains $LatestVersionLog) {
    Write-Log -Level Info -Path $LogDirectory -Message 'DC is missing latest HIBP hashes.'
    
    Write-Log -Level Info -Path $LogDirectory -Message 'Downloading HIBP hashes.'

    Start-BitsTransfer -Source $StoreFilesInDBFormatLink -Destination $StoreFilesInDBFormatFile

    Write-Log -Level Info -Path $LogDirectory -Message 'Extracting HIBP hashes'
    try {
        Expand-ZIP -ZipFile $StoreFilesInDBFormatFile -OutPath 'C:\Program Files\Lithnet\Active Directory Password Protection' -ErrorAction Stop

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
        Write-Log -Level Warn -Path $LogDirectory -Message "Ran into an issue extracting the file $StoreFilesInDBFormatFile"
        $Errors += "Ran into an issue extracting the file $StoreFilesInDBFormatFile"
    }
    Remove-Item $StoreFilesInDBFormatFile -Force -ErrorAction SilentlyContinue
} else {
    Write-Log -Level Info -Path $LogDirectory -Message 'DC already has latest HIBP hashes. Script will exit'
    exit
}
