Function Update-ADPasswordProtection {
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
    
    $LogDirectory = 'C:\Windows\Temp\PasswordProtection.log'
    
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
                } catch {
                    if ($WriteLogRetrycount -gt 5) {
                        $StopWriteLogloop = $true
                    } else {
                        Start-Sleep -Milliseconds 500
                        $WriteLogRetrycount++
                    }
                }
            }While ($StopWriteLogloop -eq $false)
        }
        End {
        }
    }
    
    #Check if computer is a DC
    if ((Get-WmiObject Win32_ComputerSystem).domainRole -lt 4) {
        Write-Log -Level Warn -Path $LogDirectory -Message 'Computer is not a DC. Script will exit'
        Start-Process $LogDirectory
        exit
    }
    
   <#
    #Check if DC has enough free space
    if ((Get-PSDrive C).free -lt 30GB) {
        Write-Log -Level Warn -Path $LogDirectory -Message 'DC has less than 30 GB free. Script will exit'
        Start-Process $LogDirectory
        exit 
    }
    #>
    
    if ((Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq "Lithnet Password Protection for Active Directory" }).Version -lt '1.1.53.0') {
        $URI = "https://github.com/lithnet/ad-password-protection/releases/latest"

        $latestRelease = Invoke-WebRequest $URI -Headers @{"Accept" = "application/json" }
        $json = $latestRelease.Content | ConvertFrom-Json
        $latestVersion = $json.tag_name

        $BuildExe = $latestVersion.Replace('v', 'LithnetPasswordProtection-') + '.exe'
        $BuildURI = "https://github.com/lithnet/ad-password-protection/releases/download/$latestVersion/" + $BuildExe

        (New-Object System.Net.WebClient).DownloadFile("$BuildURI", "c:\temp\$BuildExe")


        Write-Log -Level Info -Path $LogDirectory -Message 'Installing Password Protection MSI'
        Start-Process -FilePath C:\temp\$BuildExe -ArgumentList "/exenoui" -Wait;

        Sync-HashesFromHibp

        Write-Log -Level Info -Path $LogDirectory -Message "The Password Protection application has been installed. Restart the computer for the change to take effect."
    } else { Sync-HashesFromHibp }
}
