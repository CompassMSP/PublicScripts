<#
Place at: c:\BIN\CertRenew\Install-LeExchangeCertificate.ps1

Andy Morales
#>
param($result)

Set-Alias ps64 'C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe'

ps64 -args $result -command {

    function Write-Log {
        <#
        .Synopsis
        Write-Log writes a message to a specified log file with the current time stamp.
        .DESCRIPTION
        The Write-Log function is designed to add logging capability to other scripts.
        In addition to writing output and/or verbose you can write to a log file for
        later debugging.
        .NOTES
        Created by: Jason Wasser @wasserja
        Modified: 11/24/2015 09:30:19 AM

        Changelog:
            * Code simplification and clarification - thanks to @juneb_get_help
            * Added documentation.
            * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks
            * Revised the Force switch to work as it should - thanks to @JeffHicks

        To Do:
            * Add error handling if trying to create a log file in a inaccessible location.
            * Add ability to write $Message to $Verbose or $Error pipelines to eliminate
            duplicates.
        .PARAMETER Message
        Message is the content that you wish to add to the log file.
        .PARAMETER Path
        The path to the log file to which you would like to write. By default the function will
        create the path and file if it does not exist.
        .PARAMETER Level
        Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational)
        .PARAMETER NoClobber
        Use NoClobber if you do not wish to overwrite an existing file.
        .EXAMPLE
        Write-Log -Message 'Log message'
        Writes the message to c:\Logs\PowerShellLog.log.
        .EXAMPLE
        Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log
        Writes the content to the specified log file and creates the path and file specified.
        .EXAMPLE
        Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error
        Writes the message to the specified log file as an error message, and writes the message to the error pipeline.
        .LINK
        https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0
        #>
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

    $LogPath = 'C:\BIN\CertRenew\ExchangeCertLog.txt'

    $ExchangeFQDN = $Env:COMPUTERNAME + '.' + (Get-WmiObject Win32_ComputerSystem).Domain

    Write-Log -Path $LogPath -Level Info -Message "Updating Exchange certificates on $($ExchangeFQDN)"

    $result = $args[0]

    $pfxThumbprintHash = $result.ManagedItem.CertificateThumbprintHash

    # Import the RemoteDesktop module
    try {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
    }
    catch {
        Write-Log -Path $LogPath -Level Error -Message 'Unable to import Exchange PSSnapin. Script will exit'
        Exit
    }

    try {
        Enable-ExchangeCertificate -Thumbprint $pfxThumbprintHash -Services POP, IMAP, SMTP, IIS -Force -ErrorAction Stop
        Write-Log -Path $LogPath -Level Info -Message "Sucessfully applied certificate to POP, IMAP, SMTP, IIS on $($ExchangeFQDN)"
    }
    catch {
        Write-Log -Path $LogPath -Level Error -Message "Unable to apply certificate to POP, IMAP, SMTP, IIS on $($ExchangeFQDN)"
    }
}