<#
.SYNOPSIS
This script can be used to disable PowerShell 2 on a machine. The script will only run if WMF 5 or higher has been installed. In addition, special services like
Exchange must not be installed.

.LINK
Original Script
https://github.com/robwillisinfo/Disable-PSv2/blob/master/Disable-PSv2.ps1

Andy Morales
#>

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
    Write-Log -Message 'Restarting Server.' -Path c:\Logs\ScriptOutput.log
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
        #try to write to the log file. Retry if it is locked
        $StopWriteLogLoop = $false
        [int]$WriteLogRetryCount = "0"
        do {
            try {
                "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append -ErrorAction Stop
                $StopWriteLogLoop = $true
            }
            catch {
                if ($WriteLogRetryCount -gt 5) {
                    $StopWriteLogLoop = $true
                }
                else {
                    Start-Sleep -Milliseconds 500
                    $WriteLogRetryCount++
                }
            }
        }While ($StopWriteLogLoop -eq $false)
    }
    End {
    }
}

$LogPath = "c:\Windows\Temp\PS2RemovalScript.txt"

$PS2CanBeDisabled = $true

#run through machine checks
Switch ($true) {
    ($PSVersionTable.PSVersion -lt [version]('{0}.{1}.{2}.{3}' -f '5.0.0.0'.split('.'))) {
        Write-log -Path $LogPath -Level Error -Message 'WMF 5 is not installed. Script should exit'
        $PS2CanBeDisabled = $false
        BREAK
    }
    ((Get-PSSnapin -Registered | Select-Object -ExpandProperty name) -match 'Microsoft.Exchange.Management.PowerShell') {
        Write-log -Path $LogPath -Level Error -Message 'Exchange is installed. Do not modify any PS settings.'
        $PS2CanBeDisabled = $false
        BREAK
    }
    (Test-Path "$env:ProgramFiles\Common Files\Microsoft Shared\Web Server Extensions") {
        Write-log -Path $LogPath -Level Error -Message "Sharepoint is installed. Do not modify any PS settings"
        $PS2CanBeDisabled = $false
        BREAK
    }
    (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Real-Time Communications\{A593FD00-64F1-4288-A6F4-E699ED9DCA35}') {
        Write-log -Path $LogPath -Level Error -Message "WMF Should not be upgraded when Lync is installed"
        $PS2CanBeDisabled = $false
        BREAK
    }
    ((Test-Path "$env:ProgramFiles\Microsoft System Center 2012") -or (Test-Path "$env:ProgramFiles\Microsoft System Center 2012 R2") -or (Test-Path "$env:ProgramFiles\Microsoft Configuration Manager")) {
        Write-log -Path $LogPath -Level Error -Message "WMF Should not be upgraded when System Center is installed"
        $PS2CanBeDisabled = $false
        BREAK
    }
    ((Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductName).ProductName -notmatch "(?i)10|2012|2016|2019"){
        Write-log -Path $LogPath -Level Info -Message "Computer is not Srv 2012+ or Windows 10. Script will exit."
        $PS2CanBeDisabled = $false
        BREAK
    }
}

if ($PS2CanBeDisabled){
    #Check current PS2 status
    $PSv2PreCheck = dism.exe /Online /Get-Featureinfo /FeatureName:"MicrosoftWindowsPowerShellv2" | findstr "State"

    If ( $PSv2PreCheck -like "State : Enabled" ) {

        Write-log -Path $LogPath -Level Info -Message "PowerShell v2 appears to be enabled, disabling via dism..."

        #Remove PS2
        dism.exe /Online /Disable-Feature /FeatureName:"MicrosoftWindowsPowerShellv2" /NoRestart

        $PSv2PostCheck = dism.exe /Online /Get-Featureinfo /FeatureName:"MicrosoftWindowsPowerShellv2" | findstr "State"

        If ( $PSv2PostCheck -like "State : Enabled" ) {
            Write-log -Path $LogPath -Level Error -Message "PowerShell v2 still seems to be enabled, try removing manually."
        }
        Else {
            Write-log -Path $LogPath -Level Info -Message "PowerShell v2 disabled successfully."
        }
    }
    Else {
        Write-log -Path $LogPath -Level Info -Message "PowerShell v2 is already disabled, no changes will be made."
    }
}