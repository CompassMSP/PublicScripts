<#

This script adds an IP printer to windows. Run as administrator/SYSTEM for best results.

.PARAMETER PrinterName
Name of the printer

.PARAMETER IPAddress
IP Address/hostname of printer

.PARAMETER DriverDownloadLink
ZIP file that contains the driver

.PARAMETER SkipPing
Does not check for connectivity to printer before running the script

.PARAMETER DriverInfFolder
Location where the .inf file is located

TO FIND: Extract the zip and find the location (usually under an x64 folder)

.PARAMETER DriverName
Name of the driver that will be used

TO FIND: Open the .inf file and find the exact driver name

.PARAMETER DriverInfPath

Location in windows where the .inf is located

TO FIND:

Run the pnputil.exe command
Go to "C:\Windows\System32\DriverStore\FileRepository" and sort by newest folder.
The .inf will usually be in there

.LINK
https://www.pdq.com/blog/using-powershell-to-install-printers/

Andy Morales
#>
#Requires -RunAsAdministrator

$LogPath = 'C:\Windows\Temp\PrinterInstall.log'
$PrinterName = 'TASKalfa 6003i'
$IPAddress = '10.8.11.239'
$DriverDownloadLink = 'https://cdn.kyostatics.net/dlc/eu/driver/all/kx702415_upd_signed.-downloadcenteritem-Single-File.downloadcenteritem.tmp/KX_Universal_Pr...nter_Driver.zip'
$DriverInfFolder = 'C:\Windows\Temp\PrintDriver\Kx_8.1.1109_UPD_Signed_EU\en\64bit\*.inf'
$DriverName = 'Kyocera TASKalfa 6003i KX'
$DriverInfPath = 'C:\Windows\System32\DriverStore\FileRepository\oemsetup.inf_amd64_6bff917e8a9060a5\OEMSETUP.INF'
$SkipPing = $false
#might only work for current user
$SetDefault = $true

function Expand-ZIP {
    <#
    Extracts a ZIP file to a directory. The contents of the destination will be deleted if they already exist.

    Andy Morales
    #>
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

#region preFlight
if ($SkipPing) {
    Write-Log -Path $LogPath -Message "SkipPing is enabled. Script will not check for connectivity to printer"
}
else {
    if ((Test-NetConnection $IPAddress).PingSucceeded) {
        Write-Log -Path $LogPath -Message "Connection to printer succeeded"
    }
    else {
        Write-Log -Path $LogPath -Level Error -Message "Unable to reach printer. Script will exit"
        EXIT
    }
}

$CurrentPrinters = @(Get-Printer | Where-Object { $_.Name -eq $PrinterName })

if ($CurrentPrinters.count -ge 1) {
    Write-Log -Path $LogPath -Level Error -Message "A printer with the name $($PrinterName) already exists. Script will exit."
    EXIT
}
#endregion preFlight

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#region driverInstall
Write-Log -Path $LogPath -Message "Downloading driver from $($DriverDownloadLink)"

(New-Object System.Net.WebClient).DownloadFile($DriverDownloadLink , 'C:\Windows\Temp\PrintDriver.zip')

Expand-ZIP -ZipFile 'C:\Windows\Temp\PrintDriver.zip' -OutPath 'C:\Windows\Temp\PrintDriver'

Write-Log -Path $LogPath -Message "Installing driver from $($DriverInfFolder)"

C:\Windows\System32\pnputil.exe /a $DriverInfFolder

Write-Log -Path $LogPath -Message "Adding driver"
Add-PrinterDriver -Name $DriverName -InfPath $DriverInfPath
#endRegion driverInstall

#region AddPort
$PortName = $IPAddress

$CurrentPorts = @(Get-PrinterPort | Where-Object { $_.Name -eq $IPAddress })

if ($CurrentPorts.count -ge 1) {
    if ($CurrentPorts.PrinterHostAddress -eq $IPAddress -or $CurrentPorts.PrinterHostIP -eq $IPAddress) {
        Write-Log -Path $LogPath -Message "Port $($PortName) already exists with the correct IP"
    }
    else {
        $PortName = $PortName + '-' + (Get-Random -Maximum 99)

        Write-Log -Path $LogPath -Message "Port name already exists with a different IP. Adding as $($PortName)"
        try {
            Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -ErrorAction Stop
        }
        catch {
            Write-Log -Path $LogPath -Level Error -Message "Ran into an error adding the printer port. Script will exit"
            EXIT
        }
    }
}
else {
    Write-Log -Path $LogPath -Message "Adding Port $($PortName)"
    try {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $IPAddress -ErrorAction Stop
    }
    catch {
        Write-Log -Path $LogPath -Level Error -Message "Ran into an error adding the printer port. Script will exit"
        EXIT
    }
}
#endregion AddPort

#region AddPrinter
try {
    Add-Printer -DriverName $DriverName -Name $PrinterName -PortName $PortName -ErrorAction Stop
    Write-Log -Path $LogPath -Message "Successfully installed printer $($PrinterName)"
}
catch {
    Write-Log -Path $LogPath -Level Error -Message "Ran into an error adding the printer $($PrinterName). Script will exit"
    EXIT
}

if ($SetDefault) {
    Write-Log -Path $LogPath -Message "Setting Printer as the default."
    (New-Object -ComObject WScript.Network).SetDefaultPrinter($PrinterName)
}
#endregion AddPrinter