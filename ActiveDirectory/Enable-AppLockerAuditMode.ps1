<#
This script will enable AppLocker in audit mode to all computers if there is no current policy in place.

The goal is to enable this GPO for a few weeks and then enable it for servers.

Andy Morales
#>

#Requires -Module ActiveDirectory

$LogLocation = 'C:\Windows\Temp\AppLockerAuditScript.txt'

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

#Check if computer is the PDC
if ((Get-WmiObject Win32_ComputerSystem).domainRole -lt 5) {
    Write-Log -Level Info -Path $LogLocation -Message 'Computer is not the PDC. Script will exit'
    break
}

Import-Module -Name ActiveDirectory

#Get all GPOs
$allGPOs = Get-GPO -All

$AppLockerFound = $false

foreach ($GPO in $AllGpos) {
    [xml]$GPOReport = $GPO | Get-GPOReport -ReportType Xml

    $AppLockerStatus = $GPOReport.GPO.Computer.ExtensionData.extension.rulecollection.enforcementmode.mode

    #Check to see if AppLocker is enabled
    #'Disabled' means it's in audit mode
    if ($AppLockerStatus -contains 'Disabled' -or $AppLockerStatus -contains 'Enabled') {
        $AppLockerFound = $true

        Write-Log -Path $LogLocation -Level Info -Message 'There is already an AppLocker GPO in place. Script will exit'

        #exit the script if AppLocker is found
        break
    }
}

#Download and apply AppLocker if it was not found
if (!$AppLockerFound) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $AppLockerZipPath = 'C:\Windows\Temp\AppLockerAudit.zip'
    $AppLockerGPOFolder = ($AppLockerZipPath.Split('.'))[0]

    #Download the GPO
    (New-Object System.Net.WebClient).DownloadFile('https://github.com/CompassMSP/PublicScripts/raw/master/ActiveDirectory/GPOBackups/AppLocker%20RDS%20AUDIT%20ONLY.zip', $AppLockerZipPath)

    Expand-ZIP -ZipFile $AppLockerZipPath -OutPath $AppLockerGPOFolder

    #Generate a GPO Report
    $GPOReportPath = Get-ChildItem $AppLockerGPOFolder -Recurse | Where-Object name -EQ gpreport.xml
    [XML]$GPOReportXML = Get-Content -Path $GPOReportPath.FullName

    #Prefix an underscore to make it easier to identify later
    $GPOName = $($GPOReportXML.GPO.Name)
    $GPOPrefixedName = "_$($GPOName)"

    #Import and apply the GPO
    New-GPO -Name $GPOName -ErrorAction SilentlyContinue
    Import-GPO -Path "$AppLockerGPOFolder\$GPOName" -TargetName $GPOPrefixedName -BackupGpoName $GPOName -ErrorAction Stop
    New-GPLink -Name 'Password Protection' -Target $((Get-ADDomain).DistinguishedName) -LinkEnabled Yes -ErrorAction Stop
}