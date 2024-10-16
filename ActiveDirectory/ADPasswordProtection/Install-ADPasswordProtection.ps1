Function Install-ADPasswordProtection {
    <#
    .SYNOPSIS
    This script installs Lithnet AD Password Protection on the DC it is run on.

    .DESCRIPTION
    The goal of this application is to prevent users from setting known compromised passwords (P@ssw0rd) in AD.

    The script will do the following:
        Install the application on the DC
        Create the GPO (if the server is the PDC)
        Update the HIBP DB into the Store location

    .PARAMETER SMTPRelay
    SMTP server that will be used to send notifications if the script runs into any issues.

    .PARAMETER NotificationEmail
    Email address that will receive a notification if the script runs into any issues

    .PARAMETER FromEmail
    "From" email for notifications

    .EXAMPLE
    Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Install-ADPasswordProtection'); Install-ADPasswordProtection -NotificationEmail 'alerts@example.com' -SMTPRelay 'example.mail.protection.outlook.com' -FromEmail 'ADPasswordNotifications@example.com'

    .LINK
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

    $GPOPath = 'C:\Windows\Temp\PasswordProtection.zip'
    $LogDirectory = 'C:\Windows\Temp\PasswordProtection.log'

    $Errors = @()

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
        Write-Log -Level Info -Path $LogDirectory -Message 'Computer is not a DC. Script will exit'
        exit
    }

    #region Check For Existing components
    $GPOExistsWithCorrectSettings = $false

    #Check for GPO
    if (Get-GPO -Name 'Password Protection' -ErrorAction SilentlyContinue) {
        [XML]$GPOReport = Get-GPO -Name 'Password Protection' -ErrorAction SilentlyContinue | Get-GPOReport -ReportType Xml

        $GPOSetting = $GPOReport.GPO.Computer.ExtensionData.Extension.Policy | Where-Object { $_.Name -eq 'Reject passwords found in the compromised password store' }
        if ($GPOSetting.State -eq 'Enabled') {
            $GPOExistsWithCorrectSettings = $true
            Write-Log -Level Info -Path $LogDirectory -Message "The GPO $($GPOReport.GPO.Name) will not be created since it already exists"
        }
    }

    #Check if DC has enough free space
    if ((Get-PSDrive C).free -lt 30GB) {
        Write-Log -Level Warn -Path $LogDirectory -Message 'DC has less than 30 GB free. Script will exit'
        Start-Process $LogDirectory
        exit 
    } else {
        $FreeSpace = 'yes'
    }

    if ($FreeSpace -eq 'yes') {
        
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        if ((Test-Path 'C:\Temp' ) -eq $false) {
            New-Item -Path 'C:\Temp' -ItemType Directory
        }

        #region Download and Install
        $URI = "https://github.com/lithnet/ad-password-protection/releases/latest"

        $latestRelease = Invoke-WebRequest $URI -Headers @{"Accept" = "application/json" } -UseBasicParsing
        $json = $latestRelease.Content | ConvertFrom-Json
        $latestVersion = $json.tag_name

        $BuildExe = $latestVersion.Replace('v', 'LithnetPasswordProtection-') + '.exe'
        $BuildURI = "https://github.com/lithnet/ad-password-protection/releases/download/$latestVersion/" + $BuildExe

        (New-Object System.Net.WebClient).DownloadFile("$BuildURI", "c:\temp\$BuildExe")
 
        Write-Log -Level Info -Path $LogDirectory -Message 'Installing Password Protection'

        Start-Process -FilePath C:\temp\$BuildExe -Wait

        Import-Module LithnetPasswordProtection -Force
        
        Sync-HashesFromHibp

        Write-Log -Level Info -Path $LogDirectory -Message "The Password Protection application has been installed. Restart the computer for the change to take effect."
        
        #endregion DownloadInstallMSI

        #region ImportGPO
        if (-not $GPOExistsWithCorrectSettings) {
            if ((Get-WmiObject Win32_ComputerSystem).domainRole -eq 5) {
                #region Copy ADM files to central store
                if (Test-Path 'C:\Windows\SYSVOL') {
                    $SYSVOLPath = 'C:\Windows\SYSVOL'
                } elseif (Test-Path 'C:\Windows\SYSVOL_DFSR') {
                    $SYSVOLPath = 'C:\Windows\SYSVOL_DFSR'
                }

                #Update PolicyDefinitions if it does not Exists
                if (!(Test-Path "$($SYSVOLPath)\domain\Policies\PolicyDefinitions")) {
                    (New-Object System.Net.WebClient).DownloadFile('https://github.com/CompassMSP/PublicScripts/raw/master/ActiveDirectory/PolicyDefinitions.zip', 'C:\Windows\Temp\PolicyDefinitions.zip')

                    Expand-Archive -LiteralPath 'C:\Windows\Temp\PolicyDefinitions.zip' -DestinationPath "C:\Windows\Temp\ADMX"

                    ROBOCOPY "C:\Windows\Temp\ADMX\PolicyDefinitions" "$($SYSVOLPath)\domain\Policies\PolicyDefinitions" /R:0 /W:0 /E /xo /MT:32 /np
                }


                $FilesToCopy = @(
                    'C:\Windows\PolicyDefinitions\lithnet.admx',
                    'C:\Windows\PolicyDefinitions\lithnet.activeDirectory.passwordFilter.admx',
                    'C:\Windows\PolicyDefinitions\en-US\lithnet.activeDirectory.passwordFilter.adml',
                    'C:\Windows\PolicyDefinitions\en-US\lithnet.adml'
                )

                Foreach ($File in $FilesToCopy) {
                    if (Test-Path -Path $File) {
                        if ($File -like '*.adml') {
                            $Destination = "$($SYSVOLPath)\domain\Policies\PolicyDefinitions\en-US"
                        } elseif ($File -like '*.admx') {
                            $Destination = "$($SYSVOLPath)\domain\Policies\PolicyDefinitions"
                        }

                        Write-Log -Level Info -Path $LogDirectory -Message "Moving $File to $Destination"

                        Move-Item -Path $File -Destination $Destination -Force
                    }
                }
                #endregion Copy ADM files to central store

                Write-Log -Level Info -Path $LogDirectory -Message 'Downloading GPO'

                (New-Object System.Net.WebClient).DownloadFile('https://github.com/CompassMSP/PublicScripts/raw/master/ActiveDirectory/GPOBackups/Password%20Protection.zip', "$GPOPath")

                $GPOFolder = $GPOPath.Replace('.zip', '')

                Expand-Archive -LiteralPath $GPOPath -DestinationPath $GPOFolder

                $GPOBackupFolder = (Get-ChildItem $GPOFolder).FullName

                #Get the Name of the GPO from the content of the XML
                $GPOReportPath = Get-ChildItem $GPOFolder -Recurse | Where-Object name -EQ gpreport.xml
                [XML]$GPOReportXML = Get-Content -Path $GPOReportPath.FullName
                [string]$GPOBackupName = $GPOReportXML.GPO.Name

                try {
                    New-GPO -Name 'Password Protection' -ErrorAction Stop
                    Import-GPO -Path $GPOBackupFolder -TargetName 'Password Protection' -BackupGpoName $GPOBackupName -ErrorAction Stop

                    New-GPLink -Name 'Password Protection' -Target "OU=Domain Controllers,$((Get-ADDomain).DistinguishedName)" -LinkEnabled Yes -ErrorAction Stop
                } catch {
                    Write-Log -Level Error -Path $LogDirectory -Message "Ran into an issue importing the GPO from $GPOBackupFolder"
                    $Errors += "Ran into an issue importing the GPO from $GPOBackupFolder"
                }
            } else {
                Write-Log -Level Info -Path $LogDirectory -Message 'Computer is not the PDC. GPO will not be imported'
            }
        }
        #endregion ImportGPO

        #region RunInvoke-ADPasswordAudit
        $PDC = (Get-ADForest | Select-Object -ExpandProperty RootDomain | Get-ADDomain).PDCEmulator

        $LocalDC = [System.Net.Dns]::GetHostByName($env:computerName).HostName

        if ($PDC -eq $LocalDC) {
            Invoke-Expression (New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Invoke-ADPasswordAudit.ps1'); Invoke-ADPasswordAudit -NotificationEmail $NotificationEmail -SMTPRelay $SMTPRelay -FromEmail $FromEmail

            if ((Test-Path 'C:\Scripts' ) -eq $false) { 
                New-Item -Path 'C:\Scripts' -ItemType Directory 
            } else {
                if ((Test-Path 'C:\Scripts\Invoke-ADPasswordAudit.ps1' ) -eq $true) { 
                    Remove-Item -Path "C:\Scripts\Invoke-ADPasswordAudit.ps1" -Force 
                }
            }
                
            $taskName = “Invoke-ADPasswordAudit”
            $task = Get-ScheduledTask | Where-Object { $_.TaskName -eq $taskName } | Select-Object -First 1
            if ($null -ne $task) { $task | Unregister-ScheduledTask -Confirm:$false }

            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/CompassMSP/PublicScripts/master/ActiveDirectory/ADPasswordProtection/Invoke-ADPasswordAudit.ps1", "C:\Scripts\Invoke-ADPasswordAudit.ps1")
            
            $TaskArgument = "-Command '. C:\Scripts\Invoke-ADPasswordAudit.ps1; Invoke-ADPasswordAudit -NotificationEmail $NotificationEmail -SMTPRelay $SMTPRelay -FromEmail $FromEmail'"
            $taskTrigger = New-ScheduledTaskTrigger -Daily -At '4:00 AM'
            $taskAction = New-ScheduledTaskAction -Execute "PowerShell" -Argument $TaskArgument -WorkingDirectory $ScriptsFolder
            Register-ScheduledTask 'Invoke-ADPasswordAudit' -Action $taskAction -Trigger $taskTrigger -User "System" -RunLevel Highest
        }
        #endregion RunInvoke-ADPasswordAudit
    } else {
        Write-Log -Level Error -Path $LogDirectory -Message 'Errors have accord while processing. Please refer to log output.'
        $Errors += 'Errors have accord while processing. Please refer to log output.'
        Start-Process $LogDirectory
    }

    if ($Errors.count -gt 0) {
        $EmailBody = $Errors | ForEach-Object { [PSCustomObject]@{'Errors' = $_ } } | ConvertTo-Html -Fragment -Property 'Errors' | Out-String

        $SendMailMessageParams = @{
            To         = $NotificationEmail
            From       = $FromEmail
            Subject    = "Ran into error installing AD Password Protection on $ENV:COMPUTERNAME"
            Body       = $EmailBody
            BodyAsHtml = $true
            SmtpServer = $SMTPRelay
        }

        Send-MailMessage @SendMailMessageParams
    }

}