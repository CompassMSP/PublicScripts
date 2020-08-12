Function Install-ADPasswordProtection {
    <#
    .SYNOPSIS
    This script installs Lithnet AD Password Protection on the DC it is run on.

    .DESCRIPTION
    The goal of this application is to prevent users from setting known compromised passwords (P@ssw0rd) in AD.

    The script will do the following:
        Install the application on the DC
        Create the GPO (if the server is the PDC)
        Copy the HIBP DB into the Store location

    .PARAMETER StoreFilesInDBFormatLink
    A URL to the ZIP file where the HIBP DB files will be hosted. The script will download this ZIP and extract it directly to its store.

    This file will need to be manually updated any time there is a HIBP DB update (about once or twice a year)

    .PARAMETER SMTPRelay
    SMTP server that will be used to send notifications if the script runs into any issues.

    .PARAMETER NotificationEmail
    Email address that will recieve a notification if the script runs into any issues

    .PARAMETER FromEmail
    "From" email for notifications

    .EXAMPLE
    Install-ADPasswordProtection -StoreFilesInDBFormatLink 'https://example.com/ADPasswordStore.zip' -NotificationEmail 'alerts@example.com' -SMTPRelay 'example.mail.protection.outlook.com' -FromEmail 'ADPasswordNotifications@example.com'

    .LINK
    https://github.com/lithnet/ad-password-protection
    https://github.com/CompassMSP/PublicScripts/blob/master/ActiveDirectory/Install-ADPasswordProtection.ps1

    Andy Morales
    #>
    #Requires -Version 5 -RunAsAdministrator

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            HelpMessage = 'https://example.com/ADPasswordStore.zip')]
        [string]$StoreFilesInDBFormatLink,

        [Parameter(Mandatory = $true,
            HelpMessage = 'example.mail.protection.outlook.com')]
        [string]$SMTPRelay,

        [Parameter(Mandatory = $true)]
        [string]$NotificationEmail,

        [Parameter(Mandatory = $true)]
        [string]$FromEmail
    )

    $StoreFilesInDBFormatFile = 'C:\Temp\ADPasswordAuditStore.zip'
    $PasswordProtectionMSIFile = 'C:\Windows\Temp\Lithnet.ActiveDirectory.PasswordProtection.msi'
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
            [switch]$DaiyMode
        )

        Begin {
            # Set VerbosePreference to Continue so that verbose messages are displayed.
            $VerbosePreference = 'Continue'
            if ($DaiyMode) {
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

    Function Remove-OldFiles {
        $ItemsToDelete = @(
            $GPOPath,
            $GPOFolder,
            $StoreFilesInDBFormatFile,
            $PasswordProtectionMSIFile
        )

        Write-Log -Level Info -Path $LogDirectory -Message 'Deleting old files if they exist.'

        foreach ($item in $ItemsToDelete) {
            if(Test-Path -Path $item){
                Remove-Item -Path $item -Force -Recurse -ErrorAction SilentlyContinue
            }
        }
    }

    Function Get-InstalledApplications {
        $InstalledApplications = @()
        $UninstallKeys = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
        Foreach ($SubKey in $UninstallKeys) {

            $DisplayName = (Get-ItemProperty -Path "Registry::$($SubKey.Name)" -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
            if ([string]::IsNullOrEmpty($DisplayName)) {
            }
            else {
                $InstalledApplications += [PSCustomObject]@{
                    DisplayName = $DisplayName
                }
            }
        }

        Return $InstalledApplications
    }

    #Check if computer is a DC
    if ((Get-WmiObject Win32_ComputerSystem).domainrole -lt 4) {
        Write-Log -Level Info -Path $LogDirectory -Message 'Computer is not a DC. Script will exit'
        exit
    }

    #region Check For Existing components
    $GPOExistsWithCorrectSettings = $false
    $ADPasswordProtectionAlreadyInstalled = $false
    $HIBPDBFilesExist = $false

    #Check for GPO
    if (Get-GPO -Name 'Password Protection' -ErrorAction SilentlyContinue) {
        [XML]$GPOReport = Get-GPO -Name 'Password Protection' -ErrorAction SilentlyContinue | Get-GPOReport -ReportType Xml

        $GPOSetting = $GPOReport.GPO.Computer.ExtensionData.Extension.Policy | Where-Object { $_.Name -eq 'Reject passwords found in the compromised password store' }
        if ($GPOSetting.State -eq 'Enabled') {
            $GPOExistsWithCorrectSettings = $true
            Write-Log -Level Info -Path $LogDirectory -Message "The GPO $($GPOReport.GPO.Name) will not be created since it already exists"
        }
    }

    #Check for app installed
    if ((Get-InstalledApplications).displayname -contains 'Lithnet Password Protection for Active Directory') {
        $ADPasswordProtectionAlreadyInstalled = $true
        Write-Log -Level Info -Path $LogDirectory -Message "The Password Protection application is already installed on this computer"
    }

    #Check to see if DB files exist
    if (Test-Path -Path 'C:\Program Files\Lithnet\Active Directory Password Protection\Store\v3\p\FFFF.db') {
        $HIBPDBFilesExist = $true
        Write-Log -Level Info -Path $LogDirectory -Message "The HIBP DB files already exist in the default location"
    }
    #endregion Check For Existing components

    #Clean up any old files
    Remove-OldFiles

    if ((Get-PSDrive C).free -gt 20GB) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        #Download HIBP Hashes
        if (-not $HIBPDBFilesExist) {
            Write-Log -Level Info -Path $LogDirectory -Message 'Downloading HIBP hashes'
            New-Item -Path 'C:\Temp' -ItemType Directory -Force

            (New-Object System.Net.WebClient).DownloadFile("$StoreFilesInDBFormatLink", "$StoreFilesInDBFormatFile")

            #Extract HIBP Hashes
            Write-Log -Level Info -Path $LogDirectory -Message 'Extracting HIBP hashes'

            try {
                Expand-Archive -Path $StoreFilesInDBFormatFile -DestinationPath 'C:\Program Files\Lithnet\Active Directory Password Protection' -Force
            }
            catch {
                Write-Log -Level Error -Path $LogDirectory -Message "Ran into an issue extracting the file $StoreFilesInDBFormatFile"
                $Errors += "Ran into an issue extracting the file $StoreFilesInDBFormatFile"
            }

            Remove-Item $StoreFilesInDBFormatFile -Force -ErrorAction SilentlyContinue
        }


        #Download and install MSI
        if (-not $ADPasswordProtectionAlreadyInstalled) {
            (New-Object System.Net.WebClient).DownloadFile('https://github.com/lithnet/ad-password-protection/releases/latest/download/Lithnet.ActiveDirectory.PasswordProtection.msi', "$PasswordProtectionMSIFile")

            Write-Log -Level Info -Path $LogDirectory -Message 'Installing Password Protection MSI'
            Start-Process msiexec.exe -Wait -ArgumentList "/i $($PasswordProtectionMSIFile) /qn" -PassThru

            Write-Log -Level Info -Path $LogDirectory -Message "The Password Protection application has been installed. Restart the computer for the change to take effect."
        }

        #region Import GPO
        if (-not $GPOExistsWithCorrectSettings) {
            if ((Get-WmiObject Win32_ComputerSystem).domainrole -eq 5) {
                #region Copy ADM files to central store
                if (Test-Path 'C:\Windows\SYSVOL') {
                    $SYSVOLPath = 'C:\Windows\SYSVOL'
                }
                elseif (Test-Path 'C:\Windows\SYSVOL_DFSR') {
                    $SYSVOLPath = 'C:\Windows\SYSVOL_DFSR'
                }

                $FilesToCopy = @(
                    'C:\Windows\PolicyDefinitions\lithnet.admx',
                    'C:\Windows\PolicyDefinitions\lithnet.activedirectory.passwordfilter.admx',
                    'C:\Windows\PolicyDefinitions\en-US\lithnet.activedirectory.passwordfilter.adml',
                    'C:\Windows\PolicyDefinitions\en-US\lithnet.adml'
                )

                Foreach ($File in $FilesToCopy) {
                    if (Test-Path -Path $File){
                        if ($File -like '*.adml') {
                            $Destination = "$($SYSVOLPath)\domain\Policies\PolicyDefinitions\en-US"
                        }
                        elseif ($File -like '*.admx') {
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

                Expand-Archive -LiteralPath  $GPOPath -DestinationPath $GPOFolder -Force

                $GPOBackupFolder = (Get-ChildItem $GPOFolder).FullName

                #Get the Name of the GPO from the content of the XML
                $GPOReportPath = Get-ChildItem $GPOFolder -Recurse | Where-Object name -EQ gpreport.xml
                [XML]$GPOReportXML = Get-Content -Path $GPOReportPath.FullName
                [string]$GPOBackupName = $GPOReportXML.GPO.Name

                try {
                    New-GPO -Name 'Password Protection' -ErrorAction Stop
                    Import-GPO -Path $GPOBackupFolder -TargetName 'Password Protection' -BackupGpoName $GPOBackupName -ErrorAction Stop

                    New-GPLink -Name 'Password Protection' -Target (Get-ADDomain).DistinguishedName -LinkEnabled Yes -ErrorAction Stop
                }
                catch {
                    Write-Log -Level Error -Path $LogDirectory -Message "Ran into an issue importing the GPO from $GPOBackupFolder"
                    $Errors += "Ran into an issue importing the GPO from $GPOBackupFolder"
                }
            }
            else {
                Write-Log -Level Info -Path $LogDirectory -Message 'Computer is not the PDC. GPO will not be imported'
            }
        }
        #endregion Import GPO
    }
    else {
        Write-Log -Level Error -Path $LogDirectory -Message 'Not enough free space on the C drive. At least 20GB required.'
        $Errors += 'Not enough free space on the C drive. At least 20GB required.'
    }

    if ($Errors.count -gt 0) {
        $EmailBody = $Errors | ForEach-Object { [PSCustomObject]@{'Errors' = $_ } } | ConvertTo-Html -Fragment -Property 'Errors' | Out-String



        Send-MailMessage -To $NotificationEmail -From 'BUIWUpdates@compassmsp.com' -Subject "Ran into error installing AD Password Protection on $ENV:COMPUTERNAME" -BodyAsHtml $EmailBody -SmtpServer $SMTPRelay

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

    #Cleanup
    Remove-OldFiles
}